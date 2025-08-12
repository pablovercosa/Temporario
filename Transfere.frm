VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTransfere 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Transferência entre Filiais"
   ClientHeight    =   8445
   ClientLeft      =   195
   ClientTop       =   375
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Transfere.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   15270
   Begin VB.CommandButton cmd_verEstoqueAtual 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Estoque dos produtos da Transferência"
      Height          =   435
      Left            =   12930
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CheckBox O_Estoque 
      Appearance      =   0  'Flat
      Caption         =   "Permitir fazer transferência mesmo sem estoque suficiente"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   10440
      TabIndex        =   43
      Top             =   3600
      Width           =   4575
   End
   Begin VB.CommandButton cmd_gravar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gravar"
      Height          =   435
      Left            =   10425
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Efetuar Transferência"
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Height          =   1620
      Left            =   2100
      TabIndex        =   20
      Top             =   1920
      Width           =   13125
      Begin VB.TextBox Filial_Destino 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2235
         TabIndex        =   25
         Top             =   630
         Visible         =   0   'False
         Width           =   5490
      End
      Begin VB.TextBox txtFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   9720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   150
         Width           =   3315
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   9720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   570
         Width           =   3315
      End
      Begin VB.TextBox txtOperEntrada 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3315
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1230
         Width           =   4410
      End
      Begin VB.TextBox txtOperSaida 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3315
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   870
         Width           =   4410
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
         Bindings        =   "Transfere.frx":4E95A
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2235
         TabIndex        =   26
         Top             =   510
         Width           =   1065
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
         BevelColorFrame =   -2147483633
         BevelColorHighlight=   -2147483633
         BackColorOdd    =   8454143
         Columns(0).Width=   3200
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBCombo cboTabela 
         Bindings        =   "Transfere.frx":4E96E
         Height          =   315
         Left            =   9090
         TabIndex        =   27
         Top             =   1185
         Width           =   3945
         DataFieldList   =   "Tabela"
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
         BevelColorFrame =   -2147483633
         BevelColorHighlight=   -2147483633
         BackColorOdd    =   8454143
         Columns(0).Width=   3200
         _ExtentX        =   6959
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Tabela"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCliente 
         Bindings        =   "Transfere.frx":4E986
         Height          =   315
         Left            =   8610
         TabIndex        =   28
         ToolTipText     =   "É necessário que a Filial de Destino esteja cadastrada como Cliente"
         Top             =   570
         Width           =   1065
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
         BevelColorFrame =   -2147483633
         BevelColorHighlight=   -2147483633
         BackColorOdd    =   8454143
         Columns(0).Width=   3200
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "Transfere.frx":4E99F
         Height          =   315
         Left            =   8610
         TabIndex        =   29
         ToolTipText     =   "É necessário que a Filial de Saída esteja cadastrada como Fornecedor"
         Top             =   150
         Width           =   1065
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
         BevelColorFrame =   -2147483633
         BevelColorHighlight=   -2147483633
         BackColorOdd    =   8454143
         Columns(0).Width=   3200
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperSaida 
         Bindings        =   "Transfere.frx":4E9BB
         Height          =   315
         Left            =   2235
         TabIndex        =   30
         Top             =   855
         Width           =   1065
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
         BevelColorFrame =   -2147483633
         BevelColorHighlight=   -2147483633
         BackColorOdd    =   8454143
         Columns(0).Width=   3200
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperEntrada 
         Bindings        =   "Transfere.frx":4E9D6
         Height          =   315
         Left            =   2235
         TabIndex        =   31
         Top             =   1200
         Width           =   1065
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
         BevelColorFrame =   -2147483633
         BackColorOdd    =   8454143
         Columns(0).Width=   3200
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Código"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial de saída de produtos"
         Height          =   195
         Left            =   90
         TabIndex        =   41
         Top             =   210
         Width           =   1860
      End
      Begin VB.Label Filial_Saída 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2235
         TabIndex        =   40
         Top             =   165
         Width           =   5490
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filial de entrada de produtos"
         Height          =   195
         Left            =   90
         TabIndex        =   39
         Top             =   570
         Width           =   2055
      End
      Begin VB.Label Filial_Entrada 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3315
         TabIndex        =   38
         Top             =   525
         Width           =   4410
      End
      Begin VB.Label Nome_Filial 
         AutoSize        =   -1  'True
         Caption         =   "Filial de Destino"
         Height          =   195
         Left            =   90
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tabela de Preços"
         Height          =   195
         Left            =   7800
         TabIndex        =   36
         Top             =   1245
         Width           =   1230
      End
      Begin VB.Label lblFornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   7770
         TabIndex        =   35
         Top             =   210
         Width           =   825
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   7800
         TabIndex        =   34
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblOpEntrada 
         AutoSize        =   -1  'True
         Caption         =   "Operação de Entrada"
         Height          =   195
         Left            =   90
         TabIndex        =   33
         Top             =   1260
         Width           =   1545
      End
      Begin VB.Label lblOpSaida 
         AutoSize        =   -1  'True
         Caption         =   "Operação de Saída"
         Height          =   195
         Left            =   90
         TabIndex        =   32
         Top             =   915
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Realizadas  ou  Em andamento"
      Height          =   1905
      Left            =   0
      TabIndex        =   19
      Top             =   30
      Width           =   15195
      Begin VB.ComboBox cmb_tipo 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "Transfere.frx":4E9F3
         Left            =   9180
         List            =   "Transfere.frx":4E9FD
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   240
         Width           =   2925
      End
      Begin VB.ComboBox cmb_status 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "Transfere.frx":4EA2F
         Left            =   5010
         List            =   "Transfere.frx":4EA3C
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   240
         Width           =   2865
      End
      Begin VB.CommandButton cmd_calendarioDtIni 
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         Picture         =   "Transfere.frx":4EA79
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   210
         Width           =   465
      End
      Begin VB.CommandButton cmd_calendarioDtFim 
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3750
         Picture         =   "Transfere.frx":4F35B
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   210
         Width           =   465
      End
      Begin VB.CommandButton cmdPesquisar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   180
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid gridTransf 
         Height          =   1245
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   2196
         _Version        =   393216
         Rows            =   1
         Cols            =   16
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   8454143
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16250871
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   285
         Left            =   375
         TabIndex        =   48
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   285
         Left            =   2565
         TabIndex        =   49
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Transferência"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8130
         TabIndex        =   55
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4500
         TabIndex        =   53
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "até"
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2175
         TabIndex        =   51
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "De"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   50
         Top             =   270
         Width           =   195
      End
   End
   Begin VB.CommandButton cmd_consultarProduto 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar produto"
      Height          =   435
      Left            =   10425
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Faz a Leitura de Arquivo de Transferência"
      Top             =   4650
      Width           =   2295
   End
   Begin VB.Data datOperSaida 
      Caption         =   "datOperSaida"
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
      Left            =   3150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM [Operações Saída] ORDER BY Código"
      Top             =   8310
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Data datOperEntrada 
      Caption         =   "datOperEntrada"
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
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM [Operações Entrada] ORDER BY Código"
      Top             =   8310
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Data datCliente 
      Caption         =   "datCliente"
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
      Left            =   1620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome, Tipo FROM Cli_For ORDER BY Código"
      Top             =   8310
      Visible         =   0   'False
      Width           =   1500
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
      Left            =   90
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome, Tipo FROM Cli_For ORDER BY Código"
      Top             =   8310
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Data datTabela 
      Caption         =   "datTabela"
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
      Left            =   4650
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Tabela FROM [Tabela de Preços] ORDER BY Tabela"
      Top             =   8310
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdGrade 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Digitar Grade"
      Height          =   435
      Left            =   12930
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Faz a Leitura de Arquivo de Transferência"
      Top             =   5670
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdReadTransf 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Ler Arquivo"
      Height          =   435
      Left            =   12930
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Faz a Leitura de Arquivo de Transferência"
      Top             =   6180
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   4200
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir Transferência"
      Height          =   435
      Left            =   10425
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Transferência"
      Top             =   6180
      Width           =   2295
   End
   Begin VB.CommandButton B_Nova 
      BackColor       =   &H00F7F7F7&
      Caption         =   "Limpar Tela"
      Height          =   435
      Left            =   12930
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Nova Transferência"
      Top             =   4650
      Width           =   2295
   End
   Begin VB.CommandButton B_Transfere 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gravar e Transferir"
      Height          =   435
      Left            =   10425
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Efetuar Transferência"
      Top             =   5670
      Width           =   2295
   End
   Begin Threed.SSPanel L_Estoque 
      Height          =   315
      Left            =   12930
      TabIndex        =   11
      Top             =   7170
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   556
      _StockProps     =   15
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin Threed.SSPanel Mensagem 
      Height          =   765
      Left            =   10425
      TabIndex        =   10
      Top             =   7620
      Width           =   4800
      _Version        =   65536
      _ExtentX        =   8467
      _ExtentY        =   1349
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1800
   End
   Begin SSDataWidgets_B.SSDBDropDown DropDown1 
      Bindings        =   "Transfere.frx":4FC3D
      Height          =   750
      Left            =   7650
      TabIndex        =   9
      Top             =   5040
      Width           =   2430
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8467
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3281
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   4286
      _ExtentY        =   1323
      _StockProps     =   77
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
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1680
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Height          =   4830
      Left            =   30
      TabIndex        =   3
      Top             =   3600
      Width           =   10335
      _Version        =   196617
      DataMode        =   1
      Rows            =   255
      AllowAddNew     =   -1  'True
      AllowColumnSizing=   0   'False
      AllowColumnShrinking=   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   12648447
      RowHeight       =   503
      ExtraHeight     =   238
      Columns.Count   =   4
      Columns(0).Width=   3360
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   10583
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1588
      Columns(2).Caption=   "Qdade."
      Columns(2).Name =   "Quantidade"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3545
      Columns(3).Caption=   "Valor"
      Columns(3).Name =   "Valor"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).NumberFormat=   "$#,###,###,##0.00"
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      _ExtentX        =   18230
      _ExtentY        =   8520
      _StockProps     =   79
      Caption         =   "Produtos a Transferir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transferência entre filiais"
      Height          =   1620
      Left            =   30
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
      Begin VB.OptionButton O_Distante 
         Appearance      =   0  'Flat
         Caption         =   "Bases Separadas - gerar arquivo"
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   120
         TabIndex        =   2
         Top             =   930
         Width           =   1620
      End
      Begin VB.OptionButton O_Local 
         Appearance      =   0  'Flat
         Caption         =   "Cadastradas no mesmo QuickStore"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.Label VTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   12930
      TabIndex        =   58
      Top             =   6810
      Width           =   2295
   End
   Begin VB.Label lblVTotal 
      Caption         =   "$ Total"
      Height          =   195
      Left            =   12390
      TabIndex        =   57
      Top             =   6870
      Width           =   495
   End
   Begin VB.Label lbl_status 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10965
      TabIndex        =   18
      Top             =   4020
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Estoque"
      Height          =   225
      Left            =   12300
      TabIndex        =   16
      Top             =   7230
      Width           =   600
   End
   Begin VB.Label Total_Qtde 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11325
      TabIndex        =   15
      Top             =   6825
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Quantidade"
      Height          =   225
      Left            =   10425
      TabIndex        =   14
      Top             =   6870
      Width           =   855
   End
   Begin VB.Label Total_Itens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11325
      TabIndex        =   13
      Top             =   7185
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Itens"
      Height          =   195
      Left            =   10425
      TabIndex        =   12
      Top             =   7245
      Width           =   420
   End
End
Attribute VB_Name = "frmTransfere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private varInstancia_lCodigoTransf As Long

Dim rsParametros As Recordset
Dim rsOp_Saída As Recordset
Dim rsOp_Entrada As Recordset
Dim rsProdutos As Recordset
Dim rsGrade As Recordset
Dim rsEstoque As Recordset
Dim rsEstoque_Final As Recordset
Dim rsResumo As Recordset
Dim rsCores As Recordset
Dim rsTamanhos As Recordset

Dim p_linhas As Integer  'guarda a quantidades de linhas parametrizadas  13/02/2023 - Pablo

Private Type Tabela
   Código As String
   Nome As String
   Qtde As Single
   Valor As Currency
End Type
Dim Tabe(255) As Tabela

'30/03/2004 - Daniel
'Var de controle para a Sub CriarRegistros
'Case: Casagrande
Dim m_intContador As Integer
'Atualizadores da Ultima Movimentação de Parâmetros
Dim m_nSequencia            As Long
Dim m_nSequenciaStoreSaidas As Long
'05/07/2004 - Daniel
'Incluído Tratamento caso seja transferência envolvendo produtos com grade
Dim m_strTamanho  As String
Dim m_strCor      As String
Dim m_blnComGrade As Boolean

Dim dPrecoFinalSaidas As Double

Sub Ajusta_Entrada()
  Dim i As Integer
  Dim J As Integer
  Dim Criar_Registro As Integer
  Dim Estoque_Final As Single
  Dim Produto As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Aux_Str As String
  Dim Tipo As Integer
  Dim Erro As Integer


  For i = 0 To p_linhas
    If Tabe(i).Código <> "" And Tabe(i).Qtde <> 0 Then
    
      Produto = ""
      Tamanho = 0
      Cor = 0
      Edição = 0
      Call Acha_Produto(Tabe(i).Código, Produto, Tamanho, Cor, Edição, Tipo, Erro)
      If Erro <> 0 Then
        DisplayMsg "Produto " + str(Tabe(i).Código) + " não encontrado. Transferência interrompida."
        Exit Sub
      End If
      Produto = UCase(Produto)
      rsProdutos.Seek "=", Produto
     Rem Ajusta estoque de ENTRADA
     Rem Encontra a posição do estoque
     Criar_Registro = False
     Estoque_Final = 0
     rsEstoque.Index = "Produto"
     rsEstoque.Seek "=", Val(Combo_Filial.Text), Data_Atual, rsProdutos("Código"), Tamanho, Cor, Edição

     If Not rsEstoque.NoMatch Then Estoque_Final = rsEstoque("Estoque Final")
     If rsEstoque.NoMatch Then
       rsEstoque.Index = "Data"
       rsEstoque.Seek "<", Val(Combo_Filial.Text), Produto, Tamanho, Cor, Edição, Data_Atual
       If rsEstoque.NoMatch Then Criar_Registro = True
       If Not rsEstoque.NoMatch Then
          If rsEstoque("Filial") = Val(Combo_Filial.Text) And rsEstoque("Produto") = Produto And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edição") = Edição Then
             Criar_Registro = True
             Estoque_Final = rsEstoque("Estoque Final")
           Else
             Criar_Registro = True
             Estoque_Final = 0
           End If
       End If

       If Criar_Registro = True Then
         rsEstoque.AddNew
          rsEstoque("Filial") = Val(Combo_Filial.Text)
          rsEstoque("Data") = Data_Atual
          rsEstoque("Produto") = Produto
          rsEstoque("Tamanho") = Tamanho
          rsEstoque("Cor") = Cor
          rsEstoque("Edição") = Edição
          rsEstoque("Classe") = rsProdutos("Classe")
          rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
          rsEstoque("Estoque Anterior") = Estoque_Final
         rsEstoque.Update
       End If

       rsEstoque.Index = "Produto"
       rsEstoque.Seek "=", Val(Combo_Filial.Text), Data_Atual, Produto, Tamanho, Cor, Edição
      End If

      'Neste ponto esta com o registro de estoque
      'no buffer, agora soma com os valores da movimentação
      rsEstoque.Edit
         rsEstoque("Transf Entra") = rsEstoque("Transf Entra") + Tabe(i).Qtde
         
         Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
         Estoque_Final = Estoque_Final - rsEstoque("Transf Saída") + rsEstoque("Transf Entra")
         Estoque_Final = Estoque_Final - rsEstoque("Ajuste Saída") + rsEstoque("Ajuste Entra")
         Estoque_Final = Estoque_Final - rsEstoque("Grátis Saída") + rsEstoque("Grátis Entra")
         Estoque_Final = Estoque_Final - rsEstoque("Empre Saída") + rsEstoque("Empre Entra")
         Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolução")

         If rsProdutos("Estoque") = False Then
           Estoque_Final = 0
         End If

         rsEstoque("Estoque Final") = Estoque_Final
      rsEstoque.Update
     
      'Arruma Estoque Final
      Grava_Estoque_Final Val(Combo_Filial.Text), Produto, Tamanho, Cor, Edição, Estoque_Final, CDate(Data_Atual)
     
       
    End If
  Next i

End Sub

Sub Ajusta_Saída()
  Dim i              As Integer
  Dim J              As Integer
  Dim Criar_Registro As Integer
  Dim Estoque_Final  As Single
  Dim Produto        As String
  Dim Tamanho        As Integer
  Dim Cor            As Integer
  Dim Edição         As Long
  Dim Aux_Str        As String
  Dim Tipo           As Integer
  Dim Erro           As Integer
   
  Call StatusMsg("")
  
  rsProdutos.Index = "Código"
  For i = 0 To p_linhas
    If Tabe(i).Código <> "" And Tabe(i).Qtde <> 0 Then
       
      Produto = ""
      Tamanho = 0
      Cor = 0
      Edição = 0
      Call Acha_Produto(Tabe(i).Código, Produto, Tamanho, Cor, Edição, Tipo, Erro)
      If Erro <> 0 Then
        DisplayMsg "Produto " + str(Tabe(i).Código) + " não encontrado. Transferência interrompida."
        Exit Sub
      End If
      Produto = UCase(Produto)
      rsProdutos.Seek "=", Produto
             
       
      Call StatusMsg("Atualizando estoque de " + rsProdutos("Nome"))
      
      'Ajusta estoque de SAÍDA
      'Encontra a posição do estoque
      Criar_Registro = False
      Estoque_Final = 0
      rsEstoque.Index = "Produto"
      rsEstoque.Seek "=", gnCodFilial, Data_Atual, rsProdutos("Código"), Tamanho, Cor, Edição

      If Not rsEstoque.NoMatch Then Estoque_Final = rsEstoque("Estoque Final")
      If rsEstoque.NoMatch Then
       rsEstoque.Index = "Data"
       rsEstoque.Seek "<", gnCodFilial, Produto, Tamanho, Cor, Edição, Data_Atual
       If rsEstoque.NoMatch Then Criar_Registro = True
       If Not rsEstoque.NoMatch Then
          If rsEstoque("Filial") = gnCodFilial And rsEstoque("Produto") = Produto And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edição") = Edição Then
             Criar_Registro = True
             Estoque_Final = rsEstoque("Estoque Final")
          Else
             Criar_Registro = True
             Estoque_Final = 0
          End If
       End If

       rsEstoque.AddNew
        rsEstoque("Filial") = gnCodFilial
        rsEstoque("Data") = Data_Atual
        rsEstoque("Produto") = Produto
        rsEstoque("Tamanho") = Tamanho
        rsEstoque("Cor") = Cor
        rsEstoque("Edição") = Edição
        rsEstoque("Classe") = rsProdutos("Classe")
        rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
        rsEstoque("Estoque Anterior") = Estoque_Final
       rsEstoque.Update

       rsEstoque.Index = "Produto"
       rsEstoque.Seek "=", gnCodFilial, Data_Atual, Produto, Tamanho, Cor, Edição
      End If

      'neste ponto esta com o registro de estoque
      'no buffer, agora soma com os valores da movimentação
      rsEstoque.Edit
         rsEstoque("Transf Saída") = rsEstoque("Transf Saída") + Tabe(i).Qtde
         
         Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
         Estoque_Final = Estoque_Final - rsEstoque("Transf Saída") + rsEstoque("Transf Entra")
         Estoque_Final = Estoque_Final - rsEstoque("Ajuste Saída") + rsEstoque("Ajuste Entra")
         Estoque_Final = Estoque_Final - rsEstoque("Grátis Saída") + rsEstoque("Grátis Entra")
         Estoque_Final = Estoque_Final - rsEstoque("Empre Saída") + rsEstoque("Empre Entra")
         '27/10/2004 - Daniel & Maikel
         '             Descomentada a soma da coluna Devolução para resolver o problema de estoque
         '             ficando igual ao efetiva_saída
         Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolução")

         If rsProdutos("Estoque") = False Then
           Estoque_Final = 0
         End If

         rsEstoque("Estoque Final") = Estoque_Final
      rsEstoque.Update
     
      'Arruma Estoque Final
      Grava_Estoque_Final gnCodFilial, Produto, Tamanho, Cor, Edição, Estoque_Final, CDate(Data_Atual)
      
      '30/04/2004 - Daniel
      'Chamada da Private Sub CriarRegistros
      '05/07/2004 - Daniel
      'Incluído Tratamento caso seja transferência envolvendo produtos com grade
      Dim rstProdutos As Recordset
      Dim strQuery    As String
      Dim st          As String
      
      strQuery = "SELECT Código, Tipo, [Situação Tributária] "
      strQuery = strQuery & " FROM Produtos "
      strQuery = strQuery & " WHERE Código = '" & Produto & "'"
      
      Set rstProdutos = db.OpenRecordset(strQuery, dbOpenDynaset)
      
      With rstProdutos
        If Not (.BOF And .EOF) Then
          .MoveFirst
          
          st = .Fields("Situação Tributária").Value
          st = IIf(IsNull(st), "0", st)
          
          If .Fields("Tipo").Value = "G" Then 'Possui Grade
            m_blnComGrade = True
          
            m_strTamanho = Right(("000" & Tamanho), 3)
            m_strCor = Right(("000" & Cor), 3)
            
            Call CriarRegistros(Produto, Tabe(i).Qtde, st)
          Else
            m_blnComGrade = False
            
            Call CriarRegistros(Produto, Tabe(i).Qtde, st)
          End If
        End If
        .Close
      End With
      
      Set rstProdutos = Nothing
      '--------------------------------------------------------------------------
      
    End If
   Next i

   '30/03/2004 - Daniel
   'Zerar os contadores usados na Private Sub CriarRegistros
   m_intContador = 0
   m_nSequencia = 0
   m_nSequenciaStoreSaidas = 0
  '------------------------------------------------

End Sub

Sub Recalcula()
  Dim i As Integer
  
  Dim Itens As Integer
  Dim Qtde As Single
  Dim Valor As Currency
  
  Itens = 0
  Qtde = 0
  Valor = 0
  
  For i = 0 To p_linhas
   If Tabe(i).Código <> "" Then
     Itens = Itens + 1
     Qtde = Qtde + Tabe(i).Qtde
     Valor = Valor + Tabe(i).Valor
   End If
  Next i
  
  Total_Itens.Caption = Itens
  Total_Qtde.Caption = Qtde
  VTotal.Caption = Format(Valor, "#,###,###,##0.00")
  
End Sub

Private Sub B_Imprime_Click()
  Dim nX As Long
  Dim nMax As Long
  
  Call Grade1_LostFocus
  
  nMax = -1
  For nX = 0 To p_linhas
    If Tabe(nX).Código <> "" Then
      nMax = nX
    End If
  Next nX
  
  If nMax <> -1 Then
    With Grade1
      .Rows = nMax + 1
      .PrintData ssPrintAllRows, True, True
      .Rows = p_linhas
      .SetFocus
      .MoveLast
      .MoveFirst
    End With
  End If
  
End Sub

Private Sub B_Nova_Click()

  Call StatusMsg("")
  
  Erase Tabe
  Grade1.MoveLast
  Grade1.MoveFirst
  
  lbl_status.Caption = ""
  Combo_Filial.Text = ""
  Filial_Entrada.Caption = ""
  Total_Itens.Caption = ""
  Total_Qtde.Caption = ""
  Filial_Destino.Text = ""
  
  '29/03/2004 - Daniel
  cboOperSaida.Text = ""
  txtOperSaida.Text = ""
  cboOperEntrada.Text = ""
  txtOperEntrada.Text = ""
  cboCliente.Text = ""
  txtCliente.Text = ""
'''  cboFornecedor.Text = ""
'''  txtFornecedor.Text = ""
  cboTabela.Text = ""
  VTotal.Caption = ""
  '-------------------------
  
  B_Transfere.Enabled = True
  'Grade1.SetFocus
  'SendKeys "{TAB}"
  
  Grade1.AllowUpdate = True
  Combo_Filial.Enabled = True
  cboOperEntrada.Enabled = True
  cboOperSaida.Enabled = True
  cboTabela.Enabled = True
  
End Sub

Private Sub B_Transfere_Click()
  Dim J As Integer
  Dim i As Integer
  Dim Aux_Str As String
  Dim Aux_Str2 As String
  Dim Nome_Arq As String
  Dim Resp As Integer
 
 
  If lbl_status.Caption = "Concluída" Then
      MsgBox "Esta Transferência entre filiais já foi concluída."
      Exit Sub
  End If
  
  If Len(cboTabela.Text) <= 0 Then
    MsgBox "Selecione uma Tabela de Preços.", vbInformation, "Atenção"
    cboTabela.SetFocus
    Exit Sub
  End If
  
  If Len(Combo_Filial.Text) <= 0 Then
    MsgBox "Selecione a Filial para transferência", vbInformation, "Atenção"
    Combo_Filial.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(cboOperSaida.Text)) <= 0 Then
    MsgBox "Selecione uma Operação de transferência de saída.", vbInformation, "Atenção"
    cboOperSaida.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(cboOperEntrada.Text)) <= 0 Then
    MsgBox "Selecione uma Operação de transferência de entrada.", vbInformation, "Atenção"
    cboOperEntrada.SetFocus
    Exit Sub
  End If

  Call StatusMsg("")

  If O_Distante.Value = True Then
    Aux_Str = gsReportPath & "TRAN"
    Aux_Str2 = Format(Date, "dd/mm/yy")
    Aux_Str = Aux_Str + Left(Aux_Str2, 2)
    Aux_Str = Aux_Str + Mid(Aux_Str2, 4, 2)
      
    Aux_Str = Aux_Str + ".TSC"
    
    Dialog1.FileName = Aux_Str
    
On Error GoTo Erro_Gravar
      With Dialog1
        .CancelError = True
        .DialogTitle = "Salvar arquivo de transferência como"
        .DefaultExt = "TSC"
        .InitDir = gsDefaultPath
        .Filter = "Arquivo de Transferência | *.TSC"
        .Flags = cdlOFNPathMustExist & cdlOFNHideReadOnly
        .ShowSave
      End With
'      Dialog1.Action = 2
    On Error GoTo 0
    
    Nome_Arq = Dialog1.FileName
    If Dir(Nome_Arq) <> "" Then
      Resp = MsgBox("Já existe este arquivo, deseja sobrescrever ? ", vbQuestion + vbOKCancel, "Atenção")
      If Resp = vbCancel Then
         DisplayMsg "Transferência cancelada."
         Exit Sub
      End If
    End If
   
  End If
       
  If O_Distante.Value = False Then
    If Filial_Entrada.Caption = "" Then
      DisplayMsg "Filial de entrada não digitada. Verifique."
      Exit Sub
    End If
  
    If Val(Combo_Filial.Text) = gnCodFilial Then
      DisplayMsg "Filial de entrada e saída devem ser diferentes."
      Exit Sub
    End If
  End If
  
  
  J = 0
  For i = 0 To p_linhas
   If Tabe(i).Código <> "" Then J = J + 1
  Next i
  
  If J = 0 Then
    DisplayMsg "Não existe nenhum produto a transferir."
    Exit Sub
  End If
       
  '29/03/2004 - Daniel
  'Validação dos campos (Operações, Cliente/Fornecedor e Tabela)
  'verificação se algum deles está vazio
  If Len(txtOperSaida.Text) <= 0 Then
    MsgBox "O campo Operação de Saída está inválido.", vbExclamation, "Quick Store"
    cboOperSaida.SetFocus
    Exit Sub
  End If
  
  If Len(txtCliente.Text) <= 0 Then
    MsgBox "O campo Cliente está inválido.", vbExclamation, "Quick Store"
    cboCliente.SetFocus
    Exit Sub
  End If
  
    If O_Local.Value Then
      If Len(txtFornecedor.Text) <= 0 Then
        MsgBox "O campo Fornecedor está inválido.", vbExclamation, "Quick Store"
        cboFornecedor.SetFocus
        Exit Sub
      End If
      
      If Len(txtOperEntrada.Text) <= 0 Then
        MsgBox "O campo Operação de Entrada está inválido.", vbExclamation, "Quick Store"
        cboOperEntrada.SetFocus
        Exit Sub
      End If
    Else
        If Len(Filial_Destino.Text) <= 0 Then
          MsgBox "O campo Filial de Destino está inválido.", vbExclamation, "Quick Store"
          Filial_Destino.SetFocus
          Exit Sub
        End If
    End If
  
  If Len(cboTabela.Text) <= 0 Then
    MsgBox "O campo Tabela de Preços está inválido.", vbExclamation, "Quick Store"
    cboTabela.SetFocus
    Exit Sub
  End If
  '---------------------------------------------------------------------------------

       
   B_Transfere.Enabled = False
       
   'Gera arquivo texto
   
   If O_Distante.Value = True Then
     On Error GoTo Erro_Gravar
      Open Nome_Arq For Output As #1
     On Error GoTo 0
     Aux_Str = "***TRANSFEREQUICK***"
     Print #1, Aux_Str

    
    '25/10/2005 - mpdea
    'Incluído a tabela de preços no arquivo de transferência
    Print #1, "TABELA:" & Trim(cboTabela.Text)


     For i = 0 To p_linhas
       If Tabe(i).Código <> "" And Tabe(i).Qtde <> 0 Then
         Aux_Str = Tabe(i).Código + "#" + str(Tabe(i).Qtde)
         Print #1, Aux_Str
       End If
     Next i
     
     Print #1, "***FIMTRANSFERE"
     
     Close #1
   End If
       
  Call Ajusta_Saída
   
  If O_Distante.Value = False Then
    Ajusta_Entrada
  End If
  
  ' =====================================================================================
  ' Gravar na tabelas de:
  '                       TransferenciaEntreFiliais   - PAI
  '                       TransferenciaProdutos       - FILHA
  If O_Local.Value = True Then
  
      cmd_gravar_Click
      db.Execute "Update TransferenciaEntreFiliais set Status = 2 Where CodigoTransf = " & varInstancia_lCodigoTransf
  End If
  ' =====================================================================================

  Call StatusMsg("")
  lbl_status.Visible = True
  lbl_status.Caption = "Concluída"
  DisplayMsg "Transferência efetuada com Sucesso."
  cmdPesquisar_Click
  Exit Sub

Erro_Gravar:
  DisplayMsg "Impossível gravar arquivo. Transferência não executada."
  Exit Sub
  
End Sub

Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(0).Text
  cboCliente_LostFocus
End Sub

'05/12/2005 - mpdea
'Incluído tratamento de erro
Private Sub cboCliente_LostFocus()
  Dim rstCliente As Recordset
  Dim lngRet As Long

  On Error GoTo ErrHandler
  
  txtCliente.Text = ""
  
  If Not IsNumeric(cboCliente.Text) Then Exit Sub
  
  '05/12/2005 - mpdea
  'Tratamento de overflow
  Call IsDataType(dtLong, cboCliente.Text, lngRet)
  
  '''Set rstCliente = db.OpenRecordset("SELECT Código, Nome, Tipo FROM Cli_For WHERE Código = " & lngRet & " AND Tipo ='C' ORDER BY Código", dbOpenDynaset, dbReadOnly)
  Set rstCliente = db.OpenRecordset("SELECT Código, Nome, Tipo FROM Cli_For WHERE Código = " & lngRet & " ORDER BY Código", dbOpenDynaset, dbReadOnly)

  With rstCliente
    If Not (.BOF And .EOF) Then
      txtCliente.Text = .Fields("Nome") & ""
    End If
  End With

  rstCliente.Close
  Set rstCliente = Nothing

  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(0).Text
  cboFornecedor_LostFocus
End Sub

'05/12/2005 - mpdea
'Incluído tratamento de erro
Private Sub cboFornecedor_LostFocus()
  Dim rstFornecedor As Recordset
  Dim lngRet As Long

  On Error GoTo ErrHandler
  
  txtFornecedor.Text = ""
  
  If Not IsNumeric(cboFornecedor.Text) Then Exit Sub
  
  '14/12/2005 - mpdea
  'Corrigido origem de dados (cboFornecedor.Text)
  '05/12/2005 - mpdea
  'Tratamento de overflow
  Call IsDataType(dtLong, cboFornecedor.Text, lngRet)
  
  '''Set rstFornecedor = db.OpenRecordset("SELECT Código, Nome, Tipo FROM Cli_For WHERE Código = " & lngRet & " AND Tipo ='F' ORDER BY Código", dbOpenDynaset, dbReadOnly)
  Set rstFornecedor = db.OpenRecordset("SELECT Código, Nome, Tipo FROM Cli_For WHERE Código = " & lngRet & " ORDER BY Código", dbOpenDynaset, dbReadOnly)

  With rstFornecedor
    If Not (.BOF And .EOF) Then
      txtFornecedor.Text = .Fields("Nome") & ""
    End If
  End With

  rstFornecedor.Close
  Set rstFornecedor = Nothing

  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboOperEntrada_CloseUp()
  cboOperEntrada.Text = cboOperEntrada.Columns(0).Text
  cboOperEntrada_LostFocus
End Sub

'05/12/2005 - mpdea
'Incluído tratamento de erro
Private Sub cboOperEntrada_LostFocus()
  Dim rstOperEntrada As Recordset
  Dim intRet As Integer

  On Error GoTo ErrHandler
  
  txtOperEntrada.Text = ""
  
  If Not IsNumeric(cboOperEntrada.Text) Then Exit Sub
  
  '05/12/2005 - mpdea
  'Tratamento de overflow
  Call IsDataType(dtInteger, cboOperEntrada.Text, intRet)
  
  Set rstOperEntrada = db.OpenRecordset("SELECT Código, Nome FROM [Operações Entrada] WHERE Código = " & intRet & " ORDER BY Código ", dbOpenDynaset, dbReadOnly)

  With rstOperEntrada
    If Not (.BOF And .EOF) Then
      txtOperEntrada.Text = .Fields("Nome") & ""
    End If
  End With

  rstOperEntrada.Close
  Set rstOperEntrada = Nothing

  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboOperSaida_CloseUp()
  cboOperSaida.Text = cboOperSaida.Columns(0).Text
  cboOperSaida_LostFocus
End Sub

'05/12/2005 - mpdea
'Incluído tratamento de erro
Private Sub cboOperSaida_LostFocus()
  Dim rstOperSaida As Recordset
  Dim intRet As Integer

  On Error GoTo ErrHandler
  
  txtOperSaida.Text = ""
  
  If Not IsNumeric(cboOperSaida.Text) Then Exit Sub
  
  '05/12/2005 - mpdea
  'Tratamento de overflow
  Call IsDataType(dtInteger, cboOperSaida.Text, intRet)
  
  Set rstOperSaida = db.OpenRecordset("SELECT Código, Nome FROM [Operações Saída] WHERE Código = " & intRet & " ORDER BY Código ", dbOpenDynaset, dbReadOnly)

  With rstOperSaida
    If Not (.BOF And .EOF) Then
      txtOperSaida.Text = .Fields("Nome") & ""
    End If
  End With

  rstOperSaida.Close
  Set rstOperSaida = Nothing
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboTabela_CloseUp()
  cboTabela.Text = cboTabela.Columns(0).Text

  Dim i As Integer
  For i = 0 To p_linhas
    If Tabe(i).Código <> "" Then
      Tabe(i).Valor = gcGetPrecoProduto(Tabe(i).Código, cboTabela.Text) * Tabe(i).Qtde
    End If
  Next i
  Grade1.Refresh
  Call Recalcula

End Sub

Private Sub cmd_consultarProduto_Click()
    nChamaConsulta = 6
    frmPesquisaProduto.Show
End Sub

Private Sub cmd_gravar_Click()
On Error GoTo Erro_Gravar
  Dim Aux_Str As String
  Dim i As Integer

  Call StatusMsg("")
  
  If lbl_status.Caption = "Concluída" Then
      MsgBox "Esta Transferência entre filiais já foi concluída."
      Exit Sub
  End If
  
  If Len(cboTabela.Text) <= 0 Then
    MsgBox "Selecione uma Tabela de Preços.", vbInformation, "Atenção"
    cboTabela.SetFocus
    Exit Sub
  End If
  
  If Len(Combo_Filial.Text) <= 0 Then
    MsgBox "Selecione a Filial para transferência", vbInformation, "Atenção"
    Combo_Filial.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(cboOperSaida.Text)) <= 0 Then
    MsgBox "Selecione uma Operação de transferência de saída.", vbInformation, "Atenção"
    cboOperSaida.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(cboOperEntrada.Text)) <= 0 Then
    MsgBox "Selecione uma Operação de transferência de entrada.", vbInformation, "Atenção"
    cboOperEntrada.SetFocus
    Exit Sub
  End If
  
  ' =====================================================================================
  ' Gravar na tabelas de:
  '                       TransferenciaEntreFiliais   - PAI
  '                       TransferenciaProdutos       - FILHA
  If O_Local.Value = True Then
  
      If lbl_status.Caption = "" Then
          
          'Novo registro
          Dim rsTransfFiliais   As Recordset
          Aux_Str = "Insert into TransferenciaEntreFiliais "
          Aux_Str = Aux_Str & " (FilialLogada, FilialExportada, CodigoFornecedor, "
          Aux_Str = Aux_Str & " CodigoCliente, CodigoOperSaida, CodigoOperEntrada, TabelaPrecos, "
          Aux_Str = Aux_Str & " PermitirTransfEstoqueInsuf, Data, Status, CodigoUsuario, QuantidadeItens, NumItens) "
          Aux_Str = Aux_Str & " Values(" & gnCodFilial & "," & Combo_Filial & "," & cboFornecedor.Text & ","
          Aux_Str = Aux_Str & cboCliente.Text & "," & cboOperSaida.Text & "," & cboOperEntrada.Text & ",'" & Trim(cboTabela.Text) & "',"
        
          If O_Estoque.Value = vbChecked Then
              Aux_Str = Aux_Str & "1,"
          Else
              Aux_Str = Aux_Str & "0,"
          End If
        
          ' Status:
          '          1-Em aberto
          '          2-Concluída
          Aux_Str = Aux_Str & "'" & Format(Data_Atual, "dd/mm/yyyy") & "', 1," & gnUserCode & "," & Total_Qtde.Caption & "," & Total_Itens & ")"

          db.Execute Aux_Str

          Aux_Str = "Select max(CodigoTransf) From TransferenciaEntreFiliais "
          Aux_Str = Aux_Str & " Where FilialLogada = " & gnCodFilial & " And CodigoUsuario = " & gnUserCode
          
          Set rsTransfFiliais = db.OpenRecordset(Aux_Str, dbOpenDynaset, dbReadOnly)

          With rsTransfFiliais
              If Not (.BOF And .EOF) Then
                  varInstancia_lCodigoTransf = .Fields(0).Value
              End If
          End With

          rsTransfFiliais.Close
          Set rsTransfFiliais = Nothing
          
      ElseIf lbl_status.Caption = "Em aberto" Then
      
          'Atualizar registro
          'Dim rsTransfFiliais   As Recordset
          
          If gridTransf.RowSel > 0 Then
              varInstancia_lCodigoTransf = CLng(gridTransf.TextMatrix(gridTransf.RowSel, 1))
          Else
              MsgBox "Não existe uma transferência seleciona.", vbInformation, "Atenção"
              Exit Sub
          End If
          
          Aux_Str = "Update TransferenciaEntreFiliais "
          Aux_Str = Aux_Str & " set FilialLogada = " & gnCodFilial & ", "
          Aux_Str = Aux_Str & " FilialExportada = " & Combo_Filial.Text & ", "
          Aux_Str = Aux_Str & " CodigoFornecedor = " & cboFornecedor.Text & ", "
          Aux_Str = Aux_Str & " CodigoCliente = " & cboCliente.Text & ", "
          Aux_Str = Aux_Str & " CodigoOperSaida = " & cboOperSaida.Text & ", "
          Aux_Str = Aux_Str & " CodigoOperEntrada = " & cboOperEntrada.Text & ", "
          Aux_Str = Aux_Str & " TabelaPrecos = '" & cboTabela.Text & "',"
        
          If O_Estoque.Value = vbChecked Then
              Aux_Str = Aux_Str & " PermitirTransfEstoqueInsuf = 1,"
          Else
              Aux_Str = Aux_Str & " PermitirTransfEstoqueInsuf = 0,"
          End If
        
          Aux_Str = Aux_Str & " Data = '" & Format(Data_Atual, "dd/mm/yyyy") & "', "
          
          ' Status:
          '          1-Em aberto
          '          2-Concluída
          Aux_Str = Aux_Str & " Status = 1, "
          Aux_Str = Aux_Str & " CodigoUsuario = " & gnUserCode & ", "
          Aux_Str = Aux_Str & " QuantidadeItens = " & Total_Qtde.Caption & ", "
          Aux_Str = Aux_Str & " NumItens = " & Total_Itens
          
          Aux_Str = Aux_Str & " Where CodigoTransf = " & varInstancia_lCodigoTransf

          ' Update
          db.Execute Aux_Str

          ' Delete de todos os produtos da transf.
          Aux_Str = "Delete from TransferenciaProdutos Where CodigoTransf = " & varInstancia_lCodigoTransf
          db.Execute Aux_Str
      End If

      Dim sNomeProdAux As String
      
      For i = 0 To p_linhas
          If Tabe(i).Código <> "" And Tabe(i).Qtde <> 0 Then
          
              sNomeProdAux = Tabe(i).Nome
              sNomeProdAux = Replace(sNomeProdAux, "'", " ")
          
              Aux_Str = "Insert into TransferenciaProdutos(CodigoTransf, codigoProduto, nomeProduto, Quantidade) "
              Aux_Str = Aux_Str & " Values(" & varInstancia_lCodigoTransf & ",'" & Tabe(i).Código & "','" & sNomeProdAux & "'," & str(Tabe(i).Qtde) & ")"

              db.Execute Aux_Str
          End If
      Next i
  End If
  ' =====================================================================================

  Call StatusMsg("")
  lbl_status.Caption = "Em aberto"
  DisplayMsg "Salvo com sucesso."
  cmdPesquisar_Click
  Exit Sub

Erro_Gravar:
  DisplayMsg "Erro ao salvar a transferência." & Err.Number & " " & Err.Description
  Exit Sub
End Sub

Private Sub cmd_verEstoqueAtual_Click()
On Error GoTo Erro
    
  If gridTransf.RowSel > 0 Then
      Dim objTela As frmPosicaoEstoqueAtualFiliais
      Set objTela = New frmPosicaoEstoqueAtualFiliais
      
      objTela.lCodigoTransferencia = gridTransf.TextMatrix(gridTransf.RowSel, 1)
      objTela.Show 1
      
      Set objTela = Nothing
    Else
        MsgBox "Selecione uma Transferência na GRADE SUPERIOR.", vbInformation, "Atenção"
        gridTransf.SetFocus
    End If
  
  Exit Sub
Erro:
  MsgBox "Erro na pesquisa de posição de estoque de Filiais" & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo Erro

  Dim Erro As Boolean
  
  'Limpar tela
  Erase Tabe
  Grade1.MoveLast
  Grade1.MoveFirst
  
  lbl_status.Caption = ""
  O_Estoque.Value = vbUnchecked
  Combo_Filial.Text = ""
  Filial_Entrada.Caption = ""
  Total_Itens.Caption = ""
  Total_Qtde.Caption = ""
  Filial_Destino.Text = ""
  
  cboOperSaida.Text = ""
  txtOperSaida.Text = ""
  cboOperEntrada.Text = ""
  txtOperEntrada.Text = ""
  cboCliente.Text = ""
  txtCliente.Text = ""
  cboTabela.Text = ""
  '-------------------------
  
  B_Transfere.Enabled = True
  '
  
  Erro = False
  If IsNull(Data_Ini.Text) Then Erro = True
  If Not Erro Then If Not IsDate(Data_Ini.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data incorreta, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
 
  Erro = False
  If IsNull(Data_Fim.Text) Then Erro = True
  If Not Erro Then If Not IsDate(Data_Fim.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data incorreta, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual a data final."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If CDate(Data_Fim.Text) - CDate(Data_Ini.Text) > 90 Then
    DisplayMsg "Escolha um período de até 90 dias"
    Data_Fim.SetFocus
    Exit Sub
  End If

  Call StatusMsg("Pesquisando transferências...")
  Screen.MousePointer = vbHourglass
  Call PesquisarTransferencias
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")

  Exit Sub
Erro:
  MsgBox "Erro na pesquisa de transferências " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub PesquisarTransferencias()
On Error GoTo Erro

  Dim sSql As String
  Dim rsTransfAux As Recordset
  Dim sStatus As String

  gridTransf.Rows = 1

  ' ==========================================================================================
  sSql = "Select T.CodigoTransf, T.FilialLogada, T.FilialExportada, T.CodigoFornecedor, T.CodigoCliente, "
  sSql = sSql & " T.CodigoOperSaida, T.CodigoOperEntrada, T.TabelaPrecos, T.PermitirTransfEstoqueInsuf, "
  sSql = sSql & " T.Data, T.Status, T.CodigoUsuario, T.QuantidadeItens, T.NumItens, P.Nome, F.Apelido "
  sSql = sSql & " From TransferenciaEntreFiliais T, [Parâmetros Filial] P, Funcionários F "
  
  If cmb_tipo.Text = "de saída de produtos" Then
      sSql = sSql & " Where T.FilialLogada = " & gnCodFilial
      sSql = sSql & " AND T.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  T.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
      sSql = sSql & " AND T.FilialExportada = P.Filial "
  Else
      sSql = sSql & " Where T.FilialExportada = " & gnCodFilial
      sSql = sSql & " AND T.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  T.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
      sSql = sSql & " AND T.FilialLogada = P.Filial "
  End If
  
  'Todas
  '1-Transferência Em aberto
  '2-Transferência Concluída
  If cmb_status.Text = "Transferência Em aberto" Then
      sSql = sSql & " AND T.Status=1"
  ElseIf cmb_status.Text = "Transferência Concluída" Then
      sSql = sSql & " AND T.Status=2"
  End If
  
  sSql = sSql & " AND T.CodigoUsuario=F.Código"
  
  sSql = sSql & " Order by T.Data desc"
  
  Set rsTransfAux = db.OpenRecordset(sSql, dbOpenDynaset)
 
  With rsTransfAux
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do Until .EOF
        If .Fields(10).Value = 1 Then
            sStatus = "Em aberto"
        Else
            sStatus = "Concluída"
        End If
        
        gridTransf.AddItem vbTab & .Fields(0).Value & vbTab & .Fields(9).Value & vbTab & _
                sStatus & vbTab & .Fields(12).Value & vbTab & .Fields(13).Value & vbTab & .Fields(14).Value & vbTab & _
                .Fields(2).Value & vbTab & .Fields(3).Value & vbTab & .Fields(4).Value & vbTab & _
                .Fields(5).Value & vbTab & .Fields(6).Value & vbTab & .Fields(7).Value & vbTab & _
                .Fields(8).Value & vbTab & .Fields(10).Value & vbTab & .Fields(15).Value
    
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsTransfAux = Nothing
 
  Exit Sub
Erro:
  MsgBox "Erro na pesquisa de transferências " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub


Private Sub cmdReadTransf_Click()
  Dim F As Form
  Set F = New frmLeTransfe
  F.Show
End Sub

Private Sub Combo_Filial_CloseUp()
On Error GoTo Erro

  Combo_Filial.Text = Combo_Filial.Columns(1).Text
  Combo_Filial_LostFocus

  
  cboOperEntrada.Text = rsParametros("Transf_OpEntrada")
  cboOperEntrada_LostFocus
  cboOperSaida.Text = rsParametros("Transf_OpSaida")
  cboOperSaida_LostFocus
  cboTabela.Text = rsParametros("Transf_TabelaPrecos")
  
  Dim i As Integer
  For i = 0 To p_linhas
    If Tabe(i).Código <> "" Then
      Tabe(i).Valor = gcGetPrecoProduto(Tabe(i).Código, cboTabela.Text) * Tabe(i).Qtde
    End If
  Next i
  Grade1.Refresh
  Call Recalcula
  
  Exit Sub
Erro:
  MsgBox "Erro ao selecionar Filial que receberá os produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Combo_Filial_LostFocus()

  Dim Aux As Variant
  
  Screen.MousePointer = vbHourglass
  Filial_Entrada.Caption = ""
  Aux = Combo_Filial.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 99 Then Exit Sub
  
  If Aux = gnCodFilial Then
      MsgBox "A filial de entrada de produtos deve ser DIFERENTE da filial de saída de produtos.", vbInformation, "Atenção"
      Combo_Filial.Text = ""
      Screen.MousePointer = vbDefault
      Exit Sub
  End If
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Aux)
  If rsParametros.NoMatch Then Exit Sub
  
  Filial_Entrada.Caption = rsParametros("Nome")
  BuscaFilialCliente (Combo_Filial.Text)
 
  Screen.MousePointer = vbDefault


End Sub

Private Sub DropDown1_Click()
 Dim Estoque As Single
 Estoque = Acha_Estoque(gnCodFilial, DropDown1.Columns(1).Text, 0, 0, 0, 0)
 L_Estoque.Caption = "Estoque: " + CStr(Estoque)
End Sub

Private Sub DropDown1_CloseUp()
  Grade1.Columns(Grade1.Col).Text = DropDown1.Columns(1).Text
  Call StatusMsg("")
End Sub

Private Sub DropDown1_DropDown()
  Dim rsTemp As Recordset
  Set rsTemp = db.OpenRecordset("SELECT Código FROM Produtos WHERE Código = '" & Grade1.Columns("Código").Text & "'", dbOpenSnapshot)
  If rsTemp.EOF Then
    DropDown1.DataFieldList = "Nome"
  Else
    DropDown1.DataFieldList = "Código"
  End If
  rsTemp.Close
  Set rsTemp = Nothing
End Sub

Private Sub BuscaFilialFornecedora()
On Error GoTo Erro
  Dim rsCliForAux As Recordset
  Dim sSql As String
  Dim sCNPJ_aux As String
  Dim sCNPJ_Fornecedor As String
  Dim i As Integer
  
  sSql = " SELECT Filial, CGC FROM [Parâmetros Filial] where Filial = " & gnCodFilial
  Set rsCliForAux = db.OpenRecordset(sSql, dbOpenSnapshot)
  If rsCliForAux.EOF And rsCliForAux.BOF Then
      MsgBox "Erro ao buscar dados da Filial.", vbInformation, "Atenção"
      lblFornecedor.BackColor = &H8080FF
      cboFornecedor.Enabled = False
      rsCliForAux.Close
      Set rsCliForAux = Nothing
      Exit Sub
  End If
  
  sCNPJ_Fornecedor = rsCliForAux.Fields("CGC").Value
  rsCliForAux.Close
  Set rsCliForAux = Nothing
  
  sSql = " SELECT Código, CGC FROM Cli_For "
  Set rsCliForAux = db.OpenRecordset(sSql, dbOpenSnapshot)

  If rsCliForAux.EOF And rsCliForAux.BOF Then
      MsgBox "Empresa/Filial inexistente ou dados incompletos. Vá no CADASTRO DE CLIENTES/FORNECEDORES e crie esta Filial informando o CNPJ.", vbInformation, "Atenção"
      lblFornecedor.BackColor = &H8080FF
      cboFornecedor.Enabled = False
      rsCliForAux.Close
      Set rsCliForAux = Nothing
      Exit Sub
  End If
  rsCliForAux.MoveLast
  rsCliForAux.MoveFirst
  
  sCNPJ_Fornecedor = Trim(sCNPJ_Fornecedor)
  sCNPJ_Fornecedor = Replace(sCNPJ_Fornecedor, ".", "")
  sCNPJ_Fornecedor = Replace(sCNPJ_Fornecedor, ";", "")
  sCNPJ_Fornecedor = Replace(sCNPJ_Fornecedor, ",", "")
  sCNPJ_Fornecedor = Replace(sCNPJ_Fornecedor, "/", "")
  sCNPJ_Fornecedor = Replace(sCNPJ_Fornecedor, "\", "")
  sCNPJ_Fornecedor = Replace(sCNPJ_Fornecedor, " ", "")
  sCNPJ_Fornecedor = Replace(sCNPJ_Fornecedor, "-", "")
  sCNPJ_Fornecedor = Replace(sCNPJ_Fornecedor, "_", "")

  For i = 0 To rsCliForAux.RecordCount - 1
      If Not IsNull(rsCliForAux.Fields("CGC").Value) Then
          sCNPJ_aux = rsCliForAux.Fields("CGC").Value
          sCNPJ_aux = Trim(sCNPJ_aux)
          sCNPJ_aux = Replace(sCNPJ_aux, ".", "")
          sCNPJ_aux = Replace(sCNPJ_aux, ";", "")
          sCNPJ_aux = Replace(sCNPJ_aux, ",", "")
          sCNPJ_aux = Replace(sCNPJ_aux, "/", "")
          sCNPJ_aux = Replace(sCNPJ_aux, "\", "")
          sCNPJ_aux = Replace(sCNPJ_aux, " ", "")
          sCNPJ_aux = Replace(sCNPJ_aux, "-", "")
          sCNPJ_aux = Replace(sCNPJ_aux, "_", "")
      Else
          sCNPJ_aux = ""
      End If
      
      If sCNPJ_Fornecedor = sCNPJ_aux Then
          cboFornecedor.Text = rsCliForAux.Fields("Código").Value
          cboFornecedor_LostFocus
          cboFornecedor.Enabled = False
          Exit For
      End If
      rsCliForAux.MoveNext
  Next
  
  If txtFornecedor.Text = "" Then
      MsgBox "Empresa/Filial inexistente ou dados incompletos. Vá no CADASTRO DE CLIENTES/FORNECEDORES e crie esta Filial informando o CNPJ.", vbInformation, "Atenção"
      lblFornecedor.BackColor = &H8080FF
      cboFornecedor.Text = ""
      cboFornecedor.Enabled = False
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro ao buscar o cadastro da empresa filial/fornecedora em Cli_For " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub BuscaFilialCliente(pFilialCliente As Integer)
On Error GoTo Erro
  Dim rsCliForAux As Recordset
  Dim sSql As String
  Dim sCNPJ_aux As String
  Dim sCNPJ_Cliente As String
  Dim i As Integer
  
  cboCliente.Text = ""
  txtCliente.Text = ""
  
  sSql = " SELECT Filial, CGC FROM [Parâmetros Filial] where Filial = " & pFilialCliente
  Set rsCliForAux = db.OpenRecordset(sSql, dbOpenSnapshot)
  If rsCliForAux.EOF And rsCliForAux.BOF Then
      MsgBox "Erro ao buscar dados da Filial.", vbInformation, "Atenção"
      lblCliente.BackColor = &H8080FF
      cboCliente.Enabled = False
      rsCliForAux.Close
      Set rsCliForAux = Nothing
      Exit Sub
  End If
  
  If IsNull(rsCliForAux.Fields("CGC").Value) Or Trim(rsCliForAux.Fields("CGC").Value) = "" Then
      MsgBox "No cadastro da Filial/Empresa informe o CNPJ.", vbInformation, "Atenção"
      lblCliente.BackColor = &H8080FF
      cboCliente.Enabled = False
      rsCliForAux.Close
      Set rsCliForAux = Nothing
      Exit Sub
  End If
  
  sCNPJ_Cliente = rsCliForAux.Fields("CGC").Value
  rsCliForAux.Close
  Set rsCliForAux = Nothing
  
  sSql = " SELECT Código, CGC FROM Cli_For "
  Set rsCliForAux = db.OpenRecordset(sSql, dbOpenSnapshot)

  If rsCliForAux.EOF And rsCliForAux.BOF Then
      MsgBox "Empresa/Filial inexistente ou dados incompletos. Vá no CADASTRO DE CLIENTES/FORNECEDORES e crie esta Filial informando o CNPJ.", vbInformation, "Atenção"
      lblCliente.BackColor = &H8080FF
      cboCliente.Enabled = False
      rsCliForAux.Close
      Set rsCliForAux = Nothing
      Exit Sub
  End If
  rsCliForAux.MoveLast
  rsCliForAux.MoveFirst
  
  sCNPJ_Cliente = Trim(sCNPJ_Cliente)
  sCNPJ_Cliente = Replace(sCNPJ_Cliente, ".", "")
  sCNPJ_Cliente = Replace(sCNPJ_Cliente, ";", "")
  sCNPJ_Cliente = Replace(sCNPJ_Cliente, ",", "")
  sCNPJ_Cliente = Replace(sCNPJ_Cliente, "/", "")
  sCNPJ_Cliente = Replace(sCNPJ_Cliente, "\", "")
  sCNPJ_Cliente = Replace(sCNPJ_Cliente, " ", "")
  sCNPJ_Cliente = Replace(sCNPJ_Cliente, "-", "")
  sCNPJ_Cliente = Replace(sCNPJ_Cliente, "_", "")

  For i = 0 To rsCliForAux.RecordCount - 1
      If Not IsNull(rsCliForAux.Fields("CGC").Value) Then
          sCNPJ_aux = rsCliForAux.Fields("CGC").Value
          sCNPJ_aux = Trim(sCNPJ_aux)
          sCNPJ_aux = Replace(sCNPJ_aux, ".", "")
          sCNPJ_aux = Replace(sCNPJ_aux, ";", "")
          sCNPJ_aux = Replace(sCNPJ_aux, ",", "")
          sCNPJ_aux = Replace(sCNPJ_aux, "/", "")
          sCNPJ_aux = Replace(sCNPJ_aux, "\", "")
          sCNPJ_aux = Replace(sCNPJ_aux, " ", "")
          sCNPJ_aux = Replace(sCNPJ_aux, "-", "")
          sCNPJ_aux = Replace(sCNPJ_aux, "_", "")
      Else
          sCNPJ_aux = ""
      End If
      
      If sCNPJ_Cliente = sCNPJ_aux Then
          cboCliente.Text = rsCliForAux.Fields("Código").Value
          cboCliente_LostFocus
          cboCliente.Enabled = False
          Exit For
      End If
      rsCliForAux.MoveNext
  Next
  rsCliForAux.Close
  Set rsCliForAux = Nothing
  
  If txtCliente.Text = "" Then
      MsgBox "Empresa/Filial inexistente ou dados incompletos. Vá no CADASTRO DE CLIENTES/FORNECEDORES e crie esta Filial informando o CNPJ.", vbInformation, "Atenção"
      lblCliente.BackColor = &H8080FF
      cboCliente.Text = ""
      cboCliente.Enabled = False
      Exit Sub
  End If
  
  lblCliente.BackColor = &H8000000F
  
  Exit Sub
Erro:
  MsgBox "Erro ao buscar o cadastro da empresa filial/fornecedora em Cli_For " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  cmb_tipo.ListIndex = 0
  cmb_status.ListIndex = 0
  Data_Ini.Text = Format(Date - 60, "DD/MM/YYYY")
  Data_Fim.Text = Format(Date, "DD/MM/YYYY")

  gridTransf.ColWidth(0) = 0
  gridTransf.ColWidth(1) = 1200
  gridTransf.ColWidth(2) = 1250
  gridTransf.ColWidth(3) = 2500
  gridTransf.ColWidth(4) = 1500
  gridTransf.ColWidth(5) = 1500
  gridTransf.ColWidth(6) = 4800
  gridTransf.ColWidth(7) = 0
  gridTransf.ColWidth(8) = 0
  gridTransf.ColWidth(9) = 0
  gridTransf.ColWidth(10) = 0
  gridTransf.ColWidth(11) = 0
  gridTransf.ColWidth(12) = 0
  gridTransf.ColWidth(13) = 0
  gridTransf.ColWidth(14) = 0
  gridTransf.ColWidth(15) = 2000
  
  gridTransf.Row = 0
  gridTransf.TextMatrix(0, 0) = ""
  gridTransf.TextMatrix(0, 1) = "Código"
  gridTransf.TextMatrix(0, 2) = "Data"
  gridTransf.TextMatrix(0, 3) = "Status"
  gridTransf.TextMatrix(0, 4) = "Qtde Itens"
  gridTransf.TextMatrix(0, 5) = "Num Itens"
  gridTransf.TextMatrix(0, 6) = "Filial envolvida"
  gridTransf.TextMatrix(0, 15) = "Usuário"
  
  
  Data1.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  
  '29/04/2004 - Daniel
  datOperSaida.DatabaseName = gsQuickDBFileName
  datOperEntrada.DatabaseName = gsQuickDBFileName
  datCliente.DatabaseName = gsQuickDBFileName
  datFornecedor.DatabaseName = gsQuickDBFileName
  datTabela.DatabaseName = gsQuickDBFileName
  '-----------------------------------------------
  
  txtFornecedor.Enabled = False
  txtCliente.Enabled = False
  BuscaFilialFornecedora

  
  '30/01/2009 - mpdea
  'Adaptado para o novo menu
  'Key: Q7MENU
  'cmdReadTransf.Enabled = (frmMain.ActiveBar1.Tools("miMovLeTransf").Enabled = True)
  cmdReadTransf.Enabled = Val(gbGetUserPermition(gnUserCode, 130)) > 0 'miMovLeTransf
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsOp_Saída = db.OpenRecordset("Operações Saída", , dbReadOnly)
  Set rsOp_Entrada = db.OpenRecordset("Operações Entrada", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque")
  Set rsEstoque_Final = db.OpenRecordset("Estoque Final")
  ' Set rsResumo = db.OpenRecordset("Resumo Produtos")
  Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
  
  rsCores.Index = "Código"
  rsTamanhos.Index = "Código"
  rsProdutos.Index = "Código"
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then Exit Sub
  Filial_Saída.Caption = str(gnCodFilial) + " - " + rsParametros("Nome")
  
  ' UTILIZAR O LIMITADOR DE LINHAS PARAMETRIZADO 13/02/2023 - Pablo
  p_linhas = CInt(rsParametros("Linhas Digitação"))
  'ReDim Tabe(p_linhas)
  Grade1.Rows = p_linhas
  
  cmdGrade.Visible = gbGrade
  
  cmdPesquisar_Click
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Set frmTransfere = Nothing
  rsParametros.Close
  rsOp_Saída.Close
  rsOp_Entrada.Close
  rsProdutos.Close
  rsGrade.Close
  rsEstoque.Close
  rsEstoque_Final.Close
  rsCores.Close
  rsTamanhos.Close
  
  Set rsParametros = Nothing
  Set rsOp_Saída = Nothing
  Set rsOp_Entrada = Nothing
  Set rsProdutos = Nothing
  Set rsGrade = Nothing
  Set rsEstoque = Nothing
  Set rsEstoque_Final = Nothing
  Set rsCores = Nothing
  Set rsTamanhos = Nothing
  
  p_linhas = 0
End Sub

Function Acha_Linha()
'Dim Linha As Integer
'Dim J As Integer
'
'Linha = -1
'For J = 0 To p_linhas
'   If Linha = -1 And Tabe(J).Código = "" Then
'     Acha_Linha = J
'     Exit Function
'   End If
'Next J
'
'Acha_Linha = -1

  
  Dim nX As Integer
  
  For nX = 0 To p_linhas
    If Tabe(nX).Código = "" Then
      Acha_Linha = nX
      Exit Function
    End If
  Next nX
  Acha_Linha = -1
End Function

Private Sub Grade1_AfterUpdate(RtnDispErrMsg As Integer)
  Recalcula
End Sub

Private Sub Grade1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  Dim Aux2 As Variant
  Dim Cód As String
  Dim Cód1 As String
  Dim Cód_Str As String
  Dim Valor As Single
  Dim Valor_Int As Long
  Dim Aux_Str As String
  Dim Balança As Integer
  Dim Comp_Prod As Integer
  Dim Preço_Balança As Double
  Dim Início_Preço As Integer
  Dim Tam_Preço As Integer
  Dim Aux_Preço As Double
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Estoque As Double
  Dim Edição As Long
  Dim Tipo As Integer
  Dim Erro As Integer
  Dim Nome_Prod As String
  Dim Valor_Prod As Currency

'Call StatusMsg("")


  Aux = Grade1.Columns(ColIndex).Text

  If ColIndex = 0 Then 'Código
  
    If IsNull(Aux) Or Aux = "" Or Aux = "0" Then
       Grade1.Columns(1).Text = ""
       Grade1.Columns(2).Text = 0
       Exit Sub
    End If
  
    Acha_Produto CStr(Aux), Cód, Tamanho, Cor, Edição, Tipo, Erro
    
    If Erro <> 0 Then
      If Erro = 1 Then
        DisplayMsg "Produto não encontrado."
        Cancel = True
        Exit Sub
      End If
      If Erro = 2 Then
        DisplayMsg "Este produto usa grade, digite cor e tamanho."
        Cancel = True
        Grade1.SetFocus
        Exit Sub
      End If
      If Erro = 3 Then
        DisplayMsg "Este produto tem edição, digite a edição."
        Cancel = True
        Exit Sub
      End If
      If Erro = 4 Then
        DisplayMsg "Produto com grade sem produto principal."
        Cancel = True
        Exit Sub
      End If
    End If
    
    rsProdutos.Seek "=", Cód
    
      
    Nome_Prod = rsProdutos("Nome")
    
    If Tamanho <> 0 Then
      rsTamanhos.Seek "=", Tamanho
      If Not rsTamanhos.NoMatch Then
        Nome_Prod = Nome_Prod + "  " + rsTamanhos("Nome")
      End If
    End If
    
    If Cor <> 0 Then
      rsCores.Seek "=", Cor
      If Not rsCores.NoMatch Then
        Nome_Prod = Nome_Prod + "  " + rsCores("Nome")
      End If
    End If
      
      
    Grade1.Columns(1).Text = Nome_Prod
    
    
    L_Estoque.Caption = ""
    
    Estoque = Acha_Estoque(gnCodFilial, Cód, Tamanho, Cor, 0, 0)
    L_Estoque.Caption = CStr(Estoque)
    
    Valor_Prod = gcGetPrecoProduto(Cód, cboTabela.Text)
    Grade1.Columns("Valor").Text = Valor_Prod
    
    Call StatusMsg("")
  End If
  
  
  
  
  
  Rem QTDE
  Rem QTDe
  Rem QTde
  
  If ColIndex = 2 Then 'Qtde
    Rem Acha o produto
    Aux2 = Grade1.Columns(0).Text
    Cód = Aux2
    Tamanho = 0
    Cor = 0
    Estoque = -999999
    
    
    Acha_Produto CStr(Aux2), Cód, Tamanho, Cor, Edição, Tipo, Erro
    

    Estoque = Acha_Estoque(gnCodFilial, Cód, Tamanho, Cor, Edição, Erro)
  
    Valor_Prod = gcGetPrecoProduto(Cód, cboTabela.Text)
  
    'Verifica se Qtde é decimal
Cont_Qtde:
    If IsNull(Aux) Then
       Grade1.Columns(2).Text = 0
       Exit Sub
    End If
    If Aux = "" Then
       Grade1.Columns(2).Text = 0
       Exit Sub
    End If
    
    If Not IsNumeric(Aux) Then
       DisplayMsg "Quantidade inválida."
       Cancel = True
       Exit Sub
    End If
    
    If CDbl(Aux) < 0 Then
       DisplayMsg "Quantidade inválida."
       Cancel = True
       Exit Sub
    End If
    
    If CDbl(Aux) > 9999999 Then
       DisplayMsg "Quantidade inválida."
       Cancel = True
       Exit Sub
    End If
    
    
    
    If O_Estoque.Value = 0 Then
      If CDbl(Aux) > Estoque Then
        If Estoque <> -999999 Then
          DisplayMsg "Quantidade superior ao estoque. Estoque atual : " + Format(Estoque, "#########0")
          If CDbl(Aux) <> 0 Then Cancel = True
          Exit Sub
        End If
      End If
    End If
    
    If CDbl(Aux) < 0 Then
      DisplayMsg "Quantidade incorreta."
      Cancel = True
      Exit Sub
    End If
    
    
    
    Valor = Aux
    Valor_Int = Aux
    If Valor = Valor_Int Then
      Exit Sub
    End If
    
    Aux = Grade1.Columns(0).Text
    'Acha produto
    If Not IsNumeric(Aux) Or Val(Aux) < 0 Then Exit Sub
    If Val(Aux) > 999999999999999# Then Exit Sub
    If IsNull(Aux) Or Aux = "" Or Val(Aux) = 0 Then Exit Sub
    
    Cód = Aux
    rsProdutos.Index = "Código"
    rsProdutos.Seek "=", Aux
    If rsProdutos.NoMatch Then
      rsGrade.Index = "Código"
      rsGrade.Seek "=", Aux
      If rsGrade.NoMatch Then Exit Sub
      Cód = rsGrade("Código Original")
      rsProdutos.Seek "=", Cód
      If rsProdutos.NoMatch Then Exit Sub
    End If
    
    If rsProdutos("Fracionado") = False Then
      DisplayMsg "Produto não aceita quantidade fracionada."
      Cancel = True
      Exit Sub
    End If

    Grade1.Columns("Valor").Text = Valor_Prod * Valor
    'Grade1.Update

  End If
    


End Sub

Private Sub Grade1_InitColumnProps()
  Grade1.Columns(0).DropDownHwnd = DropDown1.hwnd
End Sub

Private Sub Grade1_KeyPress(KeyAscii As Integer)
  Dim sCodigo As String
  
  With Grade1
    If .Col = 0 Then 'Código
      If DropDown1.DroppedDown Then
        DropDown1.DataFieldList = "Nome"
      End If
      If KeyAscii = vbKeyReturn Then
        If Not DropDown1.DroppedDown Then
          KeyAscii = 0
          sCodigo = .Columns("Código").Text
          If sCodigo <> "" Then
            If Left(sCodigo, 1) = "*" Then
              sCodigo = Right(sCodigo, Len(sCodigo) - 1)
              .Columns("Código").Text = sCodigo
              .Columns("Nome").Text = gsGetNameProduto(sCodigo)
              SendKeys "{TAB}", True
            Else
              .Columns("Quantidade").Text = 1
              SendKeys "{DOWN}", True
            End If
            '.Columns("Valor").Text = gcGetPrecoProduto(sCodigo, cboTabela.Text)
            '.Update
          End If
        End If
      End If
    ElseIf .Col = 2 Then 'Quantidade
      If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        KeyAscii = 0
        If .Columns("Código").Text <> "" Then
          SendKeys "{DOWN}{HOME}{ESC}", True
          '.Columns("Valor").Text = gcGetPrecoProduto(.Columns("Código").Text, cboTabela.Text) * CDbl(.Columns("Quantidade").Text)
          '.Update
        End If
      End If
    End If
  End With
End Sub

Private Sub Grade1_LostFocus()
  With Grade1
    If .RowChanged Then
      .Update
    End If
  End With
End Sub

Private Sub Grade1_PrintError(ByVal PrintError As Long, Response As Integer)
  If PrintError = 30457 Then
    Response = False
  End If
End Sub

Private Sub Grade1_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
  Dim Texto As String

  Texto = "Transferência - Origem :  " + Filial_Saída.Caption + "   destino : "
  
  If Filial_Destino.Visible = False Then
     Texto = Texto + Filial_Entrada.Caption
  Else
     Texto = Texto + Filial_Destino.Text
  End If
  
  Texto = Texto + "  Data : " + Format(Date, "dd/mm/yyyy")
  
  ssPrintInfo.PageHeader = Texto

End Sub

Private Sub Grade1_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  If LastRow <> Grade1.Row Then
    L_Estoque.Caption = ""
  End If
End Sub

Private Sub Grade1_UnboundAddData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, NewRowBookmark As Variant)
  Dim Linha As Integer
  
  Linha = Grade1.Row
  
  Tabe(Linha).Código = Grade1.Columns(0).Text
  Tabe(Linha).Nome = Grade1.Columns(1).Text
  Tabe(Linha).Qtde = CDbl(Grade1.Columns(2).Text)
  'Tabe(Linha).Valor = CDbl(Grade1.Columns(3).Text)
  Tabe(Linha).Valor = gcGetPrecoProduto(Tabe(Linha).Código, cboTabela.Text) * Tabe(Linha).Qtde
End Sub

Private Sub Grade1_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade1.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p

End Sub


Private Sub Grade1_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim r, i, J, p As Integer

If IsNull(StartLocation) Then
  If ReadPriorRows Then
    p = Grade1.Rows
  Else
    p = 0
  End If
 Else
  p = StartLocation
  If ReadPriorRows Then
    p = p - 1
  Else
    p = p + 1
  End If
End If

For i = 0 To RowBuf.RowCount - 1
  If p < 0 Or p >= Grade1.Rows Then Exit For
     RowBuf.Value(i, 0) = Tabe(p).Código
     RowBuf.Value(i, 1) = Tabe(p).Nome
     RowBuf.Value(i, 2) = Tabe(p).Qtde
     RowBuf.Value(i, 3) = Tabe(p).Valor
     
     
   RowBuf.Bookmark(i) = p
   If ReadPriorRows Then
     p = p - 1
   Else
     p = p + 1
   End If
   
   r = r + 1
 Next i
 
 RowBuf.RowCount = r
   
     

End Sub


Private Sub Grade1_UnboundWriteData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, WriteLocation As Variant)
 Dim Linha As Integer
 
 On Error GoTo Erro
 Linha = WriteLocation

Tabe(Linha).Código = Grade1.Columns(0).Text
Tabe(Linha).Nome = Grade1.Columns(1).Text
Tabe(Linha).Qtde = CDbl(Grade1.Columns(2).Text)
'Tabe(Linha).Valor = CDbl(Grade1.Columns(3).Text)
Tabe(Linha).Valor = gcGetPrecoProduto(Tabe(Linha).Código, cboTabela.Text) * Tabe(Linha).Qtde

Exit Sub

Erro:
 Exit Sub

End Sub


Private Sub cmdGrade_Click()
  Dim nTotalLinhas As Integer
  Dim nAuxTamanho(14) As Integer
  Dim sTamanhos(14) As String
  Dim nQtdes(14) As Integer
  Dim nAuxQtde(14) As Integer
  Dim sCor As String
  Dim nRet As Integer
  Dim nCol As Integer
  Dim nLine As Integer
  Dim sCodigo As String
  Dim snome As String
  Dim nCor As Integer
  
  With frmDigitaGrade
    .O_Preço.Value = 0
    .O_Desconto.Value = 0
    .O_Impostos.Value = 0
    .O_Etiqueta.Value = 0
    .Limpa_Variáveis
    .Show vbModal
    If .Retorno.Caption <> "OK" Then
      Exit Sub
    End If
    nTotalLinhas = .Retorno1.Caption
    
    'Obtém os Tamanhos da grade digitada
    Call .RetornarTamanhos(nAuxTamanho())
  
    For nCol = 0 To 14
      sTamanhos(nCol) = Format(nAuxTamanho(nCol), "000")
    Next nCol
  
    For nLine = 1 To nTotalLinhas
      Call .RetornarLinhaGrade(nLine, sCodigo, snome, nCor, _
        nAuxQtde(), 0, 0, 0, 0, 0, 0)
        
      For nCol = 0 To 14
        nQtdes(nCol) = nAuxQtde(nCol)
      Next nCol
      
      sCor = Format(nCor, "000")
      
      For nCol = 0 To 14
        If nQtdes(nCol) <> 0 Then
          'Acha Ponto de Inclusão
          nRet = Acha_Linha
          If nRet <> -1 Then
            With Tabe(nRet)
              .Código = sCodigo & sTamanhos(nCol) & sCor
              .Qtde = nQtdes(nCol)
              .Nome = snome
            End With
          End If
        End If
      Next nCol
    Next nLine
  End With
  
  Call Recalcula
  
  Grade1.MoveLast
  Grade1.MoveFirst

End Sub

Private Sub gridTransf_Click()
    DetalharTransferencia
End Sub

Private Sub DetalharTransferencia()
On Error GoTo Erro
  Dim strSQL As String
  Dim rstTransfDet As Recordset
  Dim sCodigoTrans As String
  Dim sFilialLogada As String
  Dim sFilialExportada As String
  Dim sCodigoFornecedor As String
  Dim sCodigoCliente As String
  Dim sCodigoOperSaida As String
  Dim sCodigoOperEntrada As String
  Dim sTabelaPrecos As String
  Dim sPermitirTransfEstoqueInsuf As String
  Dim sData As String
  Dim sStatus As String
  Dim sQuantidadeItens As String
  Dim sNumItens As String
  Dim snome As String
  
  strSQL = ""
  
  If gridTransf.RowSel > 0 Then
  
    sCodigoTrans = gridTransf.TextMatrix(gridTransf.RowSel, 1)
    sFilialLogada = gsNomeFilial
    sFilialExportada = gridTransf.TextMatrix(gridTransf.RowSel, 7)
    sCodigoFornecedor = gridTransf.TextMatrix(gridTransf.RowSel, 8)
    sCodigoCliente = gridTransf.TextMatrix(gridTransf.RowSel, 8)
    sCodigoOperSaida = gridTransf.TextMatrix(gridTransf.RowSel, 10)
    sCodigoOperEntrada = gridTransf.TextMatrix(gridTransf.RowSel, 11)
    sTabelaPrecos = gridTransf.TextMatrix(gridTransf.RowSel, 12)
    sPermitirTransfEstoqueInsuf = gridTransf.TextMatrix(gridTransf.RowSel, 13)
    sData = gridTransf.TextMatrix(gridTransf.RowSel, 14)
    sStatus = gridTransf.TextMatrix(gridTransf.RowSel, 3)
    sQuantidadeItens = gridTransf.TextMatrix(gridTransf.RowSel, 4)
    sNumItens = gridTransf.TextMatrix(gridTransf.RowSel, 5)
    snome = gridTransf.TextMatrix(gridTransf.RowSel, 6)
    
    'Dados de cabeçalho
    lbl_status.Caption = sStatus
    
    If sPermitirTransfEstoqueInsuf = "1" Then
        O_Estoque.Value = vbChecked
    Else
        O_Estoque.Value = vbUnchecked
    End If
    
    O_Local.Value = True
    
    cboTabela.Text = sTabelaPrecos
    cboOperSaida.Text = sCodigoOperSaida
    cboOperSaida_LostFocus
    cboOperEntrada.Text = sCodigoOperEntrada
    cboOperEntrada_LostFocus
    
    Filial_Saída.Caption = gsNomeFilial
    cboFornecedor.Text = sCodigoFornecedor
    cboFornecedor_LostFocus
    Combo_Filial.Text = sFilialExportada
    Combo_Filial_LostFocus
        
    'Dados de produtos (da grid)
    strSQL = "SELECT * From TransferenciaProdutos Where CodigoTransf = " & sCodigoTrans
 
    Set rstTransfDet = db.OpenRecordset(strSQL, dbOpenDynaset)

    Dim X As Integer
    X = 0
    
    Grade1.Redraw = False
  
    Erase Tabe
    Grade1.MoveLast
    Grade1.MoveFirst
 
    With rstTransfDet
        
        If Not (.EOF And .BOF) Then
            .MoveFirst
        End If
        
        While Not .EOF
    
            Grade1.Columns(0).Text = .Fields("CodigoProduto").Value
            Grade1.Columns(1).Text = .Fields("NomeProduto").Value
            Grade1.Columns(2).Text = .Fields("Quantidade").Value
            Grade1.Columns(3).Text = gcGetPrecoProduto(.Fields("CodigoProduto").Value, cboTabela.Text) * .Fields("Quantidade").Value
            
            Tabe(X).Código = .Fields("CodigoProduto").Value
            Tabe(X).Nome = .Fields("NomeProduto").Value
            Tabe(X).Qtde = .Fields("Quantidade").Value
            Tabe(X).Valor = gcGetPrecoProduto(.Fields("CodigoProduto").Value, cboTabela.Text) * .Fields("Quantidade").Value

'Tabe(Linha).Valor = gcGetPrecoProduto(Tabe(Linha).Código, cboTabela.Text) * Tabe(Linha).Qtde

            Grade1.MoveNext
            X = X + 1
                
            .MoveNext
        Wend
        .Close
    End With
    Set rstTransfDet = Nothing
  End If
  
  Grade1.Redraw = True
  
  If sStatus = "Concluída" Then
    Grade1.AllowUpdate = False
    Combo_Filial.Enabled = False
    cboOperEntrada.Enabled = False
    cboOperSaida.Enabled = False
    cboTabela.Enabled = False
  End If

  Exit Sub
Erro:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
End Sub


Private Sub O_Distante_Click()
  Combo_Filial.Visible = False
  Filial_Entrada.Visible = False
  Label3.Visible = False
  Nome_Filial.Visible = True
  Filial_Destino.Top = 510
  Filial_Destino.Visible = True
  Nome_Filial.Top = 570
  '29/03/2004 - Daniel
  lblOpEntrada.Caption = ""
  cboOperEntrada.Visible = False
  txtOperEntrada.Visible = False
  
  lblFornecedor.Visible = False
  cboFornecedor.Visible = False
  txtFornecedor.Visible = False
  
  cboCliente.Enabled = True
  
  lblCliente.BackColor = &H8000000F
  cboCliente.Text = ""
  txtCliente.Text = ""
  lblCliente.Top = 210
  cboCliente.Top = 150
  txtCliente.Top = 150
  cmdReadTransf.Visible = True
  
  lbl_status.Visible = False
  
End Sub

Private Sub O_Estoque_Click()
  If O_Estoque.Value = 1 Then Call O_Estoque_SenhaGerente
End Sub

Private Sub O_Estoque_KeyPress(KeyAscii As Integer)
  If O_Estoque.Value = 1 Then Call O_Estoque_SenhaGerente
End Sub

Private Sub O_Estoque_SenhaGerente()
    If Not frmGerente.gbSenhaGerente Then
      O_Estoque.Value = 0 'False
      Grade1.SetFocus
      Exit Sub
    End If
End Sub

Private Sub O_Local_Click()
  cmdReadTransf.Visible = False
  Combo_Filial.Visible = True
  Filial_Entrada.Visible = True
  Label3.Visible = True
  Nome_Filial.Visible = False
  Filial_Destino.Visible = False
  '29/03/2004 - Daniel
  lblOpEntrada.Caption = "Operação Entrada"
  cboOperEntrada.Visible = True
  txtOperEntrada.Visible = True
  
  lblFornecedor.Visible = True
  cboFornecedor.Visible = True
  txtFornecedor.Visible = True
  
  cboCliente.Enabled = False
  
  Combo_Filial.Text = ""
  Filial_Entrada.Caption = ""
  cboCliente.Text = ""
  txtCliente.Text = ""
  
  lblCliente.Top = 600
  cboCliente.Top = 570
  txtCliente.Top = 570

End Sub

Private Sub CriarRegistros(ByVal strProduto As String, ByVal sngQtde As Single, Optional ByVal SituacaoTributaria As String = "0")
'Finalidade de Criar Saídas, [Saídas - Produtos], Entradas, [Entradas - Produtos] e Atualizar
'a movimentação em Parâmetros
'Case: Casagrande
  Dim rstSaidas           As Recordset
  Dim rstSaidasProdutos   As Recordset
  Dim rstParametros       As Recordset
  'Dim nSequencia         As Long Precisa ser global m_nSequencia
  Dim rstEntradas         As Recordset
  Dim rstEntradasProdutos As Recordset
  '27/04/2004 - Daniel
  'Var para tratamento da Observacao
  Dim strObs              As String
  Dim sValorTotalSaidas   As String
  
  
  '24/10/2005 - mpdea
  'Tabela de Preços
  Dim strTabelaPrecos As String
  
  strTabelaPrecos = Trim(cboTabela.Text)
  
  
  m_intContador = m_intContador + 1

  If m_intContador = 1 Then
      dPrecoFinalSaidas = 0

      m_nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("Última Movimentação") + 1

      m_nSequenciaStoreSaidas = m_nSequencia 'Para armazenar o valor da Saída
      
      'Saídas
      Set rstSaidas = db.OpenRecordset("SELECT * FROM Saídas", dbOpenDynaset)

      With rstSaidas
        .AddNew

        .Fields("Filial").Value = gnCodFilial
        .Fields("Data").Value = Data_Atual
        .Fields("Sequência").Value = m_nSequenciaStoreSaidas
        .Fields("Operação").Value = CInt(cboOperSaida.Text)
        .Fields("Caixa").Value = 1
        .Fields("Tabela").Value = Trim(cboTabela.Text)
        .Fields("Digitador").Value = gnUserCode
        .Fields("Operador").Value = gnUserCode
        .Fields("Cliente").Value = CLng(cboCliente.Text)
        
        If O_Local.Value = True Then
          strObs = "QS-TRANSF-FILIAIS:Exportado p/ " & Combo_Filial.Text & "-" & Filial_Entrada.Caption
        Else
          strObs = "QS-TRANSF-FILIAIS:Exportado p/ " & Filial_Destino.Text
        End If
        
        If Len(strObs) >= 68 Then   '70 é o limite mas estava dando problema com os 70 na Barro Queimado
          strObs = Left(strObs, 68) 'Para não passar do limite...
        End If
        
        .Fields("Observações").Value = strObs
        .Fields("Produtos").Value = sngQtde * Format(gcGetPrecoProduto(strProduto, strTabelaPrecos), FORMAT_VALUE)
        .Fields("Desconto").Value = 0
        .Fields("Serviços").Value = 0
        .Fields("Total").Value = .Fields("Produtos").Value
        .Fields("Efetivada").Value = True
        .Fields("Recebimento").Value = True
        .Fields("DescontoSubTotal").Value = 0
        .Fields("Percentual CSLL").Value = 0
        .Fields("Percentual COFINS").Value = 0
        .Fields("Percentual PIS").Value = 0
        .Fields("Percentual IRRF").Value = 0

        .Update
        .Close
      End With

      Set rstSaidas = Nothing
      'Fim Saídas

      'Parâmetros
      Set rstParametros = db.OpenRecordset(" SELECT [Última Movimentação] FROM [Parâmetros Filial] WHERE Filial =" & gnCodFilial, dbOpenDynaset)

        rstParametros.Edit
        rstParametros.Fields("Última Movimentação").Value = m_nSequencia
        rstParametros.Update
        rstParametros.Close

      Set rstParametros = Nothing
      'Fim Parâmetros

  Else
      ' Sómente atualizar o Valor Total da Transferência
      sValorTotalSaidas = Format(dPrecoFinalSaidas, FORMAT_VALUE) + (sngQtde * Format(gcGetPrecoProduto(strProduto, strTabelaPrecos), FORMAT_VALUE))
      sValorTotalSaidas = Replace(sValorTotalSaidas, ",", ".")
      db.Execute "Update Saídas set Produtos=" & sValorTotalSaidas & ", Total=" & sValorTotalSaidas & " Where Filial=" & gnCodFilial & " and Sequência=" & m_nSequenciaStoreSaidas

  End If 'if m_intContador = 1

        
  If m_intContador >= 1 Then

      '[Saídas - Produtos]
      Set rstSaidasProdutos = db.OpenRecordset(" SELECT * FROM [Saídas - Produtos]", dbOpenDynaset)

        With rstSaidasProdutos
          .AddNew

          .Fields("Filial").Value = gnCodFilial
          .Fields("Sequência").Value = m_nSequenciaStoreSaidas
          .Fields("Linha").Value = m_intContador
          '05/07/2004 - Daniel
          'Tratamento caso produto possua grade
          If m_blnComGrade Then
            .Fields("Código").Value = strProduto & m_strTamanho & m_strCor
          Else
            .Fields("Código").Value = strProduto
          End If
          '----------------------------------------------------------------
          .Fields("Qtde").Value = sngQtde
          
          
          '24/10/2005 - mpdea
          'Incluído o preço do produto de acordo com a tabela de preços selecionada
          .Fields("Preço").Value = Format(gcGetPrecoProduto(strProduto, strTabelaPrecos), FORMAT_VALUE)
          
          
          .Fields("Desconto").Value = 0
          .Fields("ICM").Value = 0
          .Fields("IPI").Value = 0
          .Fields("Preço Final").Value = .Fields("Preço").Value * .Fields("Qtde").Value
          
          If m_intContador > 1 Then
              dPrecoFinalSaidas = dPrecoFinalSaidas + .Fields("Preço Final").Value
          Else
              dPrecoFinalSaidas = .Fields("Preço Final").Value
          End If
          
          '27/04/2004 - Daniel
          'Adicionado o field Código sem Grade
          .Fields("Código sem Grade").Value = strProduto
          .Fields("Situação Tributária").Value = SituacaoTributaria
          
          .Update
          .Close
        End With

      Set rstSaidasProdutos = Nothing
      'Fim [Saídas - Produtos]

  End If 'if m_intContador >= 1 then


  If O_Local.Value = True Then '=====> MESMA BASE <=====
  
    If m_intContador = 1 Then

      m_nSequencia = gnGetNextSequencia(CByte(Combo_Filial.Text)) 'rsParametros("Última Movimentação") + 1

      'Entradas
      Set rstEntradas = db.OpenRecordset("SELECT * FROM Entradas", dbOpenDynaset)
        
          With rstEntradas
            .AddNew
            
            .Fields("Filial").Value = CByte(Combo_Filial.Text)
            .Fields("Data").Value = Data_Atual
            .Fields("Sequência").Value = m_nSequencia
            .Fields("Operação").Value = CInt(cboOperEntrada.Text)
            .Fields("Digitador").Value = gnUserCode
            .Fields("Fornecedor").Value = CLng(cboFornecedor.Text)
            
            strObs = "Importado da Empresa " & Filial_Saída.Caption & " em " & Format(Data_Atual, "dd/mm/yy")
        
            If Len(strObs) >= 68 Then   '70 é o limite mas estava dando problema com os 70 na Barro Queimado
              strObs = Left(strObs, 68) 'Para não passar do limite...
            End If
            
            .Fields("Observações").Value = strObs
            .Fields("Data Emissão").Value = Data_Atual
            .Fields("Produtos").Value = 0
            .Fields("Caixa").Value = 1
            .Fields("Efetivada").Value = True
              
            .Update
            .Close
          End With
          
      Set rstEntradas = Nothing
      'Fim Entradas
      
      'Parâmetros
      Set rstParametros = db.OpenRecordset(" SELECT [Última Movimentação] FROM [Parâmetros Filial] WHERE Filial =" & CByte(Combo_Filial.Text), dbOpenDynaset)
        
        rstParametros.Edit
        rstParametros.Fields("Última Movimentação").Value = m_nSequencia
        rstParametros.Update
        rstParametros.Close

      Set rstParametros = Nothing
      'Fim 'Parâmetros
      
    End If 'If m_intContador = 1
    
    If m_intContador >= 1 Then
        
      '[Entradas - Produtos]
      Set rstEntradasProdutos = db.OpenRecordset("SELECT * FROM [Entradas - Produtos]", dbOpenDynaset)
      
        With rstEntradasProdutos
          .AddNew
          
          .Fields("Filial").Value = CByte(Combo_Filial.Text)
          .Fields("Sequência").Value = m_nSequencia
          .Fields("Linha").Value = m_intContador
          '05/07/2004 - Daniel
          'Tratamento caso o produto possua grade
          If m_blnComGrade Then
            .Fields("Código").Value = strProduto & m_strTamanho & m_strCor
          Else
            .Fields("Código").Value = strProduto
          End If
          '--------------------------------------
          .Fields("Qtde").Value = sngQtde
          
          
          '24/10/2005 - mpdea
          'Incluído o preço do produto de acordo com a tabela de preços selecionada
          .Fields("Preço").Value = Format(gcGetPrecoProduto(strProduto, strTabelaPrecos), FORMAT_VALUE)
          
          
          .Fields("Desconto").Value = 0
          .Fields("ICM").Value = 0
          .Fields("IPI").Value = 0
          .Fields("Preço Final").Value = 0
          '27/04/2004 - Daniel
          'Adicionado o field Código sem Grade para impressão
          'correta de relatórios
          .Fields("Código sem Grade").Value = strProduto
          
          .Update
          .Close
        End With
      
      Set rstEntradasProdutos = Nothing
      'Fim [Entradas - Produtos]
      
    End If 'm_intContador >= 1
    
    
  End If 'If O_Local.Value = True Then...

End Sub
