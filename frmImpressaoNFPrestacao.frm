VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmImpressaoNFPrestacao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpressaoNFPrestacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   11535
   Begin VB.CommandButton cmdGerarEntrada 
      BackColor       =   &H8000000A&
      Caption         =   "&Gerar Entrada"
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Geração com todos os ítens da Grid."
      Top             =   5760
      Width           =   2535
   End
   Begin VB.TextBox txtTotalCompras 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "0,00"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox txtTotalPrestacao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "0,00"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCarregar 
      BackColor       =   &H0000C0C0&
      Caption         =   "Carregar"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Data datCaixa 
      Caption         =   "datCaixa"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT Caixa, Descrição FROM [Caixas em Uso] ORDER BY Caixa"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Data datConta 
      Caption         =   "datConta"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT * FROM [Contas Bancárias] ORDER BY  Código"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Data datCentroCusto 
      Caption         =   "datCentroCusto"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM [Centros de Custo] ORDER BY Código"
      Top             =   7320
      Width           =   2235
   End
   Begin VB.Data datOperacao 
      Caption         =   "datOperacao"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT Código, Nome, Tipo FROM [Operações Entrada] WHERE NOT Estoque AND Tipo = 'C' ORDER BY Código"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Data datOperacaoCompra 
      Caption         =   "datOperacaoCompra"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT Código, Nome, Tipo FROM [Operações Entrada] WHERE Estoque = FALSE AND Tipo = 'C' ORDER BY Código"
      Top             =   7680
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Frame fraPesquisa 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   11295
      Begin VB.Frame fraPeriodo 
         Appearance      =   0  'Flat
         Caption         =   " Período ( Vendas )"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   8520
         TabIndex        =   19
         Top             =   120
         Width           =   2535
         Begin MSMask.MaskEdBox mskDataFinal 
            Height          =   315
            Left            =   1320
            TabIndex        =   2
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDataInicial 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   1320
            TabIndex        =   21
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   6015
      End
      Begin VB.Data datFornecedores 
         Caption         =   "datFornecedores"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Cli_FOR WHERE Tipo = 'F'"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2340
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmImpressaoNFPrestacao.frx":058A
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2055
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
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   7805
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).DataField=   "Nome"
         Columns(0).FieldLen=   256
         Columns(1).Width=   3731
         Columns(1).Caption=   "Codigo"
         Columns(1).Name =   "Codigo"
         Columns(1).DataField=   "Código"
         Columns(1).FieldLen=   256
         _ExtentX        =   3625
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   825
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ítens"
      TabPicture(0)   =   "frmImpressaoNFPrestacao.frx":05A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdGeral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pagamento"
      TabPicture(1)   =   "frmImpressaoNFPrestacao.frx":05C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "lblNomeOperacaoCompra"
      Tab(1).Control(3)=   "lblNomeOperacao"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "lblLabels(29)"
      Tab(1).Control(6)=   "lblLabels(30)"
      Tab(1).Control(7)=   "Label4"
      Tab(1).Control(8)=   "lblNomeCC"
      Tab(1).Control(9)=   "cboOperEntradaCompra"
      Tab(1).Control(10)=   "cboOperEntrada"
      Tab(1).Control(11)=   "cboCodigoCC"
      Tab(1).Control(12)=   "grdCP"
      Tab(1).Control(13)=   "txtTotDigitado"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtADigitar"
      Tab(1).Control(15)=   "Frame2"
      Tab(1).Control(16)=   "Frame3"
      Tab(1).ControlCount=   17
      Begin VB.Frame Frame3 
         Caption         =   "Cheque Usado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   29
         Top             =   1920
         Width           =   4335
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   960
            MaxLength       =   38
            TabIndex        =   8
            Top             =   600
            Width           =   3240
         End
         Begin VB.TextBox txtValCheque 
            Height          =   330
            Left            =   3000
            MaxLength       =   12
            TabIndex        =   10
            Top             =   960
            Width           =   1185
         End
         Begin VB.TextBox txtCheque 
            Height          =   315
            Left            =   960
            MaxLength       =   10
            TabIndex        =   9
            Top             =   960
            Width           =   1275
         End
         Begin SSDataWidgets_B.SSDBCombo cboConta 
            Bindings        =   "frmImpressaoNFPrestacao.frx":05E0
            Height          =   315
            Left            =   960
            TabIndex        =   7
            Top             =   240
            Width           =   1050
            DataFieldList   =   "Código"
            MaxDropDownItems=   16
            _Version        =   196617
            BackColorOdd    =   12648447
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   3969
            Columns(0).Caption=   "Agência"
            Columns(0).Name =   "Agência"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Agência"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   2249
            Columns(1).Caption=   "Código"
            Columns(1).Name =   "Código"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Código"
            Columns(1).DataType=   2
            Columns(1).FieldLen=   256
            Columns(2).Width=   5503
            Columns(2).Caption=   "Descrição"
            Columns(2).Name =   "Descrição"
            Columns(2).CaptionAlignment=   0
            Columns(2).DataField=   "Descrição"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            _ExtentX        =   1852
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   65535
         End
         Begin MSMask.MaskEdBox medDataBomPara 
            Height          =   315
            Left            =   960
            TabIndex        =   11
            ToolTipText     =   "Pressione F2 para Calendário"
            Top             =   1320
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
         Begin VB.Label lblConta 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2040
            TabIndex        =   35
            Top             =   240
            Width           =   2145
         End
         Begin VB.Label lblLabels 
            Caption         =   "Valor:"
            Height          =   240
            Index           =   28
            Left            =   2520
            TabIndex        =   34
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Bom Para:"
            Height          =   225
            Index           =   18
            Left            =   120
            TabIndex        =   33
            Top             =   1365
            Width           =   780
         End
         Begin VB.Label lblLabels 
            Caption         =   "Descrição:"
            Height          =   225
            Index           =   17
            Left            =   135
            TabIndex        =   32
            Top             =   645
            Width           =   795
         End
         Begin VB.Label lblLabels 
            Caption         =   "Número:"
            Height          =   240
            Index           =   16
            Left            =   120
            TabIndex        =   31
            Top             =   997
            Width           =   900
         End
         Begin VB.Label lblLabels 
            Caption         =   "C&onta:"
            Height          =   225
            Index           =   15
            Left            =   150
            TabIndex        =   30
            Top             =   285
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Caixa Usado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   4335
         Begin VB.TextBox txtCxCheque 
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   6
            Top             =   960
            Width           =   2625
         End
         Begin VB.TextBox txtCxDinheiro 
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   5
            Top             =   600
            Width           =   2625
         End
         Begin SSDataWidgets_B.SSDBCombo cboCaixaUso 
            Bindings        =   "frmImpressaoNFPrestacao.frx":05F7
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1335
            DataFieldList   =   "Caixa"
            MaxDropDownItems=   16
            _Version        =   196617
            BackColorOdd    =   12648447
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   5556
            Columns(0).Caption=   "Descrição"
            Columns(0).Name =   "Descrição"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Descrição"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1323
            Columns(1).Caption=   "Caixa"
            Columns(1).Name =   "Caixa"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Caixa"
            Columns(1).DataType=   2
            Columns(1).FieldLen=   256
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   65535
         End
         Begin VB.Label lblCaixaUso 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   1560
            TabIndex        =   28
            Top             =   240
            Width           =   2625
         End
         Begin VB.Label lblLabels 
            Caption         =   "&Cheque:"
            Height          =   225
            Index           =   13
            Left            =   120
            TabIndex        =   27
            Top             =   1005
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            Caption         =   "&Dinheiro:"
            Height          =   225
            Index           =   12
            Left            =   120
            TabIndex        =   26
            Top             =   645
            Width           =   1395
         End
      End
      Begin VB.TextBox txtADigitar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   24
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtTotDigitado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2955
         Width           =   2055
      End
      Begin SSDataWidgets_B.SSDBGrid grdCP 
         Height          =   3015
         Left            =   -70440
         TabIndex        =   12
         Top             =   570
         Width           =   3015
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   8438015
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   2037
         Columns(0).Caption=   "Data"
         Columns(0).Name =   "Data"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2143
         Columns(1).Caption=   "Valor"
         Columns(1).Name =   "Valor"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).NumberFormat=   "##,###,##0.00"
         Columns(1).FieldLen=   256
         _ExtentX        =   5318
         _ExtentY        =   5318
         _StockProps     =   79
         Caption         =   "Contas a Pagar"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCodigoCC 
         Bindings        =   "frmImpressaoNFPrestacao.frx":060E
         Height          =   300
         Left            =   -67320
         TabIndex        =   13
         Top             =   795
         Width           =   1095
         DataFieldList   =   "Código"
         _Version        =   196617
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Codigo"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Código"
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Centro de custo"
         Columns(1).Name =   "Centro de custo"
         Columns(1).DataField=   "Nome"
         Columns(1).FieldLen=   256
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   93
         Text            =   "0"
         BackColor       =   65535
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperEntrada 
         Bindings        =   "frmImpressaoNFPrestacao.frx":062B
         Height          =   300
         Left            =   -67320
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
         DataFieldList   =   "Código"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   65535
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperEntradaCompra 
         Bindings        =   "frmImpressaoNFPrestacao.frx":0645
         Height          =   300
         Left            =   -67320
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
         DataFieldList   =   "Código"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   65535
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBGrid grdGeral 
         Height          =   3255
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   10965
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   12
         RowHeight       =   423
         Columns.Count   =   12
         Columns(0).Width=   3122
         Columns(0).Caption=   "Fornecedor"
         Columns(0).Name =   "Fornecedor"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   1614
         Columns(1).Caption=   "Nota"
         Columns(1).Name =   "Nota"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   1111
         Columns(2).Caption=   "Seq"
         Columns(2).Name =   "Sequencia"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   900
         Columns(3).Caption=   "Linha"
         Columns(3).Name =   "Linha"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(4).Width=   1826
         Columns(4).Caption=   "Codigo"
         Columns(4).Name =   "Codigo"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
         Columns(5).Width=   3731
         Columns(5).Caption=   "Nome"
         Columns(5).Name =   "Nome"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(5).Locked=   -1  'True
         Columns(6).Width=   873
         Columns(6).Caption=   "Qtde"
         Columns(6).Name =   "Qtde"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(6).Locked=   -1  'True
         Columns(7).Width=   1773
         Columns(7).Caption=   "Preco Custo"
         Columns(7).Name =   "Preco"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(7).Locked=   -1  'True
         Columns(8).Width=   1270
         Columns(8).Caption=   "Vendido"
         Columns(8).Name =   "Vendido"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(8).Locked=   -1  'True
         Columns(9).Width=   1561
         Columns(9).Caption=   "Comprado"
         Columns(9).Name =   "Comprado"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(9).Locked=   -1  'True
         Columns(10).Width=   1535
         Columns(10).Caption=   "Devolvido"
         Columns(10).Name=   "Devolvido"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(10).Locked=   -1  'True
         Columns(11).Width=   1191
         Columns(11).Caption=   "Acertar"
         Columns(11).Name=   "QtdeAcertada"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(11).Locked=   -1  'True
         Columns(11).HasForeColor=   -1  'True
         Columns(11).ForeColor=   255
         _ExtentX        =   19341
         _ExtentY        =   5741
         _StockProps     =   79
         Caption         =   "Ítens para Prestação de Contas com o Fornecedor"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblNomeCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   -66120
         TabIndex        =   44
         Top             =   795
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo"
         Height          =   195
         Left            =   -67320
         TabIndex        =   43
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "A Digitar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   -67320
         TabIndex        =   42
         Top             =   3420
         Width           =   930
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Digitado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   -67320
         TabIndex        =   41
         Top             =   3015
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Operações de Entrada:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -67320
         TabIndex        =   40
         Top             =   1200
         Width           =   1680
      End
      Begin VB.Label lblNomeOperacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   -66120
         TabIndex        =   39
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblNomeOperacaoCompra 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   -66120
         TabIndex        =   38
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Para Prestação (Não Somar ao Estoque)"
         Height          =   195
         Left            =   -67320
         TabIndex        =   37
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Para Compra (Não Somar ao Estoque)"
         Height          =   195
         Left            =   -67320
         TabIndex        =   36
         Top             =   2040
         Width           =   2730
      End
   End
   Begin VB.Label lblCompras 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Compras (R$)"
      Height          =   195
      Left            =   3420
      TabIndex        =   50
      Top             =   5820
      Width           =   1395
   End
   Begin VB.Label lblPrestacao 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Prestação (R$)"
      Height          =   195
      Left            =   90
      TabIndex        =   48
      Top             =   5820
      Width           =   1485
   End
End
Attribute VB_Name = "frmImpressaoNFPrestacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gbConsignacaoResultado As Boolean
'---------------------------[Nota:]------------------------------
'Apesar deste form possuir o nome de frmImpressaoNFPrestacao
'ele apenas gerará as entradas ora de compras, ora de prestações
'oriundas da prestação de contas
'----------------------------------------------------------------

Private Sub cmdCarregar_Click()
  If ValidarObjetosAoCarregar Then Exit Sub
  
  Call CarregarGrid
  
  cboFornecedor_LostFocus
  
End Sub

Private Sub cmdGerarEntrada_Click()

  '-------------------------------------------------------------------------------
  'Verificar se a grid está vazia
  '-------------------------------------------------------------------------------
  If grdGeral.Rows <= 0 Then
    MsgBox "Não há nenhum ítem carregado, verifique.", vbExclamation, "Atenção"
    Exit Sub
  End If

  If ValidacaoGeral Then Exit Sub

  Call StatusMsg("Aguarde criando o cerne de Entradas....")
  Screen.MousePointer = vbHourglass
  Call CriarEntrada
  
  Call StatusMsg("Atualizando a tabela Prestação de Contas...")
  'Atualizamos o campo PrestacaoFechada.PrestacaoFechada quando for Prestação de Contas
  'ou o campo PrestacaoFechada.CompraFechada quando for Compra
  Call AtualizarPrestacaoContas
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
  MsgBox "Entrada de " & frmManPrestacaodeContas.strCaption & " geradas com sucesso.", vbInformation, "Quick Store"
  
  txtTotalPrestacao.Text = "0,00"
  txtTotalCompras.Text = "0,00"
  
  grdGeral.Redraw = False
  grdGeral.RemoveAll
  grdGeral.Refresh
  grdGeral.Redraw = True
  
  'Limpando os objetos
  cboCaixaUso.Text = ""
  lblCaixaUso.Caption = ""
  txtCxDinheiro.Text = ""
  txtCxCheque.Text = ""
  cboConta.Text = ""
  lblConta.Caption = ""
  txtDescricao.Text = ""
  txtCheque.Text = ""
  txtValCheque.Text = ""
  medDataBomPara.Mask = ""
  medDataBomPara.Text = ""
  medDataBomPara.Mask = "##/##/####"
  grdCP.RemoveAll
  cboCodigoCC.Text = ""
  lblNomeCC.Caption = ""
  cboOperEntrada.Text = ""
  lblNomeOperacao.Caption = ""
  cboOperEntradaCompra.Text = ""
  lblNomeOperacaoCompra.Caption = ""
  txtTotDigitado.Text = ""
  txtADigitar.Text = ""
  
  SSTab1.Tab = 0
  SSTab1.TabCaption(0) = "&Ítens"
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFornecedores.DatabaseName = gsQuickDBFileName
  datCaixa.DatabaseName = gsQuickDBFileName
  datConta.DatabaseName = gsQuickDBFileName
  datCentroCusto.DatabaseName = gsQuickDBFileName
  datOperacao.DatabaseName = gsQuickDBFileName
  datOperacaoCompra.DatabaseName = gsQuickDBFileName
  
'  mskDataInicial.Text = "01/01/" & Year(Data_Atual)
'  mskDataFinal.Text = CDate(Data_Atual + 30)
  
  Me.Caption = "Geração de Entrada - " & frmManPrestacaodeContas.strCaption
  
  If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
    lblPrestacao.Visible = True
    txtTotalPrestacao.Visible = True
    lblCompras.Visible = False
    txtTotalCompras.Visible = False
    cboOperEntrada.BackColor = &HFFFF&
    cboOperEntradaCompra.BackColor = &HFFFFFF
  Else 'Compras
    lblPrestacao.Visible = False
    txtTotalPrestacao.Visible = False
    lblCompras.Visible = True
    txtTotalCompras.Visible = True
    cboOperEntrada.BackColor = &HFFFFFF
    cboOperEntradaCompra.BackColor = &HFFFF&
  End If
  
End Sub

Private Sub cboCaixaUso_CloseUp()
  cboCaixaUso.Text = cboCaixaUso.Columns(1).Text
  cboCaixaUso_LostFocus
End Sub

Private Sub cboCaixaUso_LostFocus()
  Dim rstCaixaUso As Recordset
  
  lblCaixaUso.Caption = ""
  
  If Not IsNumeric(cboCaixaUso.Text) Then Exit Sub
  
  Set rstCaixaUso = db.OpenRecordset("SELECT Caixa, Descrição FROM [Caixas em Uso] WHERE Caixa = " & CByte(cboCaixaUso.Text) & " ORDER BY Caixa ", dbOpenDynaset)

  With rstCaixaUso
    If Not (.BOF And .EOF) Then
      lblCaixaUso.Caption = .Fields("Descrição") & ""
    End If
  End With

  rstCaixaUso.Close
  Set rstCaixaUso = Nothing

End Sub

Private Sub cboCodigoCC_CloseUp()
  cboCodigoCC.Text = cboCodigoCC.Columns(0).Text
  cboCodigoCC_LostFocus
End Sub

Private Sub cboCodigoCC_LostFocus()
  Dim rstCodigoCC As Recordset
  
  lblNomeCC.Caption = ""
  
  If Not IsNumeric(cboCodigoCC.Text) Then Exit Sub
  
  Set rstCodigoCC = db.OpenRecordset("SELECT Código, Nome FROM [Centros de Custo] WHERE Código = " & CByte(cboCodigoCC.Text) & " ORDER BY Código ", dbOpenDynaset)

  With rstCodigoCC
    If Not (.BOF And .EOF) Then
      lblNomeCC.Caption = .Fields("Nome") & ""
    End If
  End With

  rstCodigoCC.Close
  Set rstCodigoCC = Nothing

End Sub

Private Sub cboConta_CloseUp()
  cboConta.Text = cboConta.Columns(1).Text
  cboConta_LostFocus
End Sub

Private Sub cboConta_LostFocus()
  Dim rstConta As Recordset
  
  lblConta.Caption = ""
  
  If Not IsNumeric(cboConta.Text) Then Exit Sub
  
  Set rstConta = db.OpenRecordset("SELECT Código, Descrição FROM [Contas Bancárias] WHERE Código = " & CByte(cboConta.Text) & " ORDER BY Código ", dbOpenDynaset)

  With rstConta
    If Not (.BOF And .EOF) Then
      lblConta.Caption = .Fields("Descrição") & ""
    End If
  End With

  rstConta.Close
  Set rstConta = Nothing

End Sub

Private Sub cboOperEntrada_CloseUp()
  cboOperEntrada.Text = cboOperEntrada.Columns(0).Text
  cboOperEntrada_LostFocus
End Sub

Private Sub cboOperEntrada_LostFocus()
  Dim rstOper As Recordset
  
  lblNomeOperacao.Caption = ""
  
  If Not IsNumeric(cboOperEntrada.Text) Then Exit Sub
  
  Set rstOper = db.OpenRecordset("SELECT Código, Nome FROM [Operações Entrada] WHERE Código = " & CInt(cboOperEntrada.Text) & " ORDER BY Código ", dbOpenDynaset)

  With rstOper
    If Not (.BOF And .EOF) Then
      lblNomeOperacao.Caption = .Fields("Nome") & ""
    End If
  End With

  rstOper.Close
  Set rstOper = Nothing

End Sub

Private Sub cboOperEntradaCompra_CloseUp()
  cboOperEntradaCompra.Text = cboOperEntradaCompra.Columns(0).Text
  cboOperEntradaCompra_LostFocus
End Sub

Private Sub cboOperEntradaCompra_LostFocus()
  Dim rstOper As Recordset
  
  lblNomeOperacaoCompra.Caption = ""
  
  If Not IsNumeric(cboOperEntradaCompra.Text) Then Exit Sub
  
  Set rstOper = db.OpenRecordset("SELECT Código, Nome FROM [Operações Entrada] WHERE Código = " & CInt(cboOperEntradaCompra.Text) & " ORDER BY Código ", dbOpenDynaset)

  With rstOper
    If Not (.BOF And .EOF) Then
      lblNomeOperacaoCompra.Caption = .Fields("Nome") & ""
    End If
  End With

  rstOper.Close
  Set rstOper = Nothing

End Sub

Private Sub grdCP_AfterDelete(RtnDispErrMsg As Integer)
  Call RecalculaPagar
End Sub

Private Sub grdCP_KeyDown(KeyCode As Integer, Shift As Integer)
  If grdCP.Col = 0 Then
    Call HandleKeyDown(KeyCode, Shift)
    Select Case KeyCode
      Case vbKeyF2
        grdCP.Columns("Data").Text = frmCalendario.gsDateCalender(grdCP.Columns("Data").Text)
    End Select
  End If
End Sub

Private Sub grdCP_LostFocus()
  grdCP.Update
  Call RecalculaPagar
End Sub

Private Sub RecalculaPagar()
  Dim nRow      As Long
  Dim nTotal    As Double
  Dim nValorDif As Double
  Dim bm        As Variant
  
  On Error Resume Next
  
  nTotal = 0
  For nRow = 0 To grdCP.Rows - 1
    bm = grdCP.AddItemBookmark(nRow)
    If IsDate(grdCP.Columns("Data").CellValue(bm)) Then
      If IsNumeric(grdCP.Columns("Valor").CellValue(bm)) Then
        nTotal = nTotal + CDbl(grdCP.Columns("Valor").CellValue(bm))
      Else
        gsTitle = LoadResString(201)
        gsMsg = "Valor Incorreto em Pagamentos."
        gnStyle = vbOKOnly + vbCritical
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Exit Sub
      End If
    End If
  Next nRow
  
  'nTotal = Round(nTotal, 2)
   
  nTotal = nTotal + CDbl(gsHandleNull(txtCxDinheiro.Text & ""))
  nTotal = nTotal + CDbl(gsHandleNull(txtCxCheque.Text & ""))
  nTotal = nTotal + CDbl(gsHandleNull(txtValCheque.Text & ""))
   
  txtTotDigitado.Text = gsFormatCurrency(nTotal, True)
  
  If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
    nValorDif = CDbl(gsHandleNull(txtTotalPrestacao.Text)) - CDbl(gsHandleNull(txtTotDigitado.Text & ""))
  Else
    nValorDif = CDbl(gsHandleNull(txtTotalCompras.Text)) - CDbl(gsHandleNull(txtTotDigitado.Text & ""))
  End If
  
  txtADigitar.Text = gsFormatCurrency(nValorDif, False)

End Sub

Private Sub medDataBomPara_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    medDataBomPara.Text = frmCalendario.gsDateCalender(medDataBomPara.Text)
  End If
End Sub

Private Sub medDataBomPara_LostFocus()
  medDataBomPara.Text = Ajusta_Data(medDataBomPara.Text)
End Sub

Private Sub txtCxCheque_LostFocus()
  Call RecalculaPagar
End Sub

Private Sub txtCxDinheiro_LostFocus()
  Call RecalculaPagar
End Sub

Private Sub txtValCheque_LostFocus()
  Call RecalculaPagar
End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(1).Text
  txtNomeFornecedor.Text = FindFornecedor
End Sub

Private Sub cboFornecedor_LostFocus()
  txtNomeFornecedor.Text = FindFornecedor
End Sub

Private Function FindFornecedor(Optional lngFornecedor As Long)
  Dim strSQL    As String
  Dim rstForn   As Recordset
  Dim lng_Fornecedor As Long
  
  txtNomeFornecedor.Text = ""
  
  If lngFornecedor <= 0 Then
    If Len(Trim(cboFornecedor.Text)) <= 0 Then Exit Function
    If Not IsNumeric(Trim(cboFornecedor.Text)) Then Exit Function
    
    lng_Fornecedor = CLng(cboFornecedor.Text)
  Else
    lng_Fornecedor = lngFornecedor
  End If
  
  strSQL = " SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' AND "
  strSQL = strSQL & " Código = " & lng_Fornecedor
  
  Set rstForn = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstForn
    If Not (.BOF And .EOF) Then
      FindFornecedor = .Fields("Nome").Value & ""
    Else
      FindFornecedor = ""
    End If
    .Close
  End With
  
  Set rstForn = Nothing
End Function

Private Function FindProduto(strCodigo As String)
  Dim strSQL    As String
  Dim rstProd   As Recordset
  
  strSQL = " SELECT Nome FROM Produtos WHERE Código = '" & strCodigo & "'"
  
  Set rstProd = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstProd
    If Not (.BOF And .EOF) Then
      FindProduto = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  
  Set rstProd = Nothing
End Function

Private Function ValidarObjetosAoCarregar() As Boolean
  
  If Len(txtNomeFornecedor.Text) <= 0 Then
    ValidarObjetosAoCarregar = True
    MsgBox "Fornecedor inválido, verifique.", vbExclamation, "Atenção"
    cboFornecedor.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataInicial.Text) Then
    ValidarObjetosAoCarregar = True
    MsgBox "Data Inicial inválida, verifique.", vbExclamation, "Atenção"
    mskDataInicial.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataFinal.Text) Then
    ValidarObjetosAoCarregar = True
    MsgBox "Data Final inválida, verifique.", vbExclamation, "Atenção"
    mskDataFinal.SetFocus
    Exit Function
  End If
  
  If CDate(mskDataFinal.Text) < CDate(mskDataInicial.Text) Then
    ValidarObjetosAoCarregar = True
    MsgBox "Data Final menor que a Inicial, verifique.", vbExclamation, "Atenção"
    mskDataFinal.SetFocus
    Exit Function
  End If
  
End Function

Private Sub CarregarGrid()
  Dim rstPrestacao    As Recordset
  Dim strSQL          As String
  Dim dblTotCompras   As Double
  Dim dblTotPrestacao As Double
  
  strSQL = "SELECT * FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Fornecedor = " & CLng(cboFornecedor.Text)
  strSQL = strSQL & " AND PeriodoVenda >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND PeriodoVenda <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  
  If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
    strSQL = strSQL & " AND NOT PrestacaoFechada "
    strSQL = strSQL & " AND NOT CompraFechada " '13/10/2004 AND NOT CompraFechada
    strSQL = strSQL & " AND ( Resultado = 3 OR Resultado = 5 OR Resultado = 6 )"
  Else 'Compras
    strSQL = strSQL & " AND NOT CompraFechada "
    strSQL = strSQL & " AND NOT PrestacaoFechada " '13/10/2004 AND NOT PrestacaoFechada
    strSQL = strSQL & " AND ( Resultado = 2 OR Resultado = 4 OR Resultado = 6 ) "
  End If
  
  strSQL = strSQL & " ORDER BY Filial, Sequencia, Linha "
  
  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstPrestacao.RecordCount = 0 Then
    MsgBox "Não existe informação dentro do intervalo solicitado, verifique.", vbInformation, "Quick Store"
    rstPrestacao.Close
    Set rstPrestacao = Nothing
    Exit Sub
  End If
  
  grdGeral.Redraw = False
  grdGeral.RemoveAll
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        grdGeral.AddNew
        grdGeral.Columns("Fornecedor").Value = .Fields("Fornecedor").Value & " - " & FindFornecedor(CLng(.Fields("Fornecedor").Value))
        grdGeral.Columns("Nota").Value = .Fields("NotaFiscal").Value
        grdGeral.Columns("Sequencia").Value = .Fields("Sequencia").Value
        grdGeral.Columns("Linha").Value = .Fields("Linha").Value
        grdGeral.Columns("Codigo").Value = .Fields("Produto").Value
        grdGeral.Columns("Nome").Value = FindProduto(CStr(.Fields("Produto").Value))
        grdGeral.Columns("Qtde").Value = .Fields("QtdeOriginal").Value
        grdGeral.Columns("Preco").Value = Format(.Fields("Custo").Value, FORMAT_VALUE)
        grdGeral.Columns("Vendido").Value = .Fields("QtdeVendida").Value
        grdGeral.Columns("Devolvido").Value = .Fields("QtdeDevolvida").Value
        grdGeral.Columns("Comprado").Value = .Fields("QtdeComprada").Value
        grdGeral.Columns("QtdeAcertada").Value = .Fields("QtdeAcertada").Value
        
        grdGeral.Update
      
        If .Fields("QtdeComprada").Value <> 0 Then dblTotCompras = dblTotCompras + (.Fields("QtdeComprada").Value * .Fields("Custo").Value)
        If .Fields("QtdeAcertada").Value <> 0 Then dblTotPrestacao = dblTotPrestacao + (.Fields("QtdeAcertada").Value * .Fields("Custo").Value)
      
      .MoveNext
      Loop
      
    End If
    
    txtTotalCompras.Text = Format(dblTotCompras, FORMAT_VALUE)
    txtTotalPrestacao.Text = Format(dblTotPrestacao, FORMAT_VALUE)
    
    .Close
  End With
  
  grdGeral.MoveFirst
  grdGeral.Redraw = True
  
  Set rstPrestacao = Nothing
  
End Sub

Private Function ValidacaoGeral() As Boolean

  If Len(txtADigitar.Text) <= 0 Then
    ValidacaoGeral = True
    MsgBox "Valor do Pagamento incorreto, verifique.", vbExclamation, "Atenção"
    Exit Function
  End If

  If IsNumeric(txtADigitar.Text) Then
    If CDbl(txtADigitar.Text) <> 0 Then
      ValidacaoGeral = True
      MsgBox "Valor do Pagamento incorreto, verifique.", vbExclamation, "Atenção"
      Exit Function
    End If
  End If

  If Len(lblCaixaUso.Caption) <= 0 Then
    ValidacaoGeral = True
    MsgBox "Caixa incorreto, verifique.", vbExclamation, "Atenção"
    cboCaixaUso.SetFocus
    Exit Function
  End If
  
  If Len(lblNomeCC.Caption) <= 0 Then
    ValidacaoGeral = True
    MsgBox "Centro de Custo incorreto, verifique.", vbExclamation, "Atenção"
    cboCodigoCC.SetFocus
    Exit Function
  End If
  
  If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
    If Len(lblNomeOperacao.Caption) <= 0 Then
      ValidacaoGeral = True
      MsgBox "Operação de Entrada para Prestação de Contas incorreta, verifique.", vbExclamation, "Atenção"
      cboOperEntrada.SetFocus
      Exit Function
    End If
  Else
    If Len(lblNomeOperacaoCompra.Caption) <= 0 Then
      ValidacaoGeral = True
      MsgBox "Operação de Entrada para Compra incorreta, verifique.", vbExclamation, "Atenção"
      cboOperEntradaCompra.SetFocus
      Exit Function
    End If
  End If

End Function

Private Sub CriarEntrada()
  Dim rstParametros   As Recordset
  Dim rstEntradas     As Recordset
  Dim rstEntraProdu   As Recordset
  Dim rsMovi_Parcelas As Recordset

  Dim strSQL          As String
  Dim intAuxi         As Integer
  Dim dblICMS         As Double
  Dim nSequencia      As Long
  Dim blnTransaction  As Boolean
  Dim nRet            As Integer
  Dim nRow            As Long
  Dim bm              As Variant
  Dim bytLinha        As Byte
  Dim dblPercICM      As Double
  Dim dblPercIPI      As Double
  Dim dblQtdeAcertada As Double
  
  If Not IsNumeric(cboFornecedor.Text) Then
    MsgBox "Fornecedor inválido, processo cancelado.", vbExclamation, "Quick Store"
    Exit Sub
  End If


  On Error GoTo Err_Handlel

  '-------------------------------------
  'Abrir a transação
  '-------------------------------------
  ws.BeginTrans
  blnTransaction = True

      '*** Operações com o DB

      'Buscar uma próxima Sequência
      nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("Última Movimentação") + 1

      'Abrimos Entradas
      Set rstEntradas = db.OpenRecordset("Entradas", dbOpenDynaset)

      With rstEntradas
        .AddNew
        .Fields("Filial").Value = gnCodFilial
        .Fields("Data").Value = Data_Atual
        .Fields("Sequência").Value = nSequencia
        If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
          .Fields("Operação").Value = CInt(cboOperEntrada.Text)
        Else
          .Fields("Operação").Value = CInt(cboOperEntradaCompra.Text)
        End If
        .Fields("Digitador").Value = gnUserCode
        .Fields("Fornecedor").Value = CLng(cboFornecedor.Text)
        If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
          .Fields("Observações").Value = "Prestação de Contas criada em " & Now
        Else
          .Fields("Observações").Value = "Compra criada em " & Now
        End If
        .Fields("Nota Fiscal").Value = ""
        .Fields("Data Emissão").Value = Data_Atual
        .Fields("Pedido").Value = ""
        .Fields("Forma Pagto").Value = 0
        If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
          .Fields("Produtos").Value = CDbl(txtTotalPrestacao.Text)
        Else
          .Fields("Produtos").Value = CDbl(txtTotalCompras.Text)
        End If
        .Fields("Desconto").Value = 0
        .Fields("IPI").Value = 0
        .Fields("Frete").Value = 0
        If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
          .Fields("Base ICM").Value = CDbl(txtTotalPrestacao.Text)
        Else
          .Fields("Base ICM").Value = CDbl(txtTotalCompras.Text)
        End If
        'Buscar o valor do ICMS
        If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
          Call BuscarICMS((CDbl(txtTotalPrestacao.Text)), CLng(cboFornecedor.Text), dblICMS)
        Else
          Call BuscarICMS(CDbl(txtTotalCompras.Text), CLng(cboFornecedor.Text), dblICMS)
        End If
        .Fields("Valor ICM").Value = dblICMS
        .Fields("Base ICM Subs").Value = 0
        .Fields("Valor ICM Subs").Value = 0
        If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
          .Fields("Total").Value = CDbl(txtTotalPrestacao.Text)
        Else
          .Fields("Total").Value = CDbl(txtTotalCompras.Text)
        End If
        
        If IsNumeric(txtCxDinheiro.Text) Then
          .Fields("Dinheiro Caixa").Value = CDbl(txtCxDinheiro.Text)
        Else
          .Fields("Dinheiro Caixa").Value = 0
        End If
        
        If IsNumeric(txtCxCheque.Text) Then
          .Fields("Cheque Caixa").Value = CDbl(txtCxCheque.Text)
        Else
          .Fields("Cheque Caixa").Value = 0
        End If
        
        .Fields("Caixa").Value = CByte(cboCaixaUso.Text)
        
        If Len(lblConta.Caption) > 0 Then
          .Fields("Conta").Value = CByte(cboConta.Text)
        Else
          .Fields("Conta").Value = 0
        End If
        
        .Fields("Num Cheque").Value = txtCheque.Text
        If IsDate(medDataBomPara.Text) Then
          .Fields("Bom para").Value = CDate(medDataBomPara.Text)
        End If
        
        If Len(txtValCheque.Text) > 0 Then
          .Fields("Valor Cheque").Value = CDbl(txtValCheque.Text)
        Else
          .Fields("Valor Cheque").Value = 0
        End If
        .Fields("Descrição").Value = txtDescricao.Text
        .Fields("Efetivada").Value = False 'Em Efetiva Entrada ficará True
        .Fields("Nota Impressa").Value = 0
        .Fields("Nota Cancelada").Value = False
        '.Fields("Data Acerto Empréstimo").Value
        '.Fields("WebOrderFormID").Value
        .Fields("CentroCusto").Value = CInt(cboCodigoCC.Text)
        '.Fields("ConsignacaoMestre").Value
        '.Fields("obs_Obs1").Value
        '.Fields("obs_Obs2").Value
        '.Fields("obs_Obs3").Value
        '.Fields("obs_Obs4").Value
        '.Fields("obs_Obs5").Value
        '.Fields("obs_Obs6").Value
        '.Fields("obs_Obs7").Value
        '.Fields("obs_Obs8").Value
        '.Fields("obs_Transportadora").Value
        '.Fields("obs_Placa").Value
        '.Fields("obs_Uf").Value
        '.Fields("obs_Qtde").Value
        '.Fields("obs_Especie").Value
        '.Fields("obs_Marca").Value
        '.Fields("obs_PesoLiquido").Value
        '.Fields("obs_PesoBruto").Value
        '.Fields("obs_FretePago").Value
        .Fields("ConsignacaoFechada").Value = False
        .Update

        .Close
      End With

      Set rstEntradas = Nothing
      
      'Abrimos [Movimento - Parcelas]
      Set rsMovi_Parcelas = db.OpenRecordset("Movimento - Parcelas", dbOpenDynaset)

      gbConsignacaoResultado = True 'Var para distinção em Efetiva_Entrada
    
      grdCP.Update
      
      Call EraseTypeMoviment(tmMovimentoParcelas, gnCodFilial, nSequencia)
      
      For nRow = 0 To grdCP.Rows - 1
        bm = grdCP.AddItemBookmark(nRow)
        If IsDate(grdCP.Columns("Data").CellText(bm)) Then
          If IsNumeric(grdCP.Columns("Valor").CellValue(bm)) Then
            With rsMovi_Parcelas
              .AddNew
              .Fields("Filial") = gnCodFilial
              .Fields("Sequência") = nSequencia
              .Fields("Ordem") = nRow + 1
              .Fields("Bom") = grdCP.Columns("Data").CellText(bm)
              .Fields("Valor") = grdCP.Columns("Valor").CellValue(bm)
              .Update
            End With
          End If
        End If
      Next nRow
      
      rsMovi_Parcelas.Close
      Set rsMovi_Parcelas = Nothing

      'Abrimos [Entradas - Produtos]
      Set rstEntraProdu = db.OpenRecordset("Entradas - Produtos", dbOpenDynaset)

      grdGeral.MoveFirst

      For intAuxi = 0 To (grdGeral.Rows - 1)
      
        bytLinha = bytLinha + 1
      
        With rstEntraProdu
          .AddNew
          .Fields("Filial").Value = gnCodFilial
          .Fields("Sequência").Value = nSequencia
          .Fields("Linha").Value = bytLinha
          .Fields("Código").Value = Trim(grdGeral.Columns("Codigo").Value)
          If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
            Call BuscarQtdeAcertada(gnCodFilial, grdGeral.Columns("Sequencia").Value, grdGeral.Columns("Linha").Value, dblQtdeAcertada)
            '.Fields("Qtde").Value = CSng(grdGeral.Columns("Vendido").Value) old..
            .Fields("Qtde").Value = CSng(grdGeral.Columns("QtdeAcertada").Value)
          Else
            .Fields("Qtde").Value = CSng(grdGeral.Columns("Comprado").Value)
          End If
          
          .Fields("Preço").Value = CSng(grdGeral.Columns("Preco").Value)
          .Fields("Desconto").Value = 0
          Call BuscarICM(CLng(cboFornecedor.Text), Trim(grdGeral.Columns("Codigo").Value), dblPercICM)
          .Fields("ICM").Value = dblPercICM
          Call BuscarIPI(Trim(grdGeral.Columns("Codigo").Value), dblPercIPI)
          .Fields("IPI").Value = dblPercIPI
          
          If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
            .Fields("Preço Final").Value = Format(CSng(grdGeral.Columns("QtdeAcertada").Value) * CSng(grdGeral.Columns("Preco").Value), FORMAT_VALUE)
          Else
            .Fields("Preço Final").Value = Format((CSng(grdGeral.Columns("Comprado").Value) * CSng(grdGeral.Columns("Preco").Value)), FORMAT_VALUE)
          End If
          
          .Fields("Etiqueta").Value = False
          .Fields("Código sem Grade").Value = Trim(grdGeral.Columns("Codigo").Value)
          .Fields("InGeradoViaConsig").Value = False
          .Fields("ConsignacaoFechada").Value = False
          .Fields("IndiceFinanceiro").Value = 0
          If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
            .Fields("QtdeAtual").Value = CSng(grdGeral.Columns("Qtde").Value) - CSng(grdGeral.Columns("Vendido").Value)
          Else
            .Fields("QtdeAtual").Value = CSng(grdGeral.Columns("Comprado").Value)
          End If
          .Fields("Selecionado").Value = False
          .Fields("Acertado").Value = False
          .Fields("EntradaConsignada").Value = False
          .Update
        End With

      grdGeral.MoveNext
      Next intAuxi
      
      bytLinha = 0
      
      rstEntraProdu.Close
      Set rstEntraProdu = Nothing

      '----------------------------------------------------------
      'Chamada da função do Quick Efetiva_Entrada para Efetivação
      '----------------------------------------------------------
      Call StatusMsg("Efetivando entrada...")
        nRet = Efetiva_Entrada(gnCodFilial, nSequencia)
        If nRet <> 0 Then
          Select Case nRet
            Case -1 'Ação cancelada
              Call StatusMsg("Ação cancelada.")
            Case 1
              Call DisplayMsg("Código da operação inexistente.")
            Case 2
              Call DisplayMsg("Funcionário inexistente.")
            Case 3
              Call DisplayMsg("Fornecedor inexistente.")
            Case Else
              Call DisplayMsg("Operação NÃO efetivada. Erro" & str(nRet))
          End Select
          'Screen.MousePointer = vbDefault
          blnTransaction = True
          gbConsignacaoResultado = False 'Var para distinção em Efetiva_Entrada
          'ws.Rollback
          Call StatusMsg("")
          Exit Sub
        Else
          'Call ws.CommitTrans
          Call StatusMsg("Movimentação de Entrada realizada com sucesso.")
        End If
  

      'Abrimos Parâmetros
      Set rstParametros = db.OpenRecordset(" SELECT [Última Movimentação] FROM [Parâmetros Filial] WHERE Filial =" & gnCodFilial, dbOpenDynaset)

        rstParametros.Edit
        rstParametros.Fields("Última Movimentação").Value = nSequencia
        rstParametros.Update
        rstParametros.Close

      Set rstParametros = Nothing
      
      gbConsignacaoResultado = False 'Var para distinção em Efetiva_Entrada

  '-------------------------------------
  'Fechar a transação
  '-------------------------------------
  ws.CommitTrans
  blnTransaction = False

  Call StatusMsg("")

  Exit Sub

Err_Handlel:
  If blnTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Erro na transação"

End Sub

Private Sub BuscarICMS(ByVal Total As Double, ByVal Fornecedor As Long, ByRef ICMS)
  Dim rstEstado           As Recordset
  Dim rstCliFor           As Recordset
  Dim rstProdutos         As Recordset
  Dim strSQL              As String
  Dim blnPR               As Boolean
  Dim blnUFVazio          As Boolean
  Dim strUF               As String
  Dim intAuxi             As Integer
  Dim dblTotPercICMEntra  As Double

  strSQL = "SELECT Estado FROM Cli_For "
  strSQL = strSQL & " WHERE Código = " & Fornecedor

  Set rstCliFor = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstCliFor
    If Not (.BOF And .EOF) Then
      .MoveFirst
      If .Fields("Estado").Value = "PR" Then blnPR = True
      If .Fields("Estado").Value = "" Then blnUFVazio = True
      If .Fields("Estado").Value <> "PR" And .Fields("Estado").Value <> "" Then strUF = .Fields("Estado").Value
    End If
    .Close
  End With

  Set rstCliFor = Nothing

  If blnUFVazio Then ICMS = 0

  If blnPR Then 'Pegar para cada produto da grid o valor
    
    grdGeral.MoveFirst
  
    For intAuxi = 0 To (grdGeral.Rows - 1)
      
      Set rstProdutos = db.OpenRecordset("SELECT [Percentual Icm Entrada] FROM Produtos WHERE Código = '" & Trim(grdGeral.Columns("Codigo").Value) & "'", dbOpenDynaset)
      
      With rstProdutos
        If Not (.BOF And .EOF) Then
          .MoveFirst
          dblTotPercICMEntra = dblTotPercICMEntra + .Fields("Percentual Icm Entrada").Value
        End If
        .Close
      End With
      
      Set rstProdutos = Nothing

    grdGeral.MoveNext
    Next intAuxi

    ICMS = Format(((Total * dblTotPercICMEntra) / 100), FORMAT_VALUE)
    dblTotPercICMEntra = 0
    
  Else
    Set rstEstado = db.OpenRecordset("SELECT * FROM Estados WHERE Estado = '" & strUF & "'", dbOpenDynaset)

    With rstEstado
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        ICMS = Format(((Total * .Fields("ICM").Value) / 100), FORMAT_VALUE)
      End If
      .Close
    End With
    
    Set rstEstado = Nothing

  End If

End Sub

Private Sub BuscarICM(ByVal Fornecedor As Long, ByVal Produto As String, ByRef PercICM)
  Dim rstCliFor   As Recordset
  Dim rstProdutos As Recordset
  Dim rstEstados  As Recordset
  Dim strUF       As String
  
  Set rstCliFor = db.OpenRecordset("SELECT Estado FROM Cli_For WHERE Código = " & Fornecedor, dbOpenDynaset)
  
  With rstCliFor
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      strUF = .Fields("Estado").Value & ""
    End If
    .Close
  End With
  
  Set rstCliFor = Nothing
  
  If strUF = "" Then
    PercICM = 0
    Exit Sub
  End If
  
  If strUF = "PR" Then
  
    Set rstProdutos = db.OpenRecordset("SELECT [Percentual ICM] FROM Produtos WHERE Código = '" & Produto & "'", dbOpenDynaset)
  
    With rstProdutos
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        PercICM = .Fields("Percentual ICM").Value
        
      End If
      .Close
    End With
  
    Set rstProdutos = Nothing
    
  Else
  
    Set rstEstados = db.OpenRecordset("SELECT * FROM Estados WHERE Estado = '" & strUF & "'", dbOpenDynaset)
    
    With rstEstados
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        PercICM = .Fields("ICM").Value
        
      End If
      .Close
    End With
    
    Set rstEstados = Nothing
    
  End If
  
End Sub

Private Sub BuscarIPI(ByVal Produto As String, ByRef PercIPI)
  Dim rstProdutos As Recordset
  
  Set rstProdutos = db.OpenRecordset("SELECT [Percentual IPI] FROM Produtos WHERE Código = '" & Produto & "'", dbOpenDynaset)

  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      PercIPI = .Fields("Percentual IPI").Value
      
    End If
    .Close
  End With

  Set rstProdutos = Nothing
  
End Sub

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinal.Text = frmCalendario.gsDateCalender(mskDataFinal.Text)
  End If
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal.Text = Ajusta_Data(mskDataFinal.Text)
End Sub

Private Sub mskDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicial.Text = frmCalendario.gsDateCalender(mskDataInicial.Text)
  End If
End Sub

Private Sub mskDataInicial_LostFocus()
  mskDataInicial.Text = Ajusta_Data(mskDataInicial.Text)
End Sub

Private Sub AtualizarPrestacaoContas()
  Dim rstPrestacaoContas As Recordset
  Dim strSQL             As String
  Dim intAuxi            As Integer
  
  
  grdGeral.MoveFirst
  
  For intAuxi = 0 To (grdGeral.Rows - 1)
    
    strSQL = ""
    strSQL = "SELECT PrestacaoFechada, CompraFechada, QtdeVendida, QtdeAcertada, QtdeOriginal, Sequencia, Linha, Resultado FROM PrestacaoContas "
    strSQL = strSQL & " WHERE Filial = " & gnCodFilial
    strSQL = strSQL & " AND Sequencia = " & CLng(grdGeral.Columns("Sequencia").Value)
    strSQL = strSQL & " AND Linha = " & CByte(grdGeral.Columns("Linha").Value)
    strSQL = strSQL & " AND NOT PrestacaoFechada "
    
    Set rstPrestacaoContas = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rstPrestacaoContas
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do Until .EOF
          If frmManPrestacaodeContas.strCaption = "Prestação de Contas" Then
            
            If CDbl(grdGeral.Columns("Vendido").Value) <> 0 Then
              .Edit
              .Fields("PrestacaoFechada").Value = True
              'Tratamento do Campo Selecionado
              'Caso não feche integralmente abrimos o Selecionado para False
              'para uma segunda prestação quando for Resultado => (3 - Prestação de Contas)
              If .Fields("Resultado") = 3 Then
                If .Fields("QtdeOriginal").Value <> .Fields("QtdeVendida").Value Then
                  Call AtualizarEntraProd(gnCodFilial, .Fields("Sequencia").Value, .Fields("Linha").Value)
                End If
              
              End If
              '--------------------------------------------------------------
              '14/10/2004 - Refaz caso seja: 6 - Comprar e Prestar Contas
              'Fecharemos na Compra
              'If .Fields("Resultado") = 6 Then
              '  .Fields("PrestacaoFechada").Value = False
              'End If
              
              .Update
            End If
            
          
          Else 'Compras
        
            If CDbl(grdGeral.Columns("Comprado").Value) <> 0 Then
              .Edit
              .Fields("CompraFechada").Value = True
              .Update
            End If
            
          End If
        
        .MoveNext
        Loop
        
      End If
      .Close
    End With
    
    Set rstPrestacaoContas = Nothing
  
  grdGeral.MoveNext
  Next intAuxi

End Sub

Private Sub BuscarQtdeAcertada(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte, ByRef QtdeAcertada As Double)
  Dim rstPrestacaoContas As Recordset
  Dim strQuery           As String

  QtdeAcertada = 0

  strQuery = "SELECT QtdeAcertada FROM PrestacaoContas "
  strQuery = strQuery & " WHERE Filial = " & Filial
  strQuery = strQuery & " AND Sequencia = " & Sequencia
  strQuery = strQuery & " AND Linha = " & Linha

  Set rstPrestacaoContas = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstPrestacaoContas.RecordCount = 0 Then
    rstPrestacaoContas.Close
    Set rstPrestacaoContas = Nothing
    Exit Sub
  End If

  With rstPrestacaoContas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        QtdeAcertada = QtdeAcertada + .Fields("QtdeAcertada").Value
        
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstPrestacaoContas = Nothing

End Sub

Private Sub AtualizarEntraProd(ByVal bytFilial As Byte, ByVal lngSequencia As Long, ByVal bytLinha As Byte)
  Dim rstEntraProd As Recordset
  Dim strSQL       As String
  
  strSQL = "SELECT Selecionado FROM [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & bytFilial
  strSQL = strSQL & " AND Sequência = " & lngSequencia
  strSQL = strSQL & " AND Linha = " & bytLinha
  
  Set rstEntraProd = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntraProd
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      .Edit
      .Fields("Selecionado").Value = False
      .Update
    End If
    .Close
  End With
  
  Set rstEntraProd = Nothing

End Sub

