VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmMovCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Movimentação do Caixa"
   ClientHeight    =   8505
   ClientLeft      =   270
   ClientTop       =   795
   ClientWidth     =   14445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1250
   Icon            =   "frmMovimCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8505
   ScaleWidth      =   14445
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   9030
      ScaleHeight     =   855
      ScaleWidth      =   2235
      TabIndex        =   87
      Top             =   8340
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Frame Frame4 
      Caption         =   "Extrato detalhado dos movimentos"
      Height          =   795
      Left            =   90
      TabIndex        =   76
      Top             =   7620
      Width           =   14280
      Begin VB.CommandButton cmd_posicaoCrediario 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Posição Crediário"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12060
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   240
         Width           =   2085
      End
      Begin VB.CommandButton cmd_calendario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1890
         Picture         =   "frmMovimCaixa.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   247
         Width           =   465
      End
      Begin VB.CommandButton cmd_detalharMovimentoDia 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Detalhar Movimento Dia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9630
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   240
         Width           =   2205
      End
      Begin MSMask.MaskEdBox txt_data_extratoMov 
         Height          =   315
         Left            =   720
         TabIndex        =   79
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1140
         _ExtentX        =   2011
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
      Begin SSDataWidgets_B.SSDBCombo SSDBCaixa02 
         Bindings        =   "frmMovimCaixa.frx":4F23C
         DataSource      =   "Data4"
         Height          =   330
         Left            =   3420
         TabIndex        =   80
         Top             =   292
         Width           =   1080
         DataFieldList   =   "Descrição"
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
         BevelColorFrame =   12648447
         BevelColorShadow=   16777215
         BevelColorFace  =   12648447
         BackColorOdd    =   16777152
         Columns(0).Width=   3200
         _ExtentX        =   1905
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
      End
      Begin Crystal.CrystalReport Rel 
         Left            =   3480
         Top             =   3240
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
      Begin VB.Label Nome_caixa02 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4560
         TabIndex        =   83
         Top             =   285
         Width           =   4800
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Caixa"
         Height          =   195
         Left            =   2850
         TabIndex        =   82
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Data"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   81
         Top             =   345
         Width           =   405
      End
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   60000
      Left            =   14070
      Top             =   4560
   End
   Begin TabDlg.SSTab sstItem 
      Height          =   7005
      Left            =   90
      TabIndex        =   10
      Top             =   570
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   12356
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Depósitos e Retiradas no Caixa"
      TabPicture(0)   =   "frmMovimCaixa.frx":4F250
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Nome_Caixa"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Nome"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo_Caixa"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Data4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Data3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Senha"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "B_Novo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "B_Confirma"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmd_imprimirLancamento"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "&Lançamentos do Dia"
      TabPicture(1)   =   "frmMovimCaixa.frx":4F26C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sdbCaixa"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmd_imprimirLancamento 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprimir Lançamento da Tela"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   9630
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   6390
         Width           =   4500
      End
      Begin VB.CommandButton B_Confirma 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Confirmar Movimentação do Caixa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Confirma Movimentação do Caixa"
         Top             =   6390
         Width           =   4500
      End
      Begin VB.CommandButton B_Novo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Novo Lançamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Novo Lançamento"
         Top             =   6390
         Width           =   4500
      End
      Begin VB.Frame Frame3 
         Caption         =   "Totalizadores Gerais"
         Height          =   2025
         Left            =   -74190
         TabIndex        =   58
         Top             =   4725
         Width           =   12390
         Begin VB.CommandButton cmd_detalhaCartoes 
            BackColor       =   &H0080FFFF&
            Caption         =   "Detalhar Cartão"
            Height          =   435
            Left            =   6270
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   990
            Width           =   1395
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
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
            Index           =   0
            Left            =   1860
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   503
            Width           =   1395
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   3330
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   503
            Width           =   1395
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   503
            Width           =   1395
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
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
            Index           =   3
            Left            =   6270
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   503
            Width           =   1395
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
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
            Index           =   4
            Left            =   7740
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   503
            Width           =   1395
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
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
            Index           =   5
            Left            =   9210
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   503
            Width           =   1395
         End
         Begin VB.TextBox txtTotalGeral 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
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
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   510
            Width           =   1395
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Dinheiro"
            Height          =   195
            Index           =   0
            Left            =   2670
            TabIndex        =   72
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Cheque"
            Height          =   195
            Index           =   1
            Left            =   4170
            TabIndex        =   71
            Top             =   240
            Width           =   555
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Cheque Pré"
            Height          =   195
            Index           =   2
            Left            =   5340
            TabIndex        =   70
            Top             =   240
            Width           =   840
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Cartão"
            Height          =   195
            Index           =   3
            Left            =   7170
            TabIndex        =   69
            Top             =   240
            Width           =   555
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Vale"
            Height          =   195
            Index           =   4
            Left            =   8820
            TabIndex        =   68
            Top             =   240
            Width           =   300
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Parcelamento"
            Height          =   195
            Index           =   5
            Left            =   9600
            TabIndex        =   67
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Total Geral"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   11160
            TabIndex        =   66
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Registrar o pagamento no Contas a Pagar"
         Height          =   2265
         Left            =   150
         TabIndex        =   44
         Top             =   4050
         Width           =   13995
         Begin VB.Data Data6 
            Caption         =   "Custo"
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
            Left            =   13830
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT Nome, Código FROM [Centros de Custo] WHERE Ativo ORDER BY Nome"
            Top             =   1890
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Data Data5 
            Caption         =   "Forn"
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
            Left            =   13800
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Con_Fornecedor"
            Top             =   1470
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox Desc_Conta 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            MaxLength       =   30
            TabIndex        =   55
            Top             =   1740
            Width           =   7350
         End
         Begin VB.TextBox Nota 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   330
            Left            =   8310
            MaxLength       =   15
            TabIndex        =   54
            Top             =   1740
            Width           =   1995
         End
         Begin VB.CheckBox O_Gera_Conta 
            Appearance      =   0  'Flat
            Caption         =   "Quero registrar"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   1470
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
            Bindings        =   "frmMovimCaixa.frx":4F288
            DataSource      =   "Data5"
            Height          =   330
            Left            =   240
            TabIndex        =   48
            Top             =   1020
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
            BevelColorFrame =   12648447
            BevelColorShadow=   16777215
            BevelColorFace  =   12648447
            BackColorOdd    =   16777152
            Columns(0).Width=   3200
            _ExtentX        =   1879
            _ExtentY        =   582
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Centro 
            Bindings        =   "frmMovimCaixa.frx":4F29C
            DataSource      =   "Data6"
            Height          =   330
            Left            =   7200
            TabIndex        =   51
            Top             =   1020
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
            BevelColorFrame =   12648447
            BevelColorShadow=   16777215
            BevelColorFace  =   12648447
            BackColorOdd    =   16777152
            Columns(0).Width=   3200
            _ExtentX        =   1879
            _ExtentY        =   582
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
            Enabled         =   0   'False
         End
         Begin VB.Label L_Conta 
            Caption         =   "Descrição"
            Enabled         =   0   'False
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   57
            Top             =   1500
            Width           =   1170
         End
         Begin VB.Label L_Conta 
            Caption         =   "Nº do Documento"
            Enabled         =   0   'False
            Height          =   225
            Index           =   4
            Left            =   8310
            TabIndex        =   56
            Top             =   1500
            Width           =   1335
         End
         Begin VB.Label L_Conta 
            Caption         =   "Centro de Custo"
            Enabled         =   0   'False
            Height          =   225
            Index           =   2
            Left            =   7200
            TabIndex        =   53
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Nome_Centro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8310
            TabIndex        =   52
            Top             =   1020
            Width           =   5535
         End
         Begin VB.Label L_Conta 
            Caption         =   "Fornecedor"
            Enabled         =   0   'False
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   50
            Top             =   780
            Width           =   1170
         End
         Begin VB.Label Nome_Fornecedor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1350
            TabIndex        =   49
            Top             =   1020
            Width           =   5535
         End
         Begin VB.Label L_Conta 
            AutoSize        =   -1  'True
            Caption         =   "Valor da Conta"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   2010
            TabIndex        =   47
            Top             =   375
            Width           =   1065
         End
         Begin VB.Label Valor_Conta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "0,00"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3150
            TabIndex        =   46
            Top             =   315
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFA324&
         BorderStyle     =   0  'None
         Caption         =   "Saldo Atual"
         Height          =   2805
         Left            =   150
         TabIndex        =   20
         Top             =   1170
         Width           =   8925
         Begin VB.TextBox Desc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   345
            Left            =   180
            MaxLength       =   60
            TabIndex        =   42
            Top             =   2310
            Width           =   8565
         End
         Begin VB.CommandButton cmdTransferValue 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   0
            Left            =   750
            Picture         =   "frmMovimCaixa.frx":4F2B0
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   840
            Width           =   405
         End
         Begin VB.CommandButton cmdTransferValue 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   1
            Left            =   6015
            Picture         =   "frmMovimCaixa.frx":5BBD2
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   840
            Width           =   405
         End
         Begin VB.CommandButton cmdTransferValue 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   2
            Left            =   7770
            Picture         =   "frmMovimCaixa.frx":684F4
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   840
            Width           =   405
         End
         Begin VB.CommandButton cmdTransferValue 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   3
            Left            =   2505
            Picture         =   "frmMovimCaixa.frx":74E16
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   840
            Width           =   405
         End
         Begin VB.CommandButton cmdTransferValue 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   4
            Left            =   4260
            Picture         =   "frmMovimCaixa.frx":81738
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   840
            Width           =   405
         End
         Begin MSMask.MaskEdBox Campo 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   31
            Top             =   1590
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Campo 
            Height          =   315
            Index           =   1
            Left            =   5415
            TabIndex        =   32
            Top             =   1590
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Campo 
            Height          =   315
            Index           =   2
            Left            =   7170
            TabIndex        =   33
            Top             =   1590
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Campo 
            Height          =   315
            Index           =   3
            Left            =   1920
            TabIndex        =   34
            Top             =   1590
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Campo 
            Height          =   315
            Index           =   4
            Left            =   3675
            TabIndex        =   35
            Top             =   1590
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFA324&
            Caption         =   "Descrição movimentação do caixa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   180
            TabIndex        =   43
            Top             =   2010
            Width           =   3480
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFA324&
            Caption         =   "Para pagamentos ou retiradas use o sinal 'de menos' -"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   180
            TabIndex        =   36
            Top             =   1290
            Width           =   5505
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFA324&
            Caption         =   "Dinheiro"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   600
            TabIndex        =   30
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFA324&
            Caption         =   "Cheques"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   5820
            TabIndex        =   29
            Top             =   120
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFA324&
            Caption         =   "Cheques Pré"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   7380
            TabIndex        =   28
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFA324&
            Caption         =   "Cartões"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2370
            TabIndex        =   27
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFA324&
            Caption         =   "Vales"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   4200
            TabIndex        =   26
            Top             =   120
            Width           =   525
         End
         Begin VB.Label lblTotalDia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   25
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label lblTotalDia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   5415
            TabIndex        =   24
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label lblTotalDia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   7170
            TabIndex        =   23
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label lblTotalDia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   1920
            TabIndex        =   22
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label lblTotalDia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   3675
            TabIndex        =   21
            Top             =   390
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Dinheiro retirado vai para"
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   9120
         TabIndex        =   19
         Top             =   1170
         Width           =   4950
         Begin VB.Data Data1 
            Caption         =   "C/C"
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
            Left            =   4770
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Con_Conta"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.OptionButton Din_Nada 
            Appearance      =   0  'Flat
            Caption         =   "Não determinado"
            ForeColor       =   &H80000008&
            Height          =   215
            Left            =   120
            TabIndex        =   5
            Top             =   1020
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Din_Conta 
            Appearance      =   0  'Flat
            Caption         =   "Conta bancária"
            ForeColor       =   &H80000008&
            Height          =   215
            Left            =   120
            TabIndex        =   3
            Top             =   300
            Width           =   1485
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Din 
            Bindings        =   "frmMovimCaixa.frx":8E05A
            DataSource      =   "Data1"
            Height          =   345
            Left            =   480
            TabIndex        =   4
            Top             =   540
            Width           =   1200
            DataFieldList   =   "Descrição"
            _Version        =   196617
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "WeblySleek UI Semilight"
               Size            =   9.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelColorFrame =   12648447
            BevelColorShadow=   16777215
            BevelColorFace  =   12648447
            BackColorOdd    =   16777152
            RowHeight       =   529
            Columns.Count   =   3
            Columns(0).Width=   5768
            Columns(0).Caption=   "Descrição"
            Columns(0).Name =   "Descrição"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Descrição"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3651
            Columns(1).Caption=   "Conta"
            Columns(1).Name =   "Conta"
            Columns(1).CaptionAlignment=   0
            Columns(1).DataField=   "Conta"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   1746
            Columns(2).Caption=   "Código"
            Columns(2).Name =   "Código"
            Columns(2).Alignment=   1
            Columns(2).CaptionAlignment=   1
            Columns(2).DataField=   "Código"
            Columns(2).DataType=   2
            Columns(2).FieldLen=   256
            _ExtentX        =   2117
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label Nome_Din 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1740
            TabIndex        =   18
            Top             =   540
            Width           =   3105
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cheques retirados vão para"
         Height          =   1335
         Left            =   9120
         TabIndex        =   17
         Top             =   2640
         Width           =   4950
         Begin VB.OptionButton Che_Conta 
            Appearance      =   0  'Flat
            Caption         =   "Conta bancária"
            ForeColor       =   &H80000008&
            Height          =   215
            Left            =   150
            TabIndex        =   6
            Top             =   270
            Width           =   1455
         End
         Begin VB.OptionButton Che_Nada 
            Appearance      =   0  'Flat
            Caption         =   "Não determinado"
            ForeColor       =   &H80000008&
            Height          =   215
            Left            =   150
            TabIndex        =   8
            Top             =   990
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Data Data2 
            Caption         =   "C/C"
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
            Left            =   4740
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Con_Conta"
            Top             =   960
            Visible         =   0   'False
            Width           =   1785
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Che 
            Bindings        =   "frmMovimCaixa.frx":8E06E
            DataSource      =   "Data2"
            Height          =   345
            Left            =   510
            TabIndex        =   7
            Top             =   510
            Width           =   1200
            DataFieldList   =   "Descrição"
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
            BevelColorFrame =   12648447
            BevelColorShadow=   16777215
            BevelColorFace  =   12648447
            BackColorOdd    =   16777152
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   5556
            Columns(0).Caption=   "Descrição"
            Columns(0).Name =   "Descrição"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Descrição"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3334
            Columns(1).Caption=   "Conta"
            Columns(1).Name =   "Conta"
            Columns(1).CaptionAlignment=   0
            Columns(1).DataField=   "Conta"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   1429
            Columns(2).Caption=   "Código"
            Columns(2).Name =   "Código"
            Columns(2).Alignment=   1
            Columns(2).CaptionAlignment=   1
            Columns(2).DataField=   "Código"
            Columns(2).DataType=   2
            Columns(2).FieldLen=   256
            _ExtentX        =   2117
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label Nome_Che 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1770
            TabIndex        =   16
            Top             =   510
            Width           =   3105
         End
      End
      Begin VB.TextBox Senha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   12510
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   1560
      End
      Begin VB.Data Data3 
         Caption         =   "Funcionario"
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
         Left            =   14040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Nome, Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE AND isPrestServ = false ORDER BY Nome"
         Top             =   3660
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data Data4 
         Caption         =   "Caixa"
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
         Left            =   13950
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Caixas"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1905
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Caixa 
         Bindings        =   "frmMovimCaixa.frx":8E082
         DataSource      =   "Data4"
         Height          =   330
         Left            =   150
         TabIndex        =   0
         Top             =   600
         Width           =   870
         DataFieldList   =   "Descrição"
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
         BevelColorFrame =   12648447
         BevelColorShadow=   16777215
         BevelColorFace  =   12648447
         BackColorOdd    =   16777152
         RowHeight       =   476
         Columns(0).Width=   3200
         _ExtentX        =   1535
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBCombo Combo 
         Bindings        =   "frmMovimCaixa.frx":8E096
         DataSource      =   "Data3"
         Height          =   330
         Left            =   6075
         TabIndex        =   1
         Top             =   600
         Width           =   990
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
         BevelColorFrame =   12648447
         BevelColorShadow=   16777215
         BevelColorFace  =   12648447
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   5001
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3466
         Columns(1).Caption=   "Apelido"
         Columns(1).Name =   "Apelido"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Apelido"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1773
         Columns(2).Caption=   "Código"
         Columns(2).Name =   "Código"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Código"
         Columns(2).DataType=   3
         Columns(2).FieldLen=   256
         _ExtentX        =   1746
         _ExtentY        =   582
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBGrid sdbCaixa 
         Height          =   4110
         Left            =   -74190
         TabIndex        =   73
         Top             =   510
         Width           =   12390
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   8
         stylesets.count =   1
         stylesets(0).Name=   "Total"
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "frmMovimCaixa.frx":8E0AA
         AllowUpdate     =   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   238
         Columns.Count   =   8
         Columns(0).Width=   2699
         Columns(0).Caption=   "Caixas"
         Columns(0).Name =   "Caixas"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2566
         Columns(1).Caption=   "Dinheiro"
         Columns(1).Name =   "Dinheiro"
         Columns(1).Alignment=   1
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2566
         Columns(2).Caption=   "Cheque"
         Columns(2).Name =   "Cheque"
         Columns(2).Alignment=   1
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2566
         Columns(3).Caption=   "Cheque Pré"
         Columns(3).Name =   "ChequePre"
         Columns(3).Alignment=   1
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   2566
         Columns(4).Caption=   "Cartão"
         Columns(4).Name =   "Cartao"
         Columns(4).Alignment=   1
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   2566
         Columns(5).Caption=   "Vale"
         Columns(5).Name =   "Vale"
         Columns(5).Alignment=   1
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   2566
         Columns(6).Caption=   "Parcelamento"
         Columns(6).Name =   "Parcelamento"
         Columns(6).Alignment=   1
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   2593
         Columns(7).Caption=   "Total"
         Columns(7).Name =   "Total"
         Columns(7).Alignment=   1
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(7).HeadStyleSet=   "Total"
         Columns(7).StyleSet=   "Total"
         _ExtentX        =   21855
         _ExtentY        =   7250
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Movimentação de caixa"
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
         Left            =   150
         TabIndex        =   75
         Top             =   930
         Width           =   2265
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Autorizado por"
         Height          =   195
         Left            =   6075
         TabIndex        =   15
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7110
         TabIndex        =   14
         Top             =   600
         Width           =   5250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Senha "
         Height          =   195
         Left            =   12510
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Caixa"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Nome_Caixa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1065
         TabIndex        =   11
         Top             =   600
         Width           =   4890
      End
   End
   Begin VB.CommandButton cmdAtualizarTotalDia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Atualizar Informações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Atualiza as informações de Totais a serem exibidas"
      Top             =   60
      Width           =   14280
   End
End
Attribute VB_Name = "frmMovCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'05/05/2005 - Daniel
'
'Projeto: Melhorias para o Centro de Custo
'
'A partir da versão 6.52.0.38 todo campo de Centro de Custo
'estará carregando apenas os Centros que estão ativos no sistema

Dim rsCaixa As Recordset
Dim rsFuncionarios As Recordset
Dim rsContas As Recordset
Dim rsLançamentos As Recordset
Dim rsCaixas_Uso As Recordset
Dim rsCentros As Recordset
Dim rsFornecedores As Recordset
Dim rsContas_Pagar As Recordset

Private Sub Recalcula_Conta()
  Dim nValor As Double
  Dim nX As Integer
  
  For nX = 0 To 4
    If IsNull(Campo(nX).Text) Then Campo(nX).Text = 0
    If Campo(nX).Text = "" Then Campo(nX).Text = 0
    If IsNumeric(Campo(nX).Text) Then
      If CCur(Campo(nX).Text) < 0 Then
        nValor = nValor + CCur(Campo(nX).Text)
      End If
    End If
  Next nX
  
  Valor_Conta.Caption = Format(Abs(nValor), FORMAT_VALUE)

End Sub

Private Sub B_Confirma_Click()
  Dim Tot_Dinheiro As Double
  Dim Tot_Cheques As Double
  Dim Tot_Pré As Double
  Dim Tot_Cartões As Double
  Dim Tot_Vales As Double
  Dim Saldo_Ant As Double
  Dim Ordem As Long
  Dim Erro As Integer
  Dim Criar_Zero As Integer
  Dim nX As Integer
  Dim bHasValue As Boolean
  
  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked As Integer
  
  On Error GoTo Trata_Erro:
  
  Call StatusMsg("")
  
  If IsNull(Desc.Text) Then
    Desc.Text = " "
  ElseIf Desc.Text = "" Then
    Desc.Text = " "
  End If
  If IsNull(Valor_Conta.Caption) Then
    Valor_Conta.Caption = Format(0, FORMAT_VALUE)
  ElseIf Valor_Conta.Caption = "" Then
    Valor_Conta.Caption = Format(0, FORMAT_VALUE)
  End If
  
  If Nome_Caixa.Caption = "" Then
    DisplayMsg "Digite o caixa."
    sstItem.Tab = 0
    Combo_Caixa.SetFocus
    Exit Sub
  ElseIf Nome.Caption = "" Then
    DisplayMsg "Digite o código do funcionário."
    sstItem.Tab = 0
    Combo.SetFocus
    Exit Sub
  ElseIf IsNull(Senha.Text) Then
    DisplayMsg "Senha inválida."
    sstItem.Tab = 0
    Senha.SetFocus
    Exit Sub
  End If
  
  'Verifica funcionário
  With rsFuncionarios
    .Index = "Código"
    .Seek "=", Val(Combo.Text)
    If .NoMatch Then
      DisplayMsg "Funcionário inválido."
      sstItem.Tab = 0
      Combo.SetFocus
      Exit Sub
    ElseIf !ValorP <> CriptografaSenha(Senha.Text) Then
      DisplayMsg "Senha incorreta."
      sstItem.Tab = 0
      Senha.SetFocus
      Exit Sub
    ElseIf ![Movimentar Caixa] = False Then
      DisplayMsg "Este funcionário não pode movimentar o caixa."
      sstItem.Tab = 0
      Combo.SetFocus
      Exit Sub
    End If
  End With
 
  For nX = 0 To 4
    If IsNumeric(Campo(nX).Text) Then
      If CDbl(Campo(nX).Text) <> 0 Then
        bHasValue = True
        Exit For
      End If
    End If
  Next nX
  
  If Not bHasValue Then
    DisplayMsg "Nenhum valor válido, caixa inalterado."
    sstItem.Tab = 0
    Exit Sub
  End If
  
  For nX = 0 To 4  'arruma campos
    If Not IsNumeric(Campo(nX).Text) Then
      Campo(nX).Text = 0
    End If
    Campo(nX).Text = Format(CDbl(Campo(nX).Text), FORMAT_VALUE)
  Next nX
 
  If Din_Conta.Value = True Then
    If CDbl(Campo(0).Text) > 0 Then
      DisplayMsg "Só é possivel passar dinheiro para a conta nas retiradas."
      sstItem.Tab = 0
      Din_Conta.SetFocus
      Exit Sub
    ElseIf Nome_Din.Caption = "" Then
      DisplayMsg "Digite o código da conta ou selecione <Não determinado>."
      sstItem.Tab = 0
      Combo_Din.SetFocus
      Exit Sub
    End If
  End If
  
  If Che_Conta.Value = True Then
    If CDbl(Campo(1).Text) > 0 Then
      DisplayMsg "Só é possivel passar cheques para a conta nas retiradas."
      sstItem.Tab = 0
      Che_Conta.SetFocus
      Exit Sub
    ElseIf Nome_Che.Caption = "" Then
      DisplayMsg "Digite o código da conta ou selecione <Não determinado>."
      sstItem.Tab = 0
      Combo_Che.SetFocus
      Exit Sub
    End If
  End If
  
  If O_Gera_Conta.Value = 1 Then
   If CDbl(Valor_Conta.Caption) = 0 Then
     DisplayMsg "Não é possível gerar uma conta paga sem retiradas. Favor verificar."
     sstItem.Tab = 0
     Exit Sub
   ElseIf Nome_Fornecedor.Caption = "" Then
     DisplayMsg "Por favor encontre o fornecedor."
     sstItem.Tab = 1
     Combo_Fornecedor.SetFocus
     Exit Sub
   End If
  End If
  
  Dim tTotais As tpPaymentType
  Dim nOrdem As Integer
  Dim dblSaldoAnterior As Double
  
  'Verifica o início do caixa, abertura do dia e retorna os últimos valores
  If Not gbCheckOpenCaixa(Val(Combo_Caixa.Text), Val(Combo.Text), dblSaldoAnterior, nOrdem, tTotais) Then
    'Ocorreu erro e a mensagem é exibida pela função
    Exit Sub
  End If
  
  With rsCaixa
    ws.BeginTrans
    blnInTransaction = True
    
    nOrdem = nOrdem + 1
    .AddNew
    !Filial = gnCodFilial
    !Data = Data_Atual
    !Caixa = Val(Combo_Caixa.Text)
    !Funcionário = Val(Combo.Text)
    !Hora = Format(CStr(Time), "hh:mm:ss")
    !Ordem = nOrdem
    !Descrição = Desc.Text
    !Dinheiro = CDbl(Campo(0).Text)
    ![Total Dinheiro] = tTotais.dblDinheiro + CDbl(Campo(0).Text)
    !Cheques = CDbl(Campo(1).Text)
    ![Total Cheques] = tTotais.dblCheque + CDbl(Campo(1).Text)
    ![Cheques Pré] = CDbl(Campo(2).Text)
    ![Total Cheques Pré] = tTotais.dblChequePre + CDbl(Campo(2).Text)
    !Cartões = CDbl(Campo(3).Text)
    ![Total Cartões] = tTotais.dblCartao + CDbl(Campo(3).Text)
    !Vales = CDbl(Campo(4).Text)
    ![Total Vales] = tTotais.dblVale + CDbl(Campo(4).Text)
    ![Saldo Anterior] = dblSaldoAnterior
    !Final = Format(![Saldo Anterior] + !Dinheiro + !Cheques + _
      ![Cheques Pré] + !Cartões + !Vales, FORMAT_VALUE)
    .Update
  End With

  'Atualiza Dinheiro na conta, se necessário
  If Din_Conta.Value Then
    With rsLançamentos
      .Index = "Conta"
      .Seek "<", Val(Combo_Din.Text), CDate(Data_Atual), 99999999#
      If .NoMatch Then
        Saldo_Ant = 0
      End If
      If Not .NoMatch Then
        Saldo_Ant = 0
        If !Conta = Val(Combo_Din.Text) Then
          Saldo_Ant = ![Saldo Atual]
        End If
      End If
      .AddNew
      !Conta = Val(Combo_Din.Text)
      !Data = Data_Atual
      !Descrição = "Depósito de dinheiro do caixa."
      ![Saldo Anterior] = Saldo_Ant
      !Crédito = -(CDbl(Campo(0).Text))
      ![Saldo Atual] = Saldo_Ant + Abs(CDbl(Campo(0).Text))
      .Update
    End With
  End If
     
  'Atualiza Cheques na conta, se necessário
  If Che_Conta.Value Then
    With rsLançamentos
      .Index = "Conta"
      .Seek "<", Val(Combo_Che.Text), Data_Atual, 999999999#
      If .NoMatch Then
        Saldo_Ant = 0
      End If
      If Not .NoMatch Then
        Saldo_Ant = 0
        If !Conta = Val(Combo_Che.Text) Then
          Saldo_Ant = ![Saldo Atual]
        End If
      End If
      .AddNew
      !Conta = Val(Combo_Che.Text)
      !Data = Data_Atual
      !Descrição = "Depósito de cheque(s) do caixa."
      ![Saldo Anterior] = Saldo_Ant
      !Crédito = -(CDbl(Campo(1).Text))
      ![Saldo Atual] = Saldo_Ant + Abs(CDbl(Campo(1).Text))
      .Update
    End With
  End If
  
  'Gera conta paga
  If O_Gera_Conta.Value = vbChecked Then
    With rsContas_Pagar
      .AddNew
      !Filial = gnCodFilial
      !Fornecedor = Val(Combo_Fornecedor.Text)
      ![Data Emissão] = Data_Atual
      !Descrição = Desc_Conta.Text
      !Vencimento = Data_Atual
      !Valor = Abs(CDbl(Valor_Conta.Caption))
      !Desconto = 0
      !Acréscimo = 0
      ![Valor Pago] = Abs(CDbl(Valor_Conta.Caption))
      !Pagamento = Data_Atual
      !Sequência = 0
      !Nota = Nota.Text
      If Nome_Centro.Caption <> "" Then
        ![Centro de Custo] = Val(Combo_Centro.Text)
      End If
      ![Data Alteração] = Format(Date, "dd/mm/yyyy")
      .Update
    End With
  End If
  ws.CommitTrans
  blnInTransaction = False
  
  sstItem.Tab = 0
  DisplayMsg "Caixa Atualizado."
  Call tmrRefresh_Timer
  B_Confirma.Enabled = False
  
  Exit Sub
Trata_Erro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Resume
        Else
          'Cancelamento da transação
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
      End If
    Case Else
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Manutenção - Contas a receber")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Exit Sub
        Case 3 'Encerrar
          End
      End Select
  End Select
End Sub

Private Sub B_Novo_Click()
  Dim nX As Integer
  
  For nX = 0 To 4
    Campo(nX).Text = ""
    Call Campo_LostFocus(nX)
  Next nX
  
  Combo.Text = ""
  Nome.Caption = ""
  Senha.Text = ""
  Desc.Text = ""
  B_Confirma.Enabled = True
  
  O_Gera_Conta.Value = vbUnchecked
  
  Valor_Conta.Caption = ""
  Combo_Fornecedor.Text = ""
  Combo_Fornecedor_LostFocus
  
  Combo_Centro.Text = ""
  Combo_Centro_LostFocus
  
  Desc_Conta.Text = ""
  Nota.Text = ""
  
  sstItem.Tab = 0
  Combo.SetFocus
  Call StatusMsg("")
  
End Sub

Private Sub Campo_Change(Index As Integer)
  Call Recalcula_Conta
End Sub

Private Sub Campo_GotFocus(Index As Integer)
  Call SelectAllText(Campo(Index))
End Sub

Private Sub Campo_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub Campo_LostFocus(Index As Integer)
  Call FormataValorCor(Campo(Index), False)
End Sub

Private Sub Che_Conta_Click()
  If Che_Conta.Value = True Then
    Combo_Che.Enabled = True
    Nome_Che.Enabled = True
  End If
  If Din_Conta.Value = False Then
    Combo_Din.Enabled = False
    Nome_Din.Enabled = False
  End If
End Sub

Private Sub Che_Nada_Click()
  Combo_Che.Enabled = False
End Sub

Private Sub cmd_calendario_Click()
    txt_data_extratoMov.Text = frmCalendario.gsDateCalender(txt_data_extratoMov.Text)
End Sub

Private Sub cmd_detalhaCartoes_Click()
    Dim objLancCartaoPosiDiaria As frmRelLancCartaoPosiDiaria
    
    Set objLancCartaoPosiDiaria = New frmRelLancCartaoPosiDiaria
    
    objLancCartaoPosiDiaria.paramCodFilial = gnCodFilial
    objLancCartaoPosiDiaria.Show
    
    Set objLancCartaoPosiDiaria = Nothing
End Sub

Private Sub cmd_detalharMovimentoDia_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 
 Call StatusMsg("")

 Erro = False
 If IsNull(txt_data_extratoMov.Text) Then Erro = True
 If Not Erro Then If Not IsDate(txt_data_extratoMov.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   txt_data_extratoMov.SetFocus
   Exit Sub
 End If
 
 If SSDBCaixa02.Text = "" Then
   DisplayMsg "Selecione um caixa."
   SSDBCaixa02.SetFocus
   Exit Sub
 End If
 
 'Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 'Saída
 Rel.Destination = 0

 'Nome do arquivo .rpt
 Str1 = gsReportPath & "CAIXA.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd1 Rel

 'Seleção
 Str_Data1 = "Date" + Format$(txt_data_extratoMov.Text, "(yyyy,mm,dd)")

 Str_Rel = "{Caixa.Filial} =" + CStr(gnCodFilial)
 Str_Rel = Str_Rel + " And {Caixa.Data} ="
 Str_Rel = Str_Rel + Str_Data1
 
 Rel.SelectionFormula = Str_Rel
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel.Formulas(0) = Str_Rel

 Str_Rel = "dia = '"
 Str_Rel = Str_Rel + txt_data_extratoMov.Text + "'"
 Rel.Formulas(1) = Str_Rel

 If Nome_caixa02.Caption <> "" Then
   Str_Rel = " And {Caixa.Caixa} =" + SSDBCaixa02.Text
   Rel.SelectionFormula = Rel.SelectionFormula + Str_Rel
 End If

 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
  
 'Seta a impressora para relatório
 Call SetPrinterName("REL", Rel)

 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault
End Sub

Private Sub cmd_imprimirLancamento_Click()
  On Error GoTo Erro
  
  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome As String
  Dim strNomeLPT As String
  Dim strPortaLPT As String
  Dim intX As Integer
  Dim i As Integer
  
  strNome = "TICKET"
  strNomeLPT = "NOME IMPRESSORA TICKET"
  strPortaLPT = "PORTA IMPRESSORA TICKET"

  strImpressora = GetSetting("QuickStore", "ConfigLPT", strNomeLPT, "")
  strPorta = GetSetting("QuickStore", "ConfigLPT", strPortaLPT, "")
      
  If Len(Trim(strImpressora)) > 0 And Len(Trim(strPorta)) > 0 Then
      For Each objPrinter In Printers
        If objPrinter.DeviceName = strImpressora And objPrinter.Port = strPorta Then
            Set Printer = objPrinter
            Exit For
        End If
      Next objPrinter
  End If

  Dim sCodigoProduto As String
  Dim sNomeProduto As String
  Dim sCodigoEntrada As String
  Dim sNumItens As String
  Dim sValorUnitario As String
  Dim sValorTotal As String
  Dim sLinha As String
  Dim lContador As Long
  Dim sDataAtual As String
  
  sDataAtual = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  
  Printer.Font = "LUCIDA CONSOLE"

  ' Modelo Ticket...42 colunas
  Dim sStrAux As String
      
  Printer.Print "__________________________________________"
  If Picture1.Picture <> 0 Then
      Printer.PaintPicture Picture1, 1200, 500, 2300, 1000
      Printer.Print ""
      Printer.Print ""
      Printer.Print ""
      Printer.Print ""
      Printer.Print ""
      Printer.Print ""
      Printer.Print ""
      Printer.Print ""
      Printer.Print ""
  End If
      
  Printer.Print ""
  Printer.Print "    LANÇAMENTO MANUAL DE CAIXA"
  Printer.Print ""
      
  sLinha = "Emissão   : " & sDataAtual
  Printer.Print sLinha

  sLinha = "N° Filial : " & gnCodFilial
  Printer.Print sLinha

  If Len(gsNomeFilial) > 30 Then
      sLinha = "Nome      : " & Mid(gsNomeFilial, 1, 30)
      Printer.Print sLinha
  Else
      sLinha = "Nome      : " & gsNomeFilial
      Printer.Print sLinha
  End If

  sLinha = "N° Caixa  : " & Combo_Caixa.Text
  Printer.Print sLinha
  
  If Len(Nome_Caixa.Caption) > 30 Then
      sLinha = "Nome      : " & Mid(Nome_Caixa.Caption, 1, 30)
      Printer.Print sLinha
  Else
      sLinha = "Nome      : " & Nome_Caixa.Caption
      Printer.Print sLinha
  End If

  sLinha = "Autorizado: " & Combo.Text
  Printer.Print sLinha
  
  If Len(Nome.Caption) > 30 Then
      sLinha = "Nome      : " & Mid(Nome.Caption, 1, 30)
      Printer.Print sLinha
  Else
      sLinha = "Nome      : " & Nome.Caption
      Printer.Print sLinha
  End If

  sLinha = "Dinheiro  : " & Campo(0).Text
  Printer.Print sLinha

  sLinha = "Cartões   : " & Campo(3).Text
  Printer.Print sLinha

  sLinha = "Vale      : " & Campo(4).Text
  Printer.Print sLinha

  sLinha = "Cheques   : " & Campo(1).Text
  Printer.Print sLinha

  sLinha = "ChequesPré: " & Campo(2).Text
  Printer.Print sLinha

  sLinha = "Descrição Movimentação "
  Printer.Print sLinha

  If Len(Desc.Text) > 30 Then
      sLinha = Mid(Desc.Text, 1, 30)
      Printer.Print sLinha
      Printer.Print Mid(Desc.Text, 30, Len(Desc.Text) - 30)
  Else
      sLinha = Desc.Text
      Printer.Print sLinha
  End If


  sLinha = "Registrar Fornecedor "
  Printer.Print sLinha

  sLinha = "Fornecedor  : " & Combo_Fornecedor.Text
  Printer.Print sLinha

  If Len(Nome_Fornecedor.Caption) > 30 Then
      sLinha = Mid(Nome_Fornecedor.Caption, 1, 30)
      Printer.Print sLinha
  Else
      sLinha = Nome_Fornecedor.Caption
      Printer.Print sLinha
  End If

  sLinha = "CentroCusto : " & Combo_Centro.Text
  Printer.Print sLinha

  If Len(Nome_Centro.Caption) > 30 Then
      sLinha = Mid(Nome_Centro.Caption, 1, 30)
      Printer.Print sLinha
  Else
      sLinha = Nome_Centro.Caption
      Printer.Print sLinha
  End If

  If Len(Desc_Conta.Text) > 30 Then
      sLinha = Mid(Desc_Conta.Text, 1, 30)
      Printer.Print sLinha
      Printer.Print Mid(Desc_Conta.Text, 30, Len(Desc_Conta.Text) - 30)
  Else
      sLinha = Desc_Conta.Text
      Printer.Print sLinha
  End If

  sLinha = "Nº Documento: " & Nota.Text
  Printer.Print sLinha

  Printer.Print ""
  Printer.Print "Assinatura do Atendente e carimbo da loja"
  Printer.Print ""
  Printer.Print ""
  Printer.Print ""
  Printer.Print "_____________________________________"
  Printer.Print gsUserName
  Printer.Print ""
  Printer.Print " - - - - - - - - - - - - - - - - - -"
  Printer.Print "|                                   |"
  Printer.Print "|                                   |"
  Printer.Print "|                                   |"
  Printer.Print "|                                   |"
  Printer.Print " - - - - - - - - - - - - - - - - - -"
  Printer.Print ""
  Printer.Print ""
  Printer.Print ""
  Printer.Print "_____________________________________"
  Printer.Print "Outro (se necessário)"
    
  Printer.EndDoc
  
  Exit Sub
Erro:
    MsgBox "Erro na impressão do Vale " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_posicaoCrediario_Click()
On Error GoTo Erro

    If IsNull(txt_data_extratoMov.Text) Then
        DisplayMsg "Data incorreta, verifique."
        txt_data_extratoMov.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txt_data_extratoMov.Text) Then
        DisplayMsg "Data incorreta, verifique."
        txt_data_extratoMov.SetFocus
        Exit Sub
    End If
    
'    If SSDBCaixa02.Text = "" Then
'      DisplayMsg "Selecione um caixa."
'      SSDBCaixa02.SetFocus
'      Exit Sub
'    End If

    Dim objCrediario_PosiDiaria As frmMovCaixa_PosicaoCrediario
    Set objCrediario_PosiDiaria = New frmMovCaixa_PosicaoCrediario
    
    objCrediario_PosiDiaria.DataPosicao = txt_data_extratoMov.Text
'    objCrediario_PosiDiaria.CaixaPosicao = SSDBCaixa02.Text
'    objCrediario_PosiDiaria.NomeCaixaPosicao = Nome_caixa02.Caption
    objCrediario_PosiDiaria.Show
    
    Set objCrediario_PosiDiaria = Nothing

    Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cmdAtualizarTotalDia_Click(Index As Integer)
  Call RefreshValues(False)
  Call RefreshCaixasValues
End Sub


'''Private Sub cmdFecharCaixa_Click()
'''  If Nome_Caixa.Caption = "" Then
'''    DisplayMsg "Escolha o caixa."
'''    Combo_Caixa.SetFocus
'''  Else
'''    Call B_Novo_Click
'''    Call RefreshValues(True)
'''    Desc.Text = "Fechamento de Caixa"
'''    Combo.SetFocus
'''  End If
'''End Sub

Private Sub cmdTransferValue_Click(Index As Integer)
  Campo(Index).Text = -lblTotalDia(Index).Caption
  Call FormataValorCor(Campo(Index), False)
End Sub

Private Sub Combo_Caixa_CloseUp()
  Combo_Caixa.Text = Combo_Caixa.Columns(1).Text
  Combo_Caixa_LostFocus
End Sub

Private Sub Combo_Caixa_GotFocus()
  sstItem.Tab = 0
End Sub

Private Sub Combo_Caixa_LostFocus()
  Nome_Caixa.Caption = ""
  
  Call RefreshValues(False)
  
  If IsNull(Combo_Caixa.Text) Then Exit Sub
  If Combo_Caixa.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Caixa.Text) Then Exit Sub
  If Val(Combo_Caixa.Text) < 1 Then Exit Sub
  
  rsCaixas_Uso.Index = "Caixa"
  rsCaixas_Uso.Seek "=", Val(Combo_Caixa.Text)
  If rsCaixas_Uso.NoMatch Then Exit Sub
  
  Nome_Caixa.Caption = rsCaixas_Uso("Descrição") & ""
End Sub

Private Sub Combo_Centro_CloseUp()
  Combo_Centro.Text = Combo_Centro.Columns(1).Text
  Combo_Centro_LostFocus
End Sub

Private Sub Combo_Centro_InitColumnProps()
  '04/09/2002 - mpdea
  'Incluído o redimensionamento das colunas
  With Combo_Centro
    .Columns(0).Width = 5000
    .Columns(1).Width = 1000
  End With
End Sub

Private Sub Combo_Centro_LostFocus()
  Nome_Centro.Caption = ""
  
  If IsNull(Combo_Centro.Text) Then Exit Sub
  If Combo_Centro.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Centro.Text) Then Exit Sub
  
  rsCentros.Index = "Código"
  rsCentros.Seek "=", Val(Combo_Centro.Text)
  If rsCentros.NoMatch Then Exit Sub
  
  Nome_Centro.Caption = rsCentros("Nome") & ""
End Sub

Private Sub Combo_Che_CloseUp()
  Combo_Che.Text = Combo_Che.Columns(2).Text
  Combo_Che_LostFocus
End Sub

Private Sub Combo_Che_LostFocus()
   Nome_Che.Caption = ""
   If IsNull(Combo_Che.Text) Then Exit Sub
   If Combo_Che.Text = "" Then Exit Sub
   If Not IsNumeric(Combo_Che.Text) Then Exit Sub
   If Val(Combo_Che.Text) < 1 Then Exit Sub
  '28/11/2006 - Anderson
  'Alteração do número de contas bancárias de 99 para 255
  'Solicitado por: 2227883 - SANTA FÉ DO ARAGUAIA PREFEITURA MUNICIPAL
   If Val(Combo_Che.Text) > 255 Then Exit Sub
   
   rsContas.Index = "Código"
   rsContas.Seek "=", Val(Combo_Che.Text)
   If rsContas.NoMatch Then Exit Sub
   Nome_Che.Caption = rsContas("Descrição")

End Sub

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(2).Text
  Combo_LostFocus
End Sub

Private Sub Combo_Din_CloseUp()
  Combo_Din.Text = Combo_Din.Columns(2).Text
  Combo_Din_LostFocus
End Sub

Private Sub Combo_Din_LostFocus()
   Nome_Din.Caption = ""
   If IsNull(Combo_Din.Text) Then Exit Sub
   If Combo_Din.Text = "" Then Exit Sub
   If Not IsNumeric(Combo_Din.Text) Then Exit Sub
   If Val(Combo_Din.Text) < 1 Then Exit Sub
  '28/11/2006 - Anderson
  'Alteração do número de contas bancárias de 99 para 255
  'Solicitado por: 2227883 - SANTA FÉ DO ARAGUAIA PREFEITURA MUNICIPAL
   If Val(Combo_Din.Text) > 255 Then Exit Sub
   
   rsContas.Index = "Código"
   rsContas.Seek "=", Val(Combo_Din.Text)
   If rsContas.NoMatch Then Exit Sub
   Nome_Din.Caption = rsContas("Descrição")
   
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
  Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_InitColumnProps()
  '04/09/2002 - mpdea
  'Incluído o redimensionamento das colunas
  With Combo_Fornecedor
    .Columns(0).Width = 5000
    .Columns(1).Width = 1000
  End With
End Sub

Private Sub Combo_Fornecedor_LostFocus()
  Nome_Fornecedor.Caption = ""
  If IsNull(Combo_Fornecedor.Text) Then Exit Sub
  If Combo_Fornecedor.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
  
  rsFornecedores.Index = "Código"
  rsFornecedores.Seek "=", Val(Combo_Fornecedor.Text)
  If rsFornecedores.NoMatch Then Exit Sub
  
  Nome_Fornecedor.Caption = rsFornecedores("Nome") & ""
End Sub

Private Sub Combo_LostFocus()
   Nome.Caption = ""
   If IsNull(Combo.Text) Then Exit Sub
   If Combo.Text = "" Then Exit Sub
   If Not IsNumeric(Combo.Text) Then Exit Sub
   If Val(Combo.Text) < 1 Then Exit Sub
   If Val(Combo.Text) > 9999 Then Exit Sub
   
   rsFuncionarios.Index = "Código"
   rsFuncionarios.Seek "=", Val(Combo.Text)
   If rsFuncionarios.NoMatch Then Exit Sub
   Nome.Caption = rsFuncionarios("Apelido")
End Sub

Private Sub Din_Conta_Click()
  If Din_Conta.Value = True Then
    Combo_Din.Enabled = True
    Nome_Din.Enabled = True
  End If
  If Din_Conta.Value = False Then
    Combo_Din.Enabled = False
    Nome_Din.Enabled = False
  End If
End Sub

Private Sub Din_Nada_Click()
  Combo_Din.Enabled = False
End Sub

Private Sub Form_Activate()
  Call Campo_LostFocus(0)
  Call RefreshValues(False)
  
'  B_Confirma.Visible = False
'  B_Novo.Visible = False
  sstItem.Tab = 0
End Sub

Private Sub Form_Load()
  Dim nX As Byte
  
  Call CenterForm(Me)
  
  On Error Resume Next
  Picture1.Picture = LoadPicture(App.Path & "\Imagens\logotipo.bmp")
  
  
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários")
  Set rsCaixas_Uso = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  Set rsCentros = db.OpenRecordset("Centros de Custo", , dbReadOnly)
  Set rsFornecedores = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsContas_Pagar = db.OpenRecordset("Contas a Pagar")

  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  Data6.DatabaseName = gsQuickDBFileName
  
  If gbCaixas = False Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
    Combo_Caixa.Enabled = False
  End If
  
  txt_data_extratoMov.Text = Format(Now, "dd/mm/yyyy")
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsCaixa.Close
  rsFuncionarios.Close
  rsContas.Close
  rsLançamentos.Close
  rsCaixas_Uso.Close
  rsCentros.Close
  rsFornecedores.Close
  rsContas_Pagar.Close
  
  Set rsCaixa = Nothing
  Set rsFuncionarios = Nothing
  Set rsContas = Nothing
  Set rsLançamentos = Nothing
  Set rsCaixas_Uso = Nothing
  Set rsCentros = Nothing
  Set rsFornecedores = Nothing
  Set rsContas_Pagar = Nothing
End Sub



Private Sub O_Gera_Conta_Click()
  Dim bEnabled As Boolean
  
  If O_Gera_Conta.Value = vbChecked Then
    bEnabled = True
  ElseIf O_Gera_Conta.Value = vbUnchecked Then
    bEnabled = False
  End If
  L_Conta(0).Enabled = bEnabled
  L_Conta(1).Enabled = bEnabled
  L_Conta(2).Enabled = bEnabled
  L_Conta(3).Enabled = bEnabled
  L_Conta(4).Enabled = bEnabled
  Desc_Conta.Enabled = bEnabled
  Nota.Enabled = bEnabled
  Combo_Fornecedor.Enabled = bEnabled
  Nome_Fornecedor.Enabled = bEnabled
  Combo_Centro.Enabled = bEnabled
  Nome_Centro.Enabled = bEnabled
  Valor_Conta.Enabled = bEnabled
End Sub

'Atualiza os valores do caixa com opção de valores para o fechamento
Private Sub RefreshValues(ByVal bFechar As Boolean)
  Dim sSql As String
  Dim rsTotalCaixa As Recordset
  Dim rsDataDoSaldoAnterior As Recordset '20/11/2006 - Anderson - Utilizado para verificar qual a data do último saldo anterior
  Dim strDataDoSaldoAnterior As String '20/11/2006 - Anderson - Utilizado para verificar qual a data do último saldo anterior
  Dim nX As Integer
  
  '20/11/2006 - Anderson
  'Verifica se a opção está selecionada nos parametros da filial para considerar saldo anterior na abertura do caixa
  If gbSaldoAnterior Then
    sSql = "SELECT Top 1 Data " & _
           "FROM Caixa " & _
           "WHERE Data<=#" & Format(Data_Atual, "mm/dd/yyyy") & "# " & _
           "GROUP BY Data " & _
           "ORDER BY Data Desc"

    Set rsDataDoSaldoAnterior = db.OpenRecordset(sSql, dbOpenSnapshot)
    
    Do Until rsDataDoSaldoAnterior.EOF
      strDataDoSaldoAnterior = Format(rsDataDoSaldoAnterior("Data"), "mm/dd/yyyy")
      rsDataDoSaldoAnterior.MoveNext
    Loop
    
    rsDataDoSaldoAnterior.Close
    
  Else
    strDataDoSaldoAnterior = Format(Data_Atual, "mm/dd/yyyy")
  End If
  
  If strDataDoSaldoAnterior = "" Then
      strDataDoSaldoAnterior = Format(Data_Atual, "mm/dd/yyyy")
  End If

  sSql = "SELECT Sum(Dinheiro) AS Total0, Sum(Cheques) AS Total1, Sum([Cheques Pré]) AS Total2, " & _
    "Sum(Cartões) AS Total3, Sum(Vales) AS Total4 FROM Caixa WHERE Filial = " & gnCodFilial & _
    " AND Caixa = " & Val(Combo_Caixa.Text) & " AND Data = #" & _
    strDataDoSaldoAnterior & "#"
  
  Set rsTotalCaixa = db.OpenRecordset(sSql, dbOpenSnapshot)
  
  With rsTotalCaixa
    If .RecordCount > 0 Then
      For nX = 0 To 4
        If IsNull(rsTotalCaixa("Total" & nX)) Then
          If bFechar Then
            Campo(nX).Text = Format(0, "#0.00")
          Else
            lblTotalDia(nX).Caption = Format(0, "#0.00")
            Call FormataValorCor(lblTotalDia(nX), False)
          End If
        Else
          If bFechar Then
            Campo(nX).Text = -Format(rsTotalCaixa("Total" & nX), "#0.00")
          Else
            lblTotalDia(nX).Caption = Format(rsTotalCaixa("Total" & nX), "#0.00")
            Call FormataValorCor(lblTotalDia(nX), False)
          End If
        End If
        If bFechar Then
          Call Campo_LostFocus(nX)
        End If
      Next nX
    End If
    .Close
  End With
  Set rsTotalCaixa = Nothing
End Sub

'Atualiza as informações de valores dos caixas
Private Sub RefreshCaixasValues()
  Dim sSql As String
  Dim rsCaixas As Recordset
  Dim rsTotalCaixa As Recordset
  Dim sItem As String
  Dim nX As Integer
  Dim nTotal As Double
  Dim nTotalFinal(5) As Double
  Dim nTotalGeral As Double
  
  Call StatusMsg("Atualizando valores...")
  
  sdbCaixa.RemoveAll
  
  sSql = "SELECT * FROM [Caixas em Uso]"
  Set rsCaixas = db.OpenRecordset(sSql, dbOpenSnapshot)
  
  With rsCaixas
    If .RecordCount > 0 Then
      Do Until .EOF
        sSql = "SELECT Sum(Dinheiro) AS Total0, Sum(Cheques) AS Total1, Sum([Cheques Pré]) AS Total2, " & _
          "Sum(Cartões) AS Total3, Sum(Vales) AS Total4 , Sum(Parcelamento) AS Total5 FROM Caixa " & _
          "WHERE Filial = " & gnCodFilial & "AND Caixa = " & !Caixa & " AND Data = #" & _
          Format(Data_Atual, "mm/dd/yyyy") & "#"
        Set rsTotalCaixa = db.OpenRecordset(sSql, dbOpenSnapshot)
        With rsTotalCaixa
          If .RecordCount > 0 Then
            Do Until .EOF
              sItem = ""
              nTotal = 0
              sItem = rsCaixas!Descrição
              'Verifica se houve movimentação no Caixa
              If Not gbHasMovimentCaixa(rsCaixas!Caixa) Then
                sItem = "* " & sItem
              End If
              For nX = 0 To 5
                If IsNull(rsTotalCaixa("Total" & nX)) Then
                  sItem = sItem & vbTab & Format(0, "#0.00")
                Else
                  sItem = sItem & vbTab & Format(rsTotalCaixa("Total" & nX), FORMAT_VALUE)
                  'Total do caixa
                  If nX <> 5 Then
                    nTotal = nTotal + Format(rsTotalCaixa("Total" & nX), "#0.00")
                  End If
                  nTotalFinal(nX) = nTotalFinal(nX) + Format(rsTotalCaixa("Total" & nX), "#0.00")
                End If
              Next nX
              sdbCaixa.AddItem sItem & vbTab & Format(nTotal, FORMAT_VALUE)
              .MoveNext
            Loop
          End If
          .Close
        End With
        .MoveNext
      Loop
      'Adiciona o total de cada finalizadora
      For nX = 0 To 5
        txtTotal(nX).Text = Format(nTotalFinal(nX), FORMAT_VALUE)
        Call FormataValorCor(txtTotal(nX))
        If nX <> 5 Then
          nTotalGeral = nTotalGeral + Format(nTotalFinal(nX), "#0.00")
        End If
      Next nX
      txtTotalGeral.Text = Format(nTotalGeral, FORMAT_VALUE)
      Call FormataValorCor(txtTotalGeral)
    End If
    .Close
  End With
  
  Set rsTotalCaixa = Nothing
  Set rsCaixas = Nothing
  
  Call StatusMsg("")
  
End Sub

Private Sub O_Gera_Conta_GotFocus()
'''  sstItem.Tab = 1
End Sub

Private Sub sdbCaixa_GotFocus()
  sstItem.Tab = 1
End Sub

Private Sub sdbCaixa_InitColumnProps()
  With sdbCaixa
    .StyleSets.Add "Red"
    .StyleSets("Red").ForeColor = vbRed
    .StyleSets("Red").AlignmentText = ssCaptionAlignmentRight
    .StyleSets.Add "Blue"
    .StyleSets("Blue").ForeColor = vbBlue
    .StyleSets("Blue").AlignmentText = ssCaptionAlignmentRight
    .StyleSets.Add "WinTextColor"
    .StyleSets("WinTextColor").ForeColor = vbWindowText
    .StyleSets("WinTextColor").AlignmentText = ssCaptionAlignmentRight
  End With
End Sub

Private Sub sdbCaixa_RowLoaded(ByVal Bookmark As Variant)
  Dim nX As Integer
  Dim nValue(7) As Double
  
  With sdbCaixa
    For nX = 1 To 7
      nValue(nX) = .Columns(nX).CellText(Bookmark)
      If nValue(nX) < 0 Then
        .Columns(nX).CellStyleSet "Red"
      ElseIf nValue(nX) > 0 Then
        .Columns(nX).CellStyleSet "Blue"
      Else
        .Columns(nX).CellStyleSet "WinTextColor"
      End If
    Next nX
  End With
End Sub

Private Sub SSDBCaixa02_CloseUp()
  SSDBCaixa02.Text = SSDBCaixa02.Columns(1).Text
  SSDBCaixa02_LostFocus
End Sub

Private Sub SSDBCaixa02_LostFocus()
  Nome_caixa02.Caption = ""
  
  Call RefreshValues(False)
  
  If IsNull(SSDBCaixa02.Text) Then Exit Sub
  If SSDBCaixa02.Text = "" Then Exit Sub
  If Not IsNumeric(SSDBCaixa02.Text) Then Exit Sub
  If Val(SSDBCaixa02.Text) < 1 Then Exit Sub
  
  rsCaixas_Uso.Index = "Caixa"
  rsCaixas_Uso.Seek "=", Val(SSDBCaixa02.Text)
  If rsCaixas_Uso.NoMatch Then Exit Sub
  
  Nome_caixa02.Caption = rsCaixas_Uso("Descrição") & ""
End Sub

Private Sub sstItem_Click(PreviousTab As Integer)
'  If sstItem.Tab = 0 Then
'    B_Confirma.Visible = True
'    B_Novo.Visible = True
'    Call RefreshValues(False)
'  ElseIf sstItem.Tab = 1 Then
'    B_Confirma.Visible = True
'    B_Novo.Visible = True
'  ElseIf sstItem.Tab = 2 Then
'    B_Confirma.Visible = False
'    B_Novo.Visible = False
'    Call RefreshCaixasValues
'  End If
End Sub

Private Sub tmrRefresh_Timer()
  Call RefreshValues(False)
  Call RefreshCaixasValues
End Sub
