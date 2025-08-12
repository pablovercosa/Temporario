VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Entradas"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEntradas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15480
   Begin VB.Frame Frame5 
      Height          =   1065
      Left            =   12510
      TabIndex        =   128
      Top             =   7410
      Width           =   2925
      Begin VB.Label lblEfetivada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Operação Efetivada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   60
         TabIndex        =   129
         Top             =   330
         Visible         =   0   'False
         Width           =   2805
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
      Height          =   1065
      Left            =   30
      TabIndex        =   119
      Top             =   7410
      Width           =   12435
      Begin VB.TextBox txtObsValeCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   570
         MaxLength       =   60
         TabIndex        =   126
         Top             =   660
         Width           =   6705
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   10080
         ScaleHeight     =   855
         ScaleWidth      =   2235
         TabIndex        =   123
         Top             =   120
         Width           =   2265
      End
      Begin VB.CommandButton cmd_imprimir 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Imprime Vale"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3900
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   150
         Width           =   3375
      End
      Begin VB.OptionButton opt_ticket 
         Appearance      =   0  'Flat
         Caption         =   "Modelo Ticket"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   150
         TabIndex        =   121
         Top             =   180
         Width           =   1455
      End
      Begin VB.OptionButton opt_relatorio 
         Appearance      =   0  'Flat
         Caption         =   "Modelo Relatório"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1650
         TabIndex        =   120
         Top             =   180
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000000&
         Caption         =   "Obs:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   125
         Top             =   690
         Width           =   435
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000000&
         Caption         =   """logotipo.bmp"" no diretório QuickStore\Imagens"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7770
         TabIndex        =   124
         Top             =   330
         Width           =   2235
      End
   End
   Begin VB.TextBox txt_informacoesComplNFe 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   1980
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   118
      Top             =   1710
      Width           =   9795
   End
   Begin VB.TextBox txtRef 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   6480
      TabIndex        =   113
      Top             =   2520
      Width           =   5295
   End
   Begin VB.ComboBox cboFinalidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmEntradas.frx":4E95A
      Left            =   1140
      List            =   "frmEntradas.frx":4E96A
      TabIndex        =   112
      Text            =   "1=NFe normal"
      Top             =   2520
      Width           =   3405
   End
   Begin VB.ComboBox cboConsumidorFinal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmEntradas.frx":4E9C6
      Left            =   11850
      List            =   "frmEntradas.frx":4E9D0
      TabIndex        =   108
      Text            =   "1=Sim"
      Top             =   1905
      Width           =   3585
   End
   Begin VB.ComboBox cboPresencaComprador 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmEntradas.frx":4E9E2
      Left            =   11850
      List            =   "frmEntradas.frx":4E9F8
      TabIndex        =   107
      Text            =   "1 =Operação presencial"
      Top             =   2535
      Width           =   3585
   End
   Begin VB.TextBox txtNrSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   12990
      MaxLength       =   3
      TabIndex        =   8
      ToolTipText     =   "Entre com o nº da nota fiscal"
      Top             =   1290
      Width           =   1155
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
      Left            =   10050
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Tabela FROM [Tabela de Preços] ORDER BY Tabela"
      Top             =   8190
      Visible         =   0   'False
      Width           =   1740
   End
   Begin TabDlg.SSTab tabItens 
      Height          =   4455
      Left            =   30
      TabIndex        =   12
      Top             =   2940
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   7858
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Itens"
      TabPicture(0)   =   "frmEntradas.frx":4EACC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(22)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(21)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(24)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabels(23)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabels(26)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabels(25)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabels(27)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabels(20)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTotalDesonerado"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "grdItens"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTotDescontos"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtTotProdutos"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtTotBaseICMS"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtTotIPI"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtTotBaseICMSSubst"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTotValICMS"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtTotValICMSSubst"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtTotalAPagar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ddwProduto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdPreencher"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtTotICMSDesonerado"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmd_marcarEtiquetasParaTodasLinhas"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "&Pagamentos"
      TabPicture(1)   =   "frmEntradas.frx":4EAE8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Declaração de Importação"
      TabPicture(2)   =   "frmEntradas.frx":4EB04
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDI"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmd_marcarEtiquetasParaTodasLinhas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Marcar etiquetas para todos os produtos"
         Height          =   405
         Left            =   12060
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   4020
         Width           =   3165
      End
      Begin VB.TextBox txtTotICMSDesonerado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9615
         TabIndex        =   116
         Top             =   3630
         Width           =   1005
      End
      Begin VB.Frame fraDI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -74790
         TabIndex        =   84
         Top             =   480
         Width           =   14985
         Begin VB.Frame fraAdicao 
            Caption         =   "Adição"
            Height          =   825
            Left            =   180
            TabIndex        =   98
            Top             =   1620
            Width           =   14595
            Begin VB.TextBox txtDescontoItemAdicao 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   9840
               TabIndex        =   106
               Top             =   300
               Width           =   1335
            End
            Begin VB.TextBox txtCodigoFabricante 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   7560
               TabIndex        =   104
               Top             =   300
               Width           =   1215
            End
            Begin VB.TextBox txtNumeroSequenciaItem 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   4560
               TabIndex        =   102
               Top             =   300
               Width           =   975
            End
            Begin VB.TextBox txtNumeroAdicao 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   780
               TabIndex        =   100
               Top             =   300
               Width           =   1215
            End
            Begin VB.Label lblDesconto 
               Caption         =   "Desconto"
               Height          =   255
               Left            =   9000
               TabIndex        =   105
               Top             =   345
               Width           =   765
            End
            Begin VB.Label lblCodigoFabricante 
               Caption         =   "Código do Fabricante"
               Height          =   255
               Left            =   5820
               TabIndex        =   103
               Top             =   345
               Width           =   1725
            End
            Begin VB.Label lblNumeroSequencialItem 
               Caption         =   "Número Sequencial do Item"
               Height          =   315
               Left            =   2400
               TabIndex        =   101
               Top             =   315
               Width           =   2175
            End
            Begin VB.Label lblNumeroAdicao 
               Caption         =   "Número"
               Height          =   255
               Left            =   120
               TabIndex        =   99
               Top             =   345
               Width           =   645
            End
         End
         Begin VB.Frame fraDesembaracoDI 
            Caption         =   "Desembaraço"
            Height          =   855
            Left            =   180
            TabIndex        =   91
            Top             =   690
            Width           =   14595
            Begin VB.TextBox txtUFDesembaracoDI 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   10560
               TabIndex        =   97
               Top             =   330
               Width           =   615
            End
            Begin VB.TextBox txtLocalDesembaracoDI 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   2880
               TabIndex        =   95
               Top             =   330
               Width           =   7095
            End
            Begin MSMask.MaskEdBox medDataDesembaracoDI 
               Height          =   315
               Left            =   750
               TabIndex        =   92
               ToolTipText     =   "Pressione F2 para Calendário"
               Top             =   330
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   15066597
               MaxLength       =   18
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
            Begin VB.Label lblUFDesembaraco 
               Caption         =   "UF"
               Height          =   255
               Left            =   10230
               TabIndex        =   96
               Top             =   360
               Width           =   285
            End
            Begin VB.Label lblLocalDesembaraco 
               Caption         =   "Local "
               Height          =   255
               Left            =   2370
               TabIndex        =   94
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lblDataDesembaraco 
               AutoSize        =   -1  'True
               Caption         =   "Data"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   93
               Top             =   375
               Width           =   345
            End
         End
         Begin VB.TextBox txtCodigoExportadorDI 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            Height          =   330
            Left            =   4560
            MaxLength       =   40
            TabIndex        =   89
            Top             =   240
            Width           =   3210
         End
         Begin VB.TextBox txtNumeroDI 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            Height          =   330
            Left            =   900
            MaxLength       =   10
            TabIndex        =   85
            Top             =   240
            Width           =   1305
         End
         Begin MSMask.MaskEdBox medDataRegistroDI 
            Height          =   315
            Left            =   10170
            TabIndex        =   87
            ToolTipText     =   "Pressione F2 para Calendário"
            Top             =   240
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   15066597
            MaxLength       =   17
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
         Begin VB.Label lblCodigoExportador 
            AutoSize        =   -1  'True
            Caption         =   "Código do Exportador"
            Height          =   195
            Index           =   19
            Left            =   2790
            TabIndex        =   90
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label lblDataRegistroDI 
            AutoSize        =   -1  'True
            Caption         =   "Data de Registro"
            Height          =   195
            Index           =   19
            Left            =   8850
            TabIndex        =   88
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label lblNumeroDI 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   19
            Left            =   210
            TabIndex        =   86
            Top             =   285
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdPreencher 
         Caption         =   "!!250"
         Height          =   255
         Left            =   9180
         TabIndex        =   77
         Top             =   3930
         Visible         =   0   'False
         Width           =   615
      End
      Begin SSDataWidgets_B.SSDBDropDown ddwProduto 
         Bindings        =   "frmEntradas.frx":4EB20
         Height          =   1125
         Left            =   120
         TabIndex        =   67
         Top             =   1440
         Width           =   11040
         DataFieldList   =   "Código"
         ListAutoValidate=   0   'False
         MaxDropDownItems=   16
         _Version        =   196617
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   53
         Columns.Count   =   5
         Columns(0).Width=   9472
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3334
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Codigo"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2408
         Columns(2).Caption=   "Qtde/Tipo"
         Columns(2).Name =   "Qtde"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2487
         Columns(3).Caption=   "Preço"
         Columns(3).Name =   "Preco"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   2037
         Columns(4).Caption=   "Unid"
         Columns(4).Name =   "Unidade"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         _ExtentX        =   19473
         _ExtentY        =   1984
         _StockProps     =   77
      End
      Begin VB.TextBox txtTotalAPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   13530
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   20
         Top             =   3585
         Width           =   1695
      End
      Begin VB.TextBox txtTotValICMSSubst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5970
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   19
         Top             =   3975
         Width           =   1305
      End
      Begin VB.TextBox txtTotValICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3255
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   17
         Top             =   3975
         Width           =   1305
      End
      Begin VB.TextBox txtTotBaseICMSSubst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5970
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   18
         Top             =   3585
         Width           =   1305
      End
      Begin VB.TextBox txtTotIPI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10950
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   21
         Top             =   3615
         Width           =   1395
      End
      Begin VB.TextBox txtTotBaseICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3255
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   16
         Top             =   3585
         Width           =   1305
      End
      Begin VB.TextBox txtTotProdutos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   14
         Top             =   3585
         Width           =   1305
      End
      Begin VB.TextBox txtTotDescontos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   15
         Top             =   3975
         Width           =   1305
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3945
         Left            =   -74895
         TabIndex        =   51
         Top             =   375
         Width           =   15180
         Begin VB.Data datCheques 
            Caption         =   "datCheques"
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
            Left            =   12540
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   2  'Snapshot
            RecordSource    =   $"frmEntradas.frx":4EB3A
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
         End
         Begin SSDataWidgets_B.SSDBDropDown ddwCheques 
            Bindings        =   "frmEntradas.frx":4EC1A
            Height          =   735
            Left            =   10620
            TabIndex        =   83
            Top             =   1440
            Width           =   2175
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
            BackColorOdd    =   12640511
            RowHeight       =   423
            Columns.Count   =   4
            Columns(0).Width=   3200
            Columns(0).Caption=   "Cheque"
            Columns(0).Name =   "Cheque"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Cheque"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   2064
            Columns(1).Caption=   "Vencimento"
            Columns(1).Name =   "Vencimento"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Vencimento"
            Columns(1).DataType=   7
            Columns(1).FieldLen=   256
            Columns(2).Width=   2170
            Columns(2).Caption=   "Valor"
            Columns(2).Name =   "Valor"
            Columns(2).Alignment=   1
            Columns(2).CaptionAlignment=   1
            Columns(2).DataField=   "Valor"
            Columns(2).DataType=   5
            Columns(2).FieldLen=   256
            Columns(3).Width=   1429
            Columns(3).Caption=   "Banco"
            Columns(3).Name =   "Banco"
            Columns(3).Alignment=   1
            Columns(3).CaptionAlignment=   1
            Columns(3).DataField=   "Banco"
            Columns(3).DataType=   3
            Columns(3).FieldLen=   256
            _ExtentX        =   3836
            _ExtentY        =   1296
            _StockProps     =   77
         End
         Begin VB.TextBox txtTroco 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   12360
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   2520
            Width           =   2130
         End
         Begin VB.Frame Frame3 
            Caption         =   "Cheque Usado"
            Height          =   1845
            Left            =   210
            TabIndex        =   71
            Top             =   1800
            Width           =   6405
            Begin VB.TextBox txtCheque 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   315
               Left            =   1380
               MaxLength       =   10
               TabIndex        =   30
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox txtValCheque 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   31
               Top             =   1080
               Width           =   1245
            End
            Begin VB.TextBox txtDescricao 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   315
               Left            =   1380
               MaxLength       =   40
               TabIndex        =   29
               Top             =   720
               Width           =   3000
            End
            Begin SSDataWidgets_B.SSDBCombo cboConta 
               Bindings        =   "frmEntradas.frx":4EC33
               Height          =   315
               Left            =   1380
               TabIndex        =   28
               Top             =   360
               Width           =   1020
               DataFieldList   =   "Código"
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
               _ExtentX        =   1799
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   12648447
            End
            Begin MSMask.MaskEdBox medDataBomPara 
               Height          =   315
               Left            =   1380
               TabIndex        =   32
               ToolTipText     =   "Pressione F2 para Calendário"
               Top             =   1440
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   15066597
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
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "C&onta"
               Height          =   195
               Index           =   15
               Left            =   255
               TabIndex        =   27
               Top             =   405
               Width           =   435
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Número"
               Height          =   195
               Index           =   16
               Left            =   240
               TabIndex        =   76
               Top             =   1110
               Width           =   555
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Descrição"
               Height          =   195
               Index           =   17
               Left            =   255
               TabIndex        =   75
               Top             =   765
               Width           =   690
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Bom Para"
               Height          =   195
               Index           =   18
               Left            =   240
               TabIndex        =   74
               Top             =   1485
               Width           =   675
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Height          =   195
               Index           =   28
               Left            =   2700
               TabIndex        =   73
               Top             =   1125
               Width           =   360
            End
            Begin VB.Label lblConta 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   2430
               TabIndex        =   72
               Top             =   360
               Width           =   3585
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Caixa Usado"
            Height          =   1485
            Left            =   210
            TabIndex        =   69
            Top             =   270
            Width           =   6405
            Begin VB.TextBox txtCxDinheiro 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   1380
               MaxLength       =   15
               TabIndex        =   24
               Top             =   720
               Width           =   3000
            End
            Begin VB.TextBox txtCxCheque 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               Height          =   330
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   26
               Top             =   1080
               Width           =   3000
            End
            Begin SSDataWidgets_B.SSDBCombo cboCaixaUso 
               Bindings        =   "frmEntradas.frx":4EC4A
               Height          =   315
               Left            =   270
               TabIndex        =   22
               ToolTipText     =   "Para colocar o cursor através do TAB em Caixa usado, tecle F7"
               Top             =   360
               Width           =   1050
               DataFieldList   =   "Caixa"
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
               _ExtentX        =   1852
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   12648447
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "&Dinheiro"
               Height          =   195
               Index           =   12
               Left            =   270
               TabIndex        =   23
               Top             =   765
               Width           =   585
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "&Cheque"
               Height          =   195
               Index           =   13
               Left            =   270
               TabIndex        =   25
               Top             =   1125
               Width           =   555
            End
            Begin VB.Label lblCaixaUso 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1380
               TabIndex        =   70
               Top             =   360
               Width           =   4650
            End
         End
         Begin VB.TextBox txtADigitar 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   12360
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   38
            Top             =   3240
            Width           =   2130
         End
         Begin VB.TextBox txtTotDigitado 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   12360
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   2880
            Width           =   2130
         End
         Begin SSDataWidgets_B.SSDBGrid grdCP 
            Height          =   2535
            Left            =   6960
            TabIndex        =   33
            Top             =   360
            Width           =   3165
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
            BackColorOdd    =   12648384
            RowHeight       =   423
            ExtraHeight     =   53
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
            _ExtentX        =   5583
            _ExtentY        =   4471
            _StockProps     =   79
            Caption         =   "Contas a Pagar"
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
         Begin SSDataWidgets_B.SSDBCombo cboCodigoCC 
            Bindings        =   "frmEntradas.frx":4EC65
            Height          =   300
            Left            =   6960
            TabIndex        =   35
            Top             =   3240
            Width           =   1245
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
            BackColorOdd    =   16777152
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
            _ExtentX        =   2196
            _ExtentY        =   529
            _StockProps     =   93
            Text            =   "0"
            BackColor       =   12648447
            DataFieldToDisplay=   "Nome"
         End
         Begin SSDataWidgets_B.SSDBGrid Grade_ChequesEmCaixa 
            Height          =   2055
            Left            =   10440
            TabIndex        =   36
            Top             =   360
            Width           =   4035
            _Version        =   196617
            DataMode        =   2
            Col.Count       =   4
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BackColorOdd    =   12648384
            RowHeight       =   423
            ExtraHeight     =   53
            Columns.Count   =   4
            Columns(0).Width=   1826
            Columns(0).Caption=   "Número"
            Columns(0).Name =   "NumeroCheque"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1455
            Columns(1).Caption=   "Valor"
            Columns(1).Name =   "Valor"
            Columns(1).Alignment=   1
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(2).Width=   1746
            Columns(2).Caption=   "Bom Para"
            Columns(2).Name =   "Bom Para"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(2).Locked=   -1  'True
            Columns(3).Width=   1032
            Columns(3).Caption=   "Banco"
            Columns(3).Name =   "Banco"
            Columns(3).Alignment=   1
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(3).Locked=   -1  'True
            _ExtentX        =   7117
            _ExtentY        =   3625
            _StockProps     =   79
            Caption         =   "Cheques em Caixa Utilizados para Pagar"
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
         Begin SSDataWidgets_B.SSDBDropDown ddwCartoes 
            Bindings        =   "frmEntradas.frx":4EC82
            Height          =   735
            Left            =   11220
            TabIndex        =   80
            Top             =   1680
            Width           =   1455
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
            Columns(0).Width=   3200
            _ExtentX        =   2566
            _ExtentY        =   1296
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblTroco 
            AutoSize        =   -1  'True
            Caption         =   "Troco"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   19
            Left            =   11160
            TabIndex        =   82
            Top             =   2535
            Width           =   495
         End
         Begin VB.Label lblNomeCC 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8160
            TabIndex        =   68
            Top             =   3240
            Width           =   1965
         End
         Begin VB.Label Label1 
            Caption         =   "Centro de c&usto"
            Height          =   255
            Left            =   6960
            TabIndex        =   34
            Top             =   3000
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "A Digitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   30
            Left            =   11160
            TabIndex        =   55
            Top             =   3255
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Digitado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   29
            Left            =   11160
            TabIndex        =   54
            Top             =   2895
            Width           =   690
         End
      End
      Begin SSDataWidgets_B.SSDBGrid grdItens 
         Height          =   3165
         Left            =   60
         TabIndex        =   13
         Top             =   390
         Width           =   15180
         ScrollBars      =   3
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   24
         stylesets.count =   1
         stylesets(0).Name=   "Font12"
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "frmEntradas.frx":4EC96
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         UseExactRowCount=   0   'False
         ForeColorEven   =   0
         BackColorOdd    =   12648384
         RowHeight       =   423
         ExtraHeight     =   185
         Columns.Count   =   24
         Columns(0).Width=   3519
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1138
         Columns(1).Caption=   "Qtde"
         Columns(1).Name =   "Qtde"
         Columns(1).Alignment=   1
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   5
         Columns(1).FieldLen=   256
         Columns(1).StyleSet=   "Font12"
         Columns(2).Width=   1217
         Columns(2).Caption=   "Índice"
         Columns(2).Name =   "Indice"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   4974
         Columns(3).Caption=   "Descrição"
         Columns(3).Name =   "Descricao"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(4).Width=   767
         Columns(4).Caption=   "UN"
         Columns(4).Name =   "Unidade"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   2514
         Columns(5).Caption=   "Preço"
         Columns(5).Name =   "Preco"
         Columns(5).Alignment=   1
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   2514
         Columns(6).Caption=   "Preço Total"
         Columns(6).Name =   "PrecoTotal"
         Columns(6).Alignment=   1
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   1032
         Columns(7).Caption=   "Des %"
         Columns(7).Name =   "Desconto"
         Columns(7).Alignment=   1
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   4
         Columns(7).FieldLen=   256
         Columns(8).Width=   1058
         Columns(8).Caption=   "ICM %"
         Columns(8).Name =   "ICMS"
         Columns(8).Alignment=   1
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   4
         Columns(8).FieldLen=   256
         Columns(9).Width=   900
         Columns(9).Caption=   "IPI %"
         Columns(9).Name =   "IPI"
         Columns(9).Alignment=   1
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   4
         Columns(9).FieldLen=   256
         Columns(10).Width=   2699
         Columns(10).Caption=   "Preço Final"
         Columns(10).Name=   "PrecoFinal"
         Columns(10).Alignment=   1
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(10).StyleSet=   "Font12"
         Columns(11).Width=   714
         Columns(11).Caption=   "Etiq"
         Columns(11).Name=   "Etiqueta"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   11
         Columns(11).FieldLen=   256
         Columns(11).Style=   2
         Columns(11).Nullable=   0
         Columns(12).Width=   3200
         Columns(12).Visible=   0   'False
         Columns(12).Caption=   "Base_ICM"
         Columns(12).Name=   "Base_ICM"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         Columns(13).Width=   3200
         Columns(13).Visible=   0   'False
         Columns(13).Caption=   "Valor_ICM"
         Columns(13).Name=   "Valor_ICM"
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   8
         Columns(13).FieldLen=   256
         Columns(14).Width=   3200
         Columns(14).Visible=   0   'False
         Columns(14).Caption=   "Valor_Base_Unit"
         Columns(14).Name=   "Valor_Base_Unit"
         Columns(14).DataField=   "Column 14"
         Columns(14).DataType=   8
         Columns(14).FieldLen=   256
         Columns(15).Width=   3200
         Columns(15).Visible=   0   'False
         Columns(15).Caption=   "Redução_ICM"
         Columns(15).Name=   "Redução_ICM"
         Columns(15).DataField=   "Column 15"
         Columns(15).DataType=   8
         Columns(15).FieldLen=   256
         Columns(16).Width=   3200
         Columns(16).Visible=   0   'False
         Columns(16).Caption=   "Tipo_ICM"
         Columns(16).Name=   "Tipo_ICM"
         Columns(16).DataField=   "Column 16"
         Columns(16).DataType=   8
         Columns(16).FieldLen=   256
         Columns(17).Width=   3200
         Columns(17).Visible=   0   'False
         Columns(17).Caption=   "IndicePrecoEntrada"
         Columns(17).Name=   "IndicePrecoEntrada"
         Columns(17).DataField=   "Column 17"
         Columns(17).DataType=   8
         Columns(17).FieldLen=   256
         Columns(17).Locked=   -1  'True
         Columns(18).Width=   3200
         Columns(18).Visible=   0   'False
         Columns(18).Caption=   "PrecoSemIndiceEntrada"
         Columns(18).Name=   "PrecoSemIndiceEntrada"
         Columns(18).DataField=   "Column 18"
         Columns(18).DataType=   8
         Columns(18).FieldLen=   256
         Columns(18).Locked=   -1  'True
         Columns(19).Width=   3200
         Columns(19).Visible=   0   'False
         Columns(19).Caption=   "IndiceIcmsRetido"
         Columns(19).Name=   "IndiceIcmsRetido"
         Columns(19).DataField=   "Column 19"
         Columns(19).DataType=   8
         Columns(19).FieldLen=   256
         Columns(19).Locked=   -1  'True
         Columns(20).Width=   3200
         Columns(20).Visible=   0   'False
         Columns(20).Caption=   "ValorIcmsRetido"
         Columns(20).Name=   "ValorIcmsRetido"
         Columns(20).DataField=   "Column 20"
         Columns(20).DataType=   8
         Columns(20).FieldLen=   256
         Columns(20).Locked=   -1  'True
         Columns(21).Width=   3200
         Columns(21).Visible=   0   'False
         Columns(21).Caption=   "IcmsSaida"
         Columns(21).Name=   "IcmsSaida"
         Columns(21).DataField=   "Column 21"
         Columns(21).DataType=   8
         Columns(21).FieldLen=   256
         Columns(21).Locked=   -1  'True
         Columns(22).Width=   1138
         Columns(22).Caption=   "ICMS Deson."
         Columns(22).Name=   "Valor Desonerado"
         Columns(22).Alignment=   1
         Columns(22).DataField=   "Column 22"
         Columns(22).DataType=   8
         Columns(22).FieldLen=   256
         Columns(23).Width=   1217
         Columns(23).Caption=   "% Dif."
         Columns(23).Name=   "% Diferimento"
         Columns(23).Alignment=   1
         Columns(23).DataField=   "Column 23"
         Columns(23).DataType=   8
         Columns(23).FieldLen=   256
         _ExtentX        =   26776
         _ExtentY        =   5583
         _StockProps     =   79
         ForeColor       =   0
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
      Begin VB.Label lblTotalDesonerado 
         Caption         =   "ICMS Deson"
         Height          =   255
         Left            =   8640
         TabIndex        =   115
         Top             =   3675
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar"
         Height          =   195
         Index           =   20
         Left            =   12465
         TabIndex        =   66
         Top             =   3660
         Width           =   960
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Valor ICMS Subst"
         Height          =   195
         Index           =   27
         Left            =   4650
         TabIndex        =   65
         Top             =   4050
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Valor ICMS"
         Height          =   195
         Index           =   25
         Left            =   2370
         TabIndex        =   64
         Top             =   4050
         Width           =   780
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Base ICMS Subst"
         Height          =   195
         Index           =   26
         Left            =   4650
         TabIndex        =   63
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "IPI"
         Height          =   195
         Index           =   23
         Left            =   10695
         TabIndex        =   62
         Top             =   3690
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Base ICMS"
         Height          =   195
         Index           =   24
         Left            =   2370
         TabIndex        =   61
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Produtos"
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   60
         Top             =   3660
         Width           =   645
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Descontos"
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   59
         Top             =   4050
         Width           =   750
      End
   End
   Begin VB.Data datCentroCusto 
      Caption         =   "CC"
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
      Left            =   7380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM [Centros de Custo] WHERE Ativo ORDER BY Nome"
      Top             =   7770
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data datProdutos 
      Caption         =   "Data1"
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
      Left            =   7920
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT Nome, Código FROM Produtos WHERE Desativado = False ORDER BY Nome"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data datConta 
      Caption         =   "Data1"
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
      Left            =   4260
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Contas Bancárias"
      Top             =   7860
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data datCaixasUso 
      Caption         =   "Data1"
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
      Left            =   5940
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Caixas em Uso"
      Top             =   7860
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data datFornecedor 
      Caption         =   "Data1"
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
      Left            =   5670
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT Código, Nome, Tipo, Estado FROM Cli_For WHERE Inativo = False ORDER BY Nome"
      Top             =   6450
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data datDigitador 
      Caption         =   "Data1"
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
      Left            =   3810
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT * FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE AND IsPrestServ = False ORDER BY Nome"
      Top             =   6435
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data datOper 
      Caption         =   "Data1"
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
      Left            =   1965
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT * FROM [Operações Entrada] ORDER BY Nome"
      Top             =   6435
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data datEntrada 
      Caption         =   "Data1"
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
      Left            =   2820
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Entradas"
      Top             =   7950
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtFrete 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   12990
      MaxLength       =   15
      TabIndex        =   10
      Top             =   660
      Width           =   1155
   End
   Begin VB.TextBox txtSeq 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   13680
      TabIndex        =   3
      Top             =   30
      Width           =   1755
   End
   Begin VB.TextBox txtPedido 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   11850
      MaxLength       =   15
      TabIndex        =   9
      Top             =   660
      Width           =   1035
   End
   Begin VB.TextBox txtNF 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   11850
      MaxLength       =   15
      TabIndex        =   7
      Top             =   1290
      Width           =   1035
   End
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1980
      MaxLength       =   70
      TabIndex        =   4
      Top             =   1290
      Width           =   9795
   End
   Begin SSDataWidgets_B.SSDBCombo cboOper 
      Bindings        =   "frmEntradas.frx":4ECB2
      Height          =   345
      Left            =   780
      TabIndex        =   0
      Top             =   480
      Width           =   1155
      DataFieldList   =   "Nome"
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   6429
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1085
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   900
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1323
      Columns(3).Caption=   "Estoque"
      Columns(3).Name =   "Estoque"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Estoque"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      Columns(4).Width=   2910
      Columns(4).Caption=   "Somar Frete ao Total"
      Columns(4).Name =   "Somar Frete ao Total"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "Somar Frete ao Total"
      Columns(4).DataType=   11
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      _ExtentX        =   2037
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboDigitador 
      Bindings        =   "frmEntradas.frx":4ECC8
      Height          =   345
      Left            =   6780
      TabIndex        =   1
      Top             =   450
      Width           =   1095
      DataFieldList   =   "Nome"
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   5847
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1984
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
      Bindings        =   "frmEntradas.frx":4ECE3
      Height          =   345
      Left            =   6780
      TabIndex        =   2
      Top             =   30
      Width           =   1095
      DataFieldList   =   "Nome"
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   8731
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2302
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   873
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1164
      Columns(3).Caption=   "Estado"
      Columns(3).Name =   "Estado"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "Estado"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin MSMask.MaskEdBox medDataAcerto 
      Height          =   270
      Left            =   780
      TabIndex        =   11
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   945
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   476
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   15066597
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox medDataEmissao 
      Height          =   300
      Left            =   14250
      TabIndex        =   6
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   1290
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   15066597
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   9795
      Top             =   6885
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin SSDataWidgets_B.SSDBCombo cboTabela 
      Bindings        =   "frmEntradas.frx":4ECFF
      Height          =   345
      Left            =   7860
      TabIndex        =   5
      ToolTipText     =   $"frmEntradas.frx":4ED17
      Top             =   870
      Width           =   3915
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
      BackColorOdd    =   12648447
      Columns(0).Width=   3200
      _ExtentX        =   6906
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      DataFieldToDisplay=   "Tabela"
   End
   Begin VB.Label lbl_avisoTrataComissaoVendedor 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "  Esta operação trata comissão!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3030
      TabIndex        =   127
      Top             =   960
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Label Label2 
      Caption         =   "Informações Complementares p/ NF-e"
      Height          =   510
      Left            =   60
      TabIndex        =   117
      Top             =   1710
      Width           =   1815
   End
   Begin VB.Label lblRef 
      Caption         =   "Chave Ref."
      Height          =   255
      Index           =   1
      Left            =   5550
      TabIndex        =   114
      Top             =   2565
      Width           =   885
   End
   Begin VB.Label lblFinalidade 
      Caption         =   "Finalidade NFe"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   111
      Top             =   2565
      Width           =   1065
   End
   Begin VB.Label lblConsumidorFinal 
      Caption         =   "Consumidor Final"
      Height          =   225
      Left            =   11850
      TabIndex        =   110
      Top             =   1650
      Width           =   1275
   End
   Begin VB.Label lblPresencaComprador 
      Caption         =   "Ind Presença Comprador"
      Height          =   225
      Left            =   11850
      TabIndex        =   109
      Top             =   2280
      Width           =   1995
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nr Série"
      Height          =   195
      Index           =   14
      Left            =   12990
      TabIndex        =   79
      Top             =   1050
      Width           =   570
   End
   Begin VB.Label lblTabela 
      AutoSize        =   -1  'True
      Caption         =   "Tabela de Preços"
      Height          =   195
      Left            =   6600
      TabIndex        =   78
      Top             =   960
      Width           =   1230
   End
   Begin VB.Label lblFornecedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   7890
      TabIndex        =   58
      Top             =   30
      Width           =   3885
   End
   Begin VB.Label lblDigitador 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   7890
      TabIndex        =   57
      Top             =   450
      Width           =   3885
   End
   Begin VB.Label lblOper 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   1980
      TabIndex        =   56
      Top             =   420
      Width           =   3675
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   9120
      Top             =   8070
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmEntradas.frx":4ED9E
   End
   Begin VB.Label lblToday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   14250
      TabIndex        =   53
      Top             =   660
      Width           =   1185
   End
   Begin VB.Label lblFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   780
      TabIndex        =   52
      Top             =   30
      Width           =   4875
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Pedido"
      Height          =   195
      Index           =   11
      Left            =   11850
      TabIndex        =   50
      Top             =   420
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sequência"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   12630
      TabIndex        =   49
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Index           =   9
      Left            =   14250
      TabIndex        =   48
      Top             =   420
      Width           =   345
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Fornec/Cliente"
      Height          =   195
      Index           =   8
      Left            =   5700
      TabIndex        =   47
      Top             =   90
      Width           =   1050
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Digitador"
      Height          =   195
      Index           =   7
      Left            =   6105
      TabIndex        =   46
      Top             =   540
      Width           =   645
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "NF"
      Height          =   195
      Index           =   6
      Left            =   11850
      TabIndex        =   45
      Top             =   1050
      Width           =   195
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Obs"
      Height          =   195
      Index           =   5
      Left            =   60
      TabIndex        =   44
      Top             =   1410
      Width           =   285
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Operação"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   43
      Top             =   540
      Width           =   705
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Frete"
      Height          =   195
      Index           =   3
      Left            =   12990
      TabIndex        =   42
      Top             =   435
      Width           =   390
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Acerto"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Emissão"
      Height          =   195
      Index           =   1
      Left            =   14250
      TabIndex        =   40
      Top             =   1050
      Width           =   570
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   39
      Top             =   90
      Width           =   300
   End
End
Attribute VB_Name = "frmEntrada"
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

Private gbShowEstoque As Boolean
Private rsEntradas As Recordset
Private rsProdutos As Recordset
Private rsOp_Entrada  As Recordset
Private rsGrade As Recordset
Private rsMovi_Parcelas As Recordset
Private rsParametros  As Recordset
Private rsPrecos  As Recordset
Private rsEntra_Prod  As Recordset
Private rsCotacoes  As Recordset
Private rsContas  As Recordset
Private rsEntradas2  As Recordset
Private rsLog  As Recordset
Private rsEstados As Recordset

'07/01/2010 - Andrea
Private rsMovi_Cheques As Recordset

Private gbSomarFrete As Boolean
Private gsTipoOper As String
Private gbLiberaPagar As Boolean

Private gsSql As String
Private gsOrder As String
Private gsWhere As String

Private gbBaseICMSomadoIPI As Boolean
Private gbIPI As Boolean
Private gbIPI_TOT As Boolean

Private Erro_Data2 As Boolean
Private Erro As Boolean
Private Num_Registro As Variant

Private sEstado As String

'13/12/2005 - mpdea
'Indica a utilização de 5 casas decimais na tela de Entradas
Private m_bln_5CasasEntrada As Boolean

'24/08/2004 - Daniel
'Var para controlar o indice financeiro
'sempre será True quando ocorrer a altração de preço
Private m_blnIndice    As Boolean
'14/09/2004 - Daniel
'Flag para indicação se é o cliente Livraria Resultado
Private m_blnResultado  As Boolean
Private m_blnEmprestimo As Boolean
'29/11/2004 - Daniel
'Esta var é utilizada para identificar o cliente
'Teknika que possui tratamento especial para o ICMS
Dim m_blnTeknika As Boolean
'14/02/2005 - Daniel
'
'Solicitante: Daring - RJ
'
'Se ocorre devolução e esta devolução implica em abatimento de
'comissão do vendedor, o Quick estava descontando erroneamente
'da comissão para casos em que a venda possuia descontos.
Public gsTabelaVenda As String
'06/05/2005 - Daniel
'
'Implementação.: Trabalhar com o código para fornecedor cadastrado na tela de produtos.
'                Impacto: Ao entrar com o código para o fornecedor no campo código do produto
'                o sistema deverá trazer o código do produto que estiver amarrado nele
'Solicitação...: Cristiano Pavinato - PSI RS
Private m_blnUsaCodFornec As Boolean
'19/05/2005 - Daniel
'Var modular que tratará o foco para não ficar amarrado no objeto txtNF caso ocorra
'criação de nota manual já que a gravação da saída ocorre antes do txtNF_LostFocus
Private m_blnFocoNF As Boolean
'25/06/2013-Alexandre Afornali
Dim gravada As Boolean
Dim Total_Valor_Desonerado As Double

Public bTelaChamadoraDevolucaoProdutos As Boolean       ' Tela de saída botão Devoluções => Tela Devolução Produtos
Public sCodEntradaDevolucaoProdutos As String
Public bGerarDevolucaoAfentandoComissao As Boolean
Public bTelaChamadoraDevolucao_ValeCredito As Boolean   ' Aba Movimentacao opção => Devoluções


Private Sub MostraRegistro()
  Dim nI As Integer
  Dim l As Integer
  Dim Linha As Integer
  Dim Aux As String
  Dim Cód As String
  Dim Aux_Prod As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim sValor As String
  Dim rs As Recordset
  Dim Tool As ActiveBarLibraryCtl.Tool
  
  On Error GoTo ErrMostra
  
  'Na mudança de registro o Altera Totais é desmarcado
  Set Tool = ActiveBar1.Tools("miComplAlteraTotais")
  If Tool.Checked Then
    Call ActiveBar1_Click(Tool)
  End If

  lblToday.Caption = Format(rsEntradas("Data"), "dd/mm/yyyy")
  lblEfetivada.Visible = rsEntradas("Efetivada")

  cboOper.Text = rsEntradas("Operação")
  cboOper_LostFocus

  '14/02/2005 - Daniel
  '
  'Solicitante: Daring - RJ
  '
  'Se ocorre devolução e esta devolução implica em abatimento de
  'comissão do vendedor, o Quick estava descontando erroneamente
  'da comissão para casos em que a venda possuia descontos.
'''  lblTabela.Visible = False
'''  cboTabela.Visible = False
  If ControlarComisao = True Then
      lbl_avisoTrataComissaoVendedor.Visible = True
  Else
      lbl_avisoTrataComissaoVendedor.Visible = False
  End If
  '-------------------------------------------------------------

  cboDigitador.Text = rsEntradas("Digitador")
  cboDigitador_LostFocus
  
  cboFornecedor.Text = rsEntradas("Fornecedor")
  cboFornecedor_LostFocus
  
  cboCodigoCC.Text = rsEntradas("CentroCusto") & ""
  cboCodigoCC_LostFocus
  
  txtSeq.Text = rsEntradas("Sequência") & ""
  txtObs.Text = rsEntradas("Observações") & ""
  txtNF.Text = rsEntradas("Nota Fiscal") & ""
  
  Select Case rsEntradas("Consumidor_Final").Value
    Case "1"
      cboConsumidorFinal.Text = "1=Sim"
    Case Else
      cboConsumidorFinal.Text = "0=Não"
  End Select
  Select Case rsEntradas("Presenca_Comprador").Value
    Case "0"
      cboPresencaComprador.Text = "0=Não se aplica"
    Case "1"
      cboPresencaComprador.Text = "1 =Operação presencial"
    Case "2"
      cboPresencaComprador.Text = "2=Operação não presencial, pela Internet"
    Case "3"
      cboPresencaComprador.Text = "3=Operação não presencial, Teleatendimento"
    Case "4"
      cboPresencaComprador.Text = "4=NFC-e em operação com entrega em domicílio"
    Case "9"
      cboPresencaComprador.Text = "9=Operação não presencial, outros"
    Case Else
      cboPresencaComprador.Text = "1 =Operação presencial"
  End Select
  Select Case rsEntradas("FinalidadeNFe").Value
    Case "1"
      cboFinalidade.Text = "1=NFe normal"
    Case "2"
      cboFinalidade.Text = "2=NF-e complementar;"
    Case "3"
      cboFinalidade.Text = "3=NF-e de ajuste; "
    Case "4"
      cboFinalidade.Text = "4=Devolução de mercadoria."
    Case Else
      cboFinalidade.Text = "1=NFe normal"
  End Select
  If Len(rsEntradas("ChaveReferenciada").Value) > 0 Then
    txtRef.Text = rsEntradas("ChaveReferenciada").Value
  Else
    txtRef.Text = ""
  End If
  '19/05/2005 - Daniel
  '
  'Solicitante: Pedágio Calçados - Otimização liberada
  '             para todos usuários do Quick Store
  '
  'Tratamento para o campo Nr Série da NF
  txtNrSerie.Text = rsEntradas.fields("SerieNF").Value & ""
  '-------------------------------------------------------------
  
  txtPedido.Text = rsEntradas("Pedido") & ""

  medDataEmissao.Mask = ""
  medDataEmissao.Text = ""
  If IsDate(rsEntradas("Data Emissão")) Then
    medDataEmissao.Text = gsFormatDate(rsEntradas("Data Emissão"))
  End If
  medDataEmissao.Mask = "##/##/####"

  
  '07/11/2002 - mpdea
  'Incluído exibição da Data de Acerto do Empréstimo
  medDataAcerto.Mask = ""
  medDataAcerto.Text = ""
  If IsDate(rsEntradas("Data Acerto Empréstimo")) Then
    medDataAcerto.Text = gsFormatDate(rsEntradas("Data Acerto Empréstimo"))
  End If
  medDataAcerto.Mask = "##/##/####"
    

  Call LoadGridItens
  
  With rsEntradas
    txtTotProdutos.Text = gsFormatCurrency(.fields("Produtos"), True)
    txtTotDescontos.Text = gsFormatCurrency(.fields("Desconto"), True)
    txtTotIPI.Text = gsFormatCurrency(.fields("IPI"), True)
    txtFrete.Text = gsFormatCurrency(.fields("Frete"), True)
    txtTotBaseICMS.Text = gsFormatCurrency(.fields("Base ICM"), True)
    txtTotValICMS.Text = gsFormatCurrency(.fields("Valor ICM"), True)
    txtTotBaseICMSSubst.Text = gsFormatCurrency(.fields("Base ICM Subs"), True)
    txtTotValICMSSubst.Text = gsFormatCurrency(.fields("Valor ICM Subs"), True)
    txtTotalAPagar.Text = gsFormatCurrency(.fields("Total"), True)
    If Len(.fields("Caixa") & "") > 0 Then
      cboCaixaUso.Text = .fields("Caixa")
      cboCaixaUso_LostFocus
    End If
    txtCxDinheiro.Text = gsFormatCurrency(.fields("Dinheiro Caixa"), True)
    txtCxCheque.Text = gsFormatCurrency(.fields("Cheque Caixa"), True)
    
    txtTotICMSDesonerado.Text = gsHandleNull(.fields("TotalDesoneracaoICMS"))
    
    '08/01/2010 - Andrea
    txtTroco.Text = gsFormatCurrency((.fields("Troco") * -1), True)
    
    cboConta.Text = .fields("Conta")
    txtCheque.Text = .fields("Num Cheque") & ""
    txtDescricao.Text = .fields("Descrição") & ""
    txtValCheque.Text = gsFormatCurrency(.fields("Valor Cheque"), True)
    If IsDate(.fields("Bom Para")) Then
      medDataBomPara.Text = .fields("Bom Para")
    End If
    
    '15/01/2010 - Andrea
    '-------------------------------------------------------------------------------------------------
    txtNumeroDI.Text = gsHandleNull(.fields("NumeroDI") & "")
    txtCodigoExportadorDI.Text = gsHandleNull(.fields("CodigoExportador") & "")
    txtUFDesembaracoDI.Text = gsHandleNull(.fields("UFDesembaracoDI") & "")
    txtLocalDesembaracoDI.Text = gsHandleNull(.fields("LocalDesembaracoDI") & "")
    txtNumeroAdicao.Text = gsHandleNull(.fields("NumeroAdicaoDI"))
    txtNumeroSequenciaItem.Text = gsHandleNull(.fields("NumeroSeqItemAdicaoDI"))
    txtCodigoFabricante.Text = gsHandleNull(.fields("CodigoFabricanteAdicaoDI") & "")
    txtDescontoItemAdicao.Text = gsHandleNull(.fields("DescontoAdicaoDI"))
    
    If IsDate(.fields("DataDeRegistroDI")) Then
      medDataRegistroDI.Text = .fields("DataDeRegistroDI")
    End If
    
    If IsDate(.fields("DataDesembaracoDI")) Then
      medDataDesembaracoDI.Text = .fields("DataDesembaracoDI")
    End If
    '-------------------------------------------------------------------------------------------------

    
  End With

  Call LoadGridCP
  
  Call LoadGridCheques
  
  Call RecalculaPagar
  
  Num_Registro = rsEntradas.Bookmark

  Exit Sub

ErrMostra:
  gsTitle = LoadResString(201)
  gsMsg = "Erro na Apresentação de Movimentação de Entrada."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Public Sub SearchRecord_porSequencia(lngSequencia As Long)
On Error GoTo ErrHandler
  
  gsWhere = ""
  gsWhere = gsWhere & " AND Sequência >= " & lngSequencia
  
  Set rsEntradas = db.OpenRecordset(gsSql & gsWhere & gsOrder, dbOpenDynaset)
  If Not rsEntradas.EOF Then
    Call MostraRegistro
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em função dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Sub

'27/03/2006 - mpdea
'Implementado tratamento de erro
Public Sub SearchRecord()
  Dim lngSequencia As Long

  
  On Error GoTo ErrHandler
  
  
  If Not IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Apague todos os campos da tela com o botão NOVO."
    gsMsg = gsMsg & vbCrLf & "Preencha para a pesquisa uma ou mais das seguintes informações:"
    gsMsg = gsMsg & vbCrLf & "Operação, Digitador, Fornecedor, Seqüência, Nota Fiscal, Data de Emissão"
    gsMsg = gsMsg & vbCrLf & "E pressione novamente este botão PROCURAR."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  gsWhere = ""
  If Len(cboOper.Text) > 0 Then
    gsWhere = gsWhere & " AND Operação = " & cboOper.Text
  End If
  
  If Len(cboDigitador.Text) > 0 Then
    gsWhere = gsWhere & " AND Digitador = " & cboDigitador.Text
  End If
  
  If Len(cboFornecedor.Text) > 0 Then
    gsWhere = gsWhere & " AND Fornecedor = " & cboFornecedor.Text
  End If
  
  If Len(txtNF.Text) > 0 Then
    gsWhere = gsWhere & " AND [Nota Fiscal] Like '" & txtNF.Text & "*'"
  End If
    
  If Len(txtSeq.Text) > 0 Then
    '27/03/2006 - mpdea
    'Implementado validação de dados
    If Not IsDataType(dtLong, txtSeq.Text, lngSequencia) Then
      DisplayMsg "Número de sequência para pesquisa inválida."
      Exit Sub
    End If
    gsWhere = gsWhere & " AND Sequência >= " & lngSequencia
  End If
  
  If IsDate(medDataEmissao.Text) Then
    gsWhere = gsWhere & " AND [Data Emissão] >= #" & medDataEmissao.Text & "#"
  End If
  
  Set rsEntradas = db.OpenRecordset(gsSql & gsWhere & gsOrder, dbOpenDynaset)
  If Not rsEntradas.EOF Then
    Call MostraRegistro
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em função dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub SearchProdutos()
  Dim F As Form
  Call StatusMsg("Aguarde....")
  gsCodProduto = grdItens.Columns("Codigo").Text
  '31/08/2006 - Anderson
  'Implementação de pesquisa avançada na tela de consulta do produto
  'Set F = New frmConsultaProd
  'F.Show vbModal
'  frmConsultaProd.Show
  frmPesquisaProduto.Show
  'Set F = Nothing
  Call StatusMsg("")
End Sub

Private Sub LoadGridItens()
  Dim rsEntradaProd As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim Aux_Prod As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim sDescricao As String
  Dim sUnidVenda As String
  Dim sText As String
  Dim nRows As Long
  Dim nRow As Long
  On Error GoTo 0
  
  Screen.MousePointer = vbHourglass
  
  grdItens.Redraw = False
  
  bAllow = grdItens.AllowAddNew
  grdItens.AllowAddNew = True
  grdItens.AllowUpdate = True
  grdItens.RemoveAll
  
  Set rsEntradaProd = db.OpenRecordset("SELECT * FROM [Entradas - Produtos] WHERE Filial = " & CStr(gnCodFilial) & " AND Sequência = " & rsEntradas("Sequência") & " ORDER BY Linha", dbOpenDynaset)

  If Not rsEntradaProd.EOF Then
    With rsEntradaProd
      .MoveLast
      .MoveFirst
      If .RecordCount = 1 Then
        sText = "&Item"
      Else
        sText = "&Itens"
      End If
      tabItens.TabCaption(0) = CStr(.RecordCount) & " " & sText
      Do While Not .EOF
        sCodProd = .fields("Código").Value
        Call Acha_Produto(sCodProd, Aux_Prod, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Tipo, Aux_Erro)
        If Aux_Erro = 0 Then
          If Aux_Tipo = 1 Or Aux_Tipo = 2 Then
            Call gsGetDescProd(Aux_Prod, sDescricao, sUnidVenda)
          Else
            Call gsGetDescProd(sCodProd, sDescricao, sUnidVenda)
          End If
          '24/08/2004 - Daniel
          'Adicionado a linha: "1" & vbTab &
          If m_blnIndice Then
            sRecord = sCodProd & vbTab & _
                .fields("Qtde").Value & vbTab & _
                .fields("IndiceFinanceiro").Value & vbTab & _
                sDescricao & vbTab & _
                sUnidVenda & vbTab & _
                .fields("Preço").Value & vbTab & _
                (.fields("Qtde").Value * .fields("Preço").Value) & vbTab & _
                .fields("Desconto").Value & vbTab & _
                .fields("ICM").Value & vbTab & _
                .fields("IPI").Value & vbTab & _
                .fields("Preço Final").Value & vbTab & _
                .fields("Etiqueta").Value
                '.Fields("Código sem Grade").Value
          Else
            sRecord = sCodProd & vbTab & _
                .fields("Qtde").Value & vbTab & _
                "1" & vbTab & _
                sDescricao & vbTab & _
                sUnidVenda & vbTab & _
                .fields("Preço").Value & vbTab & _
                (.fields("Qtde").Value * .fields("Preço").Value) & vbTab & _
                .fields("Desconto").Value & vbTab & _
                .fields("ICM").Value & vbTab & _
                .fields("IPI").Value & vbTab & _
                .fields("Preço Final").Value & vbTab & _
                .fields("Etiqueta").Value
                '.Fields("Código sem Grade").Value
          End If
          
          grdItens.AddItem sRecord
        End If
        .MoveNext
      Loop
      nRows = 11 - grdItens.Rows
      For nRow = 0 To nRows
        grdItens.AddItem ""
      Next nRow
      .MoveFirst
    End With
    'grdItens.Scroll -99, -99
  Else
    tabItens.TabCaption(0) = "&Itens"
  End If

  'grdItens.MoveLast
  'grdItens.MoveFirst
  
  grdItens.AllowAddNew = bAllow
  grdItens.AllowUpdate = bAllow

  grdItens.Redraw = True
  
  rsEntradaProd.Close
  Set rsEntradaProd = Nothing
  
  Screen.MousePointer = vbDefault

End Sub

'07/01/2010 - Andrea
Private Sub LoadGridCheques()

 Dim rsMoviCheques As Recordset

 Dim Erro As Integer
 Dim Ordem As Integer
 Dim sRecord As String
 Dim strSQL As String
 
 '11/12/2009 - Andrea
 '--------------------------------------------------------------------------------------------
 Grade_ChequesEmCaixa.RemoveAll
 Ordem = 0
 Erro = False
 
 strSQL = "SELECT * "
 strSQL = strSQL & "FROM [Movimento - Cheques] WHERE [Movimento - Cheques].Filial = " & gnCodFilial & "  AND "
 strSQL = strSQL & "[Movimento - Cheques].Sequência = " & txtSeq.Text & " ORDER BY [Movimento - Cheques].Ordem "
 Set rsMoviCheques = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
 With rsMoviCheques
   If Not (.BOF And .EOF) Then
     Do Until .EOF
      '11/01/2010 - mpdea
      'Formatação do valor
       sRecord = rsMoviCheques.fields("Cheque").Value & vbTab & _
            Format(rsMoviCheques.fields("Valor").Value, FORMAT_VALUE) & vbTab & _
            rsMoviCheques.fields("Bom").Value & vbTab & _
            rsMoviCheques.fields("Banco").Value
      
       Grade_ChequesEmCaixa.AddItem sRecord
     
     .MoveNext
      
    Loop
   End If
   .Close
 End With
 Set rsMoviCheques = Nothing

End Sub

Private Sub LoadGridCP()
  Dim rsMoviParcelas As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim Aux_Prod As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim sDescricao As String
  Dim sUnidVenda As String
  
  On Error GoTo 0
  
  bAllow = grdCP.AllowAddNew
  grdCP.AllowAddNew = True
  grdCP.AllowUpdate = True
  grdCP.RemoveAll
  
  Set rsMoviParcelas = db.OpenRecordset("SELECT * FROM [Movimento - Parcelas] WHERE Filial = " & CStr(gnCodFilial) & " AND Sequência = " & rsEntradas("Sequência") & " ORDER BY Ordem", dbOpenDynaset)

  If Not rsMoviParcelas.EOF Then
    grdCP.Redraw = False
    With rsMoviParcelas
      .MoveFirst
      Do While Not .EOF
        sRecord = .fields("Bom") & vbTab & _
          .fields("Valor")
        grdCP.AddItem sRecord
        .MoveNext
      Loop
      .MoveFirst
    End With
    grdCP.Scroll -99, -99
    grdCP.Redraw = True
  End If

  grdCP.AllowAddNew = bAllow
  grdCP.AllowUpdate = bAllow

  rsMoviParcelas.Close
  Set rsMoviParcelas = Nothing

End Sub

Private Function gbWriteGridItens() As Boolean
  Dim rsEntradaProd As Recordset
  Dim nRow As Long
  Dim sSql As String
  Dim bm As Variant
  Dim sCodProd As String
  Dim Cód As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Tipo As Integer
  Dim Erro As Integer
  '24/08/2004 - Daniel
  Dim dblIndice As Double
  Dim dblCUSTO  As Double

  Dim intRepeatUpdateLocked As Integer
  Dim blnInTransaction As Boolean
  
  gbWriteGridItens = False
  
  sSql = "SELECT * FROM [Entradas - Produtos] "
  sSql = sSql & " WHERE Filial = " & CStr(gnCodFilial) & " AND Sequência = " & txtSeq.Text
  Set rsEntradaProd = db.OpenRecordset(sSql, dbOpenDynaset)
  
  On Error GoTo ErrTrans
  
  Call ws.BeginTrans
  blnInTransaction = True
  
  With rsEntradaProd
    '
    If Not .EOF Then
    
      .MoveFirst
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
      
    End If

    For nRow = 0 To grdItens.Rows - 1
      '
      bm = grdItens.AddItemBookmark(nRow)
      sCodProd = grdItens.Columns("Codigo").CellText(bm)
      If Len(sCodProd) > 0 Then
        .AddNew
        .fields("Filial").Value = gnCodFilial
        .fields("Sequência").Value = CLng(gsHandleNull(txtSeq.Text))
        .fields("Linha").Value = nRow + 1
        .fields("Código").Value = sCodProd
        .fields("Qtde").Value = CSng(gsHandleNull(grdItens.Columns("Qtde").CellText(bm)))
        '14/09/2004 - Daniel
        'Case Resultado: Adicionado campos QtdeAtual e EntradaConsignada
        If m_blnResultado Then
          .fields("QtdeAtual").Value = CSng(gsHandleNull(grdItens.Columns("Qtde").CellText(bm)))
          
          If m_blnEmprestimo Then
            .fields("EntradaConsignada").Value = True
          Else
            .fields("EntradaConsignada").Value = False
          End If
          
        Else
          .fields("QtdeAtual").Value = 0
          .fields("EntradaConsignada").Value = False
        End If
        '05/05/2004 - Daniel
        'Personalização Embalavi
        If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
          .fields("Preço").Value = CSng(Format((gsHandleNull(grdItens.Columns("Preco").CellText(bm))), "##,###,##0.00000"))
          '26/08/2004 - Daniel
          'Tratamento para Indice Financeiro
          If (gsHandleNull(grdItens.Columns("Qtde").CellText(bm))) > 1 Then
            dblCUSTO = CSng(gsHandleNull(grdItens.Columns("PrecoFinal").CellText(bm))) / (gsHandleNull(grdItens.Columns("Qtde").CellText(bm)))
          Else 'Qtde = 1
            dblCUSTO = CSng(gsHandleNull(grdItens.Columns("PrecoFinal").CellText(bm)))
          End If
        
          dblCUSTO = Format(dblCUSTO, "##,###,##0.00000")
        '30/04/2007 - Anderson - Implementação de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
          .fields("Preço").Value = CSng(Format((gsHandleNull(grdItens.Columns("Preco").CellText(bm))), "##,###,##0.000"))
          'Tratamento para Indice Financeiro
          If (gsHandleNull(grdItens.Columns("Qtde").CellText(bm))) > 1 Then
            dblCUSTO = CSng(gsHandleNull(grdItens.Columns("PrecoFinal").CellText(bm))) / (gsHandleNull(grdItens.Columns("Qtde").CellText(bm)))
          Else 'Qtde = 1
            dblCUSTO = CSng(gsHandleNull(grdItens.Columns("PrecoFinal").CellText(bm)))
          End If
        
          dblCUSTO = Format(dblCUSTO, "##,###,##0.000")
        Else
          .fields("Preço").Value = CSng(gsHandleNull(grdItens.Columns("Preco").CellText(bm)))
          '26/08/2004 - Daniel
          'Tratamento para Indice Financeiro
          If (gsHandleNull(grdItens.Columns("Qtde").CellText(bm))) > 1 Then 'Deve-se pegar sempre o valor unitário
            dblCUSTO = CSng(gsHandleNull(grdItens.Columns("PrecoFinal").CellText(bm))) / (gsHandleNull(grdItens.Columns("Qtde").CellText(bm)))
          Else 'Qtde = 1
            dblCUSTO = CSng(gsHandleNull(grdItens.Columns("PrecoFinal").CellText(bm)))
          End If
        End If
        
        .fields("Desconto").Value = CSng(gsHandleNull(grdItens.Columns("Desconto").CellText(bm)))
        .fields("ICM").Value = CSng(gsHandleNull(grdItens.Columns("ICMS").CellText(bm)))
        .fields("IPI").Value = CSng(gsHandleNull(grdItens.Columns("IPI").CellText(bm)))
        .fields("Preço Final").Value = CSng(gsHandleNull(grdItens.Columns("PrecoFinal").CellText(bm)))
        
        If IsNull(grdItens.Columns("Etiqueta").CellText(bm)) Then
          .fields("Etiqueta").Value = False
        Else
          If grdItens.Columns("Etiqueta").CellText(bm) = "" Then
            .fields("Etiqueta").Value = False
          Else
            .fields("Etiqueta").Value = grdItens.Columns("Etiqueta").CellText(bm)
          End If
        End If
        
        .fields("Código Sem Grade") = ""
        Call Acha_Produto(sCodProd, Cód, Tamanho, Cor, Edição, Tipo, Erro)
        If Erro = 0 Then
          .fields("Código Sem Grade") = Cód
        End If
        
        '24/08/2004 - Daniel
        'Adicionado Campo IndiceFinanceiro
        If m_blnIndice Then
          .fields("IndiceFinanceiro").Value = CDbl(gsHandleNull(grdItens.Columns("Indice").CellText(bm)))
          dblIndice = CDbl(gsHandleNull(grdItens.Columns("Indice").CellText(bm)))
        Else
          .fields("IndiceFinanceiro").Value = 0
        End If
        
        
        '17/05/2006 - mpdea
        'Armazenda o valor do ICMS Retido
        .fields("ValorIcmsRetido").Value = CDbl(gsHandleNull(grdItens.Columns("ValorIcmsRetido").CellText(bm)))
        
        .fields("ValorICMSDesonerado").Value = CDbl(gsHandleNull(grdItens.Columns("Valor Desonerado").CellText(bm)))
        .fields("Percentual_Diferimento") = CDbl(gsHandleNull(grdItens.Columns("% Diferimento").CellText(bm)))
        
        .Update
        
        '24/08/2004 - Daniel
        'Private para atualização do Preço de Venda
        'Passamos como parâmetro o Código do Produto, o valor do Índice e a Operação de Entrada
        If m_blnIndice Then Call AtualizarPrecoDeVenda(sCodProd, dblIndice, cboOper.Text, dblCUSTO)
        
    End If
      '
    Next nRow
  '
  End With
  
  '22/09/2004 - Daniel
  'Case: Resultado
  'Variável modular m_blnEmprestimo volta ao estado inicial 'False'
  m_blnEmprestimo = False

  Call ws.CommitTrans
  blnInTransaction = False
  
  rsEntradaProd.Close
  Set rsEntradaProd = Nothing
  
  gbWriteGridItens = True
  
  Exit Function
  
ErrTrans:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
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
          Exit Function
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
          Exit Function
        Case 3 'Encerrar
          End
      End Select
  End Select
End Function

'Private Sub WriteGridCP()
'  Dim sSql As String
'  Dim bm As Variant
'  Dim nRow As Long
'
'  On Error GoTo ErrHandler
'
'  grdCP.Update
'
'  sSql = "DELETE * FROM [Movimento - Parcelas] "
'  sSql = sSql & " WHERE Filial = " & CStr(gnCodFilial) & " AND Sequência = " & CStr(txtSeq.Text)
'  Call db.Execute(sSql, dbFailOnError)
'
'  For nRow = 0 To grdCP.Rows - 1
'    bm = grdCP.AddItemBookmark(nRow)
'    If IsDate(grdCP.Columns("Data").CellText(bm)) Then
'      If IsNumeric(grdCP.Columns("Valor").CellValue(bm)) Then
'        With rsMovi_Parcelas
'          .AddNew
'          .Fields("Filial") = gnCodFilial
'          .Fields("Sequência") = Val(txtSeq.Text)
'          .Fields("Ordem") = nRow + 1
'          .Fields("Bom") = grdCP.Columns("Data").CellText(bm)
'          .Fields("Valor") = grdCP.Columns("Valor").CellValue(bm)
'          .Update
'        End With
'      End If
'    End If
'  Next nRow
'
'  Exit Sub
'
'ErrHandler:
'  gsTitle = LoadResString(201)
'  gsMsg = "Erro ao Atualizar Pagamentos."
'  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
'  gnStyle = vbOKOnly + vbExclamation
'  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'  Exit Sub
'
'End Sub

Private Sub CalculaLinha(ByVal bm As Variant)
  Dim nPrecoTotal As Double
  Dim nPrecoFinal As Double
  Dim nQtde As Single
  Dim nPreco As Single
  Dim nDesconto As Single
  Dim nValorDesconto As Single
  Dim nIPI As Single
  Dim nValorIPI As Single
  
  nQtde = grdItens.Columns("Qtde").CellText(bm)
  '05/05/2004 - Daniel
  'Personalização Embalavi
  If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
    nPreco = Format((grdItens.Columns("Preco").CellText(bm)), "##,###,##0.00000")
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    nPreco = Format((grdItens.Columns("Preco").CellText(bm)), "##,###,##0.000")
  Else
    nPreco = grdItens.Columns("Preco").CellText(bm)
  End If
  
  nDesconto = gsHandleNull(grdItens.Columns("Desconto").CellText(bm) & "")
  nIPI = grdItens.Columns("IPI").CellText(bm)
  
  nPrecoTotal = gsFormatCurrency((nQtde * nPreco), True)
  grdItens.Columns("PrecoTotal").Text = nPrecoTotal
  
  nValorDesconto = nPrecoTotal * nDesconto / 100 'Round((nPrecoTotal * nDesconto / 100), 2)
  nPrecoFinal = (nPrecoTotal - nValorDesconto)
  nValorIPI = nPrecoFinal * nIPI / 100 'Round((nPrecoFinal * nIPI / 100), 2)
  nPrecoFinal = nPrecoFinal + nValorIPI
  
  grdItens.Columns("PrecoFinal").Text = nPrecoFinal
End Sub

'17/05/2006 - mpdea
'Incluído tratamento de erro
'bln_update_icms_retido => Se True atualiza no grid os valores de ICMS Retido
Public Sub Recalcula(Optional ByVal bln_update_icms_retido As Boolean = False)
  Dim nI As Integer
  Dim bm As Variant
  Dim nRow As Long
  Dim nTotProd As Double
  Dim nTotDesc As Double
  Dim nTotIPI As Double
  Dim nTotPagar As Double
  Dim nValDesc As Double
  Dim nValIPI As Double
  Dim nTemp As Double
'  Dim nBaseICMS As Currency
'  Dim nValorICMS As Currency
'  Dim nIndICMS As Integer
'  Dim ICMSInd() As Single
'  Dim ICMSVal() As Currency
  Dim nPrecoTotal As Double
  Dim nDesconto As Single
  Dim nIPI As Single
  Dim nICMS As Single
  
  Dim nQtde As Single
  Dim cValorBaseUnit As Double
  Dim nPercReducao As Single
  Dim cAux As Double
  
  Dim cBaseICMS As Double
  Dim cValorICMS As Double
  Dim cBaseICMSSubs As Double
  Dim cValorICMSSubs As Double
  
  Dim Tot_Desoneracao As Double
  Dim Temp_Desoneracao As Double
  Dim ValorTotalDesoneracao As Double
  
  Dim iQtdeItens As Integer '30/11/2006 - Anderson - Variável criada para contabilizar a quantidade de itens da nota para fazer o rateio dos produtos
  Dim dTotalNotaIPI As Double '04/12/2006 - Anderson - Variável auxiliar para contabilizar o valor do IPI no total da nota
  
  '-----------------------------------------------------------------------------
  '16/09/2005 - mpdea
  'Índice para cálculo do Preço de Entrada
  Dim dblIndicePrecoEntrada As Double
  'Preço sem Índice para cálculo do Preço de Entrada
  Dim dblPrecoSemIndiceEntrada As Double
  'Preço Final sem Índice para cálculo do Preço de Entrada
  Dim dblPrecoFinalSemIndiceEntrada As Double
  'Quantidade
  Dim sngQuantidade As Single
  '-----------------------------------------------------------------------------
  
  '-------------------------------------------
  '16/05/2006 - mpdea
  'Frete
  Dim sngPercFrete As Single
  Dim dblValorFrete As Double
  Dim dblValorTotal As Double
  Dim dblValorIPI As Double
  '
  'Preço para cálculo de Custo
  Dim dblPrecoFinalCusto As Double
  '
  Static bln_2X As Boolean
  '-------------------------------------------
  
  '17/05/2006 - mpdea
  'Índice ICMS Retido - Base complementar
  Dim dblIndiceIcmsRetido As Double
  '
  'Alíquota do fornecedor (se for do mesmo
  'estado é utilizada a do produto)
  Dim intAliquotaFornecedor As Integer
  
  '18/05/2006 - mpdea
  'ICMS de Saída
  Dim intIcmsSaida As Integer
  'Valor do ICMS Retido
  Dim dblValorIcmsRetido As Double
  
  
  On Error GoTo ErrHandler
  
  Tot_Desoneracao = 0#
  ValorTotalDesoneracao = 0#
  
  '16/05/2006 - mpdea
  'É necessário a chamada em duplicidade da função Recalcula para
  'o correto cálculo de impostos
  bln_2X = Not bln_2X
  If bln_2X Then Recalcula
  
  
  'Caso esteja com altera totais pressionado
  'não executa o recálculo dos totais
  If ActiveBar1.Tools("miComplAlteraTotais").Checked Then Exit Sub
  
  If lblEfetivada.Visible Then
    Exit Sub
  End If
  
  
  gbLiberaPagar = (rsOp_Entrada("Tipo") = "C" Or rsOp_Entrada("Tipo") = "D")
  
  'grdItens.Redraw = False
  
  
  '--------------------------------------------------------------------------------
  '16/05/2006 - mpdea
  'Percentual de Frete para cálculos de custo
  If rsOp_Entrada.fields("SomarFreteCustoProduto").Value Then
    'Valor do frete
    Call IsDataType(dtDouble, Me.txtFrete.Text, dblValorFrete)
    'Valor total
    Call IsDataType(dtDouble, Me.txtTotalAPagar.Text, dblValorTotal)
    'Verifica se o frete soma no total
    If rsOp_Entrada.fields("Somar Frete ao Total").Value Then
      dblValorTotal = dblValorTotal - dblValorFrete
    End If
    'Verifica se não soma o IPI a Base de ICMS
    If Not rsOp_Entrada.fields("Base ICM com IPI").Value Then
      Call IsDataType(dtDouble, Me.txtTotIPI.Text, dblValorIPI)
      dblValorTotal = dblValorTotal - dblValorIPI
    End If
    'Percentual de frete
    If dblValorTotal > 0 Then
      sngPercFrete = dblValorFrete / dblValorTotal
    End If
  End If
  '--------------------------------------------------------------------------------
  
  
  nTotProd = 0
  nTotDesc = 0
  nTotIPI = 0
  iQtdeItens = 0
'  nIndICMS = -1
  
  
  '17/06/2005 - mpdea
  'Atualiza no grid os valores de ICMS Retido
  If bln_update_icms_retido Then
    grdItens.MoveFirst
    grdItens.Redraw = False
  End If
  
  
  For nRow = 0 To grdItens.Rows - 1
    bm = grdItens.AddItemBookmark(nRow)
    
    Temp_Desoneracao = CDbl(gsHandleNull(grdItens.Columns("Valor Desonerado").CellValue(bm)))
    Tot_Desoneracao = Tot_Desoneracao + Temp_Desoneracao
    'ValorTotalDesoneracao = ValorTotalDesoneracao + grdItens.Columns("Valor Desonerado").CellValue(bm)
    
    '-----------------------------------------------------------------------------
    '19/09/2005 - mpdea
    'Índice para cálculo do Preço de Entrada
    If g_blnIndicePrecoEntrada Then
      'Obtém o Índice para cálculo do Preço de Entrada
      Call IsDataType(dtDouble, grdItens.Columns("IndicePrecoEntrada").CellValue(bm), dblIndicePrecoEntrada)
      'Preço sem Índice para cálculo do Preço de Entrada
      Call IsDataType(dtDouble, grdItens.Columns("PrecoSemIndiceEntrada").CellValue(bm), dblPrecoSemIndiceEntrada)
      'Quantidade
      Call IsDataType(dtSingle, grdItens.Columns("Qtde").CellValue(bm), sngQuantidade)
     End If
    '-----------------------------------------------------------------------------
    
    nPrecoTotal = CDbl(gsHandleNull(grdItens.Columns("PrecoTotal").CellValue(bm)))
    
    '05/05/2004 - Daniel
    'Personalização Embalavi
    If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
      nTotProd = Format((nTotProd + nPrecoTotal), "##,###,##0.00000")
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      nTotProd = Format((nTotProd + nPrecoTotal), "##,###,##0.000")
    Else
      nTotProd = nTotProd + nPrecoTotal
    End If
        
    nDesconto = CSng(gsHandleNull(grdItens.Columns("Desconto").CellValue(bm)))
    nValDesc = nPrecoTotal * nDesconto / 100 'Round(nPrecoTotal * nDesconto / 100, 2)
    nTotDesc = nTotDesc + nValDesc
    
    nICMS = CSng(gsHandleNull(grdItens.Columns("ICMS").CellValue(bm)))
    nIPI = CSng(gsHandleNull(grdItens.Columns("IPI").CellValue(bm)))
    
    nTemp = Format(nPrecoTotal - nValDesc, FORMAT_VALUE)
    
    '-----------------------------------------------------------------------------
    '19/09/2005 - mpdea
    'Índice para cálculo do Preço de Entrada
    'Valor do IPI é calculado sobre o Preço Unitário
    'sem ter indexado pelo Índice para cálculo do Preço de Entrada
    If g_blnIndicePrecoEntrada Then
      nTemp = sngQuantidade * dblPrecoSemIndiceEntrada
      nValDesc = (nTemp * nDesconto / 100)
      dblPrecoFinalSemIndiceEntrada = (nTemp - nValDesc)
      nValIPI = (dblPrecoFinalSemIndiceEntrada * nIPI / 100)
    Else
      nValIPI = Format(nTemp * (nIPI / 100), "#0.00") 'Round(nTemp * nIPI / 100, 2)
    End If
    '-----------------------------------------------------------------------------
    
    If gbIPI Then
       nTotIPI = nTotIPI + nValIPI
    Else
       nTotIPI = 0
    End If
'    If nICMS <> 0# Then
'      nIndICMS = nIndICMS + 1
'      ReDim Preserve ICMSInd(nIndICMS), ICMSVal(nIndICMS)
'      ICMSInd(nIndICMS) = nICMS
'      ICMSVal(nIndICMS) = ICMSVal(nIndICMS) + nTemp
'    End If
    
'    If nICMS <> 0# Then
'      nBaseICMS = nBaseICMS + nTemp
'      nValorICMS = nValorICMS + (nTemp * nICMS / 100#)
'    End If
    
        
      
    '16/05/2006 - mpdea
    'Corrigido cálculos de Base de Cálculo que estavam utilizando
    'tanto a base por valor quanto a base por percentual
    '(ICMS Retido e ICMS com Base Reduzida)
    With grdItens
      nQtde = CSng(gsHandleNull(.Columns("Qtde").CellValue(bm)))
      cValorBaseUnit = CDbl(gsHandleNull(.Columns("Valor_Base_Unit").CellValue(bm)))
      nPercReducao = CSng(gsHandleNull(.Columns("Redução_ICM").CellValue(bm)))
      
      '18/06/2006 - mpdea
      'ICMS de Saída
      Call IsDataType(dtDouble, grdItens.Columns("IcmsSaida").CellValue(bm), intIcmsSaida)
    
     If .Columns("ICMS").CellValue(bm) <> "" Then
      iQtdeItens = iQtdeItens + 1
     End If
    
      'Calculo do ICM
      
      Select Case UCase(.Columns("Tipo_ICM").CellValue(bm))
        Case "N"
        
          'ICM Normal
          If gbBaseICMSomadoIPI Then
            cBaseICMS = cBaseICMS + nTemp + nValIPI
            cValorICMS = cValorICMS + ((nTemp + nValIPI) * nICMS / 100)
          Else
            cBaseICMS = cBaseICMS + nTemp
            cValorICMS = cValorICMS + (nTemp * nICMS / 100)
          End If
          If rsOp_Entrada("Tipo") = "T" And rsOp_Entrada("ICM").Value = False Then
            cBaseICMS = 0
            cValorICMS = 0
          End If
          
        Case "R" 'ICM Retido
          '01/02/2007 - Anderson - A variavel está sendo zerada porque estava acumulando o valor e somando ao total da base de ICMS
          cAux = 0
          If cValorBaseUnit <> 0 Then
            'Base Fixa
            cAux = nQtde * cValorBaseUnit
          ElseIf nPercReducao <> 0 Then
            'Base Reduzida
            
            '17/03/2009 - mpdea
            'Retornado calculo anterior, pois gerava valores incorretos
            '10/10/2007 - Anderson
            'Retirado condição, pois ICMS retido deve sempre somar o valor do frete
            '16/05/2006 - mpdea
            'Soma IPI a base de cálculo
            If gbBaseICMSomadoIPI Then
              cAux = (nTemp + nValIPI)
            Else
              cAux = nTemp
            End If
            
            'Realiza a redução da base de cálculo
            cAux = cAux * nPercReducao / 100
          End If
          
          '16/05/2006 - mpdea
          'Soma frete a base de cálculo
          If rsOp_Entrada.fields("SomarFreteCustoProduto").Value Then
            cAux = cAux * (1 + sngPercFrete)
          End If
            
          '17/05/2006 - mpdea
          'Índice ICMS Retido - Base complementar
          Call IsDataType(dtDouble, .Columns("IndiceIcmsRetido").CellValue(bm), dblIndiceIcmsRetido)
          cAux = cAux * dblIndiceIcmsRetido
          
          'Alíquota do fornecedor
          intAliquotaFornecedor = m_intGetAliquotaIcmsEstado()
          If intAliquotaFornecedor = -1 Then intAliquotaFornecedor = intIcmsSaida
          
          '17/03/2009 - mpdea
          'Retornado calculo anterior, pois gerava valores incorretos
          '10/10/2007 - Anderson
          'Alteração realizada para calcular o valor da Base de ICMS por Substituição
          'Base acumulada
          cBaseICMSSubs = cBaseICMSSubs + cAux
'          cBaseICMSSubs = cBaseICMSSubs + (nTemp + nValIPI) + cAux
          
          '17/03/2009 - mpdea
          'Retornado calculo anterior, pois gerava valores incorretos
          '10/10/2007 - Anderson
          'Alteração realizada para calcular o valor da Base de ICMS por Substituição
          'Valor do imposto
          dblValorIcmsRetido = Format((cAux * intIcmsSaida / 100) - (nTemp * intAliquotaFornecedor / 100), "#0.00")
'          dblValorIcmsRetido = Format((cBaseICMSSubs * intIcmsSaida / 100) - (nTemp * intAliquotaFornecedor / 100), "#0.00")
          
          'Valor do imposto acumulado
          cValorICMSSubs = cValorICMSSubs + dblValorIcmsRetido
          
          '17/06/2005 - mpdea
          'Valor do imposto de ICMS Retido para o produto
          If bln_update_icms_retido Then
            .Columns("ValorIcmsRetido").Value = dblValorIcmsRetido
          End If
          
        Case "Z"
        
          'ICM Reduzido
          If cValorBaseUnit <> 0 Then
            'Base Fixa
            cAux = nQtde * cValorBaseUnit
            cBaseICMS = cBaseICMS + cAux
            cValorICMS = cValorICMS + (cAux * nICMS / 100)
          ElseIf nPercReducao <> 0 Then
            'Base Reduzida
            
            '29/11/2004 - Daniel
            'Tratamento especial para a Teknika
            'pegará o valor em cheio do ICM sem reduções
            If Not m_blnTeknika Then 'Demais clientes
              cAux = nTemp * nPercReducao / 100
              cBaseICMS = cBaseICMS + cAux
              cValorICMS = cValorICMS + (cAux * nICMS / 100)
            Else 'Teknika
            
              If gbBaseICMSomadoIPI Then
                cBaseICMS = cBaseICMS + nTemp + nValIPI
                cValorICMS = cValorICMS + ((nTemp + nValIPI) * nICMS / 100)
              Else
                cBaseICMS = cBaseICMS + nTemp
                cValorICMS = cValorICMS + (nTemp * nICMS / 100)
              End If
            End If
          End If
      End Select
    End With
  
    '17/06/2005 - mpdea
    'Move para a próxima linha
    If bln_update_icms_retido Then grdItens.MoveNext
  Next nRow
  
  '17/06/2005 - mpdea
  'Restaura posição e atualização de desenho do grid
  If bln_update_icms_retido Then
    grdItens.MoveFirst
    grdItens.Redraw = True
  End If
    
'  If nIndICMS > -1 Then
'    For nI = LBound(ICMSInd) To UBound(ICMSInd)
'      nBaseICMS = nBaseICMS + ICMSVal(nI)
'      nValorICMS = nValorICMS + (ICMSVal(nI) * ICMSInd(nI) / 100#)
'    Next nI
'  End If
  
  If gbIPI Then
     nTotPagar = nTotProd - nTotDesc + nTotIPI
  Else
     nTotPagar = nTotProd - nTotDesc
  End If
  
  '30/11/2006 - Anderson
  'Alteração realizada para poder calcular o valor do frete no cálculo do ICMS
  'If gbSomarFrete = True Then
  If gbSomarFrete = True Or rsOp_Entrada.fields("BaseICMSFrete").Value Then
    nTotPagar = nTotPagar + CDbl(gsHandleNull(txtFrete.Text))
  End If
  
  txtTotProdutos.Text = gsFormatCurrency(nTotProd, True)
  txtTotDescontos.Text = gsFormatCurrency(nTotDesc, True)
  
  '04/12/2006 - Anderson
  'Variavel auxiliar para contabilizar o valor do IPI no total da nota
  dTotalNotaIPI = 0
  
  '04/12/2006 - Anderson
  ' Verifica se ICMS Incide sobre IPI
  If rsOp_Entrada.fields("ICMSSobreIPI").Value Then
    nTotIPI = 0
    For nRow = 0 To grdItens.Rows - 1
      bm = grdItens.AddItemBookmark(nRow)
      '04/12/2006 - Anderson
      ' Verifica se ICMS Incide sobre IPI
      nTotIPI = nTotIPI + ((CDbl(gsHandleNull(grdItens.Columns("PrecoTotal").CellValue(bm))) + CDbl(gsHandleNull(txtFrete.Text))) * CSng(gsHandleNull(grdItens.Columns("IPI").CellValue(bm))) / 100)
      If grdItens.Columns("IPI").CellValue(bm) <> "" Then
        dTotalNotaIPI = dTotalNotaIPI + (nTotIPI - (CDbl(gsHandleNull(grdItens.Columns("PrecoTotal").CellValue(bm))) * CSng(gsHandleNull(grdItens.Columns("IPI").CellValue(bm))) / 100))
      End If
    Next
  End If
  txtTotIPI.Text = gsFormatCurrency(nTotIPI, True)
  
  '18/05/2006 - mpdea
  'Soma no total a pagar o imposto de ICMS retido
  txtTotalAPagar.Text = gsFormatCurrency(nTotPagar + cValorICMSSubs + dTotalNotaIPI, True)
  
  '30/11/2006 - Anderson - Alteração de acordo com exemplos enviados pelo Rodrigo da Technomax
  '24/11/2006 - Anderson - Aguardando exemplos do Rodrigo Technomax
  '17/11/2006 - Anderson - Solicitação - Technomax
  'Utilizado para somar o valor do frete no calculo de icms para movimentos de entrada.
  If rsOp_Entrada.fields("BaseICMSFrete").Value Then
    For nRow = 0 To grdItens.Rows - 1
      bm = grdItens.AddItemBookmark(nRow)
      If grdItens.Columns("ICMS").CellValue(bm) <> "" Then
       cValorICMS = cValorICMS + Round(((CDbl(gsHandleNull(txtFrete.Text)) / txtTotProdutos.Text) * CDbl(gsHandleNull(grdItens.Columns("PrecoTotal").CellValue(bm)))) * (grdItens.Columns("ICMS").CellValue(bm) / 100), 2)
      End If
    Next
    txtTotBaseICMS.Text = gsFormatCurrency(cBaseICMS + CDbl(gsHandleNull(txtFrete.Text)), True)
    txtTotValICMS.Text = gsFormatCurrency(cValorICMS, True)
  Else
    txtTotBaseICMS.Text = gsFormatCurrency(cBaseICMS, True)
    txtTotValICMS.Text = gsFormatCurrency(cValorICMS, True)
  End If
  
  Me.txtTotBaseICMSSubst.Text = gsFormatCurrency(cBaseICMSSubs, True)
  txtTotValICMSSubst.Text = gsFormatCurrency(cValorICMSSubs, True)
  
  txtTotICMSDesonerado.Text = gsFormatCurrency(Tot_Desoneracao, True)
  
  If gbLiberaPagar Then
    Call RecalculaPagar
  End If
  
  Exit Sub
  
ErrHandler:
  '17/06/2005 - mpdea
  'Restaura posição e atualização de desenho do grid
  If bln_update_icms_retido Then
    grdItens.MoveFirst
    grdItens.Redraw = True
  End If
  MsgBox "Erro no recálculo dos itens: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub RecalculaPagar()
  Dim nRow As Long
  Dim nTotal As Double
  Dim nValorDif As Double
  Dim bm As Variant
  
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
  
  '-----------------------------------------------------------------------------------------------------------
  '07/01/2010 - Andrea
  'Recalcula o valor recebido em cheques do contas a receber
  With Grade_ChequesEmCaixa
    'Verifica ocorrência
    If .Rows > 0 Then
      
      Dim lng_row As Long
      Dim var_book As Variant
      Dim dbl_valor_recebido_cheques As Double
      Dim dbl_valor As Double
      Dim str_numero_cheque As String
      
      dbl_valor_recebido_cheques = 0
      
      For lng_row = 0 To .Rows - 1
          
        var_book = .AddItemBookmark(lng_row)
              
        'Verifica registro informado
        Call IsDataType(dtString, .Columns("NumeroCheque").CellText(var_book), str_numero_cheque)
        If str_numero_cheque <> "" Then
          'Valores
          Call IsDataType(dtDouble, .Columns("Valor").CellText(var_book), dbl_valor)
          
          dbl_valor_recebido_cheques = dbl_valor_recebido_cheques + dbl_valor
        End If
      Next lng_row
      '11/01/2010 - mpdea
      'Formatação do valor
      txtCxCheque.Text = Format(dbl_valor_recebido_cheques, FORMAT_VALUE)
      'nTotal = nTotal + CDbl(dbl_valor_recebido_cheques)
    Else
      '11/01/2010 - mpdea
      'Zera valor
      txtCxCheque.Text = Format(0, FORMAT_VALUE)
    End If
  End With

  'nTotal = Round(nTotal, 2)
   
  nTotal = nTotal + CDbl(gsHandleNull(txtCxDinheiro.Text & ""))
  nTotal = nTotal + CDbl(gsHandleNull(txtCxCheque.Text & ""))
  nTotal = nTotal + CDbl(gsHandleNull(txtValCheque.Text & ""))
   
  txtTotDigitado.Text = gsFormatCurrency(nTotal, True)
     
  nValorDif = CDbl(gsHandleNull(txtTotalAPagar.Text)) - CDbl(gsHandleNull(txtTotDigitado.Text & ""))
  
  '08/01/2010 - Andrea
  txtTroco.Text = 0
  If CDbl(txtCxCheque.Text) > 0 Then 'Tem valores recebido em cheque
    If nValorDif < 0 Then ' e tem diferença, quer dizer que tem troco (que vai entrar em dinheiro no caixa).
      txtTroco.Text = gsFormatCurrency(nValorDif, False)
      nValorDif = 0
    End If
  End If
  
  txtADigitar.Text = gsFormatCurrency(nValorDif, False)

End Sub

Private Sub cmd_imprimir_Click()
  On Error GoTo Erro
  
  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome As String
  Dim strNomeLPT As String
  Dim strPortaLPT As String
  Dim intX As Integer
  Dim i As Integer
  
  If txtSeq.Text <> "" Then
      If Not IsNumeric(txtSeq.Text) Or Trim(cboOper.Text) = "" Then
          MsgBox "Selecione uma Sequência de Entrada válida", vbInformation, "Atenção"
          Exit Sub
      End If
  Else
      MsgBox "Selecione uma Sequência de Entrada válida", vbInformation, "Atenção"
      Exit Sub
  End If
  
  If opt_ticket.Value = True Then
      strNome = "TICKET"
      strNomeLPT = "NOME IMPRESSORA TICKET"
      strPortaLPT = "PORTA IMPRESSORA TICKET"
  Else
      strNome = "REL"
      strNomeLPT = "NOME IMPRESSORA REL"
      strPortaLPT = "PORTA IMPRESSORA REL"
  End If

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

  If strNome = "REL" Then
      Printer.Print "   _________________________________________________________________________________________________________________"
      Printer.Print ""
      Printer.Print "                                                     VALE CRÉDITO"
      Printer.Print ""
      
      If Picture1.Picture <> 0 Then
          Printer.PaintPicture Picture1, 9000, 500, 2300, 1000
      End If
      
      sLinha = "   Emissão Vale : " & sDataAtual
      Printer.Print sLinha
      
      sLinha = "   Empresa      : " & gsNomeFilial
      Printer.Print sLinha
    
      sLinha = "   Atendente    : " & gsUserName
      Printer.Print sLinha
    
      sLinha = "   Sequência    : " & txtSeq.Text
      Printer.Print sLinha
    
      sLinha = "   Operação     : " & cboOper.Text & " " & lblOper.Caption
      Printer.Print sLinha
    
      sLinha = "   Cliente      : " & cboFornecedor.Text & " " & lblFornecedor.Caption
      Printer.Print sLinha
    
      Printer.Print ""
    
      sLinha = "   Código Produto       Nome                                       Cód.Entrada  Núm.itens Valor unitário Valor total"
      Printer.Print sLinha
    
      Printer.Print "   _________________________________________________________________________________________________________________"
      Printer.Print ""
    
      Dim bm As Variant
      Dim nRow As Long
    
      With grdItens
          For nRow = 0 To .Rows - 1
              bm = .AddItemBookmark(nRow)
    
              ' ************************** ATENÇÃO ***********************************
              ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
              ' De preferência com o mesmo nome da impressora !!!
    
              sCodigoProduto = .Columns("Codigo").CellValue(bm)
              If Len(sCodigoProduto) < 20 Then
                For i = Len(sCodigoProduto) To 19
                    sCodigoProduto = " " & sCodigoProduto
                Next
              End If
    
              sNomeProduto = .Columns("Descricao").CellValue(bm)
              If Len(sNomeProduto) < 40 Then
                For i = Len(sNomeProduto) To 39
                    sNomeProduto = sNomeProduto & " "
                Next
              Else
                  sNomeProduto = Mid(sNomeProduto, 1, 40)
              End If
    
              sCodigoEntrada = ""
              If Len(sCodigoEntrada) < 11 Then
                For i = Len(sCodigoEntrada) To 10
                    sCodigoEntrada = " " & sCodigoEntrada
                Next
              End If
    
              sNumItens = .Columns("Qtde").CellValue(bm)
              If Len(sNumItens) < 9 Then
                For i = Len(sNumItens) To 8
                    sNumItens = " " & sNumItens
                Next
              End If
              
              sValorUnitario = Format(.Columns("Preco").CellValue(bm), FORMAT_VALUE)
              If Len(sValorUnitario) < 14 Then
                For i = Len(sValorUnitario) To 13
                    sValorUnitario = " " & sValorUnitario
                Next
              End If
    
              sValorTotal = Format(.Columns("PrecoTotal").CellValue(bm), FORMAT_VALUE)
              If Len(sValorTotal) < 11 Then
                For i = Len(sValorTotal) To 10
                    sValorTotal = " " & sValorTotal
                Next
              End If
    
              sLinha = sCodigoProduto
              sLinha = sLinha & " " & sNomeProduto
              sLinha = sLinha & "   " & sCodigoEntrada
              sLinha = sLinha & "  " & sNumItens
              sLinha = sLinha & " " & sValorUnitario
              sLinha = sLinha & " " & sValorTotal
    
              If Trim(sCodigoProduto) <> "" Then
                  Printer.Print "   " & sLinha
              End If
          Next nRow
      End With
    
      Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      Printer.Print "   TOTAL DO VALE CRÉDITO : " & Format(txtTotalAPagar.Text, FORMAT_VALUE)
      Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      
      If Trim(txtObsValeCredito.Text) <> "" Then
          Printer.Print ""
          Printer.Print ""
          Printer.Print "   OBS: " & Trim(txtObsValeCredito.Text)
          Printer.Print ""
      End If
      
      Printer.Print "                                                      _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ "
      Printer.Print "   Assinatura do Atendente e carimbo da loja         |                                                              |"
      Printer.Print ""
      Printer.Print "                                                     |                                                              |"
      Printer.Print ""
      Printer.Print "   __________________________________________        |                                                              |"
      Printer.Print "   " & gsUserName
      Printer.Print "                                                     |_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ |"
      Printer.Print "   _________________________________________________________________________________________________________________"
    
      Printer.EndDoc
  Else
  
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
      Printer.Print "               VALE CRÉDITO"
      Printer.Print ""
      
      sLinha = "Emissão  : " & sDataAtual
      Printer.Print sLinha
      
      If Len(gsNomeFilial) > 30 Then
          sLinha = "Empresa  : " & Mid(gsNomeFilial, 1, 30)
          Printer.Print sLinha
          Printer.Print Mid(gsNomeFilial, 30, Len(gsNomeFilial) - 30)
      Else
          sLinha = "Empresa  : " & gsNomeFilial
          Printer.Print sLinha
      End If

      If Len(gsUserName) > 30 Then
          sLinha = "Atendente: " & Mid(gsUserName, 1, 30)
          Printer.Print sLinha
          Printer.Print Mid(gsUserName, 30, Len(gsUserName) - 30)
      Else
          sLinha = "Atendente: " & gsUserName
          Printer.Print sLinha
      End If

      sLinha = "Sequência: " & txtSeq.Text
      Printer.Print sLinha
    
      sStrAux = cboOper.Text & " " & lblOper.Caption
      If Len(sStrAux) > 30 Then
          sLinha = "Operação : " & Mid(sStrAux, 1, 30)
          Printer.Print sLinha
          Printer.Print Mid(sStrAux, 30, Len(sStrAux) - 30)
      Else
          sLinha = "Operação : " & sStrAux
          Printer.Print sLinha
      End If
    
      sStrAux = cboFornecedor.Text & " " & lblFornecedor.Caption
      If Len(sStrAux) > 30 Then
          sLinha = "Cliente  : " & Mid(sStrAux, 1, 30)
          Printer.Print sLinha
          Printer.Print Mid(sStrAux, 30, Len(sStrAux) - 30)
      Else
          sLinha = "Cliente  : " & sStrAux
          Printer.Print sLinha
      End If
    
      Printer.Print ""
    
      sLinha = "Produto Entrada Itens VlrUnitário VlrTotal"
      Printer.Print sLinha
    
      Printer.Print "__________________________________________"
      Printer.Print ""
    
'''      Dim bm As Variant
'''      Dim nRow As Long
    
      With grdItens
          For nRow = 0 To .Rows - 1
              bm = .AddItemBookmark(nRow)
    
              ' ************************** ATENÇÃO ***********************************
              ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
              ' De preferência com o mesmo nome da impressora !!!
    
              sCodigoProduto = .Columns("Codigo").CellValue(bm)
              If Len(sCodigoProduto) < 20 Then
                For i = Len(sCodigoProduto) To 19
                    sCodigoProduto = " " & sCodigoProduto
                Next
              End If
    
              sNomeProduto = .Columns("Descricao").CellValue(bm)
              If Len(sNomeProduto) > 42 Then
                  sNomeProduto = Mid(sNomeProduto, 1, 42)
              End If
      
              sCodigoEntrada = ""
              If Len(sCodigoEntrada) < 11 Then
                For i = Len(sCodigoEntrada) To 10
                    sCodigoEntrada = " " & sCodigoEntrada
                Next
              End If
    
              sNumItens = .Columns("Qtde").CellValue(bm)
              If Len(sNumItens) < 9 Then
                For i = Len(sNumItens) To 8
                    sNumItens = " " & sNumItens
                Next
              End If
              
              sValorUnitario = Format(.Columns("Preco").CellValue(bm), FORMAT_VALUE)
              If Len(sValorUnitario) < 14 Then
                For i = Len(sValorUnitario) To 13
                    sValorUnitario = " " & sValorUnitario
                Next
              End If
    
              sValorTotal = Format(.Columns("PrecoTotal").CellValue(bm), FORMAT_VALUE)
              If Len(sValorTotal) < 11 Then
                For i = Len(sValorTotal) To 10
                    sValorTotal = " " & sValorTotal
                Next
              End If
    
              If Trim(sCodigoProduto) <> "" Then
                  sLinha = sCodigoProduto
                  Printer.Print sLinha
                  
                  sLinha = sNomeProduto
                  Printer.Print sLinha
                  
                  sLinha = sCodigoEntrada
                  Printer.Print sLinha
                  
                  sLinha = sNumItens
                  sLinha = sLinha & " " & sValorUnitario
                  sLinha = sLinha & " " & sValorTotal
                  Printer.Print sLinha
              End If
          
          Next nRow
      End With
      
      Printer.Print "------------------------------------------"
      Printer.Print "TOTAL VALE CRÉDITO: " & Format(txtTotalAPagar.Text, FORMAT_VALUE)
      Printer.Print "------------------------------------------"
      
      If Trim(txtObsValeCredito.Text) <> "" Then
          Printer.Print ""
          Printer.Print ""
          
          If Len(Trim(txtObsValeCredito.Text)) > 35 Then
              Printer.Print "OBS: " & Mid(Trim(txtObsValeCredito.Text), 1, 35)
              Printer.Print Mid(Trim(txtObsValeCredito.Text), 36, Len(Trim(txtObsValeCredito.Text)) - 35)
          Else
              Printer.Print "OBS: " & Trim(txtObsValeCredito.Text)
          End If
          Printer.Print ""
      End If
      
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
    
      Printer.EndDoc
  End If
  
  
  Exit Sub
Erro:
    MsgBox "Erro na impressão do Vale " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_marcarEtiquetasParaTodasLinhas_Click()
On Error GoTo Erro

  Dim nRow As Integer

  'grdItens.Redraw = False
  grdItens.Height = 100000

  With grdItens
      For nRow = 0 To .Rows - 1
          .Row = nRow

          If Trim(.Columns("Código").Value) <> "" Then
              .Columns("Etiqueta").Text = "1"
          End If
      Next nRow
  End With
  grdItens.Height = 3165
  'grdItens.Redraw = True
  
  Exit Sub

Erro:
  
  
  MsgBox "Erro ao marcar etiquetas para todos os produtos " & Err.Number & " - " & Err.Description, vbInformation, "Atenção"
End Sub

'11/01/2010 - mpdea
Private Sub ddwCheques_RowLoaded(ByVal Bookmark As Variant)
  ddwCheques.Columns("Valor").Text = Format(ddwCheques.Columns("Valor").Text, FORMAT_VALUE)
End Sub

Private Sub Grade_ChequesEmCaixa_AfterDelete(RtnDispErrMsg As Integer)
  RtnDispErrMsg = False
  Call RecalculaPagar
End Sub

Private Sub Grade_ChequesEmCaixa_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Call RecalculaPagar
End Sub

Private Sub Grade_ChequesEmCaixa_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  If Len(Trim(Grade_ChequesEmCaixa.ActiveCell.Text)) = 0 Then
    If bGridBeforeDelete() = True Then
      Cancel = False
    Else
      Cancel = True
    End If
  Else
    Cancel = True
  End If
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case 0
          gsOrder = " ORDER BY Sequência"
        Case 1
          gsOrder = " ORDER BY Data, Sequência"
        Case 2
          gsOrder = " ORDER BY Fornecedor, Sequência"
        Case 3
          gsOrder = " ORDER BY [Nota Fiscal]"
      End Select
  End Select
  Set rsEntradas = db.OpenRecordset(gsSql & gsWhere & gsOrder, dbOpenDynaset)
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Dim bLocked As Boolean
  
  Select Case Tool.Name
    Case "miOpFirst"
      Call MoveFirst
      
    Case "miOpPrevious"
      Call MovePrevious
      
    Case "miOpNext"
      Call MoveNext
      
    Case "miOpLast"
      Call MoveLast
      
    Case "miOpClear"
      Call ClearScreen
      
    Case "miOpUpdate"
      Call UpdateRecord
      If (txtSeq.Text <> "" And gravada = True) Then
        Call UpdateTotalNCM
      End If
      
    Case "miOpDelete"
      Call DeleteRecord
      
    Case "miOpSearch"
      Call SearchRecord
      
'    Case "miOpPassword"
    Case "miComplInfo"
      Call GetInformation
      
    Case "miComplConsultaProdutos"
      '04/11/2009 - mpdea
      'Incluído opção incluir produto na tela de origem
      nChamaConsulta = 3
      Call SearchProdutos
      
    Case "miComplSelectPrint"
'
    Case "miComplFindNextPedido"
      Call FindNextPedido
      
    Case "miComplTransformPedidoCompra"
      Call TransformaPedidoEmCompra
      
    Case "miComplUndoMovim"
      Call UndoMovimEntrada
      Call ClearScreen
      
    Case "miComplPrintNotaFiscal"
      Call PrintNotaEntrada
      
    Case "miComplPrintReport"
      Call PrintReport
      
    Case "miComplAlteraTotais"
      Call AlteraTotais
      
    Case "miComplTypeGrade"
      Call TypeGrade
    '
    'Opções
    Case "miComplLeitorOtico"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigENTRADAS", "Scanner", Tool.Checked)
      
    Case "miOpFreezeOper"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigENTRADAS", "Mantem Operacao", Tool.Checked)
      
    Case "miOpFreezeDigitador"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigENTRADAS", "Mantem Digitador", Tool.Checked)
      
    Case "miOpFreezeFornecedor"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigENTRADAS", "Mantem Fornecedor", Tool.Checked)
      
    Case "miOpFreezeFormaPagto"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigENTRADAS", "Mantem Forma Pagto", Tool.Checked)
      
    Case "miOpReplicaEnt"
      Call ReplicaEnt
      
    '20/06/2007 - Anderson
    'Exportar entradas para Excel
    Case "miOpExportarExcel"
      Call ExportarExcel
      
  End Select

End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsEntradas
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call MostraRegistro
    End If
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  With rsEntradas
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call MostraRegistro
    End If
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  With rsEntradas
    .MovePrevious
    If Not .BOF Then
      Call MostraRegistro
    Else
      Beep
      .MoveNext
    End If
  End With
End Sub

Private Sub MoveNext()
  On Error Resume Next
  With rsEntradas
    .MoveNext
    If Not .EOF Then
      Call MostraRegistro
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Private Sub AlteraTotais()
  
  Dim Tool As ActiveBarLibraryCtl.Tool
  Dim bLocked As Boolean
  
  Call StatusMsg("")
  
'  If IsNull(Num_Registro) Then
'    gsTitle = LoadResString(201)
'    gsMsg = "Encontre a movimentação de entrada antes."
'    gnStyle = vbOKOnly + vbExclamation
'    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'    Exit Sub
'  End If
  
  Set Tool = ActiveBar1.Tools("miComplAlteraTotais")
  
  If Not Tool.Checked Then
    If lblEfetivada.Visible = True Then
'      gsTitle = LoadResString(201)
'      gsMsg = "Esta operação já foi efetivada e não pode ser alterada."
'      gnStyle = vbOKOnly + vbExclamation
'      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      frmEfetivada.Show vbModal
      Exit Sub
    Else
      If Not frmGerente.gbSenhaGerente Then
        Exit Sub
      End If
    End If
  End If
  Tool.Checked = Not Tool.Checked
  
  bLocked = Not Tool.Checked
  
  txtTotProdutos.Locked = bLocked
  txtTotDescontos.Locked = bLocked
  txtTotIPI.Locked = bLocked
  txtTotBaseICMS.Locked = bLocked
  txtTotValICMS.Locked = bLocked
  txtTotBaseICMSSubst.Locked = bLocked
  txtTotValICMSSubst.Locked = bLocked
  txtTotalAPagar.Locked = bLocked
  
  If Not bLocked Then
    txtTotProdutos.SetFocus
  End If
  
End Sub

Private Sub DeleteRecord()
  Dim Sai_Loop As Integer
  Dim Fim As Integer
  Dim Ordem As Long
  Dim Resposta As Integer
  Dim sSql As String

  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Encontre a movimentação de entrada antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  If lblEfetivada.Visible = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Esta operação já foi efetivada e não pode ser apagada por aqui, veja a ajuda."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente apagar esta movimentação"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    gsTitle = LoadResString(201)
    gsMsg = "Movimentação não apagada."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If


  Call ws.BeginTrans

  Call StatusMsg("Apagando movimentação (produtos)...")

  sSql = "DELETE * FROM [Entradas - Produtos] WHERE Filial = " & CStr(gnCodFilial) & " AND Sequência = " & CStr(txtSeq.Text)
  Call db.Execute(sSql, dbFailOnError)
  '
  Call StatusMsg("Apagando movimentação (parcelas)...")

  sSql = "DELETE * FROM [Movimento - Parcelas] WHERE Filial = " & CStr(gnCodFilial) & " AND Sequência = " & CStr(txtSeq.Text)
  Call db.Execute(sSql, dbFailOnError)
  '
  txtSeq.Text = ""
  Num_Registro = Null

  Call StatusMsg("Apagando movimentação de entrada...")

  rsEntradas.Delete

  Call ws.CommitTrans
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Oper:" & cboOper.Text & " Cli:" & cboFornecedor.Text & " ChaveRef:" & txtRef.Text & " Tot:" & txtTotalAPagar.Text, 80) & "', 'DESF_ENTRADA')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************

  Call ClearScreen

  gsTitle = LoadResString(201)
  gsMsg = "Operação apagada."
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)

  Exit Sub

ErrDelete:
  gsTitle = LoadResString(201)
  gsMsg = "Erro na Operação de Eliminação de Registros."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Call ws.Rollback
  Exit Sub

End Sub

Private Sub UpdateRecord()
  Dim nSequencia As Long
  'Variáveis de Tratamento de Erro
  Dim bSequencia As Boolean
  Dim bSeqChanged As Boolean
  Dim nRepeatUpdate3022 As Integer
  Dim nRepeatUpdateLocked As Integer
  
  Dim nRet As Integer
  
  Dim i As Integer
  Dim bm As Variant
  Dim nRow As Long
  Dim sMsg As String
  
  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked As Integer
  
  Dim dblTotalPagar As Double
  
  Dim nChaveRefNFe As String
  
  
  On Error Resume Next
  grdItens.Row = 2
  grdItens.Row = 1
  DoEvents
  
  On Error GoTo ErrTransaction
  
  If Len(txtRef.Text) > 0 Then
      nChaveRefNFe = txtRef.Text
      nChaveRefNFe = Replace(nChaveRefNFe, " ", "")
      nChaveRefNFe = Replace(nChaveRefNFe, "-", "")
      nChaveRefNFe = Replace(nChaveRefNFe, "_", "")
      nChaveRefNFe = Replace(nChaveRefNFe, ".", "")
      nChaveRefNFe = Replace(nChaveRefNFe, ",", "")
      nChaveRefNFe = Replace(nChaveRefNFe, ";", "")
      nChaveRefNFe = Replace(nChaveRefNFe, "/", "")
      nChaveRefNFe = Replace(nChaveRefNFe, "|", "")
      nChaveRefNFe = Replace(nChaveRefNFe, "\", "")
      
      If Len(nChaveRefNFe) <> 44 Then
          MsgBox "Chave de Referência deve conter 44 caracteres (APENAS NÚMEROS)", vbInformation, "Atenção"
          Exit Sub
      End If
  End If
  
  If IsNull(Num_Registro) And gbDemoVersion Then
    rsEntradas.MoveLast
    rsEntradas.MoveFirst
    If rsEntradas.RecordCount >= NMAXREGDEMO Then
      gsTitle = LoadResString(201)
      gsMsg = LoadResString(13)
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Exit Sub
    End If
  End If
  
  
  ' **********************************
  If Not IsNull(Num_Registro) Then
      'Verifica se alterou o codigo da Operacao de Entrada
      If rsEntradas.fields("Sequência").Value = txtSeq.Text And rsEntradas.fields("Operação").Value <> cboOper.Text Then
          MsgBox "Não é permitido alterar o Código da Operação.", vbInformation, "Atenção"
          cboOper.Text = rsEntradas.fields("Operação").Value
          cboOper_LostFocus
          Exit Sub
      End If
  End If
  ' **********************************
  
  If lblEfetivada.Visible = True Then
    Beep
    frmEfetivada.Show vbModal
    Exit Sub
  End If

  Call StatusMsg("")
  
  'Caso esteja no modo de alteração de totais não força o recálculo a pagar
  'do evento cboOper_LostFocus
  If Not ActiveBar1.Tools("miComplAlteraTotais").Checked Then
    Call cboOper_LostFocus
  End If
  If Len(Trim(cboOper.Text & "")) = 0 Or lblOper.Caption = "" Then
    DisplayMsg "Operação incorreta, verifique."
    cboOper.SetFocus
    Exit Sub
  End If
  
  '14/02/2005 - Daniel
  '
  'Solicitante: Daring - RJ
  '
  'Se ocorre devolução e esta devolução implica em abatimento de
  'comissão do vendedor, o Quick estava descontando erroneamente
  'da comissão para casos em que a venda possuia descontos.
  If ControlarComisao Then
    If Len(cboTabela.Text) <= 0 Then 'Está vazio
      Dim strMensagem As String
          
        strMensagem = "Esta Operação de Entrada implica no desconto na comissão do vendedor " & vbCrLf
        strMensagem = strMensagem & "será necessário informar a tabela pela qual o(s) produto(s) foram vendidos para que" & vbCrLf
        strMensagem = strMensagem & "ocorra o cálculo coerente na redução da comissão." & vbCrLf
        
        MsgBox strMensagem, vbExclamation, "Redução da Comissão"
        cboTabela.SetFocus
        Exit Sub
    End If
  End If
  
  
  '23/01/2003 - mpdea
  'Centro de Custo padrão para o modo limitado
  If Not gblnQuickFull Then
    cboCodigoCC.Text = "1"
  Else
    If rsOp_Entrada("Dinheiro") = True Then
      If Len(Trim(cboCodigoCC.Text & "")) = 0 Or lblNomeCC.Caption = "" Then
        DisplayMsg "Centro de custo incorreto, verifique."
        tabItens.Tab = 1
        cboCodigoCC.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  
  Call cboDigitador_LostFocus
  If Len(Trim(cboDigitador.Text & "")) = 0 Or lblDigitador.Caption = "" Then
    DisplayMsg "Digitador incorreto, verifique."
    cboDigitador.SetFocus
    Exit Sub
  End If
  
  Call cboFornecedor_LostFocus
  If Len(Trim(cboFornecedor.Text & "")) = 0 Or lblFornecedor.Caption = "" Then
    DisplayMsg "Fornecedor incorreto, verifique."
    cboFornecedor.SetFocus
    Exit Sub
  End If

  If medDataAcerto.Visible = True Then
    If Not IsDate(medDataAcerto.Text) Then
      DisplayMsg "Digite a data de acerto do empréstimo."
      medDataAcerto.SetFocus
      Exit Sub
    End If
  End If

  If grdItens.Rows = 0 Then
    DisplayMsg "Movimento de Entrada sem detalhamento de itens de produtos, verifique."
    grdItens.SetFocus
    Exit Sub
  End If
  
  If Not gbCheckGridItens Then
    DisplayMsg "Nenhum produto digitado ou quantidades zeradas, impossível gravar."
    grdItens.SetFocus
    Exit Sub
  End If
  
  
  '12/12/2005 - mpdea
  If rsOp_Entrada.fields("Dinheiro").Value Then
    'Verifica total zerado para operações com saída de dinheiro
    Call IsDataType(dtDouble, txtTotalAPagar.Text, dblTotalPagar)
    If dblTotalPagar <= 0 Then
      DisplayMsg "Total a Pagar menor ou igual a zero."
      tabItens.Tab = 0
      grdItens.SetFocus
      Exit Sub
    End If
    '13/12/2005 - mpdea
    'Verifica se há produto com Preço Total/Final inválido
    If m_blnItemPrecoInvalido Then
      DisplayMsg "Produto com Valor inválido, verifique"
      tabItens.Tab = 0
      grdItens.SetFocus
      Exit Sub
    End If
  End If
  

  If rsOp_Entrada("Senha") = True Then
    If Not frmGerente.gbSenhaGerente Then
      Exit Sub
    End If
  End If

  txtCxDinheiro_LostFocus
  txtCxCheque_LostFocus
  txtValCheque_LostFocus
  
  If Retorna_Valor(txtADigitar.Text) <> 0 Then
    If rsOp_Entrada("Dinheiro") = True Then
      DisplayMsg "Forma de Pagamento incompleta, verifique"
      tabItens.Tab = 1
      If cboCaixaUso.Enabled Then
        cboCaixaUso.SetFocus
      Else
        txtCxDinheiro.SetFocus
      End If
      Exit Sub
    End If
  End If

  If rsOp_Entrada("Dinheiro") = True Then
    If CDbl(gsHandleNull(txtCxDinheiro.Text)) <> 0 Or CDbl(gsHandleNull(txtCxCheque.Text)) <> 0 Then
      If Len(cboCaixaUso.Text) = 0 Then
        DisplayMsg "Informe o caixa."
        tabItens.Tab = 1
        cboCaixaUso.SetFocus
        Exit Sub
      End If
    End If
  End If

  If medDataBomPara.Enabled = True Then
    If IsNumeric(txtValCheque.Text) Then
      If CCur(txtValCheque.Text) <> 0# Then
        If Not IsDate(medDataBomPara.Text) Then
          DisplayMsg "Digite a data do cheque."
          tabItens.Tab = 1
          medDataBomPara.SetFocus
          Exit Sub
        End If
        If Len(cboConta.Text) = 0 Then
          DisplayMsg "Escolha a conta."
          tabItens.Tab = 1
          cboConta.SetFocus
          Exit Sub
        End If
      End If
    End If
  End If

  '19/05/2005 - Daniel
  '
  'Solicitante: Pedágio - Esta otimização está disponível
  '             para todos usuários do Quick Store
  '
  'O sistema deverá julgar se a nota fiscal será criada
  'automaticamente ou manualmente a partir da operação escolhida
  If gbNotaManual(CInt(Trim(cboOper.Text)), "ENTRADA") Then
    If Len(txtNF.Text) < 1 Then  'Não preencheu
     If MsgBox("Deseja gravar a entrada sem o nº da nota fiscal?", vbQuestion + vbYesNo) = vbNo Then
       txtNF.SetFocus
       m_blnFocoNF = True
       Exit Sub
      End If
    End If
  End If

  '29/06/2005 - Daniel
  'Adicionado validação caso a data do "cheque usado" seja inferior a data atual o sistema avisará
  'o usuário deixando opcional a confirmação do usuário para o prosseguimento
  If IsNumeric(txtValCheque.Text) And IsDate(medDataBomPara.Text) Then
    If Data_Atual > CDate(Format(medDataBomPara.Text, "DD/MM/YYYY")) Then
      If MsgBox("A data do cheque usado é inferior a data atual, deseja continuar ?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção") = vbNo Then Exit Sub
    End If
  End If
  
  '29/06/2005 - Daniel
  'A grid de contas a pagar estava aceitando movimentações com data retroativa e depois ao efetivar o recebimento
  'não gerava o "contas a pagar" devido a validação: If rsMovi_Parcelas("Bom") >= dtData Then rsContas_Pagar.AddNew
  'Criamos um tratamento para a grid não permitir mais datas retroativas na geração das parcelas
  If VerificaDataCP Then Exit Sub

  '---------------------------------------------------------------
  ' Início da gravação da Movimentação de Entrada
  '---------------------------------------------------------------
  Call StatusMsg("Gravando Movimentação de Entrada...")
  
  '17/05/2006 - mpdea
  'Recalcula valores para atualizar no grid os valores de ICMS Retido
  Call Recalcula(True)
 

  Call ws.BeginTrans
  blnInTransaction = True
  
  Rem pega número da nova movimentação
  If IsNull(Num_Registro) Then
    nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("Última Movimentação") + 1
    
    rsParametros.Edit
    rsParametros("Última Movimentação") = nSequencia
    rsParametros.Update
  End If
  
  With rsEntradas
    If IsNull(Num_Registro) Then
      .AddNew
      .fields("Sequência") = nSequencia
      sMsg = "inserida"
    Else
'      .FindFirst "Sequência = " & nSequencia
      .Edit
      nSequencia = .fields("Sequência")
      sMsg = "atualizada"
    End If

    '22/09/2004 - Daniel
    'Case: Resultado
    'Verificar se a operação é Empréstimo caso seja
    'faremos o tratamento para o campo [Entradas - Produtos].EntradaConsignada
    'posteriormente
    If m_blnResultado Then
      Call VerificarOperacao
    End If

    .fields("Filial") = gnCodFilial
    .fields("Data") = lblToday.Caption
    .fields("Operação") = Val(cboOper.Text)
    .fields("Digitador") = Val(cboDigitador.Text)
    .fields("Fornecedor") = CLng(cboFornecedor.Text)
    .fields("Observações") = txtObs.Text & ""
    .fields("Nota Fiscal") = txtNF.Text & ""
    
    .fields("obs_infCpl1") = txt_informacoesComplNFe(0).Text & ""
    
    '19/05/2005 - Daniel
    '
    'Solicitante: Pedágio Calçados - Otimização liberada
    '             para todos usuários do Quick Store
    '
    'Tratamento para o campo Nr Série da NF
    .fields("SerieNF").Value = UCase(txtNrSerie.Text) & ""
    
    '17/09/2009 - mpdea
    'Modelo de documento fiscal
    .fields("ModeloDocumentoFiscal").Value = gstrGetModeloDocumentoFiscalOperacao(tmEntradas, .fields("Operação").Value)
    
    .fields("Pedido") = txtPedido.Text & ""
    
    If rsOp_Entrada("Dinheiro") = True Then
      If IsNumeric(Trim(cboCodigoCC.Text)) _
         And Len(lblNomeCC.Caption) > 0 Then
        .fields("CentroCusto") = Trim(cboCodigoCC.Text) & ""
      End If
    End If
    
    .fields("Produtos") = Retorna_Valor(txtTotProdutos.Text)
    .fields("Desconto") = Retorna_Valor(txtTotDescontos.Text)
    .fields("IPI") = Retorna_Valor(txtTotIPI.Text)
    .fields("Frete") = Retorna_Valor(txtFrete.Text)
    .fields("Base ICM") = Retorna_Valor(txtTotBaseICMS.Text)
    .fields("Valor ICM") = Retorna_Valor(txtTotValICMS.Text)
    .fields("Base ICM Subs") = Retorna_Valor(txtTotBaseICMSSubst.Text)
    .fields("Valor ICM Subs") = Retorna_Valor(txtTotValICMSSubst)
    .fields("Total") = Retorna_Valor(txtTotalAPagar.Text)
    .fields("Caixa") = False

    'If frmPagto_Entrada.O_Caixa = True Then .Fields("Caixa") = True

    If rsOp_Entrada("Dinheiro") = True Then
      .fields("Caixa") = 0
      If Len(cboCaixaUso.Text) > 0 Then
        .fields("Caixa") = Val(cboCaixaUso.Text)
      End If
      .fields("Dinheiro Caixa") = CDbl(gsHandleNull(txtCxDinheiro.Text & ""))
      .fields("Cheque Caixa") = CDbl(gsHandleNull(txtCxCheque.Text & ""))
      
      '08/01/2010 - Andrea
      .fields("Troco") = CDbl(gsHandleNull((txtTroco.Text * -1) & ""))
      
      If Len(cboConta.Text) > 0 Then
        .fields("Conta") = Val(cboConta.Text)
      End If
      .fields("Num Cheque") = gsHandleNull(txtCheque.Text & "")
      .fields("Descrição") = txtDescricao.Text & ""
      .fields("Valor Cheque") = gsHandleNull(txtValCheque.Text & "")
      If IsDate(medDataBomPara.Text) Then
        .fields("Bom Para") = CDate(medDataBomPara.Text)
      End If
    Else
      .fields("Dinheiro Caixa") = 0
      .fields("Cheque Caixa") = 0
      .fields("Conta") = 0
      .fields("Num Cheque") = ""
      .fields("Descrição") = ""
    End If

    If IsDate(medDataEmissao.Text) Then
      .fields("Data Emissão") = CDate(medDataEmissao.Text)
    End If

    If medDataAcerto.Visible = True Then
      .fields("Data Acerto Empréstimo") = CDate(medDataAcerto.Text)
    End If


    '15/01/2010 - Andrea
    '-------------------------------------------------------------------------------------------------
    .fields("NumeroDI") = gsHandleNull(txtNumeroDI.Text & "")
    .fields("CodigoExportador") = gsHandleNull(txtCodigoExportadorDI.Text & "")
    .fields("UFDesembaracoDI") = txtUFDesembaracoDI.Text & ""
    .fields("LocalDesembaracoDI") = txtLocalDesembaracoDI.Text & ""
    .fields("NumeroAdicaoDI") = gsHandleNull(txtNumeroAdicao.Text & "")
    .fields("NumeroSeqItemAdicaoDI") = gsHandleNull(txtNumeroSequenciaItem.Text & "")
    .fields("CodigoFabricanteAdicaoDI") = txtCodigoFabricante.Text & ""
    .fields("DescontoAdicaoDI") = gsHandleNull(txtDescontoItemAdicao.Text & "")
    
    If IsDate(medDataRegistroDI.Text) Then
      .fields("DataDeRegistroDI") = CDate(medDataRegistroDI.Text)
    End If
    
    If IsDate(medDataDesembaracoDI.Text) Then
      .fields("DataDesembaracoDI") = CDate(medDataDesembaracoDI.Text)
    End If
    '-------------------------------------------------------------------------------------------------
    
    .fields("Consumidor_Final").Value = Left(cboConsumidorFinal.Text, 1)
    .fields("Presenca_Comprador").Value = Left(cboPresencaComprador.Text, 1)
    .fields("FinalidadeNFe").Value = Left(cboFinalidade.Text, 1)
    If Len(nChaveRefNFe) > 0 Then
        .fields("ChaveReferenciada").Value = nChaveRefNFe
    End If
    .fields("TotalDesoneracaoICMS").Value = txtTotICMSDesonerado.Text
    
    bSeqChanged = False
    bSequencia = True
    .Update
    bSequencia = False
    'Grava novamente a última movimentação
    'se a mesma foi alterada
    If bSeqChanged Then
      With rsParametros
        .Edit
        .fields("Última Movimentação") = nSequencia
        .Update
      End With
    End If
    
    .Bookmark = .LastModified

  End With
  
  txtSeq.Text = nSequencia
  
  If Erro_Data2 = True Then  'grava log
    rsLog.AddNew
    rsLog("Tipo") = "MOVIMENTAÇÃO"
    rsLog("Data") = Date
    rsLog("Texto") = "Movimentação " & nSequencia & _
                     " gravada com data incorreta. Filial " & str(gnCodFilial)
    rsLog.Update
  End If

'  Call WriteGridCP
  
  grdCP.Update
  
  Call EraseTypeMoviment(tmMovimentoParcelas, gnCodFilial, CLng(txtSeq.Text))
  
  For nRow = 0 To grdCP.Rows - 1
    bm = grdCP.AddItemBookmark(nRow)
    If IsDate(grdCP.Columns("Data").CellText(bm)) Then
      If IsNumeric(grdCP.Columns("Valor").CellValue(bm)) Then
        With rsMovi_Parcelas
          .AddNew
          .fields("Filial") = gnCodFilial
          .fields("Sequência") = Val(txtSeq.Text)
          .fields("Ordem") = nRow + 1
          .fields("Bom") = grdCP.Columns("Data").CellText(bm)
          .fields("Valor") = grdCP.Columns("Valor").CellValue(bm)
          .Update
        End With
      End If
    End If
  Next nRow
    
    
  Grade_ChequesEmCaixa.Update
  
  '------------------------------------------------------------------------------------------------------------------
  '07/01/2010 - Andrea
  'Apaga Cheques
  Call EraseTypeMoviment(tmMovimentoCheques, gnCodFilial, Val(txtSeq.Text))
  'Grava Cheques
  For nRow = 0 To Grade_ChequesEmCaixa.Rows - 1
    bm = Grade_ChequesEmCaixa.AddItemBookmark(nRow)
    If IsDate(Grade_ChequesEmCaixa.Columns("Bom Para").CellText(bm)) Then
      If IsNumeric(Grade_ChequesEmCaixa.Columns("Valor").CellValue(bm)) Then
        With rsMovi_Cheques
          .AddNew
          .fields("Filial") = gnCodFilial
          .fields("Sequência") = Val(txtSeq.Text)
          .fields("Ordem") = nRow + 1
          .fields("Bom") = Grade_ChequesEmCaixa.Columns("Bom Para").CellText(bm)
          .fields("Valor") = Grade_ChequesEmCaixa.Columns("Valor").CellValue(bm)
          .fields("Banco") = Grade_ChequesEmCaixa.Columns("Banco").CellText(bm)
          .fields("Cheque") = Grade_ChequesEmCaixa.Columns("NumeroCheque").CellValue(bm)
          .Update
        End With
      End If
    End If
  Next nRow
  '------------------------------------------------------------------------------------------------------------------
  
'  Call WriteGridItens
  
  If Not gbWriteGridItens Then
    Screen.MousePointer = vbDefault
    Call StatusMsg("")
    ws.Rollback
    blnInTransaction = False
    Exit Sub
  End If
  
  If rsOp_Entrada("Tipo") <> "P" Then
    Call StatusMsg("Efetivando entrada...")
    nRet = Efetiva_Entrada(gnCodFilial, Val(txtSeq.Text))
    If nRet <> 0 Then
      Select Case nRet
        Case -1
          'Ação cancelada
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
      Screen.MousePointer = vbDefault
      ws.Rollback
      blnInTransaction = False
      Exit Sub
    Else
      lblEfetivada.Visible = True
    End If
  End If

  Call ws.CommitTrans
  blnInTransaction = False
  Num_Registro = rsEntradas.Bookmark
  
  gravada = True
  
  If lblEfetivada.Visible = True Then
    Call StatusMsg("OPERAÇÃO EFETIVADA. Movimentação de Entrada " & sMsg & " com sucesso.")
  Else
    Call StatusMsg("Movimentação de Entrada " & sMsg & " com sucesso.")
  End If
  
  tabItens.Tab = 0
  cboOper.SetFocus
  
  Exit Sub

ErrTransaction:
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

Private Sub PrintReport()
  Dim sSql As String
  Dim Str_Rel As String
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Encontre ou grave uma movimentação antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  sSql = "DELETE * FROM ENTRADAS WHERE CodUsuarioOwner = " & CStr(gnUserCode)
  Call dbTemp.Execute(sSql, dbFailOnError)
  
  Call Grava_Temp_Entradas(rsEntradas("Filial"), rsEntradas("Sequência"))
  
  Rem  Nome do BD
  With Rel1
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
  End With
  Rel1.SelectionFormula = "{Entradas.CodUsuarioOwner} = " & CStr(gnUserCode)
  
  Rem Saída
  Rel1.Destination = 0
  
  Rem Nome do arquivo .rpt
  Rel1.ReportFileName = gsReportPath & "Entrada1.RPT"
  
  Rem Seleção
  
  Str_Rel = "{Entradas.Filial} =" + str(gnCodFilial)
  Str_Rel = Str_Rel + " And {Entradas.Sequência} = " + txtSeq.Text
  
  ' Rel1.SelectionFormula = Str_Rel
  
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  
  Rel1.Formulas(0) = Str_Rel
  
  Str_Rel = "filial = '"
  Str_Rel = Str_Rel + lblFilial.Caption + "'"
  Rel1.Formulas(1) = Str_Rel
  
  Call StatusMsg("Aguarde, imprimindo..")
  MousePointer = vbHourglass
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  
  
  Rel1.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault
  
End Sub

Public Sub ClearScreen()
  Dim nRow As Long
  Dim Tool As ActiveBarLibraryCtl.Tool
  
  'Na mudança de registro o Altera Totais é desmarcado
  Set Tool = ActiveBar1.Tools("miComplAlteraTotais")
  If Tool.Checked Then
    Call ActiveBar1_Click(Tool)
  End If
  
  If ActiveBar1.Tools("miOpFreezeOper").Checked = False Then
     cboOper.Text = ""
     lblOper.Caption = ""
  End If
  If ActiveBar1.Tools("miOpFreezeDigitador").Checked = False Then
    cboDigitador.Text = ""
    lblDigitador.Caption = ""
  End If
  If ActiveBar1.Tools("miOpFreezeFornecedor").Checked = False Then
    cboFornecedor.Text = ""
    lblFornecedor.Caption = ""
  End If
  
  txtObs.Text = ""
  txtNF.Text = ""
  
  cboConsumidorFinal.Text = "1=Sim"
  cboPresencaComprador.Text = "1 =Operação presencial"
  cboFinalidade.Text = "1=NFe normal"
  txtRef.Text = ""
  txtTotICMSDesonerado.Text = ""
  '19/05/2005 - Daniel
  '
  'Solicitante: Pedágio Calçados - Otimização liberada
  '             para todos usuários do Quick Store
  '
  'Tratamento para o campo Nr Série da NF
  txtNrSerie.Text = ""
  '-------------------
  txtPedido.Text = ""
  cboCodigoCC.Text = "0"
  lblNomeCC.Caption = ""
    
  medDataEmissao.Mask = ""
  medDataEmissao.Text = ""
  medDataEmissao.Mask = "##/##/####"
    
  
  '07/11/2002 - mpdea
  'Incluído o campo de Data de Acerto de Empréstimo
  With medDataAcerto
    .Mask = ""
    .Text = ""
    .Mask = "##/##/####"
  End With
  
  
  lblToday.Caption = Format$(Data_Atual, "dd/mm/yyyy")
  
  txtFrete.Text = ""
  lblEfetivada.Visible = False
  
  '14/02/2005 - Daniel
  '
  'Solicitante: Daring - RJ
  '
  'Se ocorre devolução e esta devolução implica em abatimento de
  'comissão do vendedor, o Quick estava descontando erroneamente
  'da comissão para casos em que a venda possuia descontos.
'''  lblTabela.Visible = False
'''  cboTabela.Visible = False
  lbl_avisoTrataComissaoVendedor.Visible = False
  '-------------------------------------------------------------
  
  txtSeq.Text = ""
  
  '07/01/2010 - Andrea
  Grade_ChequesEmCaixa.RemoveAll
  '20/01/2010 - mpdea
  datCheques.Refresh

  grdItens.Redraw = False
  grdItens.RemoveAll
  For nRow = 0 To 11
    grdItens.AddItem ""
  Next nRow
  grdItens.Redraw = True

  cboCaixaUso.Text = ""
  lblCaixaUso.Caption = ""
  If gbCaixas = False Then
    cboCaixaUso.Text = "1"
  End If
  
  '03/08/2005 - Daniel
  'Correção de ajuste da grid
  With grdItens
    .Columns("PrecoFinal").Width = 1468
    .Columns("Descricao").Width = 2068
    .Columns("Indice").Visible = False
  End With
    
  txtCxDinheiro.Text = ""
  txtCheque.Text = ""
  txtDescricao.Text = ""
  txtValCheque.Text = ""
  txtCxCheque.Text = ""
  cboConta.Text = ""
  lblConta.Caption = ""
  
  medDataBomPara.Mask = ""
  medDataBomPara.Text = ""
  medDataBomPara.Mask = "##/##/####"
  
  txtTotDigitado.Text = ""
  txtADigitar.Text = ""
  
  grdCP.RemoveAll
  
  txtTotProdutos.Text = ""
  txtTotDescontos.Text = ""
  txtTotIPI.Text = ""
  txtTotBaseICMS.Text = ""
  txtTotValICMS.Text = ""
  txtTotBaseICMSSubst.Text = ""
  txtTotValICMSSubst.Text = ""
  
  txtTotalAPagar.Text = ""
  
  '15/01/2010 - Andrea
  txtNumeroDI.Text = ""
  txtCodigoExportadorDI.Text = ""
  
  medDataRegistroDI.Mask = ""
  medDataRegistroDI.Text = ""
  medDataRegistroDI.Mask = "##/##/####"

  txtUFDesembaracoDI.Text = ""
  txtLocalDesembaracoDI.Text = ""
  
  medDataDesembaracoDI.Mask = ""
  medDataDesembaracoDI.Text = ""
  medDataDesembaracoDI.Mask = "##/##/####"
  
  txtNumeroAdicao.Text = ""
  txtNumeroSequenciaItem.Text = ""
  txtCodigoFabricante.Text = ""
  txtDescontoItemAdicao.Text = ""
  
  tabItens.Tab = 0
  tabItens.TabCaption(0) = "&Itens"
  
  If Not rsEntradas.EOF Then
    On Error Resume Next
    rsEntradas.MoveFirst
    rsEntradas.MovePrevious
    On Error GoTo 0
  End If

  Num_Registro = Null
  
  cboOper.SetFocus
  gravada = False
End Sub
Public Sub ReplicaEnt()

  Dim Tool As ActiveBarLibraryCtl.Tool
  
  
  If IsNull(Num_Registro) Then
     DisplayMsg ("Encontre uma movimentação antes. ")
     Exit Sub
  End If
   
  Set Tool = ActiveBar1.Tools("miComplAlteraTotais")
  If Tool.Checked Then
    Call ActiveBar1_Click(Tool)
  End If
   
  txtNF.Text = ""
  lblToday.Caption = Format$(Data_Atual, "dd/mm/yyyy")
    
  lblEfetivada.Visible = False
  
  txtSeq.Text = ""
    
  txtCxDinheiro.Text = ""
  txtCheque.Text = ""
  txtDescricao.Text = ""
  txtValCheque.Text = ""
  txtCxCheque.Text = ""
  cboConta.Text = ""
  lblConta.Caption = ""
  
  medDataBomPara.Mask = ""
  medDataBomPara.Text = ""
  medDataBomPara.Mask = "##/##/####"
  
  txtTotDigitado.Text = ""
  txtADigitar.Text = ""
  
  grdCP.RemoveAll
  
  Num_Registro = Null

  DisplayMsg ("Movimentação Replicada. Verifique os valores e Grave.")
  
  End Sub


'28/11/2003 - mpdea
'Incluído tratamento de erro "on error"
Private Sub PrintNotaEntrada()
  Dim Aux As Variant
  Dim Nome_Arq As String
  Dim Texto As String
  Dim Final As Integer
  Dim Str_Impre As String
  Dim Num_cod As Integer
  Dim Resposta As Long
  Dim Final_Linha As Integer
  Dim Linhas As Integer
  Dim Resp As Integer
  Dim F As Form
  Dim nX As Integer
  
  Dim lngUltimaNotaFiscal As Long
  Dim blnInTransaction As Boolean
  
  '19/12/2007 - Anderson
  'Implementação do NSU para SC
  Dim blnGerarNSU As Boolean
  
  '19/12/2007 - Anderson
  'Implementação do NSU para SC
  blnGerarNSU = True
  
  On Error GoTo ErrHandler
 
  Call StatusMsg("")
  
  Aux = txtSeq.Text
  If IsNull(Aux) Or Aux = "" Then
    DisplayMsg "Ache uma entrada antes."
    Exit Sub
  End If
  
  Set F = New frmObsNota
  F.lngSequencia = rsEntradas("Sequência")
  F.bytTipoTabela = 2
  F.Show vbModal
  Set F = Nothing
  If gsRetornoDoc <> "OK" Then
    DisplayMsg "Nota não impressa."
    Exit Sub
  End If
  
  '11/08/2003 - maikel
  '             Gravação dos campos de observações na tela de saídas
  '----------------------------------------------------------------'
    rsEntradas.Edit
    
    'For nX = 0 To 7
    '  rsEntradas.Fields("obs_Obs" & nX + 1).Value = gsObsDoc(nX)
    'Next nX
    For nX = 0 To 1
      rsEntradas.fields("obs_infCpl" & nX + 1).Value = gsObsDoc(nX)
    Next nX
    
    rsEntradas.fields("obs_Transportadora") = gsTransportadora
    rsEntradas.fields("obs_Placa") = gsPlaca
    rsEntradas.fields("obs_Uf") = gsUfrmPlaca
    rsEntradas.fields("obs_Especie") = gsEspecieTrans
    rsEntradas.fields("obs_Qtde") = gsQtdeTrans
    rsEntradas.fields("obs_Marca") = gsMarcaTrans
    rsEntradas.fields("obs_PesoBruto") = IIf(IsNumeric(gsPesoBruto), gsPesoBruto, 0)
    rsEntradas.fields("obs_PesoLiquido") = IIf(IsNumeric(gsPesoLiquido), gsPesoLiquido, 0)
    
    rsEntradas.fields("obs_FretePago") = IIf(IsNumeric(gsFretePago), gsFretePago, 0)
    rsEntradas.Update
    
  '19/12/2007 - Anderson
  'Implementação do NSU
  If Not (gbNotaManual(rsEntradas.fields("Operação").Value, "ENTRADA")) Then
    Call IsDataType(dtLong, rsEntradas.fields("Nota Impressa").Value, lngUltimaNotaFiscal)
    If lngUltimaNotaFiscal <> 0 Then
      
      '18/12/2007 - Anderson
      'Implementação do NSU para SC
      blnGerarNSU = False

    End If
  
  End If
    
  '----------------------------------------------------------------'
  
  '19/05/2005 - Daniel
  '
  'Solicitante: Pedágio - Esta otimização está disponível
  '             para todos usuários do Quick Store
  '
  'O sistema deverá julgar se a nota fiscal será criada
  'automaticamente ou manualmente a partir da operação escolhida
  'Nota: Caso seja manualmente (notas de bloquinho), o sistema não
  'deverá incrementar o contador pois o sistema estava fora do ar
  If Not (gbNotaManual(rsEntradas.fields("Operação").Value, "ENTRADA")) Then
  
      'Pega próxima nota e grava no arquivo
      Aux = rsEntradas("Nota Impressa")
      If IsNull(Aux) Then Aux = 0
      If Not IsNumeric(Aux) Then Aux = 0
      If Val(Aux) = 0 Then
        
        '-------------------------------------------------------------------
        '28/11/2003 - mpdea
        'Modificado leitura e gravação do número da última nota fiscal
        'Incluído transação durante gravação
        'lngUltimaNotaFiscal = rsParametros.Fields("Última Nota").Value + 1
        lngUltimaNotaFiscal = g_lngNextNotaFiscal(rsEntradas.fields("Filial").Value)
        '
        ws.BeginTrans
        blnInTransaction = True
        '
        'With rsParametros
        '  .Edit
        '  .Fields("Última Nota").Value = lngUltimaNotaFiscal
        '  .Update
        'End With
        '
        With rsEntradas
          .Edit
          .fields("Nota Impressa").Value = lngUltimaNotaFiscal
          .Update
        End With
        '
        ws.CommitTrans
        blnInTransaction = False
        '
        'rsParametros.Edit
        'rsParametros("Última Nota") = rsParametros("Última Nota") + 1
        'rsParametros.Update
        'rsEntradas.Edit
        'rsEntradas("Nota Impressa") = rsParametros("Última Nota")
        'rsEntradas.Update
        '-------------------------------------------------------------------
      End If
  
  End If 'If Not (gbNotaManual...)
  
  '18/12/2007 - Anderson
  'Implementação do NSU
  If blnGerarNSU Then
    Call GerarNSU(rsEntradas, "Entradas")
  End If
  
  '-------------------------------------------------------------------
  'Pegar o nome do arquivo de configuração
  '-------------------------------------------------------------------
  Nome_Arq = gsConfigPath & rsParametros("Nota Entrada") + ".CNF"
  Resp = Imprime_Nota_Entrada(Nome_Arq, rsEntradas("Filial"), rsEntradas("Sequência"))
  
  If Resp = 0 Then
    DisplayMsg "Nota impressa com sucesso."
  Else
    DisplayMsg "Houve o erro " + str(Resp) + " durante a impressão da nota."
  End If
 
  Exit Sub
  
Arq_Inexiste:
  DisplayMsg "Arquivo de configuração não encontrado."
  Exit Sub
  
Final_Arquivo:
  DisplayMsg "Nota fiscal impressa."
  Exit Sub
  
ErrHandler:
  If blnInTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub cboCodigoCC_CloseUp()
  cboCodigoCC.Text = Trim(cboCodigoCC.Columns(0).Text)
  cboCodigoCC_LostFocus
End Sub

Private Sub cboCodigoCC_GotFocus()
  datCentroCusto.Refresh
End Sub

Private Sub cboCodigoCC_LostFocus()
  Dim rs As Recordset
  Dim sValor As String
  
  If Not IsNumeric(cboCodigoCC.Text) Then
    lblNomeCC.Caption = ""
    Exit Sub
  End If
  If Len(cboCodigoCC.Text) > 0 Then
    datCentroCusto.Refresh
    Set rs = datCentroCusto.Recordset.Clone
    rs.FindFirst "Código = " & cboCodigoCC.Text
    If Not rs.NoMatch Then
      lblNomeCC.Caption = rs!Nome
    Else
      lblNomeCC.Caption = ""
    End If
    rs.Close
    Set rs = Nothing
  End If
End Sub

Private Sub cboDigitador_DropDown()
  cboDigitador.SelStart = 0
  cboDigitador.SelLength = Len(cboDigitador.Text)
End Sub

Private Sub cboDigitador_KeyPress(KeyAscii As Integer)
  If Len(cboDigitador.Text) >= 4 Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub cboFornecedor_DropDown()
  cboFornecedor.SelStart = 0
  cboFornecedor.SelLength = Len(cboFornecedor.Text)
End Sub

Private Sub cboFornecedor_KeyPress(KeyAscii As Integer)
  If Len(cboFornecedor.Text) >= 8 Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub cboOper_DropDown()
  cboOper.SelStart = 0
  cboOper.SelLength = Len(cboOper.Text)
End Sub

Private Sub cboOper_KeyPress(KeyAscii As Integer)
  If Len(cboFornecedor.Text) >= 3 Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub cboTabela_LostFocus()
  '14/02/2005 - Daniel
  '
  'Solicitante: Daring - RJ
  '
  'Se ocorre devolução e esta devolução implica em abatimento de
  'comissão do vendedor, o Quick estava descontando erroneamente
  'da comissão para casos em que a venda possuia descontos.
  If Len(cboTabela.Text) > 0 Then gsTabelaVenda = Trim(cboTabela.Text) & ""
  
End Sub

Private Sub cmdPreencher_Click()
  Dim nX As Integer
  Call ClearScreen
  grdItens.SetFocus
  SendKeys "^{HOME}", True
  For nX = 1 To 250
    SendKeys "1{DOWN}", True
    grdItens.Columns("Qtde").Text = "1"
  Next nX
  SendKeys "1{UP}", True
End Sub

'07/01/2010 - Andrea
Private Sub ddwCheques_Click()
  
  Grade_ChequesEmCaixa.Columns("NumeroCheque").Text = ddwCheques.Columns("Cheque").Text
  Grade_ChequesEmCaixa.Columns("Valor").Text = gsFormatCurrency(ddwCheques.Columns("Valor").Value, True)
  Grade_ChequesEmCaixa.Columns("Bom Para").Text = ddwCheques.Columns("Vencimento").Text
  Grade_ChequesEmCaixa.Columns("Banco").Text = ddwCheques.Columns("Banco").Text
 
  Grade_ChequesEmCaixa.Col = 0
  Grade_ChequesEmCaixa.ActiveCell.SelStart = 0
  Grade_ChequesEmCaixa.ActiveCell.SelLength = 0

End Sub

'07/01/2010 - Andrea
Private Sub ddwCheques_InitColumnProps()
  
  Grade_ChequesEmCaixa.Columns("NumeroCheque").DropDownHwnd = ddwCheques.hwnd
  ddwCheques.DataFieldList = "Cheque"

End Sub

Private Sub ddwProduto_Click()
'  grdItens.Columns("Codigo").Text = ddwProduto.Columns("Codigo").Text
'  grdItens.Columns("Descricao").Text = ddwProduto.Columns("Nome").Text
'  'grdItens.Columns("Preco").Text = ddwProduto.Columns("Preco").Value
'  grdItens.Columns("Preco").Text = gsFormatCurrency(ddwProduto.Columns("Preco").Value, True)
'  grdItens.Columns("Unidade").Text = ddwProduto.Columns("Unidade").Text
'  grdItens.Col = 0
'  grdItens.ActiveCell.SelStart = 0
'  grdItens.ActiveCell.SelLength = 0
End Sub

Private Sub ddwProduto_DropDown()
'  If IsNumeric(grdItens.Columns("Codigo").Text) Then
'    ddwProduto.DataFieldList = "Código"
'  Else
'    ddwProduto.DataFieldList = "Nome"
'  End If
  Dim rsTemp As Recordset
  Set rsTemp = db.OpenRecordset("SELECT Código FROM Produtos WHERE Código = '" & grdItens.Columns("Codigo").Text & "'", dbOpenSnapshot)
  If rsTemp.EOF Then
    ddwProduto.DataFieldToDisplay = "Nome"
  Else
    ddwProduto.DataFieldToDisplay = "Código"
  End If
  rsTemp.Close
  Set rsTemp = Nothing
End Sub

Private Sub ddwProduto_InitColumnProps()
  '05/05/2004 - Daniel
  'Personalização Embalavi
  If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
    ddwProduto.Columns("Preco").NumberFormat = "##,###,##0.00000"
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    ddwProduto.Columns("Preco").NumberFormat = "##,###,##0.000"
  Else
    ddwProduto.Columns("Preco").NumberFormat = gsCurrencyFormat
  End If
End Sub

Private Sub Grade_ChequesEmCaixa_AfterUpdate(RtnDispErrMsg As Integer)
  Call RecalculaPagar
End Sub

Private Sub grdCP_AfterDelete(RtnDispErrMsg As Integer)
  Call RecalculaPagar
End Sub

Private Sub grdCP_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  If Len(Trim(grdCP.ActiveCell.Text)) = 0 Then
    If bGridBeforeDelete() = True Then
      Cancel = False
    Else
      Cancel = True
    End If
  Else
    Cancel = True
  End If
End Sub

Private Sub grdCP_BeforeUpdate(Cancel As Integer)
  Dim nRow As Long
  Dim bm As Variant
  
'  For nRow = 0 To grdCP.Rows - 1
'    bm = grdCP.AddItemBookmark(nRow)
'    If IsDate(grdCP.Columns(0).CellText(bm)) Then
'      If Len(grdCP.Columns(1).CellText(bm)) = 0 Then
'        gsTitle = LoadResString(201)
'        gsMsg = "Valor de Pagamento faltante ou incorreto. Verifique."
'        gnStyle = vbOKOnly + vbCritical
'        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'        Cancel = True
'        Exit Sub
'      End If
'    End If
'  Next nRow
'
'  Cancel = False
'
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

Private Sub grdItens_AfterColUpdate(ByVal ColIndex As Integer)
'  If ColIndex = 0 Then
'    If ddwProduto.DataFieldList = "" Then
'      ddwProduto.DataFieldList = "Codigo"
'    End If
'  End If
  Call Calcula_Linha
End Sub

Private Sub grdItens_AfterDelete(RtnDispErrMsg As Integer)
  RtnDispErrMsg = False
'  ddwProduto.Refresh
  Call Recalcula
  Call RecalculaPagar
End Sub

Private Sub grdItens_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  If Len(Trim(grdItens.ActiveCell.Text)) = 0 Then
    If bGridBeforeDelete() = True Then
      Cancel = False
    Else
      Cancel = True
    End If
  Else
    Cancel = True
  End If
End Sub

Private Sub grdItens_BeforeUpdate(Cancel As Integer)
'  Cancel = False
'  If gsHandleNull(grdItens.Columns("Qtde").Text) = 0 Then
'    DisplayMsg "Quantidade incorreta."
'    Cancel = True
'    grdItens.Col = 1
'  End If
End Sub

Private Sub grdItens_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
'  If grdItens.Col = 0 Then
'    If IsNumeric(grdItens.Columns(0).Text) Then
'      ddwProduto.DataFieldList = "Código"
'    Else
'      ddwProduto.DataFieldList = "Nome"
'    End If
'  End If
End Sub

Private Sub grdItens_RowLoaded(ByVal Bookmark As Variant)
  Dim nCol As Integer
  
  For nCol = 0 To grdItens.Cols - 1
    If grdItens.Columns(nCol).Name = "PrecoFinal" Then
      grdItens.Columns(nCol).CellStyleSet "Total", grdItens.Row
    Else
      If grdItens.Columns(nCol).Name <> "Descricao" And grdItens.Columns(nCol).Name <> "Unidade" Then
        grdItens.Columns(nCol).CellStyleSet "Normal", grdItens.Row
      End If
    End If
  Next nCol
  
  '05/05/2004 - Daniel
  'Personalização Embalavi
  If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
    grdItens.Columns("Preco").Text = Format(grdItens.Columns("Preco").Text, "##,###,##0.00000")   'gsFormatCurrency(grdItens.Columns("Preco").Text, True)
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    grdItens.Columns("Preco").Text = Format(grdItens.Columns("Preco").Text, "##,###,##0.000")   'gsFormatCurrency(grdItens.Columns("Preco").Text, True)
  Else
    grdItens.Columns("Preco").Text = Format(grdItens.Columns("Preco").Text, FORMAT_VALUE)   'gsFormatCurrency(grdItens.Columns("Preco").Text, True)
  End If
  
  grdItens.Columns("PrecoFinal").Text = Format(grdItens.Columns("PrecoFinal").Text, FORMAT_VALUE)  'gsFormatCurrency(grdItens.Columns("PrecoFinal").Text, True)
  grdItens.Columns("PrecoTotal").Text = Format(grdItens.Columns("PrecoTotal").Text, FORMAT_VALUE)   'gsFormatCurrency(grdItens.Columns("PrecoTotal").Text, True)
End Sub




Private Sub medDataAcerto_GotFocus()
  Call SelectAllText(medDataAcerto)
End Sub

Private Sub medDataBomPara_GotFocus()
  medDataBomPara.SelStart = 0
  medDataBomPara.SelLength = Len(medDataBomPara.Text)
End Sub

Private Sub medDataBomPara_LostFocus()
  medDataBomPara.Text = Ajusta_Data(medDataBomPara.Text)
End Sub

Private Sub medDataBomPara_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      medDataBomPara.Text = frmCalendario.gsDateCalender(medDataBomPara.Text)
  End Select
End Sub

Private Sub medDataEmissao_GotFocus()
  Call SelectAllText(medDataEmissao)
End Sub

Private Sub tabItens_Click(PreviousTab As Integer)
'27/04/2004 - Daniel
'Estava recalculando o que já havia sido
'calculado
'  txtADigitar.Text = txtTotalAPagar.Text
'  Call FormatCurrencyValue(txtADigitar)
End Sub

Private Sub tabItens_KeyDown(KeyCode As Integer, Shift As Integer)
  '12/11/2004 - Daniel
  'Solicitação: Primart
  Select Case KeyCode
    Case vbKeyF7
      cboCaixaUso.SetFocus
  End Select
End Sub

Private Sub txtADigitar_GotFocus()
  SendKeys "{Tab}"
End Sub

Private Sub txtCheque_GotFocus()
  txtCheque.SelStart = 0
  txtCheque.SelLength = Len(txtCheque.Text)
End Sub

Private Sub txtCxCheque_GotFocus()
  txtCxCheque.SelStart = 0
  txtCxCheque.SelLength = Len(txtCxCheque.Text)
End Sub

Private Sub txtCxCheque_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub txtCxCheque_LostFocus()
  If IsGoodNumber(txtCxCheque) Then
    Call RecalculaPagar
  End If
End Sub

Private Sub cboCaixaUso_Click()
  cboCaixaUso.Text = cboCaixaUso.Columns(1).Text
  cboCaixaUso_LostFocus
End Sub

Private Sub cboCaixaUso_LostFocus()
  Dim rs As Recordset
  Dim sValor As String
  
  Call StatusMsg("")
  
  If Len(cboCaixaUso.Text) > 0 Then
    sValor = ""
    Set rs = datCaixasUso.Recordset.Clone
    rs.FindFirst "Caixa = " & cboCaixaUso.Text
    If Not rs.NoMatch Then
      sValor = rs("Descrição")
    End If
    rs.Close
    Set rs = Nothing
    lblCaixaUso.Caption = sValor
  End If

End Sub

Private Sub cboConta_Click()
  cboConta.Text = cboConta.Columns(1).Text
  cboConta_LostFocus
End Sub

Private Sub cboConta_LostFocus()
  Dim rs As Recordset
  Dim sValor As String
  
  Call StatusMsg("")
  
  If Len(cboConta.Text) > 0 Then
    sValor = ""
    Set rs = datConta.Recordset.Clone
    rs.FindFirst "Código = " & cboConta.Text
    If Not rs.NoMatch Then
      sValor = rs("Descrição")
      lblConta.Caption = sValor
    Else
      DisplayMsg "Conta Corrente inválida. Verifique."
      lblConta.Caption = ""
      '21/06/2007 - Anderson
      'Estava causando erro pois o programa estava setando o foco quando o objeto estava desativado
      If cboConta.Enabled = True Then
        cboConta.SetFocus
      End If
    End If
    rs.Close
    Set rs = Nothing
  Else
    '29/06/2005 - Daniel
    'Adicionado a cláusula [Else] para limpeza do objeto lblConta
    lblConta.Caption = ""
  End If
  
End Sub

Private Sub medDataAcerto_LostFocus()
  medDataAcerto.Text = Ajusta_Data(medDataAcerto.Text)
End Sub

Private Sub medDataAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      medDataAcerto.Text = frmCalendario.gsDateCalender(medDataAcerto.Text)
  End Select
End Sub

Private Sub medDataEmissao_LostFocus()
  medDataEmissao.Text = Ajusta_Data(medDataEmissao.Text)
End Sub

Private Sub medDataEmissao_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      medDataEmissao.Text = frmCalendario.gsDateCalender(medDataEmissao.Text)
  End Select
End Sub

Private Sub cboDigitador_Click()
  cboDigitador.Text = cboDigitador.Columns(1).Text
  cboDigitador_LostFocus
End Sub

'05/12/2005 - mpdea
'Incluído tratamento de erro
Private Sub cboDigitador_LostFocus()
  Dim rs As Recordset
  Dim sValor As String
  Dim lngRet As Long

  On Error GoTo ErrHandler
  
  '28/10/2002 - mpdea
  'Valida se há código preenchido
  If cboDigitador.Text = "" Then
    lblDigitador.Caption = ""
    Exit Sub
  End If
  
  '05/12/2005 - mpdea
  'Tratamento de overflow
  Call IsDataType(dtLong, cboDigitador.Text, lngRet)
  cboDigitador.Text = lngRet
  
  If Not IsNumeric(cboDigitador.Text) Then
     DisplayMsg "Digitador Inválido"
     cboDigitador.SetFocus
     Exit Sub
  End If
  If Len(cboDigitador.Text) > 0 Then
    sValor = ""
    Set rs = datDigitador.Recordset.Clone
    rs.FindFirst "Código = " & lngRet
    If Not rs.NoMatch Then
      sValor = rs("Nome")
    End If
    rs.Close
    Set rs = Nothing
    lblDigitador.Caption = sValor
  End If

  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboFornecedor_Click()
  cboFornecedor.Text = cboFornecedor.Columns(1).Text
  cboFornecedor_LostFocus
End Sub

'05/12/2005 - mpdea
'Incluído tratamento de erro
Private Sub cboFornecedor_LostFocus()
  Dim rs As Recordset
  Dim sValor As String
  Dim lngRet As Long

  On Error GoTo ErrHandler
  
  sValor = ""
  
  If Len(cboFornecedor.Text) > 0 Then
    If IsNumeric(cboFornecedor.Text) Then
      
      '05/12/2005 - mpdea
      'Tratamento de overflow
      Call IsDataType(dtLong, cboFornecedor.Text, lngRet)
      cboFornecedor.Text = lngRet
      
      Set rs = db.OpenRecordset("SELECT Nome, Estado FROM Cli_For WHERE Inativo = False And Código = " & lngRet, dbOpenDynaset, dbReadOnly)
      If Not rs.EOF Then
        sValor = rs("Nome")
        If Not IsNull(rs("Estado")) Then
           sEstado = rs("Estado")
        Else
           sEstado = ""
        End If
      End If
      rs.Close
      Set rs = Nothing
    End If
  End If
  
  '28/10/2002 - mpdea
  'Limpa descrição do fornecedor caso não haja código preenchido
  lblFornecedor.Caption = sValor

  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboOper_Click()
  
  cboOper.Text = cboOper.Columns(1).Text
  cboOper_LostFocus
  gsTipoOper = cboOper.Columns(2).Text
  gbSomarFrete = cboOper.Columns("Somar Frete ao Total").Text
  
  '30/03/2006 - mpdea
  'Comentado verificação de pagamento já realizada em cboOper_LostFocus
  'tabItens.TabEnabled(1) = (gsTipoOper = "C" Or gsTipoOper = "D")
  
  medDataAcerto.Enabled = (gsTipoOper = "E")
  
  Screen.MousePointer = vbHourglass
  
'''  If gsTipoOper = "D" Then
        datFornecedor.RecordSource = "SELECT Código, Nome, Tipo, Estado FROM Cli_For WHERE Inativo = False And (Tipo = 'F' OR Tipo = 'C') ORDER BY Nome"
'''  Else
'''    datFornecedor.RecordSource = "SELECT Código, Nome, Tipo, Estado FROM Cli_For WHERE Inativo = False And Tipo = 'F' ORDER BY Nome"
'''  End If
  datFornecedor.Refresh
  
  Screen.MousePointer = vbDefault
  
End Sub

'05/12/2005 - mpdea
'Incluído tratamento de erro
Private Sub cboOper_LostFocus()
  Dim sValor As String
  Dim lngRet As Long

  On Error GoTo ErrHandler
  
  Call StatusMsg("")
  
  sValor = ""
  
  If Len(cboOper.Text) > 0 Then
    
    If IsNumeric(cboOper.Text) Then
      
      '05/12/2005 - mpdea
      'Tratamento de overflow
      Call IsDataType(dtLong, cboOper.Text, lngRet)
      cboOper.Text = lngRet
      
      gbBaseICMSomadoIPI = False
      
      rsOp_Entrada.FindFirst "Código = " & lngRet
      If Not rsOp_Entrada.NoMatch Then
        sValor = rsOp_Entrada("Nome")
      End If
      
      lblOper.Caption = sValor
      
      '30/03/2006 - mpdea
      'Corrigido verificação de pagamento
'      If rsOp_Entrada("Tipo") = "C" Or rsOp_Entrada("Tipo") = "D" Then
      If rsOp_Entrada.fields("Dinheiro").Value Then
        cboCaixaUso.Enabled = gbCaixas
        txtCxDinheiro.Enabled = True
        txtCxCheque.Enabled = True
        cboConta.Enabled = True
        txtDescricao.Enabled = True
        txtCheque.Enabled = True
        txtValCheque.Enabled = True
        medDataBomPara.Enabled = True
        grdCP.Enabled = True
        tabItens.TabEnabled(1) = True
      Else
        cboCaixaUso.Enabled = False
        txtCxDinheiro.Enabled = False
        txtCxCheque.Enabled = False
        cboConta.Enabled = False
        txtDescricao.Enabled = False
        txtCheque.Enabled = False
        txtValCheque.Enabled = False
        medDataBomPara.Enabled = False
        grdCP.Enabled = False
        tabItens.TabEnabled(1) = False
      End If
      
      gbSomarFrete = rsOp_Entrada("Somar Frete ao Total").Value
      medDataAcerto.Visible = (rsOp_Entrada("Tipo") = "E")
      lblLabels(2).Visible = medDataAcerto.Visible
      
      gbBaseICMSomadoIPI = rsOp_Entrada("Base ICM com IPI")
      gbIPI = rsOp_Entrada("IPI")
      gbIPI_TOT = rsOp_Entrada("IPI TOT")
      
      Call Recalcula
        
    End If
  End If
  
  '28/10/2002 - mpdea
  'Limpa descrição da operação caso não haja código preenchido
  lblOper.Caption = sValor
  
  
  '23/08/2004 - Daniel
  'Verificar se a Operação possui
  'tabela para cálculo de índice financeiro
  Dim rstOpEntrada As Recordset
  Dim strQuery     As String
  
  If Len(cboOper.Text) <= 0 Then
    grdItens.Columns("Indice").Visible = False
    m_blnIndice = False
  
    Exit Sub
  End If
  
  '19/07/2007 - Anderson
  'Implementado campo PermitirAlterPreco para verificar as opções selecionadas na tela de cadastro de operações de entrada.
  'strQuery = "SELECT Código, Tabela "
  strQuery = "SELECT Código, Tabela, PermitirAlterPreco "
  strQuery = strQuery & " FROM [Operações Entrada] "
  strQuery = strQuery & " WHERE Código = " & CInt(cboOper.Text)
  
  Set rstOpEntrada = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstOpEntrada
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      '19/07/2007 - Anderson
      'Implementado campo PermitirAlterPreco para verificar as opções selecionadas na tela de cadastro de operações de entrada.
      'If Len(.Fields("Tabela").Value) > 0 Then
      If .fields("PermitirAlterPreco") = -1 And Len(.fields("Tabela").Value) > 0 Then
        grdItens.Columns("Indice").Visible = True
        m_blnIndice = True
      Else
        grdItens.Columns("Indice").Visible = False
        m_blnIndice = False
      End If
    End If
    .Close
  End With
  
  Set rstOpEntrada = Nothing
  
  If m_blnIndice Then
    '03/08/2005 - Daniel
    'Correção de ajuste da grid
    With grdItens
      .Columns("PrecoFinal").Width = 1124
      .Columns("Descricao").Width = 1725
    End With
  Else
    With grdItens
      .Columns("PrecoFinal").Width = 1468
      .Columns("Descricao").Width = 2068
    End With
  End If
  '------------------------------------------------------------
  
  '14/02/2005 - Daniel
  '
  'Solicitante: Daring - RJ
  '
  'Se ocorre devolução e esta devolução implica em abatimento de
  'comissão do vendedor, o Quick estava descontando erroneamente
  'da comissão para casos em que a venda possuia descontos.
'''  lblTabela.Visible = ControlarComisao
'''  cboTabela.Visible = ControlarComisao
  If ControlarComisao = True Then
      lbl_avisoTrataComissaoVendedor.Visible = True
  End If
  '-------------------------------------------------------------
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub ddwProduto_CloseUp()
'  If Mid(ddwProduto.Columns("Qtde").Text, 2, 1) <> "-" Then
'    ddwProduto.DataFieldList = "Código"
'  Else
'    ddwProduto.DataFieldList = ""
'  End If
  With ddwProduto
    grdItens.Columns("Codigo").Text = ddwProduto.Columns("Codigo").Text
    grdItens.Columns("Qtde").Text = "1"
    grdItens.Columns("Descricao").Text = .Columns("Nome").Text
    '05/05/2004 - Daniel
    'Personalização Embalavi
    If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
      grdItens.Columns("Preco").Text = Format(.Columns("Preco").Value, "##,###,##0.00000")
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      grdItens.Columns("Preco").Text = Format(.Columns("Preco").Value, "##,###,##0.000")
    Else
      grdItens.Columns("Preco").Text = gsFormatCurrency(.Columns("Preco").Value, True)
    End If
    
    grdItens.Columns("Unidade").Text = .Columns("Unidade").Text
'    .ActiveCell.SelStart = 0
'    .ActiveCell.SelLength = 0
'    DoEvents
    Call Calcula_Linha
'    DoEvents
'    .MoveRecords (0)
'    .Refresh
'    SendKeys "{Tab}"
'    .Col = 0
  End With
End Sub

Private Sub ddwProduto_RowLoaded(ByVal Bookmark As Variant)
  Dim nEstoque As Double
  Dim sMsgEstoque As String
  Dim nErro As Integer
  'Dim Aux_Preço As Currency
  '05/05/2004 - Daniel
  'Alterado o tipo de dado da var Aux_Preço para suportar
  '5 casas após a "," quando for Embalavi
  Dim Aux_Preço As Double

  
  '15/09/2005 - mpdea
  'Índice para cálculo do Preço de Entrada
  Dim dblIndicePrecoEntrada As Double


  With ddwProduto

    If rsParametros("VR Mostrar Estoque") = False Then
      .Columns("Qtde").Width = 1
      .Columns("Preco").Width = 1
      Exit Sub
    End If
    'Estoque
    nEstoque = Acha_Estoque(gnCodFilial, .Columns("Código").Text, 0, 0, 0, nErro)
    Select Case nErro
      Case 0
        sMsgEstoque = nEstoque
      Case 1
        sMsgEstoque = "1-Não iniciado"
      Case 2
        sMsgEstoque = "2-Com grade"
      Case 3
        sMsgEstoque = "3-Com edição"
      Case 4
        sMsgEstoque = "4-Não existe"
    End Select
    .Columns("Qtde").Text = sMsgEstoque
    'Acha Preço
    rsPrecos.Index = "Tabela"
    rsPrecos.Seek "=", "CUSTO", .Columns("Codigo").Text
    If rsPrecos.NoMatch Then
      grdItens.Columns("Preco").Text = 0
      .Columns("Preco").Text = "Preço não encontrado"
    Else
      
      
      '19/09/2005 - mpdea
      'Inclusão do Índice para cálculo do Preço de Entrada
      If g_blnIndicePrecoEntrada Then
        rsProdutos.Index = "Código"
        rsProdutos.Seek "=", .Columns("Código").Text
        If rsProdutos.NoMatch Then
          dblIndicePrecoEntrada = 1
        Else
          If Not IsDataType(dtDouble, rsProdutos.fields("IndicePrecoEntrada").Value, dblIndicePrecoEntrada) Then
            'Se o valor não for válido é igualado ao padrão 1
            dblIndicePrecoEntrada = 1
          End If
        End If
      End If
      
      
      '05/05/2004 - Daniel
      'Personalização Embalavi
      If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
        Aux_Preço = Format((rsPrecos("Preço")), "##,###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Aux_Preço = Format((rsPrecos("Preço")), "##,###,##0.000")
      Else
        Aux_Preço = rsPrecos("Preço")
      End If
      
      If rsProdutos("Moeda") <> 1 Then
        rsCotacoes.Index = "Moeda"
        rsCotacoes.Seek "<=", rsProdutos("Moeda"), Data_Atual
        If Not rsCotacoes.NoMatch Then
          If rsCotacoes("Moeda") = rsProdutos("Moeda") Then
            '05/05/2004 - Daniel
            'Personalização Embalavi
            If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
              Aux_Preço = Format((Aux_Preço * rsCotacoes("Cotação")), "##,###,##0.00000")
            '30/04/2007 - Anderson - Implementação de 3 casas decimais
            ElseIf g_bln3CasasDecimais Then
              Aux_Preço = Format((Aux_Preço * rsCotacoes("Cotação")), "##,###,##0.000")
            Else
              Aux_Preço = Aux_Preço * rsCotacoes("Cotação")
            End If
          End If
        End If
      End If
      
      '19/09/2005 - mpdea
      'Inclusão do Índice para cálculo do Preço de Entrada
      If g_blnIndicePrecoEntrada Then
        '05/05/2004 - Daniel
        'Personalização Embalavi
        If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
          .Columns("Preco").Text = Format(Aux_Preço * dblIndicePrecoEntrada, "##,###,##0.00000")
        '30/04/2007 - Anderson - Implementação de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
          .Columns("Preco").Text = Format(Aux_Preço * dblIndicePrecoEntrada, "##,###,##0.000")
        Else
          .Columns("Preco").Text = gsFormatCurrency(Aux_Preço * dblIndicePrecoEntrada, False)
        End If
      Else
        '05/05/2004 - Daniel
        'Personalização Embalavi
        If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
          .Columns("Preco").Text = Format(Aux_Preço, "##,###,##0.00000")
        '30/04/2007 - Anderson - Implementação de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
          .Columns("Preco").Text = Format(Aux_Preço, "##,###,##0.000")
        Else
          .Columns("Preco").Text = gsFormatCurrency(Aux_Preço, False)
        End If
      End If
      
    End If
    .Columns("Unidade").Text = rsProdutos("Unidade Venda").Value & ""
  End With
End Sub

Public Sub CheckMovimentacao()
  If Erro_Data2 Then
    Erro_Data2 = False
    If frmErroMov.gbContinue Then
      If Not frmGerente.gbSenhaGerente Then
        Unload Me
      End If
    Else
      Unload Me
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  '26/10/2005 - mpdea
  'Adicionado atalhos
  If Shift = vbCtrlMask Then
    Select Case KeyCode
      Case vbKeyO 'Transformar Pedido em Compra
        Call TransformaPedidoEmCompra
        KeyCode = 0: Shift = 0
        Exit Sub
        
      Case vbKeyE 'Encontrar Próximo Pedido do Fornecedor
        Call FindNextPedido
        KeyCode = 0: Shift = 0
        Exit Sub
        
      Case vbKeyF 'Imprimir Nota Fiscal
        Call PrintNotaEntrada
        KeyCode = 0: Shift = 0
        Exit Sub
        
      Case vbKeyI 'Imprimir Relatório
        Call PrintReport
        KeyCode = 0: Shift = 0
        Exit Sub
    End Select
  End If
  
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
On Error GoTo Erro

  Dim nRow As Long
  Dim sRet As String
  
  Call CenterForm(Me)
  
  KeyPreview = True

  Screen.MousePointer = vbHourglass
  
  On Error Resume Next
  Picture1.Picture = LoadPicture(App.Path & "\Imagens\logotipo.bmp")
  
  With ActiveBar1.Tools("miOpOrdem")
    .CBList.Clear
    .CBList.AddItem "Por Seqüência"
    .CBList.AddItem "Por Data e Seqüência"
    .CBList.AddItem "Por Fornecedor e Seqüência"
    .CBList.AddItem "Por Nota Fiscal"
    .Text = .CBList(0)
  End With
  With ActiveBar1
    .RecalcLayout
    .Refresh
  End With

  Call ActiveBarLoadToolTips(Me)

  sRet = GetSetting("QuickStore", "ConfigENTRADAS", "Scanner", False)
  ActiveBar1.Tools("miComplLeitorOtico").Checked = CBool(sRet)
 
  sRet = GetSetting("QuickStore", "ConfigENTRADAS", "Mantem Operacao", False)
  ActiveBar1.Tools("miOpFreezeOper").Checked = CBool(sRet)
 
  sRet = GetSetting("QuickStore", "ConfigENTRADAS", "Mantem Digitador", False)
  ActiveBar1.Tools("miOpFreezeDigitador").Checked = CBool(sRet)
 
  sRet = GetSetting("QuickStore", "ConfigENTRADAS", "Mantem Fornecedor", False)
  ActiveBar1.Tools("miOpFreezeFornecedor").Checked = CBool(sRet)
 
  sRet = GetSetting("QuickStore", "ConfigENTRADAS", "Mantem Forma Pagto", False)
  ActiveBar1.Tools("miOpFreezeFormaPagto").Checked = CBool(sRet)
  
  '20/06/2007 - Anderson
  'Exportar dados para excel. Customização Candy-Clean
  ActiveBar1.Tools("miOpExportarExcel").Visible = CheckSerialCaseMod("QS37957-281")
  
  datCaixasUso.DatabaseName = gsQuickDBFileName
  datConta.DatabaseName = gsQuickDBFileName
  datDigitador.DatabaseName = gsQuickDBFileName
  datEntrada.DatabaseName = gsQuickDBFileName
  datFornecedor.DatabaseName = gsQuickDBFileName
  datOper.DatabaseName = gsQuickDBFileName
  datProdutos.DatabaseName = gsQuickDBFileName
  datCentroCusto.DatabaseName = gsQuickDBFileName
  '14/02/2005 - Daniel
  datTabela.DatabaseName = gsQuickDBFileName
  
  '07/01/2010 - Andrea
  datCheques.DatabaseName = gsQuickDBFileName
  '----------------------------------------------

  gsSql = "SELECT * FROM Entradas WHERE Filial = " & gnCodFilial
  gsWhere = ""
  gsOrder = " ORDER BY Sequência"
  
  Set rsEntradas = db.OpenRecordset(gsSql & gsWhere & gsOrder, dbOpenDynaset)
  
  ' Trocar por .Clone e rever os Seeks para FindFirst....
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  
  Set rsOp_Entrada = db.OpenRecordset("Operações Entrada", dbOpenSnapshot, dbReadOnly)
  Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsMovi_Parcelas = db.OpenRecordset("Movimento - Parcelas")
  Set rsParametros = db.OpenRecordset("Parâmetros Filial")
  Set rsPrecos = db.OpenRecordset("Preços", , dbReadOnly)
  Set rsCotacoes = db.OpenRecordset("Cotações", , dbReadOnly)
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsEntradas2 = rsEntradas.Clone
  Set rsLog = db.OpenRecordset("ZZZLog")
  Set rsEstados = db.OpenRecordset("Estados")

  '07/01/2010 - Andrea
  Set rsMovi_Cheques = db.OpenRecordset("Movimento - Cheques")
  
  lblToday.Caption = Format$(Data_Atual, "dd/mm/yyyy")

  If gbGrade = False Then
    ActiveBar1.Tools("miComplTypeGrade").Visible = False
  End If

  If gbCaixas = False Then
    cboCaixaUso.Text = "1"
    cboCaixaUso.Enabled = False
  End If

  rsEntradas.FindFirst "Data > #" & Format(Data_Atual + 1, "mm/dd/yyyy") & "#"
  If Not rsEntradas.NoMatch Then
    Erro_Data2 = True
  End If

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then
    Screen.MousePointer = vbDefault
    DisplayMsg "Filial não encontrada"
    Exit Sub
  End If
  lblFilial.Caption = gnCodFilial & "-" & gsNomeFilial
  gbShowEstoque = rsParametros("VR Mostrar Estoque")
  
  cboCaixaUso.Enabled = gbCaixas
  
'  txtTotProdutos.TabStop = False
'  txtTotDescontos.TabStop = False
'  txtTotIPI.TabStop = False
'  txtTotBaseICMS.TabStop = False
'  txtTotValICMS.TabStop = False
'  txtTotBaseICMSSubst.TabStop = False
'  txtTotValICMSSubst.TabStop = False
  
  grdItens.RowHeight = 450.1418
  
  grdItens.StyleSets("Total").Font.Size = 12
  grdItens.StyleSets("Total").Font.Bold = True
  grdItens.StyleSets("Normal").Font.Size = 10
  grdItens.StyleSets("Normal").Font.Bold = True
  
  grdItens.Redraw = False
  grdItens.RemoveAll
  For nRow = 0 To 11
    grdItens.AddItem ""
  Next nRow
  grdItens.Redraw = True
  
  
  '13/12/2005 - mpdea
  'Padronizado esquema de 5 casas decimais na tela de Entradas
  '
  '29/11/2004 - Daniel
  'Criamos em Parâmetros um novo campo chamado Permitir5Casas
  'caso este campo estiver True configuraremos o preço unitário
  'para 5 casas após a vírgula
  m_bln_5CasasEntrada = rsParametros.fields("Permitir5Casas").Value
  
  
  With grdItens
    If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
      .Columns("Preco").NumberFormat = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Columns("Preco").NumberFormat = "##,###,##0.000"
    Else
      .Columns("Preco").NumberFormat = "##,###,##0.00"
    End If
  End With
  
  With ddwProduto
    If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
      .Columns("Preco").NumberFormat = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Columns("Preco").NumberFormat = "##,###,##0.000"
    Else
      .Columns("Preco").NumberFormat = "##,###,##0.00"
    End If
  End With
  '---------------------------------------------------------------------------------------------------------
  
  '03/08/2005 - Daniel
  'Correção de ajuste da grid
  With grdItens
    .Columns("PrecoFinal").Width = 1468
    .Columns("Descricao").Width = 2068
  End With
  
  
  '14/09/2004 - Daniel
  'Case.......: Livraria Resultado
  'Finalidade.: Monitorar tratamento para o campo em [Entradas - Produtos].QtdeAtual
  'm_blnResultado = CheckSerialCaseMod("QS40590-987")
  '---------------------------------------------------------------------------------------------------------
  
  '29/11/2004 - Daniel
  'Case: Cliente Teknika
  'm_blnTeknika = CheckSerialCaseMod("QS40966-243")

 
'--------  p 1

  '23/08/2004 - Daniel
  'Tratamento IndiceFinanceiro
  grdItens.Columns("Indice").Visible = False
  
  '14/02/2005 - Daniel
  '
  'Solicitante: Daring - RJ
  '
  'Se ocorre devolução e esta devolução implica em abatimento de
  'comissão do vendedor, o Quick estava descontando erroneamente
  'da comissão para casos em que a venda possuia descontos.
'''  lblTabela.Visible = False
'''  cboTabela.Visible = False
  lbl_avisoTrataComissaoVendedor.Visible = False
  '-------------------------------------------------------------
 
  '06/05/2005 - Daniel
  '
  'Implementação.: Trabalhar com o código para fornecedor cadastrado na tela de produtos.
  '                Impacto: Ao entrar com o código para o fornecedor no campo código do produto
  '                o sistema deverá trazer o código do produto que estiver amarrado nele
  'Solicitação...: Cristiano Pavinato - PSI RS
  m_blnUsaCodFornec = g_blnVerificarUsoCodFornece
  '-------------------------------------------------------------
  
 
  Me.Show
  DoEvents
  
  If bTelaChamadoraDevolucaoProdutos = True Then
      Call ClearScreen
      txtSeq.Text = sCodEntradaDevolucaoProdutos
      bTelaChamadoraDevolucaoProdutos = False
      MsgBox "Clique no ícone BINOCLO para localizar o detalhamento da Entrada de produtos por Devolução de cliente."
  Else
      Call ClearScreen
  End If
  
  ' Aba Movimentacao opção => Devoluções
  If bTelaChamadoraDevolucao_ValeCredito = True Then
      If bGerarDevolucaoAfentandoComissao = True Then
          cboOper.Text = -2 ' COM COMISSÃO
      Else
          cboOper.Text = -1 ' SEM COMISSÃO
      End If
      
      cboOper_LostFocus
  End If
  
  Screen.MousePointer = vbDefault
  
  
'--------  p 2
  '22/01/2003 - mpdea
  'Quick em modo limitado
  If Not gblnQuickFull Then
    'Centro de custo
    Label1.Visible = False
    cboCodigoCC.Visible = False
    lblNomeCC.Visible = False
    
    'Pagamento em cheque
    Frame3.Visible = False
'    lblLabels(14).Visible = False
'    lblLabels(15).Visible = False
'    lblLabels(16).Visible = False
'    lblLabels(17).Visible = False
'    lblLabels(18).Visible = False
'    lblLabels(28).Visible = False
    
    cboConta.Visible = False
    lblConta.Visible = False
    medDataBomPara.Visible = False
    txtDescricao.Visible = False
    txtCheque.Visible = False
    txtValCheque.Visible = False
    
    With ActiveBar1
      With .Bands("tbrComplem")
        .Tools("miComplPrintNotaFiscal").Visible = False
        .Tools("miComplTypeGrade").Visible = False
      End With
      .RecalcLayout
      .Refresh
    End With
  End If
  gravada = False
  
  Exit Sub
Erro:
  MsgBox "Erro ao carregar a tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  On Error Resume Next
  
  Call StatusMsg("")
  
  rsEntradas.Close
  rsProdutos.Close
  rsOp_Entrada.Close
  rsGrade.Close
  rsMovi_Parcelas.Close
  rsMovi_Cheques.Close
  rsParametros.Close
  rsPrecos.Close
  rsEntra_Prod.Close
  rsCotacoes.Close
  rsContas.Close
  rsEntradas2.Close
  rsLog.Close
  
  Set rsEntradas = Nothing
  Set rsProdutos = Nothing
  Set rsOp_Entrada = Nothing
  Set rsGrade = Nothing
  Set rsMovi_Parcelas = Nothing
  Set rsMovi_Cheques = Nothing
  Set rsParametros = Nothing
  Set rsPrecos = Nothing
  Set rsEntra_Prod = Nothing
  Set rsCotacoes = Nothing
  Set rsContas = Nothing
  Set rsEntradas2 = Nothing
  Set rsLog = Nothing
  
  On Error GoTo 0
  
End Sub


Private Sub grdCP_AfterUpdate(RtnDispErrMsg As Integer)
  Call RecalculaPagar
End Sub

Private Sub grdCP_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
On Error GoTo Erro

  Dim Aux As Variant
  
  Aux = grdCP.Columns(ColIndex).Text
  If IsNull(Aux) Then Aux = ""
  
  Call StatusMsg("")
  
  If ColIndex = 0 Then
    If Aux = "" Then Exit Sub
    
    If IsNumeric(Aux) Then
      If Val(Aux) > 0 And Val(Aux) < 500 Then
        If IsDate(medDataEmissao.Text) Then
           grdCP.Columns(ColIndex).Text = Format(CDate(medDataEmissao.Text) + Val(Aux), "dd/mm/yyyy")
        Else
           grdCP.Columns(ColIndex).Text = Data_Atual + Val(Aux)
        End If
        Aux = grdCP.Columns(ColIndex).Text
      End If
    End If
    
    If Not IsDate(Aux) Then
      DisplayMsg "Digite uma data."
      Cancel = True
      Exit Sub
    End If
    If CDate(Aux) < Data_Atual Then
      DisplayMsg "ATENÇÃO. Data anterior a data atual, possivelmente errada."
    End If
    grdCP.Columns(ColIndex).Text = Format(Aux, "dd/mm/yyyy")
  End If
  
  If ColIndex = 1 Then
    If Not IsNumeric(Aux) Then
      DisplayMsg "Digite um valor."
      Cancel = True
      Exit Sub
     End If
  End If
  
  Exit Sub
Erro:
    MsgBox "Inconsistência em data ou valor digitado " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub grdCP_LostFocus()
  grdCP.Update
End Sub

Public Sub Calcula_Linha()
 'Calcula preço total e final da linha
 Dim Preço_Total As Double
 Dim Preço_Final As Double
 Dim Qtde As Single
 Dim Preço As Double
 Dim Desconto As Single
 Dim Valor_Desconto As Double
 Dim IPI As Single
 Dim Valor_IPI As Double
 Dim Valor_Desonerado As Double
  
  '-----------------------------------------------------------------------------
  '15/09/2005 - mpdea
  'Índice para cálculo do Preço de Entrada
  Dim dblIndicePrecoEntrada As Double
  'Preço sem Índice para cálculo do Preço de Entrada
  Dim dblPrecoSemIndiceEntrada As Double
  'Preço Final sem Índice para cálculo do Preço de Entrada
  Dim dblPrecoFinalSemIndiceEntrada As Double
  '-----------------------------------------------------------------------------
  
 If grdItens.Columns(1).Text = "" Then grdItens.Columns(1).Text = 0
 If grdItens.Columns(4).Text = "" Then grdItens.Columns(4).Text = 0
 If grdItens.Columns(6).Text = "" Then grdItens.Columns(6).Text = 0
 If grdItens.Columns(7).Text = "" Then grdItens.Columns(7).Text = 0
 
  '-----------------------------------------------------------------------------
  '19/09/2005 - mpdea
  'Índice para cálculo do Preço de Entrada
  If g_blnIndicePrecoEntrada Then
    'Obtém o Índice para cálculo do Preço de Entrada
    Call IsDataType(dtDouble, grdItens.Columns("IndicePrecoEntrada").Text, dblIndicePrecoEntrada)
    'Preço sem Índice para cálculo do Preço de Entrada
    Call IsDataType(dtDouble, grdItens.Columns("PrecoSemIndiceEntrada").Text, dblPrecoSemIndiceEntrada)
  End If
  '-----------------------------------------------------------------------------
 
 Qtde = grdItens.Columns("Qtde").Text
 If Not IsNumeric(grdItens.Columns("Preco").Text) Then
  Preço = 0
 Else
  '05/05/2004 - Daniel
  'Personalização Embalavi
  If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
    Preço = gsHandleNull(Format((grdItens.Columns("Preco").Text & ""), "##,###,##0.00000"))
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    Preço = gsHandleNull(Format((grdItens.Columns("Preco").Text & ""), "##,###,##0.000"))
  Else
    Preço = gsHandleNull(grdItens.Columns("Preco").Text & "")
  End If
 End If
 
 If grdItens.Columns("Valor Desonerado").Text = "" Then
    grdItens.Columns("Valor Desonerado").Text = 0#
 End If
 
 If grdItens.Columns("% Diferimento").Text = "" Then
    grdItens.Columns("% Diferimento").Text = 0#
 End If
 
 Valor_Desonerado = Format((grdItens.Columns("Valor Desonerado").Text), "#0.00")
 
 If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
   grdItens.Columns("Preço").Text = Format(Preço, "##,###,##0.00000")
 '30/04/2007 - Anderson - Implementação de 3 casas decimais
 ElseIf g_bln3CasasDecimais Then
   grdItens.Columns("Preço").Text = Format(Preço, "##,###,##0.000")
 Else
   grdItens.Columns("Preço").Text = Format(Preço, FORMAT_VALUE)   'gsFormatCurrency(Preço_Total, True)
 End If
 
 Desconto = gsHandleNull(grdItens.Columns("Desconto").Text & "")
 IPI = gsHandleNull(grdItens.Columns("IPI").Text & "")
 
 Preço_Total = Qtde * Preço
  
 grdItens.Columns("PrecoTotal").Text = Format(Preço_Total, FORMAT_VALUE)   'gsFormatCurrency(Preço_Total, True)
  
 Valor_Desconto = (Preço_Total * Desconto / 100)
 Preço_Final = (Preço_Total - Valor_Desconto)
 
 If Not gbIPI Then
    Valor_IPI = 0
 Else
 
    '-----------------------------------------------------------------------------
    '19/09/2005 - mpdea
    'Índice para cálculo do Preço de Entrada
    'Valor do IPI é calculado sobre o Preço Unitário
    'sem ter indexado pelo Índice para cálculo do Preço de Entrada
    If g_blnIndicePrecoEntrada Then
      Preço_Total = Qtde * dblPrecoSemIndiceEntrada
      Valor_Desconto = (Preço_Total * Desconto / 100)
      dblPrecoFinalSemIndiceEntrada = (Preço_Total - Valor_Desconto)
      Valor_IPI = (dblPrecoFinalSemIndiceEntrada * IPI / 100)
    Else
      '04/12/2006 - Anderson
      'Verifica se ICMS incide IPI
      If rsOp_Entrada.fields("ICMSSobreIPI").Value Then
       Valor_IPI = ((Preço_Total + CDbl(gsHandleNull(txtFrete.Text))) * (CDbl(gsHandleNull(grdItens.Columns("IPI").Text)) / 100))
      Else
       Valor_IPI = (Preço_Final * IPI / 100)
      End If
    End If
    '-----------------------------------------------------------------------------
    
 End If
  
 If Not gbIPI_TOT Then
    Preço_Final = (Preço_Final + Valor_IPI)
 End If

 grdItens.Columns("PrecoFinal").Text = Format(Preço_Final, FORMAT_VALUE)  'gsFormatCurrency(Preço_Final, True)
 
End Sub

Sub Calcula_Tabela(Linha As Integer)
  Rem Calcula preço total e final da linha
  Dim Preço_Total As Double
  Dim Preço_Final As Double
  Dim Qtde As Single
  Dim Preço As Single
  Dim Desconto As Single
  Dim Valor_Desconto As Single
  Dim IPI As Single
  Dim Valor_IPI As Single
  Dim bm As Variant
  
  bm = grdItens.AddItemBookmark(grdItens.Row)
  Qtde = grdItens.Columns("Qtde").CellText(bm)
  '05/05/2004 - Daniel
  'Personalização Embalavi
  If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
    Preço = Format((grdItens.Columns("Preco").CellText(bm)), "##,###,##0.00000")
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    Preço = Format((grdItens.Columns("Preco").CellText(bm)), "##,###,##0.000")
  Else
    Preço = grdItens.Columns("Preco").CellText(bm)
  End If
  
  Desconto = grdItens.Columns("Desconto").CellText(bm)
  IPI = grdItens.Columns("IPI").CellText(bm)
  
  Preço_Total = gsFormatCurrency((Qtde * Preço), gnCurrencyDecimals)
  grdItens.Columns("PrecoTotal").Text = Preço_Total
  
  Valor_Desconto = gsFormatCurrency((Preço_Total * Desconto / 100), gnCurrencyDecimals)
  Preço_Final = gsFormatCurrency((Preço_Total - Valor_Desconto), gnCurrencyDecimals)
  Valor_IPI = gsFormatCurrency((Preço_Final * IPI / 100), gnCurrencyDecimals)
  Preço_Final = gsFormatCurrency((Preço_Final + Valor_IPI), gnCurrencyDecimals)
  
  grdItens.Columns("PrecoTotal").Text = Preço_Total
  
End Sub

Private Sub grdItens_AfterUpdate(RtnDispErrMsg As Integer)
  Call Recalcula
  Call RecalculaPagar
End Sub

Public Sub grdItens_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  Dim Cód As String
  Dim Valor As Single
  Dim Valor_Int As Long
  Dim Aux_Preço As Double
  Dim Código As String
  Dim Produto As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Tipo As Integer
  Dim Erro As Integer
  Dim Aux_Str As String
  Dim sCurrCol As String
  Dim nCol As Integer
  Static sCodProd As String
  
  '15/09/2005 - mpdea
  'Índice para cálculo do Preço de Entrada
  Dim dblIndicePrecoEntrada As Double
  
  '17/05/2006 - mpdea
  'Índice ICMS Retido - Base complementar
  Dim dblIndiceIcmsRetido As Double
  
  '24/04/2007 - Anderson
  'Inclusão de código para resolver problema ao digitar um código do fornecedor igual ao código do produto
  Dim rstProdutos As Recordset
  Dim strSQL      As String
  Dim intI       As Integer
  Dim bytAsc     As Byte
  Dim strConcat  As String
  
  Cancel = False
  
  Call StatusMsg("")
  
  sCurrCol = Trim(grdItens.Columns(ColIndex).Text)

  ' 24/04/2007 - Anderson
  '---[ Loop criado para retirar o ENTER e TAB da variavel sCurrCol na emissão de nota fiscal ]---'
    For intI = 1 To Len(sCurrCol)
      bytAsc = Asc(Mid(sCurrCol, intI, 1))
      
      If Not (bytAsc = 13 Or bytAsc = 10) Then
        strConcat = strConcat & Chr(bytAsc)
      End If
    Next intI
  '---[ Loop criado para retirar o ENTER e TAB da variavel sCurrCol na emissão de nota fiscal ]---'
  
  sCurrCol = strConcat

  Select Case grdItens.Columns(ColIndex).Name
  
    Case "Codigo"
 
'      If sCurrCol = "0" Then
'        DisplayMsg "Produto não existe."
'        SendKeys "{Home}+{End}"
'        Cancel = True
'        Exit Sub
'      End If

'      If IsNumeric(sCurrCol) Then
'         If Val(sCurrCol) = 0 Then
'            grdItens.Columns(nCol).Text = ""
'            Call Calcula_Linha
'            Exit Sub
'         End If
'      End If

      '---------------------------------------------------------------------------------------------
      '06/05/2005 - Daniel
      '
      'Implementação.: Trabalhar com o código para fornecedor cadastrado na tela de produtos.
      '                Impacto: Ao entrar com o código para o fornecedor no campo código do produto
      '                o sistema deverá trazer o código do produto que estiver amarrado nele
      'Solicitação...: Cristiano Pavinato - PSI RS
      
      '-----------------------
      '24/04/2007 - Anderson
      'Inclusão de código para resolver problema ao digitar um código do fornecedor igual ao código do produto
      If m_blnUsaCodFornec Then
        Dim strCodParaFornec As String
  
        strSQL = "SELECT Código, [Código do Fornecedor] FROM Produtos WHERE Código = '" & sCurrCol & "'"
        
        Set rstProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
        
        If rstProdutos.RecordCount = 0 Then
          strCodParaFornec = sCurrCol
        Else
          '04/05/2009 - mpdea
          'Corrigido RT-94 (Invalid use of the null)
          strCodParaFornec = rstProdutos("Código do Fornecedor") & ""
        End If
        
        Set rstProdutos = Nothing
        
        If Not (strCodParaFornec = "0" Or strCodParaFornec = "") Then
          sCurrCol = g_strBuscarCodProd(strCodParaFornec)
          grdItens.Columns(ColIndex).Text = sCurrCol
          
          If sCurrCol = "" Then
            sCurrCol = strCodParaFornec
            Exit Sub
          End If
        End If
        '24/04/2007 - Anderson
        'Retirado, para readequação do código acima.
        ''Dim strCodParaFornec As String

        'strCodParaFornec = Trim(grdItens.Columns(ColIndex).Text)
        'sCurrCol = g_strBuscarCodProd(strCodParaFornec)
        'grdItens.Columns(ColIndex).Text = sCurrCol

        ''Se não existir nenhum produto amarrado
        'If sCurrCol = strCodParaFornec Then Exit Sub
      End If
      '---------------------------------------------------------------------------------------------
      
      If IsNull(sCurrCol) Or sCurrCol = "" Or sCurrCol = "0" Then
'        ddwProduto.DataFieldList = "Codigo"
        For nCol = 0 To grdItens.Cols - 1
          If grdItens.Columns(nCol).DataType = 8 Then
            grdItens.Columns(nCol).Text = ""
          Else
            If grdItens.Columns(nCol).DataType = 11 Then
              grdItens.Columns(nCol).Text = False
            Else
              grdItens.Columns(nCol).Value = 0
            End If
          End If
        Next nCol
        Exit Sub
      End If
  
      sCurrCol = UCase(sCurrCol)
      '26/05/2004 - Daniel
      'Tratamento para 0 'zero' a esquerda
      If Not gbZeroEsquerda Then
        sCurrCol = Retira_Zeros(sCurrCol)
      End If
      grdItens.Columns(ColIndex).Text = sCurrCol

      Call Acha_Produto(sCurrCol, Produto, Tamanho, Cor, Edição, Tipo, Erro)
      Select Case Erro
        Case 1
          DisplayMsg "Produto não existe."
          SendKeys "{RIGHT}"
          SendKeys "{LEFT}"
          Cancel = True
          Exit Sub
        Case 2
          DisplayMsg "Este produto usa grade. Digite o código completo: inicial(5 a 8 dígitos), mais tamanho(3 dígitos) e cor(3 dígitos)."
          Cancel = True
          Exit Sub
        Case 3
          DisplayMsg "Este produto usa edição. Digite o código completo: inicial(13 dígitos), com a edição(5 dígitos)"
          Cancel = True
          Exit Sub
        Case 4
          DisplayMsg "Código inválido."
          Cancel = True
          Exit Sub
      End Select
  
      rsProdutos.Index = "Código"
      rsProdutos.Seek "=", Produto
  
      If rsProdutos.fields("Desativado") Then
        MsgBox "Produto Inativo, verifique !", vbCritical, "Quick Store"
        grdItens.Columns(0).Text = ""
        grdItens.Columns(1).Text = ""
        Exit Sub
      End If
  
      '24/08/2004 - Daniel
      'Tratamento de IndiceFinanceiro
      If m_blnIndice Then
        If IsNumeric(rsProdutos("IndiceFinanceiro").Value) Then grdItens.Columns("Indice").Text = Format((rsProdutos("IndiceFinanceiro").Value), "##,###,##0.000")
      End If
      grdItens.Columns("Descricao").Text = rsProdutos("Nome")
      grdItens.Columns("Unidade").Text = rsProdutos("Unidade Venda") & ""
      grdItens.Columns("IPI").Text = gsHandleNull(rsProdutos("Percentual IPI"))
      grdItens.Columns("Desconto").Text = 0#
      
      
      '18/05/2006 - mpdea
      'ICMS de Saída para cálculos
      grdItens.Columns("IcmsSaida").Text = gsHandleNull(rsProdutos("Percentual ICM Saida"))
      
      
      If sEstado = "" Then
         grdItens.Columns("ICMS").Text = gsHandleNull(rsProdutos("Percentual ICM Entrada"))
      ElseIf sEstado <> "" Then
          rsEstados.Index = "Estado"
          rsEstados.Seek "=", sEstado
          If rsEstados.NoMatch Then
             grdItens.Columns("ICMS").Text = gsHandleNull(rsProdutos("Percentual ICM Entrada"))
          ElseIf Not rsEstados.NoMatch Then
             If rsEstados("ICM") = -1 Then
                grdItens.Columns("ICMS").Text = gsHandleNull(rsProdutos("Percentual ICM Entrada"))
             Else
                grdItens.Columns("ICMS").Text = rsEstados("ICM")
             End If
          End If
     End If
      
       'Mostra ICM do Estado
'      If Estado = "" Then
'        .Columns("ICM").Text = rsProdutos2("Percentual ICM")
'      ElseIf Estado <> "" Then
'        rsEstados.Index = "Estado"
'        rsEstados.Seek "=", Estado
'        If rsEstados.NoMatch Then
'          .Columns("ICM").Text = rsProdutos2("Percentual ICM")
'        ElseIf Not rsEstados.NoMatch Then
'          If rsEstados("ICM") = -1 Then
'             .Columns("ICM").Text = rsProdutos2("Percentual ICM")
'          Else
'             .Columns("ICM").Text = rsEstados("ICM")
'          End If
'        End If
'      End If

      
      
      
      '----------------------------------------------------------------------------
      '19/09/2005 - mpdea
      'Inclusão do Índice para cálculo do Preço de Entrada
      If g_blnIndicePrecoEntrada Then
        If Not IsDataType(dtDouble, rsProdutos.fields("IndicePrecoEntrada").Value, dblIndicePrecoEntrada) Then
          'Se o valor não for válido é igualado ao padrão 1
          dblIndicePrecoEntrada = 1
        End If
        grdItens.Columns("IndicePrecoEntrada").Text = dblIndicePrecoEntrada
      End If
      '----------------------------------------------------------------------------
      
      
      
      '<<<<<<<<<<<<
      With grdItens
        .Columns("Base_ICM").Text = 0
        .Columns("Valor_ICM").Text = 0
        .Columns("Valor_Base_Unit").Text = 0
        .Columns("Redução_ICM").Text = 0
        .Columns("Tipo_ICM").Text = rsProdutos("Tipo ICM") & ""
        Select Case rsProdutos("Tipo ICM")
          Case "I"
            .Columns("ICMS").Text = "0"
          
          Case "R" 'ICM Retido
            If rsProdutos("Base Cálculo") <> 0 Then
              .Columns("Valor_Base_Unit").Text = rsProdutos("Base Cálculo")
            End If
            If rsProdutos("Redução ICM") <> 0 Then
              .Columns("Redução_ICM").Text = rsProdutos("Redução ICM")
            End If
            
            '17/05/2006 - mpdea
            'Índice ICMS Retido - Base complementar
            Call IsDataType(dtDouble, rsProdutos("IndiceIcmsRetido"), dblIndiceIcmsRetido)
            .Columns("IndiceIcmsRetido").Text = dblIndiceIcmsRetido
            
          Case "Z" 'ICM Reduzido
            If rsProdutos("Base Cálculo") <> 0 Then
              .Columns("Valor_Base_Unit").Text = rsProdutos("Base Cálculo")
            End If
            If rsProdutos("Redução ICM") <> 0 Then
              .Columns("Redução_ICM").Text = rsProdutos("Redução ICM")
            End If
        End Select
      End With
      '>>>>>>>>>>>>
      
      'Acha preço
      rsPrecos.Index = "Tabela"
      '10/08/2005 - Daniel
      'Tratamento para a Busca pelo preço da tabela informada: Ora pelo "CUSTO" como sempre existiu
      'no Quick ora pela "Tabela de Venda" informada pelo usuário no momento da Devolução configurada
      'para abatimento da comissão do vendedor
      If Len(cboTabela.Text) > 0 Then
        rsPrecos.Seek "=", Trim(cboTabela.Text), Produto
      Else
        rsPrecos.Seek "=", "CUSTO", Produto
      End If
      '-------------------------------------------------
      If rsPrecos.NoMatch Then
        grdItens.Columns("Preco").Text = 0
      End If
      If Not rsPrecos.NoMatch Then
      
        '05/05/2004 - Daniel
        'Personalização Embalavi
        If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
          Aux_Preço = Format((rsPrecos("Preço")), "##,###,##0.00000")
        '30/04/2007 - Anderson - Implementação de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
          Aux_Preço = Format((rsPrecos("Preço")), "##,###,##0.000")
        Else
          Aux_Preço = rsPrecos("Preço")
        End If
        
        If rsProdutos("Moeda") <> 1 Then
          rsCotacoes.Index = "Moeda"
          rsCotacoes.Seek "<=", rsProdutos("Moeda"), Data_Atual
          If Not rsCotacoes.NoMatch Then
            If rsCotacoes("Moeda") = rsProdutos("Moeda") Then
              '05/05/2004 - Daniel
              'Personalização Embalavi
              If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
                Aux_Preço = Format((Aux_Preço * rsCotacoes("Cotação")), "##,###,##0.00000")
              '30/04/2007 - Anderson - Implementação de 3 casas decimais
              ElseIf g_bln3CasasDecimais Then
                Aux_Preço = Format((Aux_Preço * rsCotacoes("Cotação")), "##,###,##0.000")
              Else
                Aux_Preço = Aux_Preço * rsCotacoes("Cotação")
              End If
            End If
          End If
        End If
        
        '19/09/2005 - mpdea
        'Índice para cálculo do Preço de Entrada
        If g_blnIndicePrecoEntrada Then
          '05/05/2004 - Daniel
          'Personalização Embalavi
          If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
            'Preço unitário sem o Índice para cálculo do Preço de Entrada
            grdItens.Columns("PrecoSemIndiceEntrada").Text = Format(Aux_Preço, "##,###,##0.00000")
            'Preço unitário com o Índice para cálculo do Preço de Entrada
            grdItens.Columns("Preco").Text = Format(Aux_Preço * dblIndicePrecoEntrada, "##,###,##0.00000") 'gsFormatCurrency(Aux_Preço, gnCurrencyDecimals)
          '30/04/2007 - Anderson - Implementação de 3 casas decimais
          ElseIf g_bln3CasasDecimais Then
            'Preço unitário sem o Índice para cálculo do Preço de Entrada
            grdItens.Columns("PrecoSemIndiceEntrada").Text = Format(Aux_Preço, "##,###,##0.000")
            'Preço unitário com o Índice para cálculo do Preço de Entrada
            grdItens.Columns("Preco").Text = Format(Aux_Preço * dblIndicePrecoEntrada, "##,###,##0.000") 'gsFormatCurrency(Aux_Preço, gnCurrencyDecimals)
          Else
            'Preço unitário sem o Índice para cálculo do Preço de Entrada
            grdItens.Columns("PrecoSemIndiceEntrada").Text = Format(Aux_Preço, FORMAT_VALUE)
            'Preço unitário com o Índice para cálculo do Preço de Entrada
            grdItens.Columns("Preco").Text = Format(Aux_Preço * dblIndicePrecoEntrada, FORMAT_VALUE) 'gsFormatCurrency(Aux_Preço, gnCurrencyDecimals)
          End If
        Else
          '05/05/2004 - Daniel
          'Personalização Embalavi
          If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
            grdItens.Columns("Preco").Text = Format(Aux_Preço, "##,###,##0.00000")  'gsFormatCurrency(Aux_Preço, gnCurrencyDecimals)
          '30/04/2007 - Anderson - Implementação de 3 casas decimais
          ElseIf g_bln3CasasDecimais Then
            grdItens.Columns("Preco").Text = Format(Aux_Preço, "##,###,##0.000")  'gsFormatCurrency(Aux_Preço, gnCurrencyDecimals)
          Else
            grdItens.Columns("Preco").Text = Format(Aux_Preço, FORMAT_VALUE)  'gsFormatCurrency(Aux_Preço, gnCurrencyDecimals)
          End If
        End If
      End If

    Case "Qtde"
      sCodProd = ""
      'Verifica se Qtde é decimal
      If IsNull(sCurrCol) Then
         grdItens.Columns(1).Text = 0
         Call Calcula_Linha
         Exit Sub
      End If
      If sCurrCol = "" Then
         grdItens.Columns(1).Text = 0
         Call Calcula_Linha
      End If
  
      If Not IsNumeric(sCurrCol) Then
         DisplayMsg "Quantidade inválida."
         Cancel = True
         Exit Sub
      End If
      If CDbl(sCurrCol) <= 0 Then
        If grdItens.Columns(0).Text <> "" Then
          DisplayMsg "Quantidade incorreta."
          Cancel = True
        End If
        Exit Sub
      End If
      If CDbl(sCurrCol) > 9999999 Then
         DisplayMsg "Quantidade incorreta."
         Cancel = True
         Exit Sub
      End If
  
      Valor = sCurrCol
      Valor_Int = sCurrCol
      If Valor = Valor_Int Then
        Call Calcula_Linha
        Exit Sub
      End If
  
      Aux = grdItens.Columns(0).Text
      'Acha produto
  
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
  
      Call Calcula_Linha
    
    Case "Preco"
      sCodProd = ""
      If IsNull(sCurrCol) Then
         grdItens.Columns(4).Text = 0
         Call Calcula_Linha
         Exit Sub
      End If
  
      If sCurrCol = "" Then
         grdItens.Columns(4).Text = 0
         Call Calcula_Linha
         Exit Sub
      End If
  
      If Not IsNumeric(sCurrCol) Then
        DisplayMsg "Preço incorreto."
        Cancel = True
        Exit Sub
      End If
      If CDbl(sCurrCol) < 0 Then
        DisplayMsg "Preço não pode ser menor que 0."
        Cancel = True
        Exit Sub
      End If
      If CDbl(sCurrCol) > 9999999.99 Then
        DisplayMsg "Preço incorreto, máximo é 9.999.999,99"
        Cancel = True
        Exit Sub
      End If
      
      '19/09/2005 - mpdea
      'Índice para cálculo do Preço de Entrada
      'Aplica o índice quando há alteração de preço
      If g_blnIndicePrecoEntrada Then
        Call IsDataType(dtDouble, grdItens.Columns("IndicePrecoEntrada").Text, dblIndicePrecoEntrada)
        Call IsDataType(dtDouble, grdItens.Columns("Preço").Text, Aux_Preço)
        grdItens.Columns("PrecoSemIndiceEntrada").Text = Aux_Preço
        grdItens.Columns("Preço").Text = Aux_Preço * dblIndicePrecoEntrada
      End If
      
      Call Calcula_Linha

    Case "Desconto"
      sCodProd = ""
      If IsNull(sCurrCol) Then
         grdItens.Columns(8).Text = 0
         Call Calcula_Linha
         Exit Sub
      End If
  
      If sCurrCol = "" Then
         grdItens.Columns(8).Text = 0
         Call Calcula_Linha
         Exit Sub
      End If
  
      If Not IsNumeric(sCurrCol) Then
        DisplayMsg "Desconto incorreto."
        Cancel = True
        Exit Sub
      End If
      If CDbl(sCurrCol) < 0 Or CDbl(sCurrCol) > 100 Then
        DisplayMsg "Desconto não pode ser menor que 0 ou maior que 100."
        Cancel = True
        Exit Sub
      End If
      Call Calcula_Linha
  
    Case "ICMS"
      sCodProd = ""
      If IsNull(sCurrCol) Then
         grdItens.Columns(6).Text = 0
         Exit Sub
      End If
  
      If sCurrCol = "" Then
         grdItens.Columns(6).Text = 0
         Exit Sub
      End If
  
      If Not IsNumeric(sCurrCol) Then
        DisplayMsg "ICM incorreto."
        Cancel = True
        Exit Sub
      End If
      If CDbl(sCurrCol) < 0 Or CDbl(sCurrCol) > 999 Then
         DisplayMsg "ICM incorreto, deve ser entre 0 e 999."
         Cancel = True
         Exit Sub
      End If
  
      Call Calcula_Linha

    Case "IPI"
      sCodProd = ""
      If IsNull(sCurrCol) Then
         grdItens.Columns(7).Text = 0
         Call Calcula_Linha
         Exit Sub
      End If
  
      If sCurrCol = "" Then
         grdItens.Columns(7).Text = 0
         Call Calcula_Linha
         Exit Sub
      End If
  
      If Not IsNumeric(sCurrCol) Then
        DisplayMsg "IPI incorreto."
        Cancel = True
        Exit Sub
      End If
      If CDbl(sCurrCol) < 0 Or CDbl(sCurrCol) > 999 Then
         DisplayMsg "IPI incorreto, deve ser entre 0 e 999."
         Cancel = True
         Exit Sub
      End If
  
      Call Calcula_Linha

  End Select

End Sub

Private Sub grdItens_GotFocus()
  grdItens.Col = 0
End Sub

Private Sub grdItens_InitColumnProps()
  '05/05/2004 - Daniel
  'Personalização Embalavi
  If g_bln5CasasDecimais Or m_bln_5CasasEntrada Then
    grdItens.Columns("Preco").NumberFormat = "##,###,##0.00000"   'gsCurrencyFormat
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    grdItens.Columns("Preco").NumberFormat = "##,###,##0.000"   'gsCurrencyFormat
  Else
    grdItens.Columns("Preco").NumberFormat = FORMAT_VALUE   'gsCurrencyFormat
  End If
  
  grdItens.Columns("PrecoTotal").NumberFormat = FORMAT_VALUE   'gsCurrencyFormat
  grdItens.Columns("PrecoFinal").NumberFormat = FORMAT_VALUE 'gsCurrencyFormat
  grdItens.Columns("Codigo").DropDownHwnd = ddwProduto.hwnd
  grdItens.Columns("Unidade").Locked = True
  ddwProduto.DataFieldList = "Codigo"
End Sub

Private Sub grdItens_KeyPress(KeyAscii As Integer)
  If Len(grdItens.Columns("Código").Text) > 0 Then
    If Asc(grdItens.Columns("Código").Text) = 13 Then grdItens.Columns("Código").Text = ""
  End If
 With ddwProduto
    If grdItens.Col = 0 Then
      If .DroppedDown Then
        .DataFieldToDisplay = "Nome"
      End If
      If KeyAscii = vbKeyReturn Then
        If ActiveBar1.Tools("miComplLeitorOtico").Checked And Not .DroppedDown Then
'          If IsNumeric(grdItens.Columns("Codigo").Text) Then
'             If Val(grdItens.Columns("Codigo").Text) = 0 Then
'                Exit Sub
'             End If
'          End If
'          If grdItens.Columns("Codigo").Text <> "0" Then
'            grdItens.Columns("Qtde").Text = 1
'            SendKeys "{DOWN}{HOME}", True
'            grdItens.Columns("Código").Text = ""
'            KeyAscii = 0
'          End If
          If grdItens.Columns("Codigo").Text <> "" And grdItens.Columns("Codigo").Text <> "0" Then
              grdItens.Columns("Qtde").Text = 1
              'Grade1.SetFocus '07/10/2004 - Daniel - Case: Sucupira
              SendKeys "{DOWN}{HOME}", True
              
              '07/10/2004 - Daniel - Case: Sucupira
              If grdItens.Columns("Código").Text = "" Then
                grdItens.Col = 0
                'grdItens.SetFocus '21/10/2004 - Forçar o foco
              Else
                SendKeys "{HOME}", True
                'Grade1.SetFocus '21/10/2004 - Forçar o foco {comentada esta linha em 29/11/2004 - Daniel}
              End If
              
              '27/07/2004 - mpdea
              'Comentado devido a perda de performance da busca
              'pela lista de produtos (permanece como 0 - zero)
              '.Text = "" 'Replace(Grade1.Columns("Código").Text, Chr(13), "")
              
              KeyAscii = 0
            
            End If
        End If
      End If
    End If
  End With

' If grdItens.Col = 0 Then
'   If KeyAscii = 13 Then  'enter
'     If ActiveBar1.Tools("miComplLeitorOtico").Checked = True Then
'
'       grdItens.Columns(1).Text = 1 'qtde
'       SendKeys "{END}"
'       DoEvents
'       SendKeys "{RIGHT}"
'       DoEvents
'       SendKeys "{LEFT}"
'       DoEvents
'       SendKeys "{DOWN}"
'       KeyAscii = 0
'     End If
'   Else
'     If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
'       ddwProduto.DataFieldList = ""
'     End If
'   End If
'
' End If

End Sub

Private Sub grdItens_LostFocus()
  If grdItens.RowChanged = True Then
    grdItens.Update
  End If
End Sub

Private Sub grdItens_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'  Dim sCodProd As String
'  Dim Tamanho As Integer
'  Dim Cor As Integer
'  Dim Edição As Long
'  Dim Tipo As Integer
'  Dim Erro As Integer
'  Dim bm As Variant
  
  On Error GoTo ErrHandler
  
  grdItens.SetFocus
  grdItens.Col = grdItens.ColContaining(X, y)
  If grdItens.Col = 0 Then
''    grdItens.Row = grdItens.RowContaining(Y)
''    bm = grdItens.AddItemBookmark(grdItens.Row)
''    sCodProd = grdItens.Columns(0).CellText(bm)
'    If IsNumeric(sCodProd) Then
'      ddwProduto.DataFieldList = "Código"
'    Else
'      ddwProduto.DataFieldList = "Nome"
'    End If
  End If
  Exit Sub
  
ErrHandler:
  Exit Sub
End Sub

Private Sub txtCxDinheiro_GotFocus()
  txtCxDinheiro.SelStart = 0
  txtCxDinheiro.SelLength = Len(txtCxDinheiro.Text)
End Sub

Private Sub txtCxDinheiro_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub txtCxDinheiro_LostFocus()
  If IsGoodNumber(txtCxDinheiro) Then
    Call RecalculaPagar
  End If
End Sub

Private Sub txtFrete_GotFocus()
  Call SelectAllText(txtFrete)
End Sub

Private Sub txtFrete_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub txtFrete_LostFocus()
  If IsGoodNumber(txtFrete) Then
    Call Recalcula
  End If
End Sub

Private Sub txtNF_GotFocus()
  Call SelectAllText(txtNF)
End Sub

Private Sub txtNF_LostFocus()
  If IsNull(txtNF.Text) Then txtNF.Text = ""
  txtNF.Text = UCase(txtNF.Text)
  If txtNF.Text = "" Then Exit Sub
  '----------------------------------------------------------------
  '19/05/2005 - Daniel
  'Tratamento para não aparecer a mensagem que a entrada já existe
  'pois a gravação está ocorrendo antes da perda do foco
  If m_blnFocoNF Then
    m_blnFocoNF = False 'Volta ao estado normal
    Exit Sub
  End If
  '----------------------------------------------------------------
  If Len(cboOper.Text) > 0 Then
    If Len(cboFornecedor.Text) > 0 Then
      If rsOp_Entrada("Tipo") = "C" Then
        rsEntradas2.FindFirst "[Nota Fiscal] = '" & txtNF.Text & "' And Fornecedor = " & Val(cboFornecedor.Text)
        If Not rsEntradas2.NoMatch Then
          MsgBox "Atenção. Esta compra já foi lançada no dia " + str(rsEntradas2("Data")) + " como o número de seqüência " + str(rsEntradas2("Sequência"))
          Exit Sub
        End If
      End If
    End If
  End If
  
End Sub

Private Sub txtNrSerie_LostFocus()
  '19/05/2005 - Daniel
  txtNrSerie.Text = UCase(txtNrSerie.Text & "")
End Sub

Private Sub txtObs_GotFocus()
  Call SelectAllText(txtObs)
End Sub

Private Sub txtPedido_GotFocus()
  Call SelectAllText(txtPedido)
End Sub

Private Sub txtPedido_LostFocus()
  '19/05/2005 - Daniel
  txtPedido.Text = UCase(txtPedido.Text & "")
End Sub

Private Sub txtSeq_GotFocus()
  Call SelectAllText(txtSeq)
End Sub

Private Sub txtTotalAPagar_GotFocus()
'  SendKeys "{Tab}"
  Call SelectAllText(txtTotalAPagar)
End Sub

Private Sub txtTotalAPagar_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub txtTotalAPagar_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(txtTotalAPagar)
End Sub

Private Sub txtTotBaseICMS_GotFocus()
'  If txtTotBaseICMS.TabStop = False Then
'    tabItens.SetFocus
'  End If
  Call SelectAllText(txtTotBaseICMS)
End Sub

Private Sub txtTotBaseICMS_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub txtTotBaseICMS_LostFocus()
  txtTotBaseICMS.Text = gsFormatCurrency(txtTotBaseICMS.Text, gnCurrencyDecimals)
End Sub

Private Sub txtTotBaseICMSSubst_GotFocus()
'  If txtTotBaseICMSSubst.TabStop = False Then
'    tabItens.SetFocus
'  End If
  Call SelectAllText(txtTotBaseICMSSubst)
End Sub

Private Sub txtTotBaseICMSSubst_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub


Private Sub txtTotBaseICMSSubst_LostFocus()
  txtTotBaseICMSSubst.Text = gsFormatCurrency(txtTotBaseICMSSubst.Text, gnCurrencyDecimals)
End Sub

Private Sub txtTotDescontos_GotFocus()
'  If txtTotDescontos.TabStop = False Then
'    tabItens.SetFocus
'  End If
  Call SelectAllText(txtTotDescontos)
End Sub

Private Sub txtTotDescontos_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub


Private Sub txtTotDescontos_LostFocus()
  txtTotalAPagar.Text = CStr(CCur(gsHandleNull(txtTotProdutos.Text)) - CCur(gsHandleNull(txtTotDescontos.Text)) + CCur(gsHandleNull(txtTotIPI.Text)))
  If gbSomarFrete = True Then
    txtTotalAPagar.Text = CStr(CCur(gsHandleNull(txtTotalAPagar.Text)) + CCur(gsHandleNull(txtFrete.Text)))
  End If
  txtTotDescontos.Text = gsFormatCurrency(txtTotDescontos.Text, gnCurrencyDecimals)
  Call FormatCurrencyValue(txtTotalAPagar)
End Sub

Private Sub txtTotIPI_GotFocus()
'  If txtTotIPI.TabStop = False Then
'    tabItens.SetFocus
'  End If
  Call SelectAllText(txtTotIPI)
End Sub

Private Sub txtTotIPI_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub


Private Sub txtTotIPI_LostFocus()
'  txtTotalAPagar.Text = CStr(CCur(gsHandleNull(txtTotProdutos.Text)) - CCur(gsHandleNull(txtTotDescontos.Text)) + CCur(gsHandleNull(txtTotIPI.Text)))
'  If gbSomarFrete = True Then
'    txtTotalAPagar.Text = CStr(CCur(gsHandleNull(txtTotalAPagar.Text)) + CCur(gsHandleNull(txtFrete.Text)))
'  End If
  txtTotIPI.Text = gsFormatCurrency(txtTotIPI.Text, gnCurrencyDecimals)
'  Call FormatCurrencyValue(txtTotalAPagar)
End Sub

Private Sub txtTotProdutos_GotFocus()
'  If txtTotProdutos.TabStop = False Then
'    tabItens.SetFocus
'  End If
  Call SelectAllText(txtTotProdutos)
End Sub

Private Sub txtTotProdutos_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub


Private Sub txtTotProdutos_LostFocus()
'  txtTotProdutos.Text = gsFormatCurrency(Retorna_Valor(txtTotProdutos.Text), gnCurrencyDecimals)
'  txtTotalAPagar.Text = CStr(CCur(gsHandleNull(txtTotProdutos.Text)) - CCur(gsHandleNull(txtTotDescontos.Text)) + CCur(gsHandleNull(txtTotIPI.Text)))
'  If gbSomarFrete = True Then
'    txtTotalAPagar.Text = CStr(CCur(gsHandleNull(txtTotalAPagar.Text)) + CCur(gsHandleNull(txtFrete.Text)))
'  End If
  txtTotProdutos.Text = gsFormatCurrency(txtTotProdutos.Text, gnCurrencyDecimals)
'  Call FormatCurrencyValue(txtTotalAPagar)
End Sub

Private Sub FindNextPedido()
  Dim nSeq As Variant
  Dim sSql As String
  Dim nCodFornec As Integer
  
  Call StatusMsg("")
  If Len(cboFornecedor.Text) = 0 Then
    Beep
    DisplayMsg "Selecione um fornecedor antes."
    cboFornecedor.SetFocus
    Exit Sub
  End If
  
  nCodFornec = Val(cboFornecedor.Text)
  
  nSeq = gsHandleNull(txtSeq.Text & "")
  
  Call ClearScreen
  
  sSql = "SELECT Entradas.* FROM Entradas LEFT JOIN [Operações Entrada] "
  sSql = sSql & " ON Entradas.Operação = [Operações Entrada].Código "
  sSql = sSql & " WHERE Entradas.Filial = " & gnCodFilial
  sSql = sSql & " AND Entradas.Fornecedor = " & nCodFornec
  sSql = sSql & " AND Entradas.Sequência >= " & nSeq
  sSql = sSql & " AND [Operações Entrada].Tipo = 'P' "
  sSql = sSql & " ORDER BY Filial, Sequência "
  Set rsEntradas = db.OpenRecordset(sSql, dbOpenDynaset)
  
  If Not rsEntradas.EOF Then
    Num_Registro = rsEntradas.Bookmark
    Call MostraRegistro
  Else
    DisplayMsg "Não existem outros pedidos desta filial para este fornecedor."
  End If
  
End Sub

Private Sub UndoMovimEntrada()
  Dim Conta       As Integer
  Dim Sai_Loop    As Integer
  Dim Fim         As Integer
  Dim Ordem       As Integer
  Dim sSql        As String
  '07/11/2004 - Daniel
  Dim rstEntradas As Recordset
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre uma entrada antes."
    Exit Sub
  End If

  If rsEntradas("Efetivada") = False Then
    DisplayMsg "Esta operação não foi efetivada."
    Exit Sub
  End If
  
  '07/11/2004 - Daniel
  'Adicionado a verificação das movimentações
  sSql = "SELECT Data FROM Entradas WHERE Data > #" & Format(Data_Atual, "MM/DD/YYYY") & "#"
  
  Set rstEntradas = db.OpenRecordset(sSql, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then 'Encontrou registros
      If Not frmErroMov.gbContinue Then Exit Sub
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing
  sSql = ""
  '------------------------------------------
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If

  Conta = Desefetiva_Entrada(gnCodFilial, Val(txtSeq.Text))
  If Conta <> 0 Then
     DisplayMsg "Erro ao desfazer movimentação de Entrada. Erro = " & CStr(Conta)
     Exit Sub
  End If
  
  Call StatusMsg("Apagando itens de produtos do Movimento de Entrada...")
  
  On Error GoTo ErrTrans
  
  Call ws.BeginTrans
  
  
  ' Apaga Entradas - Produtos
  sSql = "DELETE * FROM [Entradas - Produtos] "
  sSql = sSql & " WHERE Filial = " & CStr(gnCodFilial) & "AND Sequência = " & Val(txtSeq.Text)
  Call db.Execute(sSql, dbFailOnError)
  
  Call StatusMsg("Apagando parcelas de pagamentos do Movimento de Entrada...")
  
  ' Apaga Parcelas existentes
  sSql = "DELETE * FROM [Movimento - Parcelas] "
  sSql = sSql & " WHERE Filial = " & CStr(gnCodFilial) & " AND Sequência = " & Val(txtSeq.Text)
  Call db.Execute(sSql, dbFailOnError)
  
  Call StatusMsg("Apagando Movimento de Entrada...")
  
  rsEntradas.Delete
  
  Call ws.CommitTrans
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Oper:" & cboOper.Text & " Cli:" & cboFornecedor.Text & " ChaveRef:" & txtRef.Text & " Tot:" & txtTotalAPagar.Text & " Seq:" & Val(txtSeq.Text), 80) & "', 'DESF_ENTRADA')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
  txtSeq.Text = ""
  Num_Registro = Null
  lblEfetivada.Visible = False
  Call StatusMsg("")

  
  DisplayMsg "Movimento de Entrada desfeito com sucesso."

  Exit Sub
  
ErrTrans:
  gsMsg = "Erro ao desfazer operação de entrada."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  Call StatusMsg("")
  Call ws.Rollback
  gsTitle = LoadResString(201)
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Num_Registro = Null
  Exit Sub

End Sub

Private Sub TypeGrade()
  Dim nTotalLinhas As Integer
  Dim sTamanhos(14) As String
  Dim nAuxTamanho(14) As Integer
  Dim nCol As Integer
  Dim nQtdes(14) As Integer
  Dim nAuxQtde(14) As Integer
  Dim nLine As Integer
  Dim sCor As String
  Dim sCodigo As String
  Dim snome As String
  Dim nCor As Integer
  Dim nPreco As Single
  Dim nDesconto As Single
  Dim nICM As Integer
  Dim nIPI As Integer
  Dim nValorTotal As Double
  Dim nEtiqueta As Integer
  Dim sRecord As String
  Dim bm As Variant
  Dim nRow As Long
  
  '22/04/2009 - mpdea
  'Descarrega tela com dados anteriores
  Unload frmDigitaGrade
  
  With frmDigitaGrade
    .O_Preço.Value = 1
    .O_Desconto.Value = 1
    .O_Impostos.Value = 1
    .O_Etiqueta.Value = 1
    .O_Mostra_Custo.Value = 0


    '07/11/2002 - mpdea
    'Modificado para que sempre mostre o preço de custo

'    If Len(cboOper.Text) > 0 Then
'      If rsOp_Entrada("Tipo") = "C" Then
        .O_Mostra_Custo.Value = 1
'      End If
'    End If


    .Limpa_Variáveis
    .Show vbModal
    If .Retorno.Caption <> "OK" Then
      Exit Sub
    End If
    nTotalLinhas = .Retorno1.Caption
  End With
    
  'Obtém os tamanhos da grade digitada
  Call frmDigitaGrade.RetornarTamanhos(nAuxTamanho())
  For nCol = 0 To 14
    sTamanhos(nCol) = Format(nAuxTamanho(nCol), "000")
  Next nCol
  
'   grdItens.Redraw = False
  
  
  '22/04/2009 - mpdea
  'Não remove os itens já digitados, adiciona os novos
'  grdItens.RemoveAll
  
  '23/04/2009 - mpdea
  'Remove as linhas em branco
  Dim Str_Aux As String
  For nRow = grdItens.Rows - 1 To 0 Step -1
    bm = grdItens.AddItemBookmark(nRow)
    grdItens.Bookmark = bm
    Str_Aux = gsHandleNull(grdItens.Columns("Código").CellText(bm))
    If (Str_Aux = "0" Or Str_Aux = "") And Not IsEmpty(bm) Then
      grdItens.RemoveItem grdItens.AddItemRowIndex(bm)
    End If
  Next nRow
  grdItens.Scroll -99, -99
  grdItens.Update
  
  For nLine = 1 To nTotalLinhas
    Call frmDigitaGrade.RetornarLinhaGrade(nLine, sCodigo, snome, nCor, _
      nAuxQtde(), nPreco, nDesconto, nICM, nIPI, nValorTotal, nEtiqueta)
    
    For nCol = 0 To 14
      nQtdes(nCol) = nAuxQtde(nCol)
    Next nCol
    
    sCor = Format(nCor, "000")
    
    For nCol = 0 To 14
      If nQtdes(nCol) <> 0 Then
        '01/12/2004 - Daniel
        'Antiga linha: nQtdes(nCol) & vbTab & sNome & vbTab & "" & vbTab...
        'Tratamento Devido a adição da coluna indice financeiro
        sRecord = sCodigo & sTamanhos(nCol) & sCor & vbTab & _
          nQtdes(nCol) & vbTab & vbTab & snome & vbTab & "" & vbTab & _
          nPreco & vbTab & "0" & vbTab & nDesconto & vbTab & _
          nICM & vbTab & nIPI & vbTab & "0" & vbTab & nEtiqueta
        grdItens.AddItem sRecord
        grdItens.Update
        DoEvents
      End If
    Next nCol
  Next nLine
  
  grdItens.MoveLast
  grdItens.MoveFirst
  
  For nRow = 0 To grdItens.Rows - 1
    DoEvents
    bm = grdItens.AddItemBookmark(nRow)
    grdItens.Bookmark = bm
    If gsHandleNull(grdItens.Columns("Código").CellText(bm)) <> "0" Then
      Call CalculaLinha(bm)
      grdItens.Update
    End If
  Next nRow
  
  grdItens.Redraw = True

'   Call Recalcula
  grdItens.Scroll -99, -99
'   grdItens.Update
  
End Sub

Private Sub ChangeLetrasPequenas()
  If ActiveBar1.Tools("miComplLetrasPequenas").Checked = True Then
    grdItens.Font.Size = 9
  Else
    grdItens.Font.Size = 7
  End If
End Sub

Private Sub GetInformation()
  Dim F As Form
  
  Call StatusMsg("")
  
  cboFornecedor.Text = gsHandleNull(cboFornecedor.Text)
  
  If Val(cboFornecedor.Text) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Selecione um Fornecedor antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    cboFornecedor.SetFocus
    Exit Sub
  End If
  
  Set F = New frmInformacoes
  Load F
  gsCodCliente = CStr(Val(cboFornecedor.Text))
  F.Show
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub TransformaPedidoEmCompra()
  Dim Resposta As Integer
  Dim Seq As Long
  Dim Sai_Loop As Integer
  Dim sSql As String
  
  Call StatusMsg("")

  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Encontre um Pedido antes."
    Exit Sub
  End If
  
  If rsOp_Entrada("Tipo") <> "P" Then
    Beep
    DisplayMsg "Esta Entrada não é um Pedido."
    Exit Sub
  End If
  
  Resposta = MsgBox("Esta operação não poderá ser desfeita, deseja realmente transformar este pedido em uma compra ? ", 1, "Atenção")
  If Resposta = 2 Then Exit Sub
  
  Call StatusMsg("Aguarde, alterando operação...")
  
  On Error GoTo ErrTrans
  
  Call ws.BeginTrans
  
  ' Apaga produtos
  sSql = "DELETE * FROM [Entradas - Produtos] "
  sSql = sSql & " WHERE Filial = " & CStr(gnCodFilial) & " AND Sequência = " & Val(txtSeq.Text)
  Call db.Execute(sSql, dbFailOnError)
  
  ' Apaga Entradas
  rsEntradas.Delete
  
  Call ws.CommitTrans
  
  lblEfetivada.Visible = False
  Num_Registro = Null
  
  cboOper.Text = 10
  cboOper_LostFocus
  cboCaixaUso.Text = 1
  cboCaixaUso_LostFocus
  cboConta.Text = ""
  
  txtSeq.Text = ""
  Call StatusMsg("")
  
  lblToday.Caption = Format$(Data_Atual, "dd/mm/yyyy")

  MsgBox ("Compra criada. Verifique o código da operação, os produtos, quantidades e preço. Informe os Detalhes do Pagamento e grave a operação.")
  
  Exit Sub
  
ErrTrans:
  gsMsg = "Erro ao realizar transformação de Pedido em Compra."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  Call StatusMsg("")
  Call ws.Rollback
  gsTitle = LoadResString(201)
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Num_Registro = Null

End Sub

Private Sub txtTotValICMS_GotFocus()
'  If txtTotValICMS.TabStop = False Then
'    tabItens.SetFocus
'  End If
  Call SelectAllText(txtTotValICMS)
End Sub

Private Sub txtTotValICMS_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub


Private Sub txtTotValICMS_LostFocus()
  txtTotValICMS.Text = gsFormatCurrency(txtTotValICMS.Text, gnCurrencyDecimals)
End Sub

Private Sub txtTotValICMSSubst_GotFocus()
'  If txtTotValICMSSubst.TabStop = False Then
'    tabItens.SetFocus
'  End If
  Call SelectAllText(txtTotValICMSSubst)
End Sub

Private Sub txtTotValICMSSubst_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub


Private Sub txtTotValICMSSubst_LostFocus()
  txtTotValICMSSubst.Text = gsFormatCurrency(txtTotValICMSSubst.Text, gnCurrencyDecimals)
End Sub

Private Sub txtValCheque_GotFocus()
  txtValCheque.SelStart = 0
  txtValCheque.SelLength = Len(txtValCheque.Text)
End Sub

Private Sub txtValCheque_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub txtValCheque_LostFocus()
  If IsGoodNumber(txtValCheque) Then
    Call RecalculaPagar
  End If
End Sub

Private Sub txtTotDigitado_GotFocus()
  SendKeys "{Tab}"
End Sub

'Verifica a existência de ao menos um produto para gravação
Private Function gbCheckGridItens() As Boolean
  Dim nRow As Long
  Dim bm As Variant
  
  gbCheckGridItens = False
  
  With grdItens
    For nRow = 0 To .Rows - 1
      bm = .AddItemBookmark(nRow)
      If Len(.Columns("Codigo").CellText(bm)) > 0 Then
        If CSng(gsHandleNull(.Columns("Qtde").CellText(bm))) > 0 Then
          gbCheckGridItens = True
          Exit Function
        End If
      End If
    Next nRow
  End With
  
End Function

'13/12/2005 - mpdea
'Verifica se há produto com Preço Total/Final inválido
Private Function m_blnItemPrecoInvalido() As Boolean
  Dim lng_row As Long
  Dim var_book As Variant
  Dim dbl_preco As Double
  Dim sng_desconto As Single
    
  With grdItens
    For lng_row = 0 To .Rows - 1
      var_book = .AddItemBookmark(lng_row)
      If Len(.Columns("Codigo").CellText(var_book)) > 0 Then
        'Verifica Preço Total menor ou igual a zero
        Call IsDataType(dtDouble, .Columns("PrecoTotal").CellText(var_book), dbl_preco)
        If dbl_preco <= 0 Then
          m_blnItemPrecoInvalido = True
          Exit Function
        End If
        'Verifica Preço Final menor que zero ou
        'igual a zero e sem desconto
        Call IsDataType(dtSingle, .Columns("Desconto").CellText(var_book), sng_desconto)
        Call IsDataType(dtDouble, .Columns("PrecoFinal").CellText(var_book), dbl_preco)
        If dbl_preco < 0 Or (dbl_preco = 0 And sng_desconto = 0) Then
          m_blnItemPrecoInvalido = True
          Exit Function
        End If
      End If
    Next lng_row
  End With
  
End Function

Private Sub AtualizarPrecoDeVenda(ByVal CodProduto As String, ByVal indice As Double, ByVal CodOpEntrada As Integer, ByVal dblCustoDigitado As Double)
  '24/08/2004 - Daniel
  Dim rstOpEntrada As Recordset
  Dim rstPrecos    As Recordset
  Dim strTabela    As String
  Dim strQuery     As String
  Dim dblCUSTO     As Double
  Dim dblAuxi      As Double
  
  strQuery = "SELECT Código, Tabela "
  strQuery = strQuery & " FROM [Operações Entrada] "
  strQuery = strQuery & " WHERE Código = " & CodOpEntrada
  
  Set rstOpEntrada = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstOpEntrada
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      strTabela = .fields("Tabela").Value & ""
    End If
    .Close
  End With

  Set rstOpEntrada = Nothing
  
  strQuery = ""
  
  'Populamos o custo
  dblCUSTO = dblCustoDigitado
  
  'Atualização
  strQuery = "SELECT * "
  strQuery = strQuery & " FROM Preços "
  strQuery = strQuery & " WHERE Produto = '" & CodProduto & "'"
  strQuery = strQuery & " AND Tabela = '" & strTabela & "'"
  
  Set rstPrecos = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstPrecos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      .Edit
      .fields("Data Alteração").Value = Data_Atual
      
      If indice >= 1 Then
        dblAuxi = indice - 1
        
        dblAuxi = dblAuxi * 100
      
        .fields("Preço").Value = Format((dblCUSTO + ((dblCUSTO * dblAuxi) / 100)), FORMAT_VALUE)
      Else
        dblAuxi = 1 - indice
      
        dblAuxi = dblAuxi * 100
      
        .fields("Preço").Value = Format((dblCUSTO - ((dblCUSTO * dblAuxi) / 100)), FORMAT_VALUE)
      End If
      
      .Update
    End If
    .Close
  End With

  Set rstPrecos = Nothing

End Sub

Private Sub VerificarOperacao()
  '22/09/2004 - Daniel
  Dim rstOperacao As Recordset
  Dim strSQL      As String
  
  strSQL = "SELECT * FROM [Operações Entrada] "
  strSQL = strSQL & " WHERE Código = " & CInt(cboOper.Text)
  
  Set rstOperacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstOperacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .fields("Tipo").Value = "E" Then m_blnEmprestimo = True
      
    End If
    .Close
  End With
  
  Set rstOperacao = Nothing
  
End Sub

Private Function ControlarComisao() As Boolean
  '14/02/2005 - Daniel
  '
  'Solicitante: Daring - RJ
  '
  'Se ocorre devolução e esta devolução implica em abatimento de
  'comissão do vendedor, o Quick estava descontando erroneamente
  'da comissão para casos em que a venda possuia descontos.
  Dim rstOpEntrada As Recordset
  Dim strSQL       As String
  
  If Len(cboOper.Text) <= 0 Then Exit Function
  
  strSQL = "SELECT Tipo, Comissão FROM [Operações Entrada] "
  strSQL = strSQL & " WHERE Código = " & CInt(cboOper.Text)
  
  Set rstOpEntrada = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstOpEntrada
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .fields("Tipo").Value = "D" And .fields("Comissão").Value = True Then ControlarComisao = True
      
    End If
    .Close
  End With
  
  Set rstOpEntrada = Nothing

End Function

Private Function VerificaDataCP() As Boolean
  '29/06/2005 - Daniel
  'A grid de contas a pagar estava aceitando movimentações com data retroativa e depois ao efetivar o recebimento
  'não gerava o "contas a pagar" devido a validação: If rsMovi_Parcelas("Bom") >= dtData Then rsContas_Pagar.AddNew
  'Criamos a Função "VerificaDataCP" para validação das datas lançadas na grid
  Dim intAuxi As Integer

  On Error GoTo TratarErro

  grdCP.MoveFirst

  If Len(grdCP.Columns("Data").Text) <= 0 Then Exit Function

  For intAuxi = 0 To (grdCP.Rows - 1)

    If Len(grdCP.Columns("Data").Text) <= 0 Then
        Exit Function
    Else
      If CDate(Format(grdCP.Columns("Data").Text, "DD/MM/YYYY")) < Data_Atual Then
        VerificaDataCP = True
        MsgBox "Contas a pagar não aceita parcelamento com data retroativa, verifique.", vbExclamation, "Atenção"
        Exit Function

'''        Criamos este codigo em 09/12/2019...porém decidimos não deixar que a entrada do parcelamento retroativo
'''        Dim retMsg As Variant
'''        retMsg = MsgBox("Contas a pagar não aceita parcelamento com data retroativa " & Format(grdCP.Columns("Data").Text, "DD/MM/YYYY") & ". Mesmo assim deseja continuar?", vbYesNo, "Atenção")
'''
'''        If retMsg = vbNo Then
'''            VerificaDataCP = True
'''            Exit Function
'''        End If

      End If
    End If
    
    grdCP.MoveNext
  Next intAuxi
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

'17/05/2006 - mpdea
'Obtém a alíquota de ICMS do estado solicitado
Private Function m_intGetAliquotaIcmsEstado() As Integer
  Dim int_ret As Integer
  
  rsEstados.Index = "Estado"
  rsEstados.Seek "=", sEstado
  
  If Not rsEstados.NoMatch Then
    Call IsDataType(dtInteger, rsEstados.fields("ICM").Value, int_ret)
  End If
  
  If rsOp_Entrada("Tipo") = "T" And rsOp_Entrada("ICM").Value = False Then
    int_ret = 0
  End If
  
  m_intGetAliquotaIcmsEstado = int_ret
End Function

'18/06/2007 - Anderson
'Função utilizada para exportar dados para excel
Private Sub ExportarExcel()

  Dim appExcel As New Excel.Application
  Dim rsExpParametros As Recordset
  Dim rsExpEntradas As Recordset
  Dim rsExpEntradasProdutos As Recordset
  Dim strSQL As String
  Dim strRange As String
  Dim intContador As Integer
  Dim strCampo As String
  Dim strValor As String
  
  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Encontre um registro antes."
    Set appExcel = Nothing
    Exit Sub
  End If
  
  If gsArquivoExcelEntrada = "" Then
    Beep
    DisplayMsg "Arquivo modelo para exportação de dados não está configurado, favor verificar as configurações no arquivo config.ini no diretório padrão do Quick Store"
    Set appExcel = Nothing
    Exit Sub
  End If
  
  If MsgBox("Deseja exportar dados atual para Excel?", vbYesNo + vbQuestion, "Exportar Dados para Excel") = vbYes Then
    
    Call StatusMsg("Aguarde, exportando dados...")
    MousePointer = vbHourglass
  
    'Inicia Excel
    'appExcel.Application.Visible = True
    appExcel.ScreenUpdating = False
    'Abre o arquivo modelo para exportação
    appExcel.Workbooks.Open gsReportPath & gsArquivoExcelEntrada
    
    'Seleciona Célula A1
    appExcel.Range("A1").Select

    DoEvents
    'Parametros da empresa Filial
    strSQL = ""
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM [Parâmetros Filial] "
    strSQL = strSQL & "WHERE Filial=" & rsEntradas("Filial")
    
    Set rsExpParametros = db.OpenRecordset(strSQL)
    
    'Exporta Cabeçalho
    With rsExpParametros
      If Not (.BOF And .EOF) Then
              
        For intContador = 0 To .fields.Count - 1
          strCampo = Mid(.fields(intContador).Name, InStr(1, .fields(intContador).Name, ".") + 1)
          strCampo = Replace(sTranslateInvalidChar(.fields(intContador).SourceTable & "_" & strCampo), " ", "_")
          If .fields(intContador).Type = dbCurrency Or .fields(intContador).Type = dbDecimal Or .fields(intContador).Type = dbDouble Or .fields(intContador).Type = dbFloat Or .fields(intContador).Type = dbSingle Then
            strValor = Replace("" & .fields(intContador), ",", ".")
          Else
            strValor = "" & .fields(intContador)
          End If
          Call ExcelSubstituirCampo("[" & strCampo & "]", strValor, gsArquivoExcelEntrada, appExcel)
        Next
        
      End If
      .Close
    End With
    
    'Cabeçalho da Entrada
    strSQL = ""
    strSQL = strSQL & "SELECT Entradas.*, Cli_For.*, Funcionários.*, [Operações Entrada].*, [Centros de Custo].*, Transportadoras.* "
    strSQL = strSQL & "FROM ([Centros de Custo] RIGHT JOIN ([Operações Entrada] RIGHT JOIN ((Entradas LEFT JOIN Cli_For ON Entradas.Fornecedor = Cli_For.Código) LEFT JOIN Funcionários ON Entradas.Digitador = Funcionários.Código) ON [Operações Entrada].Código = Entradas.Operação) ON [Centros de Custo].Código = Entradas.CentroCusto) LEFT JOIN Transportadoras ON Entradas.obs_Transportadora = Transportadoras.Nome "
    strSQL = strSQL & "WHERE Entradas.Filial=" & rsEntradas("Filial") & " AND Entradas.Sequência = " & rsEntradas("Sequência") & " "
    strSQL = strSQL & "ORDER BY Entradas.Filial, Entradas.Data, Entradas.Sequência "

    Set rsExpEntradas = db.OpenRecordset(strSQL)
    
    'Exporta Cabeçalho
    With rsExpEntradas
      If Not (.BOF And .EOF) Then
              
        For intContador = 0 To .fields.Count - 1
          strCampo = Mid(.fields(intContador).Name, InStr(1, .fields(intContador).Name, ".") + 1)
          strCampo = Replace(sTranslateInvalidChar(.fields(intContador).SourceTable & "_" & strCampo), " ", "_")
          If .fields(intContador).Type = dbCurrency Or .fields(intContador).Type = dbDecimal Or .fields(intContador).Type = dbDouble Or .fields(intContador).Type = dbFloat Or .fields(intContador).Type = dbSingle Then
            strValor = Replace("" & .fields(intContador), ",", ".")
          Else
            strValor = "" & .fields(intContador)
          End If
          Call ExcelSubstituirCampo("[" & strCampo & "]", strValor, gsArquivoExcelEntrada, appExcel)
        Next
        
      End If
      .Close
    End With
    
    'Detalhe da Entrada
    strSQL = ""
    strSQL = strSQL & "SELECT [Entradas - Produtos].*, Produtos.*, Classes.*, [Sub Classes].*, GrupoFiscal.* "
    strSQL = strSQL & "FROM ((Classes RIGHT JOIN ([Entradas - Produtos] LEFT JOIN Produtos ON [Entradas - Produtos].[Código sem Grade] = Produtos.Código) ON Classes.Código = Produtos.Classe) LEFT JOIN [Sub Classes] ON Produtos.[Sub Classe] = [Sub Classes].Código) LEFT JOIN GrupoFiscal ON Produtos.GrupoFiscal = GrupoFiscal.Código "
    strSQL = strSQL & "WHERE [Entradas - Produtos].Filial=" & rsEntradas("Filial") & " AND [Entradas - Produtos].Sequência = " & rsEntradas("Sequência") & " "
    strSQL = strSQL & "ORDER BY [Entradas - Produtos].Filial, [Entradas - Produtos].Sequência, [Entradas - Produtos].Linha "
    
    Set rsExpEntradasProdutos = db.OpenRecordset(strSQL)
    
    'Exporta Cabeçalho
    With rsExpEntradasProdutos
      If Not (.BOF And .EOF) Then
                
        Do Until .EOF
          'Localiza campo para inserir linha para acrescentar produto
          strRange = ""
          strRange = ExcelLocalizarCampo("[PROXIMO_PRODUTO]", gsArquivoExcelEntrada, appExcel)
          
          'Se não tiver campo [PROXIMO_PRODUTO], o sistema não insere os produtos
          If strRange <> "" Then
            'Seleciona [PROXIMO_PRODUTO]
            appExcel.Range(strRange).Select
            'Seleciona linha
            appExcel.Rows(appExcel.ActiveCell.Row & ":" & appExcel.ActiveCell.Row).Select
            'Insere Linha
            appExcel.Selection.Insert Shift:=xlDown
            'Seleciona [PROXIMO_PRODUTO]
            appExcel.Range(strRange).Select
            'Seleciona linha para copiar modelo
            appExcel.Rows(appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1).Select
            'Copia modelo
            appExcel.Selection.Copy
            'Seleciona [PROXIMO_PRODUTO]
            appExcel.Range(strRange).Select
            'Seleciona Linha
            appExcel.Rows(appExcel.ActiveCell.Row & ":" & appExcel.ActiveCell.Row).Select
            'Cola Linha
            appExcel.ActiveSheet.Paste
            'Desativa mode de copia
            appExcel.CutCopyMode = False
            'Seleciona [PROXIMO_PRODUTO]
            appExcel.Range(strRange).Select
            
            'Altera conteúdo das células
            For intContador = 0 To .fields.Count - 1
              strCampo = Mid(.fields(intContador).Name, InStr(1, .fields(intContador).Name, ".") + 1)
              strCampo = Replace(sTranslateInvalidChar(.fields(intContador).SourceTable & "_" & strCampo), " ", "_")
              
              If .fields(intContador).Type = dbCurrency Or .fields(intContador).Type = dbDecimal Or .fields(intContador).Type = dbDouble Or .fields(intContador).Type = dbFloat Or .fields(intContador).Type = dbSingle Then
                strValor = Replace("" & .fields(intContador), ",", ".")
              Else
                strValor = "" & .fields(intContador)
              End If

              Call ExcelSubstituirCampo("[" & strCampo & "]", strValor, gsArquivoExcelEntrada, appExcel, appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1)
            Next
            
          End If
          
          .MoveNext
          
        Loop
        
        'Seleciona [PROXIMO_PRODUTO]
        strRange = ""
        strRange = ExcelLocalizarCampo("[PROXIMO_PRODUTO]", gsArquivoExcelEntrada, appExcel)
        
        'Seleciona [PROXIMO_PRODUTO]
        appExcel.Range(strRange).Select
        'Seleciona linha
        appExcel.Rows(appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1).Select
        'Exclui linha modelo
        appExcel.Selection.Delete Shift:=xlUp
        'Seleciona [PROXIMO_PRODUTO]
        appExcel.Range(strRange).Select
        'Limpa campo [PROXIMO_PRODUTO]
        Call ExcelSubstituirCampo("[PROXIMO_PRODUTO]", "", gsArquivoExcelEntrada, appExcel)
        
      End If
      
      .Close
      
    End With
    
    If gsSaveExcelEntrada = "" Then
      appExcel.Visible = True
      With appExcel.FileDialog(2)
        .InitialFileName = rsEntradas("Sequência")
        .Show
      End With
      appExcel.ActiveWorkbook.SaveAs appExcel.FileDialog(2).InitialFileName & ".xls"
    Else
      appExcel.DisplayAlerts = False
      appExcel.ActiveWorkbook.SaveAs gsSaveExcelEntrada & rsEntradas("Sequência") & ".xls"
      appExcel.DisplayAlerts = True
    End If
    
    appExcel.ScreenUpdating = True
    
    appExcel.Application.Quit

  
    MsgBox "Operação concluída com sucesso!", vbExclamation, "Exportar Dados"
  
  End If
  
  Set rsExpParametros = Nothing
  Set rsExpEntradas = Nothing
  Set rsExpEntradasProdutos = Nothing
  Set appExcel = Nothing
  
  Call StatusMsg("")
  MousePointer = vbDefault

End Sub

Private Function UpdateTotalNCM()
  rsEntradas.Edit
  Dim totalNCM As Double 'Total em R$ de imposto pago na movimentação
  Dim Valor_Aprox_Impostos As Double
  Dim rsAliquotas As Recordset 'Tabela que filtra todos os produtos da sequencia
  Dim rsProdutos3 As Recordset 'Tabela que filtra produto por produto da movimentação
  totalNCM = 0#
  Set rsAliquotas = db.OpenRecordset("SELECT [Código Sem Grade],[Preço Final],[Valor_Aprox_Impostos],[MotivoDesoneracaoICMS] FROM [Entradas - Produtos] WHERE [Sequência] = " & txtSeq.Text, dbOpenDynaset)
  rsAliquotas.MoveFirst
  While Not rsAliquotas.EOF
    Set rsProdutos3 = db.OpenRecordset("SELECT [AliqNCM],[MotivoDesoneracaoICMS] FROM [Produtos] WHERE [Código] = '" & rsAliquotas("Código Sem Grade") & "'", dbOpenDynaset)
    rsProdutos3.MoveFirst
    If (rsProdutos3("AliqNCM") <> "" Or rsProdutos3("AliqNCM") = 0) Then
      Valor_Aprox_Impostos = (rsProdutos3("AliqNCM") * rsAliquotas("Preço Final") / 100)
      Valor_Aprox_Impostos = FormatNumber(Valor_Aprox_Impostos, 2)
      totalNCM = totalNCM + (rsProdutos3("AliqNCM") * rsAliquotas("Preço Final") / 100)
      totalNCM = FormatNumber(totalNCM, 2)
      rsAliquotas.Edit
      rsAliquotas("Valor_Aprox_Impostos") = Valor_Aprox_Impostos
      rsAliquotas("MotivoDesoneracaoICMS") = rsProdutos3("MotivoDesoneracaoICMS")
      rsAliquotas.Update
    Else
      rsAliquotas.Edit
      rsAliquotas("MotivoDesoneracaoICMS") = rsProdutos3("MotivoDesoneracaoICMS")
      rsAliquotas.Update
      'MsgBox "O produto " & rsAliquotas("Código Sem Grade") & " não possui aliquota de NCM", vbExclamation
    End If
    rsAliquotas.MoveNext
  Wend
  rsEntradas("TotalNCM") = totalNCM
  'rsEntradas("TotalNCM") = Format(rsEntradas("TotalNCM"), "##,###,##0.00")
  rsEntradas.Update
End Function


