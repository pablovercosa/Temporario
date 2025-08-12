VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmGrafico4 
   Caption         =   " CockPit de Produtos:  1-Produtos mais vendidos;  2-Curva ABC  e  3- Esboço de D.R.E."
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrafico4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   15915
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   60
      TabIndex        =   7
      Top             =   2040
      Width           =   15825
      _ExtentX        =   27914
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Maiores produtos do período"
      TabPicture(0)   =   "frmGrafico4.frx":4E95A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gridProdutos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd_detalharProdutoEntradas"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmd_detalharProdutoSaidas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalhamento de saídas"
      TabPicture(1)   =   "frmGrafico4.frx":4E976
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_abreSequencia"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txt_qtde02"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt_codProd02"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txt_nmProd02"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txt_tamanhoCor02"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "gridSaidas"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label28"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label27"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label26"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label25"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Detalhamento de entradas"
      TabPicture(2)   =   "frmGrafico4.frx":4E992
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt_tamanhoCor"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txt_nmProd01"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txt_codProd01"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txt_qtde01"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "gridEntradas"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label13"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label10"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label9"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lbl_qtde01"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Curva ABC - Pareto"
      TabPicture(3)   =   "frmGrafico4.frx":4E9AE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label15"
      Tab(3).Control(1)=   "Label16"
      Tab(3).Control(2)=   "Label17"
      Tab(3).Control(3)=   "Label18"
      Tab(3).Control(4)=   "Label19"
      Tab(3).Control(5)=   "Label20"
      Tab(3).Control(6)=   "Label21"
      Tab(3).Control(7)=   "Label22"
      Tab(3).Control(8)=   "Label23"
      Tab(3).Control(9)=   "Label24"
      Tab(3).Control(10)=   "gridProdutos_pareto"
      Tab(3).Control(11)=   "txt_paretoA"
      Tab(3).Control(12)=   "txt_paretoB"
      Tab(3).Control(13)=   "txt_paretoC"
      Tab(3).Control(14)=   "cmd_calcularCurvaABC"
      Tab(3).Control(15)=   "txt_valorTotalConsiderado"
      Tab(3).Control(16)=   "txt_numTotalItens"
      Tab(3).Control(17)=   "txt_numTotalItensA"
      Tab(3).Control(18)=   "txt_numTotalItensB"
      Tab(3).Control(19)=   "txt_numTotalItensC"
      Tab(3).ControlCount=   20
      TabCaption(4)   =   "D.R.E."
      TabPicture(4)   =   "frmGrafico4.frx":4E9CA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label47"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label51"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Line2"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmd_DRE_help"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmd_buscarInformacoesParaCalculoDRE"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmd_DRE_finalizarCalculo"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmd_DRE_limparTela"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "txt_DRE_Observacoes"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "cmd_DRE_salvar"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Frame4"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Frame5"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).ControlCount=   11
      TabCaption(5)   =   "D.R.E. Histórico"
      TabPicture(5)   =   "frmGrafico4.frx":4E9E6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame2"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame1"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmb_DRE_AnoPesquisa"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cmd_DRE_pesquisar"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "grid_DRE"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label54"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).ControlCount=   6
      Begin VB.CommandButton cmd_abreSequencia 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Detalhar Sequência"
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
         Left            =   -61470
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   6510
         Width           =   2145
      End
      Begin VB.Frame Frame5 
         Height          =   6405
         Left            =   -68070
         TabIndex        =   119
         Top             =   360
         Width           =   3165
         Begin VB.ComboBox cmb_mes 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmGrafico4.frx":4EA02
            Left            =   150
            List            =   "frmGrafico4.frx":4EA2A
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   570
            Width           =   1605
         End
         Begin VB.ComboBox cmb_ano 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmGrafico4.frx":4EA93
            Left            =   1860
            List            =   "frmGrafico4.frx":4EAD9
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   570
            Width           =   1155
         End
         Begin VB.ComboBox cmb_DRE_tipoAnexo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmGrafico4.frx":4EB61
            Left            =   150
            List            =   "frmGrafico4.frx":4EB6B
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   1350
            Width           =   2865
         End
         Begin VB.CheckBox chk_DRE_somenteSaidasComNFeNFCe 
            Appearance      =   0  'Flat
            Caption         =   "Saídas x NFe/NFCe"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1320
            TabIndex        =   124
            Top             =   1770
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Frame Frame3 
            Caption         =   "Consultar?"
            ForeColor       =   &H000000FF&
            Height          =   915
            Left            =   150
            TabIndex        =   121
            Top             =   3450
            Width           =   2865
            Begin VB.CommandButton Command1 
               BackColor       =   &H008080FF&
               Caption         =   "Contas à pagar"
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
               Left            =   1470
               Style           =   1  'Graphical
               TabIndex        =   123
               Top             =   300
               Width           =   1335
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H008080FF&
               Caption         =   "Contas Pagas"
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
               Left            =   60
               Style           =   1  'Graphical
               TabIndex        =   122
               Top             =   300
               Width           =   1335
            End
         End
         Begin VB.ListBox lst_centroCusto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   1155
            Left            =   150
            Style           =   1  'Checkbox
            TabIndex        =   120
            Top             =   2250
            Width           =   2865
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   30
            X2              =   150
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label52 
            Caption         =   "Mês                       Ano"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   150
            TabIndex        =   130
            Top             =   270
            Width           =   2295
         End
         Begin VB.Label Label53 
            Caption         =   "Tipo de empresa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   150
            TabIndex        =   129
            Top             =   1080
            Width           =   1485
         End
         Begin VB.Label Label49 
            Caption         =   "Centro Custo"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   150
            TabIndex        =   128
            Top             =   2010
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         Height          =   6405
         Left            =   -74880
         TabIndex        =   78
         Top             =   360
         Width           =   6675
         Begin VB.TextBox txt_DRE_receitaBruta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   96
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_Devolucoes 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   95
            Top             =   570
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_ImpostosSobreVendas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   94
            Top             =   900
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_receitaOperacionalLiquida 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4740
            TabIndex        =   93
            Top             =   1230
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_cmv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   92
            Top             =   1575
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_lucroBruto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4740
            TabIndex        =   91
            Top             =   1920
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_despesasAdministrativas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   90
            Top             =   2250
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_despesasComerciais 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   89
            Top             =   2580
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_despesasDepreciacao 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   88
            Top             =   2910
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_despesasFinanceiras 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   87
            Top             =   3240
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_receitasFinanceiras 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   86
            Top             =   3570
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_lucroPrejuizoOperacional 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4740
            TabIndex        =   85
            Top             =   3900
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_despesasNaoOperacionais 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   84
            Top             =   4230
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_receitasNaoOperacionais 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   83
            Top             =   4560
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_LAIR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4740
            TabIndex        =   82
            Top             =   4890
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_provisaoIR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   81
            Top             =   5220
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_provisaoCSLL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   80
            Top             =   5550
            Width           =   1800
         End
         Begin VB.TextBox txt_DRE_LucroLiquido 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4740
            TabIndex        =   79
            Top             =   5880
            Width           =   1800
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   4770
            X2              =   6600
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label30 
            Caption         =   "Receita bruta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   270
            TabIndex        =   118
            Top             =   285
            Width           =   1125
         End
         Begin VB.Label Label31 
            Caption         =   "- Devoluções"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   117
            Top             =   615
            Width           =   1125
         End
         Begin VB.Label Label32 
            Caption         =   "- Impostos sobre as Vendas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   116
            Top             =   945
            Width           =   2295
         End
         Begin VB.Label Label33 
            Caption         =   "= Receita Operacional Líquida"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   150
            TabIndex        =   115
            Top             =   1245
            Width           =   2805
         End
         Begin VB.Label Label34 
            Caption         =   "- CMV/CPV/CSP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   114
            Top             =   1635
            Width           =   1365
         End
         Begin VB.Label Label35 
            Caption         =   "= Lucro bruto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   150
            TabIndex        =   113
            Top             =   1935
            Width           =   2535
         End
         Begin VB.Label Label36 
            Caption         =   "- Despesas Administrativas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   112
            Top             =   2295
            Width           =   2415
         End
         Begin VB.Label Label37 
            Caption         =   "- Despesas Comerciais"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   111
            Top             =   2625
            Width           =   2385
         End
         Begin VB.Label Label38 
            Caption         =   "- Despesas Depreciação"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   110
            Top             =   2955
            Width           =   2385
         End
         Begin VB.Label Label39 
            Caption         =   "- Despesas Financeiras"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   109
            Top             =   3285
            Width           =   2385
         End
         Begin VB.Label Label40 
            Caption         =   "+ Receitas Financeiras"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   150
            TabIndex        =   108
            Top             =   3615
            Width           =   2385
         End
         Begin VB.Label lbl_DRE_lucroPrejuizoOperacional 
            Caption         =   "= Lucro/Prejuízo Operacional"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   150
            TabIndex        =   107
            Top             =   3915
            Width           =   2775
         End
         Begin VB.Label Label41 
            Caption         =   "- Despesas NÃO Operacionais"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   106
            Top             =   4275
            Width           =   2385
         End
         Begin VB.Label Label42 
            Caption         =   "+ Receitas NÃO Operacionais"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   150
            TabIndex        =   105
            Top             =   4605
            Width           =   2385
         End
         Begin VB.Label Label43 
            Caption         =   "= LAIR (Lucro antes do IR)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   150
            TabIndex        =   104
            Top             =   4905
            Width           =   3465
         End
         Begin VB.Label Label44 
            Caption         =   "- Provisão IR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   103
            Top             =   5265
            Width           =   2385
         End
         Begin VB.Label Label45 
            Caption         =   "- Provisão CSLL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   150
            TabIndex        =   102
            Top             =   5595
            Width           =   2385
         End
         Begin VB.Label Label46 
            Caption         =   "= Lucro Líquido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   150
            TabIndex        =   101
            Top             =   5895
            Width           =   2385
         End
         Begin VB.Label lbl_DRE_Percentual_IR 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   4800
            TabIndex        =   100
            Top             =   5265
            Width           =   525
         End
         Begin VB.Label Label48 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   5370
            TabIndex        =   99
            Top             =   5265
            Width           =   255
         End
         Begin VB.Label lbl_DRE_Percentual_CSLL 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   4800
            TabIndex        =   98
            Top             =   5595
            Width           =   525
         End
         Begin VB.Label Label50 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5370
            TabIndex        =   97
            Top             =   5595
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Toda a grade"
         Height          =   795
         Left            =   -63120
         TabIndex        =   76
         Top             =   6090
         Width           =   3825
         Begin VB.CommandButton cmd_DRE_imprimir 
            BackColor       =   &H00FF8080&
            Caption         =   "Imprima"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   630
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   240
            Width           =   2745
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selecione um e"
         Height          =   795
         Left            =   -74910
         TabIndex        =   73
         Top             =   6090
         Width           =   11685
         Begin VB.CommandButton cmd_DRE_imprimirUM 
            BackColor       =   &H00FF8080&
            Caption         =   "Imprima"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3420
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   240
            Width           =   2745
         End
         Begin VB.CommandButton cmd_DRE_excluir 
            BackColor       =   &H00FF8080&
            Caption         =   "Exclua"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   240
            Width           =   2745
         End
      End
      Begin VB.ComboBox cmb_DRE_AnoPesquisa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmGrafico4.frx":4EB8C
         Left            =   -74430
         List            =   "frmGrafico4.frx":4EBD5
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   405
         Width           =   1065
      End
      Begin VB.CommandButton cmd_DRE_pesquisar 
         BackColor       =   &H00FF8080&
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -62040
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   360
         Width           =   2745
      End
      Begin VB.CommandButton cmd_detalharProdutoSaidas 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Detalhar movimento 'de saídas' de um produto"
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
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   6450
         Width           =   4095
      End
      Begin VB.CommandButton cmd_detalharProdutoEntradas 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Detalhar movimento 'de entradas' de um produto"
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
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   6450
         Width           =   4275
      End
      Begin VB.CommandButton cmd_DRE_salvar 
         BackColor       =   &H00FF8080&
         Caption         =   "Salvar D.R.E."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -64770
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   6360
         Width           =   5355
      End
      Begin VB.TextBox txt_DRE_Observacoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   1515
         Left            =   -64770
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   4770
         Width           =   5355
      End
      Begin VB.CommandButton cmd_DRE_limparTela 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Limpar tela D.R.E."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -64770
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2040
         Width           =   5355
      End
      Begin VB.CommandButton cmd_DRE_finalizarCalculo 
         BackColor       =   &H00FF8080&
         Caption         =   "Calcular D.R.E."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -64770
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3270
         Width           =   5355
      End
      Begin VB.CommandButton cmd_buscarInformacoesParaCalculoDRE 
         BackColor       =   &H00FF8080&
         Caption         =   "Buscar informações para cálculo D.R.E."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -64770
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   2790
         Width           =   5355
      End
      Begin VB.CommandButton cmd_DRE_help 
         Height          =   585
         Left            =   -60090
         Picture         =   "frmGrafico4.frx":4EC5F
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3930
         Width           =   675
      End
      Begin VB.TextBox txt_qtde02 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -74940
         TabIndex        =   52
         Top             =   690
         Width           =   1470
      End
      Begin VB.TextBox txt_codProd02 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73245
         TabIndex        =   51
         Top             =   690
         Width           =   2130
      End
      Begin VB.TextBox txt_nmProd02 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -67800
         TabIndex        =   50
         Top             =   690
         Width           =   8475
      End
      Begin VB.TextBox txt_tamanhoCor02 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -70980
         TabIndex        =   49
         Top             =   690
         Width           =   3060
      End
      Begin VB.TextBox txt_numTotalItensC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
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
         Height          =   345
         Left            =   -60180
         MaxLength       =   20
         TabIndex        =   44
         Top             =   5250
         Width           =   915
      End
      Begin VB.TextBox txt_numTotalItensB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
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
         Height          =   345
         Left            =   -60180
         MaxLength       =   20
         TabIndex        =   43
         Top             =   4860
         Width           =   915
      End
      Begin VB.TextBox txt_numTotalItensA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
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
         Height          =   345
         Left            =   -60180
         MaxLength       =   20
         TabIndex        =   42
         Top             =   4470
         Width           =   915
      End
      Begin VB.TextBox txt_numTotalItens 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
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
         Height          =   345
         Left            =   -60180
         MaxLength       =   20
         TabIndex        =   40
         Top             =   4080
         Width           =   915
      End
      Begin VB.TextBox txt_valorTotalConsiderado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
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
         Height          =   345
         Left            =   -60960
         MaxLength       =   30
         TabIndex        =   39
         Top             =   6030
         Width           =   1695
      End
      Begin VB.CommandButton cmd_calcularCurvaABC 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Calcular Curva ABC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -60930
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2070
         Width           =   1665
      End
      Begin VB.TextBox txt_paretoC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -60120
         MaxLength       =   3
         TabIndex        =   36
         Text            =   "100"
         Top             =   1620
         Width           =   465
      End
      Begin VB.TextBox txt_paretoB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -60120
         MaxLength       =   3
         TabIndex        =   35
         Text            =   "95"
         Top             =   1260
         Width           =   465
      End
      Begin VB.TextBox txt_paretoA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -60120
         MaxLength       =   3
         TabIndex        =   34
         Text            =   "80"
         Top             =   900
         Width           =   465
      End
      Begin VB.TextBox txt_tamanhoCor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -70980
         TabIndex        =   27
         Top             =   690
         Width           =   3060
      End
      Begin VB.TextBox txt_nmProd01 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -67800
         TabIndex        =   24
         Top             =   690
         Width           =   8475
      End
      Begin VB.TextBox txt_codProd01 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73245
         TabIndex        =   19
         Top             =   690
         Width           =   2130
      End
      Begin VB.TextBox txt_qtde01 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -74940
         TabIndex        =   18
         Top             =   690
         Width           =   1470
      End
      Begin MSFlexGridLib.MSFlexGrid gridProdutos 
         Height          =   5970
         Left            =   90
         TabIndex        =   8
         Top             =   420
         Width           =   15630
         _ExtentX        =   27570
         _ExtentY        =   10530
         _Version        =   393216
         Rows            =   1
         Cols            =   12
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridEntradas 
         Height          =   5730
         Left            =   -74940
         TabIndex        =   20
         Top             =   1140
         Width           =   15630
         _ExtentX        =   27570
         _ExtentY        =   10107
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   16777152
         BackColorSel    =   12640511
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
      Begin MSFlexGridLib.MSFlexGrid gridProdutos_pareto 
         Height          =   6330
         Left            =   -74910
         TabIndex        =   29
         Top             =   540
         Width           =   13890
         _ExtentX        =   24500
         _ExtentY        =   11165
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   8454016
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16250871
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridSaidas 
         Height          =   5310
         Left            =   -74940
         TabIndex        =   53
         Top             =   1140
         Width           =   15630
         _ExtentX        =   27570
         _ExtentY        =   9366
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   12640511
         BackColorSel    =   12632064
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
      Begin MSFlexGridLib.MSFlexGrid grid_DRE 
         Height          =   5190
         Left            =   -74925
         TabIndex        =   68
         Top             =   870
         Width           =   15630
         _ExtentX        =   27570
         _ExtentY        =   9155
         _Version        =   393216
         Rows            =   1
         Cols            =   26
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   16744576
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16250871
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   -68250
         X2              =   -68070
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label54 
         Caption         =   "Ano"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -74850
         TabIndex        =   72
         Top             =   450
         Width           =   345
      End
      Begin VB.Label Label51 
         Caption         =   $"frmGrafico4.frx":501F5
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1485
         Left            =   -64770
         TabIndex        =   70
         Top             =   510
         Width           =   5475
      End
      Begin VB.Label Label47 
         Caption         =   "Observações"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -64800
         TabIndex        =   65
         Top             =   4470
         Width           =   1125
      End
      Begin VB.Label Label28 
         Caption         =   "Qtde"
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
         Left            =   -74910
         TabIndex        =   57
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   -73245
         TabIndex        =   56
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Left            =   -67815
         TabIndex        =   55
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Tamanho e Cor"
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
         Left            =   -70980
         TabIndex        =   54
         Top             =   450
         Width           =   1260
      End
      Begin VB.Label Label24 
         Caption         =   "C"
         Height          =   255
         Left            =   -60720
         TabIndex        =   48
         Top             =   5310
         Width           =   225
      End
      Begin VB.Label Label23 
         Caption         =   "B"
         Height          =   255
         Left            =   -60720
         TabIndex        =   47
         Top             =   4950
         Width           =   225
      End
      Begin VB.Label Label22 
         Caption         =   "A"
         Height          =   255
         Left            =   -60720
         TabIndex        =   46
         Top             =   4530
         Width           =   225
      End
      Begin VB.Label Label21 
         Caption         =   "Nº total"
         Height          =   255
         Left            =   -60900
         TabIndex        =   45
         Top             =   4140
         Width           =   585
      End
      Begin VB.Label Label20 
         Caption         =   "Números em Itens"
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
         Left            =   -60960
         TabIndex        =   41
         Top             =   3810
         Width           =   1665
      End
      Begin VB.Label Label19 
         Caption         =   "Valor Total Considerado"
         Height          =   255
         Left            =   -60960
         TabIndex        =   38
         Top             =   5790
         Width           =   1755
      End
      Begin VB.Label Label18 
         Caption         =   "C                %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -60690
         TabIndex        =   33
         Top             =   1650
         Width           =   1425
      End
      Begin VB.Label Label17 
         Caption         =   "B                %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -60690
         TabIndex        =   32
         Top             =   1290
         Width           =   1425
      End
      Begin VB.Label Label16 
         Caption         =   "A               %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -60690
         TabIndex        =   31
         Top             =   930
         Width           =   1425
      End
      Begin VB.Label Label15 
         Caption         =   "Classe  -  Corte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -60960
         TabIndex        =   30
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tamanho e Cor"
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
         Left            =   -70980
         TabIndex        =   26
         Top             =   450
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Left            =   -67815
         TabIndex        =   23
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   -73245
         TabIndex        =   22
         Top             =   420
         Width           =   555
      End
      Begin VB.Label lbl_qtde01 
         Caption         =   "Qtde"
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
         Left            =   -74910
         TabIndex        =   21
         Top             =   420
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmd_pesquisar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   15825
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   30
      TabIndex        =   10
      Top             =   30
      Width           =   15825
      Begin VB.CommandButton cmd_calendarioDtFim 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4770
         Picture         =   "frmGrafico4.frx":502E1
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   547
         Width           =   465
      End
      Begin VB.CommandButton cmd_calendarioDtIni 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1650
         Picture         =   "frmGrafico4.frx":50BC3
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   547
         Width           =   465
      End
      Begin VB.TextBox txtSubclasseNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   570
         Width           =   3375
      End
      Begin VB.TextBox txt_parteCodigoProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   12990
         TabIndex        =   132
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt_parteNomeProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6690
         TabIndex        =   3
         Top             =   960
         Width           =   4245
      End
      Begin VB.TextBox txtClasseNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   3375
      End
      Begin VB.Data datClasses 
         Caption         =   "datClasses"
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
         Left            =   15660
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "con_Classe"
         Top             =   210
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Data datSubClasses 
         Caption         =   "datSubClasses"
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
         Left            =   15660
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "con_Sub_Classe"
         Top             =   660
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Por grandeza de ""Número de itens vendidos"""
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
         Left            =   11205
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   3990
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Por grandeza de ""Valor (R$) faturado"""
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
         Left            =   11205
         TabIndex        =   5
         Top             =   540
         Width           =   3450
      End
      Begin VB.ComboBox Combo_Filial 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmGrafico4.frx":514A5
         Left            =   450
         List            =   "frmGrafico4.frx":514A7
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   195
         Width           =   870
      End
      Begin VB.ComboBox cmb_numProdutos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmGrafico4.frx":514A9
         Left            =   1170
         List            =   "frmGrafico4.frx":514D1
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1230
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmGrafico4.frx":5150E
         DataSource      =   "datClasses"
         Height          =   345
         Left            =   6690
         TabIndex        =   1
         Top             =   180
         Width           =   855
         DataFieldList   =   "Código"
         _Version        =   196617
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1773
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Código"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Código"
         Columns(0).DataType=   3
         Columns(0).FieldLen=   256
         Columns(1).Width=   7064
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   1508
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   285
         Left            =   3540
         TabIndex        =   137
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   585
         Width           =   1185
         _ExtentX        =   2090
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
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   285
         Left            =   435
         TabIndex        =   138
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   585
         Width           =   1185
         _ExtentX        =   2090
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
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "frmGrafico4.frx":51527
         DataSource      =   "datSubClasses"
         Height          =   345
         Left            =   6690
         TabIndex        =   139
         Top             =   570
         Width           =   855
         DataFieldList   =   "Código"
         _Version        =   196617
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1773
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Código"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Código"
         Columns(0).DataType=   3
         Columns(0).FieldLen=   256
         Columns(1).Width=   7064
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   1508
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Até"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3195
         TabIndex        =   142
         Top             =   615
         Width           =   300
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "De"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   141
         Top             =   615
         Width           =   315
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "SubClasse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   140
         Top             =   615
         Width           =   870
      End
      Begin VB.Label Label55 
         Appearance      =   0  'Flat
         Caption         =   "ou  Parte do Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11250
         TabIndex        =   133
         Top             =   1035
         Width           =   1650
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "Parte do Nome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   1035
         Width           =   1260
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Classe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Visualizar os                       produtos mais vendidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   12
         Top             =   1005
         Width           =   4380
      End
      Begin VB.Label Nome_Filial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1335
         TabIndex        =   9
         Top             =   195
         Width           =   3900
      End
      Begin VB.Label Label7 
         Caption         =   "Filial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   225
         Width           =   465
      End
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10620
      TabIndex        =   58
      Top             =   1950
      Width           =   5175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8010
      TabIndex        =   28
      Top             =   1950
      Width           =   2475
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5370
      TabIndex        =   25
      Top             =   1950
      Width           =   2505
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   16
      Top             =   1950
      Width           =   2475
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   15
      Top             =   1950
      Width           =   2475
   End
End
Attribute VB_Name = "frmGrafico4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrayTamanhos() As Variant
Dim arrayCores() As Variant
Dim contador_arrayTamanhos As Integer
Dim contador_arrayCores As Integer
Dim arrayFiliais(20, 7) As String
Dim iConta As Integer
Dim s_DRE_situacaoTributaria As String
Dim s_DRE_PIS As String
Dim s_DRE_COFINS As String
Dim arrayCentroCustos() As Variant
Dim contador_arrayCentroCustos As Integer


Private Function AchaTamanho(pTamanho As Integer) As String
  Dim i As Integer
  AchaTamanho = ""
  
  For i = 0 To contador_arrayTamanhos - 1
      If arrayTamanhos(i, 0) = pTamanho Then
          AchaTamanho = arrayTamanhos(i, 1)
          Exit For
      End If
  Next
End Function

Private Function AchaCor(pCor As Integer) As String
  Dim i As Integer
  AchaCor = ""
  For i = 0 To contador_arrayCores - 1
      If arrayCores(i, 0) = pCor Then
          AchaCor = arrayCores(i, 1)
          Exit For
      End If
  Next
End Function

Private Sub cboClasse_CloseUp()
  Call cboClasse_LostFocus
End Sub

Private Sub cboClasse_LostFocus()
  Dim intClasse As Integer
    
On Error GoTo ErrHandler
  
  txtClasseNome.Text = ""
  
  If cboClasse.Text <> "" Then
    If Not IsDataType(dtInteger, cboClasse.Text, intClasse) Then
      DisplayMsg "Classe inválida."
      cboClasse.Text = ""
      Exit Sub
    End If
    
    If intClasse < 1 Or intClasse > 9999 Then
      DisplayMsg "Classe inválida."
      cboClasse.Text = ""
      Exit Sub
    End If
    
    With datClasses.Recordset
      .FindFirst "Código = " & intClasse
      If Not .NoMatch Then
        txtClasseNome.Text = .Fields("Nome").Value & ""
      End If
    End With
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cboSubClasse_CloseUp()
  Call cboSubClasse_LostFocus
End Sub

Private Sub cboSubClasse_LostFocus()
  Dim intSubClasse As Integer

On Error GoTo ErrHandler
  
  txtSubclasseNome.Text = ""
  
  If cboSubClasse.Text <> "" Then
    If Not IsDataType(dtInteger, cboSubClasse.Text, intSubClasse) Then
      DisplayMsg "Sub Classe inválida."
      cboSubClasse.Text = ""
      Exit Sub
    End If
    
    If intSubClasse < 1 Or intSubClasse > 9999 Then
      DisplayMsg "Sub Classe inválida."
      cboSubClasse.Text = ""
      Exit Sub
    End If
    
    With datSubClasses.Recordset
      .FindFirst "Código = " & intSubClasse
      If Not .NoMatch Then
        txtSubclasseNome.Text = .Fields("Nome").Value & ""
      End If
    End With
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_abreSequencia_Click()
On Error GoTo Erro

  If gridSaidas.RowSel < 1 Then
      MsgBox "Selecione uma linha na grade.", vbInformation
      Exit Sub
  End If
  
  Dim objSaidas As frmSaidas
  Set objSaidas = New frmSaidas
  
  objSaidas.txtSeq = gridSaidas.TextMatrix(gridSaidas.RowSel, 8)
  objSaidas.SearchRecord_peloNumSeq (gridSaidas.TextMatrix(gridSaidas.RowSel, 8))
  objSaidas.Show
  
  Set objSaidas = Nothing
    
  Exit Sub
  
Erro:
  MsgBox "Erro no detalhamento da Sequência" & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_calcularCurvaABC_Click()
On Error GoTo Erro

  Dim lContador As Long
  Dim arrPareto() As Variant
  Dim numTotalItens As Long
  Dim valorTotalItens As Double

  If IsNumeric(txt_paretoA.Text) = False Then
    DisplayMsg "Informe um percentual válido para Curva A"
    txt_paretoA.SetFocus
    Exit Sub
  End If

  If IsNumeric(txt_paretoB.Text) = False Then
    DisplayMsg "Informe um percentual válido para Curva B"
    txt_paretoB.SetFocus
    Exit Sub
  End If

  If IsNumeric(txt_paretoC.Text) = False Then
    DisplayMsg "Informe um percentual válido para Curva C"
    txt_paretoC.SetFocus
    Exit Sub
  End If

  If cmb_numProdutos.Text <> "TODOS" Then
'    DisplayMsg "Para realizar o cálculo da 'Curva ABC':" & vbCrLf & vbCrLf & "          - Informe uma Filial" & vbCrLf & "          - Entre com um período de datas" & vbCrLf & "          - Na combo 'visualizar' escolha TODOS" & vbCrLf & "          - Seleciona a opção 'Por grandeza de Valor (R$) faturado" & vbCrLf & vbCrLf & "Então clique no botão Calcular"
'    cmb_numProdutos.SetFocus
'    Exit Sub
      cmb_numProdutos.Text = "TODOS"
  End If
  
  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual a data final."
    Data_Ini.SetFocus
    Exit Sub
  End If

  If CDate(Data_Fim.Text) - CDate(Data_Ini.Text) > 62 Then
    DisplayMsg "Escolha um período de até 61 dias"
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If opt2.Value = False Then
'    DisplayMsg "Para realizar o cálculo da 'Curva ABC':" & vbCrLf & vbCrLf & "          - Informe uma Filial" & vbCrLf & "          - Entre com um período de datas" & vbCrLf & "          - Na combo 'visualizar' escolha TODOS" & vbCrLf & "          - Seleciona a opção 'Por grandeza de Valor (R$) faturado" & vbCrLf & vbCrLf & "Então clique no botão Calcular"
'    opt2.SetFocus
'    Exit Sub
      opt2.Value = True
  End If

  Screen.MousePointer = vbHourglass
  
  Dim rsEstoque As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
  Dim sTamanho As String
  Dim sCor As String
 
  gridProdutos_pareto.Rows = 1
  gridProdutos_pareto.Row = 0
  
  strSQL = "SELECT Sum(E.[Valor Vendas])-Sum(E.[Valor Devolução]) as SomaPreco, Sum(E.Vendas)-Sum(E.Devolução) AS SomaDeVendas, "
  strSQL = strSQL & " E.Produto, P.Nome, E.Tamanho, E.Cor "
  strSQL = strSQL & " From Estoque E, Produtos P "
  strSQL = strSQL & " where E.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " E.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
  
  If txt_parteNomeProduto.Text <> "" Then
      strSQL = strSQL & " P.Nome like '*" & txt_parteNomeProduto.Text & "*' and "
  End If

  strSQL = strSQL & " E.Produto=P.Código and "

  If cboClasse.Text <> "" Then
      strSQL = strSQL & " P.Classe=" & cboClasse.Text & " and "
  End If

  If cboSubClasse.Text <> "" Then
      strSQL = strSQL & " P.[Sub Classe]=" & cboSubClasse.Text & " and "
  End If

  strSQL = strSQL & " E.Filial=" & Combo_Filial.Text

  strSQL = strSQL & " GROUP BY E.Produto, P.Nome, E.Tamanho, E.cor "
  strSQL = strSQL & " ORDER BY 1 DESC "

  Screen.MousePointer = vbHourglass
  
  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  lngContadorRegGrid = 1
  
  If Not (rsEstoque.EOF And rsEstoque.BOF) Then
    rsEstoque.MoveFirst
    
    ReDim arrPareto(rsEstoque.RecordCount, 9)
  End If
  
  
  ' ***********************
  ' ETAPA 1
  lContador = 0
  valorTotalItens = 0
  While Not rsEstoque.EOF
  
      If rsEstoque.Fields(4).Value <> 0 Then
          sTamanho = AchaTamanho(rsEstoque.Fields(4).Value)
      Else
          sTamanho = ""
      End If

      If rsEstoque.Fields(5).Value <> 0 Then
          sCor = AchaCor(rsEstoque.Fields(5).Value)
      Else
          sCor = ""
      End If
      If rsEstoque.Fields(1).Value > 0 Then
          arrPareto(lContador, 0) = FormataValorTexto(rsEstoque.Fields(0).Value, 2)     ' valor total da qtde vendida/saida do item
          arrPareto(lContador, 1) = rsEstoque.Fields(1).Value                           ' qtde itens vendidos
          arrPareto(lContador, 2) = rsEstoque.Fields(2).Value                           ' código do produto
          arrPareto(lContador, 3) = rsEstoque.Fields(3).Value                           ' nome do produto
          arrPareto(lContador, 4) = sTamanho                                            ' nome tamanho
          arrPareto(lContador, 5) = sCor                                                ' nome cor
          
          lContador = lContador + 1
          valorTotalItens = valorTotalItens + CDbl(rsEstoque.Fields(0).Value)
      End If

      rsEstoque.MoveNext
  Wend
  rsEstoque.Close
  Set rsEstoque = Nothing
  
  txt_valorTotalConsiderado.Text = valorTotalItens
  txt_numTotalItens.Text = lContador
  numTotalItens = lContador
  ' ***********************

  ' ***********************
  ' ETAPA 2
  Dim numTotalItensA As Long
  Dim numTotalItensB As Long
  Dim numTotalItensC As Long
  Dim dValorAux As Double
  'Dim dValorAcumuladoAux As Double
  
  dValorAux = 0
  'dValorAcumuladoAux = 0
  numTotalItensA = 0
  numTotalItensB = 0
  numTotalItensC = 0
  
  For lContador = 0 To numTotalItens - 1
      ' Coluna 7 do array (% individual)
      ' Coluna 8 do array (% acumulado)
      ' Coluna 9 do array (Classificação)
      
      ' Coluna 7 do array (% individual)
      arrPareto(lContador, 6) = FormataValorTexto(CDbl(arrPareto(lContador, 0)) / valorTotalItens, 2)
      
      ' Coluna 8 do array (% acumulado)
      dValorAux = dValorAux + (CDbl(arrPareto(lContador, 0)) / valorTotalItens)
      arrPareto(lContador, 7) = FormataValorTexto(dValorAux, 2)
      
      ' Coluna 9 do array (Classificação)
      If dValorAux <= CDbl(txt_paretoA.Text) Then
          arrPareto(lContador, 8) = "A"
          numTotalItensA = numTotalItensA + 1
      ElseIf dValorAux <= CDbl(txt_paretoB.Text) Then
          arrPareto(lContador, 8) = "B"
          numTotalItensB = numTotalItensB + 1
      Else
          arrPareto(lContador, 8) = "C"
      End If
      
      gridProdutos_pareto.AddItem lContador + 1 & vbTab & arrPareto(lContador, 0) & vbTab & _
                  arrPareto(lContador, 1) & vbTab & _
                  arrPareto(lContador, 2) & vbTab & _
                  arrPareto(lContador, 3) & vbTab & _
                  arrPareto(lContador, 4) & vbTab & _
                  arrPareto(lContador, 5) & vbTab & _
                  arrPareto(lContador, 6) & vbTab & _
                  arrPareto(lContador, 7) & vbTab & _
                  arrPareto(lContador, 8)
  Next

  ' ***********************
  
  txt_numTotalItensA.Text = numTotalItensA
  txt_numTotalItensB.Text = numTotalItensB
  txt_numTotalItensC.Text = numTotalItens - numTotalItensA - numTotalItensB

  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar cálculo de Pareto... " & Err.Number & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_buscarInformacoesParaCalculoDRE_Click()
On Error GoTo Erro
  Dim rsSaidas As Recordset
  Dim rsEntradas As Recordset
  Dim strSQL As String
  Dim dQtde As Double
  Dim dPrecoCusto As Double
  Dim dTotalPrecoCusto As Double
  
  Dim sDataInicial As String
  Dim sDataFinal As String
  Dim mes As Integer
  Dim Ano As Integer
  Dim Dia As Integer
  Dim dataAux As Date
  Dim i As Integer

  
  txt_DRE_receitaBruta.Text = ""
  txt_DRE_Devolucoes.Text = ""
  txt_DRE_ImpostosSobreVendas.Text = ""
  txt_DRE_receitaOperacionalLiquida.Text = ""
  txt_DRE_cmv.Text = ""
  txt_DRE_lucroBruto.Text = ""
  txt_DRE_despesasAdministrativas.Text = ""
  txt_DRE_despesasComerciais.Text = ""
  txt_DRE_despesasDepreciacao.Text = ""
  txt_DRE_despesasFinanceiras.Text = ""
  txt_DRE_receitasFinanceiras.Text = ""
  txt_DRE_lucroPrejuizoOperacional.Text = ""
  txt_DRE_despesasNaoOperacionais.Text = ""
  txt_DRE_receitasNaoOperacionais.Text = ""
  txt_DRE_LAIR.Text = ""
  txt_DRE_provisaoIR.Text = ""
  txt_DRE_provisaoCSLL.Text = ""
  txt_DRE_LucroLiquido.Text = ""
  txt_DRE_Observacoes.Text = ""
  
  If cmb_ano.Text = "" Then
      DisplayMsg "Escolha o ANO para o cálculo."
      cmb_ano.SetFocus
      Exit Sub
  End If
  
  If cmb_mes.Text = "" Then
      DisplayMsg "Escolha o MÊS para o cálculo."
      cmb_mes.SetFocus
      Exit Sub
  End If
    
  mes = cmb_mes.ListIndex + 1
  Ano = CInt(cmb_ano.Text)
  
  If mes <= 9 Then
      sDataInicial = "01/0" & mes & "/" & Ano
  Else
      sDataInicial = "01/" & mes & "/" & Ano
  End If
  
  dataAux = DateAdd("m", 1, CDate(sDataInicial))
  sDataFinal = DateAdd("d", -1, dataAux)
  

'  If Not IsDate(Data_Ini.Text) Then
'    DisplayMsg "Escolha um período de datas."
'    Data_Ini.SetFocus
'    Exit Sub
'  End If
'
'  If Not IsDate(Data_Fim.Text) Then
'    DisplayMsg "Escolha um período de datas."
'    Data_Fim.SetFocus
'    Exit Sub
'  End If
'
'  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
'    DisplayMsg "Data inicial deve ser menor ou igual a data final."
'    Data_Ini.SetFocus
'    Exit Sub
'  End If
'
'  If CDate(Data_Fim.Text) - CDate(Data_Ini.Text) > 365 Then
'    DisplayMsg "Escolha um período de até 365 dias"
'    Data_Fim.SetFocus
'    Exit Sub
'  End If
  
  Screen.MousePointer = vbHourglass
    
  ' ==============================================================================
  ' Buscar Total de vendas no período (PREÇO DE VENDA)
  
  strSQL = "SELECT SUM(S.Total) "
  strSQL = strSQL & " FROM Saídas S, [Operações Saída] O"
  strSQL = strSQL & " Where S.Filial = " & Combo_Filial.Text
  strSQL = strSQL & " AND S.Data >= #" & Format(sDataInicial, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(sDataFinal, "mm/dd/yyyy") & " 23:59:59# "
'  strSQL = strSQL & " AND S.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
  strSQL = strSQL & " AND S.Efetivada=-1 AND S.Recebimento=-1 AND S.[Movimentação Desfeita]=0"
  strSQL = strSQL & " AND S.Operação = O.Código"
  strSQL = strSQL & " AND O.Tipo='V'"
  
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If Not IsNull(rsSaidas.Fields(0).Value) Then
      txt_DRE_receitaBruta.Text = Format(rsSaidas.Fields(0).Value, FORMAT_VALUE)
  Else
      txt_DRE_receitaBruta.Text = "0,00"
  End If
  rsSaidas.Close
  Set rsSaidas = Nothing
  ' ==============================================================================
  
  ' ==============================================================================
  ' Buscar Total de devoluções no período
  
  strSQL = "SELECT SUM(E.Total) "
  strSQL = strSQL & " FROM Entradas E, [Operações Entrada] O"
  strSQL = strSQL & " Where E.Filial = " & Combo_Filial.Text
  strSQL = strSQL & " AND E.Data >= #" & Format(sDataInicial, "mm/dd/yyyy") & "# AND  E.Data <= #" & Format(sDataFinal, "mm/dd/yyyy") & " 23:59:59# "
'  strSQL = strSQL & " AND E.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  E.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
  strSQL = strSQL & " AND E.Efetivada=-1 "
  strSQL = strSQL & " AND E.Operação = O.Código"
  strSQL = strSQL & " AND O.Tipo='D'"
  
  Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If Not IsNull(rsEntradas.Fields(0).Value) Then
      txt_DRE_Devolucoes.Text = Format(rsEntradas.Fields(0).Value, FORMAT_VALUE)
  Else
      txt_DRE_Devolucoes.Text = "0,00"
  End If
  rsEntradas.Close
  Set rsEntradas = Nothing
  ' ==============================================================================

  ' ==============================================================================
  ' Buscar Total de vendas no período (PREÇO DE CUSTO)
  
  strSQL = "SELECT Qtde, PrecoCusto "
  strSQL = strSQL & " FROM Saídas S, [Saídas - Produtos] P, [Operações Saída] O"
  strSQL = strSQL & " Where S.Filial = " & Combo_Filial.Text
  strSQL = strSQL & " AND S.Data >= #" & Format(sDataInicial, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(sDataFinal, "mm/dd/yyyy") & " 23:59:59# "
'  strSQL = strSQL & " AND S.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
  strSQL = strSQL & " AND S.Efetivada=-1 AND S.Recebimento=-1 AND S.[Movimentação Desfeita]=0"
  strSQL = strSQL & " AND S.Operação = O.Código"
  strSQL = strSQL & " AND O.Tipo='V'"
  strSQL = strSQL & " AND S.Filial=P.Filial"
  strSQL = strSQL & " AND S.Sequência=P.Sequência"
  
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  dTotalPrecoCusto = 0
  
  If Not (rsSaidas.EOF And rsSaidas.BOF) Then
      rsSaidas.MoveFirst
  
      While Not rsSaidas.EOF
          dQtde = rsSaidas.Fields(0).Value
          dPrecoCusto = rsSaidas.Fields(1).Value
          dTotalPrecoCusto = dTotalPrecoCusto + (dQtde * dPrecoCusto)
          rsSaidas.MoveNext
      Wend
  End If
  
  txt_DRE_cmv.Text = Format(dTotalPrecoCusto, FORMAT_VALUE)
  rsSaidas.Close
  Set rsSaidas = Nothing
  ' ==============================================================================

  ' ==============================================================================
  ' Buscar despesas Administrativas pelos Centro de Custos Selecionados
  Dim rsCentroCustoCP As Recordset
  Dim d_DRE_despesasAdministrativas As Double
  
  d_DRE_despesasAdministrativas = 0
  
  If contador_arrayCentroCustos > 0 Then
      For i = 0 To contador_arrayCentroCustos - 1
          If lst_centroCusto.Selected(i) = True Then
              strSQL = "Select SUM(Valor) From [Contas a Pagar] "
              strSQL = strSQL & " Where Filial = " & Combo_Filial.Text
              strSQL = strSQL & " AND Vencimento >= #" & Format(sDataInicial, "mm/dd/yyyy") & "# AND  Vencimento <= #" & Format(sDataFinal, "mm/dd/yyyy") & " 23:59:59# "
              strSQL = strSQL & " AND [Centro de Custo] = " & arrayCentroCustos(i, 0)
              
              Set rsCentroCustoCP = db.OpenRecordset(strSQL, dbOpenDynaset)
              
              If Not IsNull(rsCentroCustoCP.Fields(0).Value) Then
                  d_DRE_despesasAdministrativas = d_DRE_despesasAdministrativas + rsCentroCustoCP.Fields(0).Value
              End If
              rsCentroCustoCP.Close
              Set rsCentroCustoCP = Nothing
          End If
      Next
  End If
  txt_DRE_despesasAdministrativas.Text = Format(d_DRE_despesasAdministrativas, FORMAT_VALUE)
  ' ==============================================================================


  ' ==============================================================================
  ' Buscar os percentuais de CSLL e IRRF cadastrados em Parametros/Filial
  For i = 0 To iConta - 1
      If arrayFiliais(i, 0) = Combo_Filial.Text Then
          lbl_DRE_Percentual_IR.Caption = Format(arrayFiliais(i, 2), FORMAT_VALUE)
          lbl_DRE_Percentual_CSLL.Caption = Format(arrayFiliais(i, 3), FORMAT_VALUE)
          s_DRE_situacaoTributaria = arrayFiliais(i, 4)
          s_DRE_PIS = arrayFiliais(i, 5)
          s_DRE_COFINS = arrayFiliais(i, 6)
          
          Exit For
      End If
  Next
  ' ==============================================================================

  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar o cálculo do D.R.E. " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_detalharProdutoEntradas_Click()
On Error GoTo Erro
 
  Dim rsEntrada As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
  Dim sTamanho As String
  Dim sCor As String
  Dim lQtde As Long
 
  If gridProdutos.RowSel < 1 Then
    MsgBox "Selecione um produto da ABA AMARELA.", vbInformation
    SSTab1.Tab = 0
    Exit Sub
  End If
  
  'lbl_qtde01.Caption = gridProdutos.TextMatrix(0, 1)
  'txt_qtde01.Text = gridProdutos.TextMatrix(gridProdutos.RowSel, 1)
  txt_codProd01.Text = gridProdutos.TextMatrix(gridProdutos.RowSel, 3)
  txt_nmProd01.Text = gridProdutos.TextMatrix(gridProdutos.RowSel, 4)
  sTamanho = gridProdutos.TextMatrix(gridProdutos.RowSel, 10)
  sCor = gridProdutos.TextMatrix(gridProdutos.RowSel, 11)

  gridEntradas.Rows = 1
  gridEntradas.Row = 0

  strSQL = " Select E.Data, E.Operação, E.Fornecedor, C.Nome, EP.Qtde, EP.Preço, O.Nome "
  strSQL = strSQL & " From Entradas E, [Entradas - Produtos] EP, Cli_for C, [Operações Entrada] O "
  
  If sTamanho <> "0" Then
      If Len(sTamanho) = 1 Then
          sTamanho = "00" & sTamanho
      ElseIf Len(sTamanho) = 2 Then
          sTamanho = "0" & sTamanho
      End If
      
      If Len(sCor) = 1 Then
          sCor = "00" & sCor
      ElseIf Len(sCor) = 2 Then
          sCor = "0" & sCor
      End If
    
      txt_tamanhoCor = " " & gridProdutos.TextMatrix(gridProdutos.RowSel, 5) & " - " & gridProdutos.TextMatrix(gridProdutos.RowSel, 6)
  
      strSQL = strSQL & " where EP.Código = '" & txt_codProd01.Text & sTamanho & sCor & "' AND "
  Else
      strSQL = strSQL & " where EP.Código = '" & txt_codProd01.Text & "' AND "
      
      txt_tamanhoCor.Text = ""
  End If
  
  strSQL = strSQL & " EP.Filial = " & Combo_Filial.Text & " AND "
  strSQL = strSQL & " EP.Filial = E.Filial AND "
  strSQL = strSQL & " EP.Sequência = E.Sequência AND "
  strSQL = strSQL & " E.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " E.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
  strSQL = strSQL & " E.Fornecedor = C.Código and "
  strSQL = strSQL & " E.Operação = O.Código "
  strSQL = strSQL & " ORDER BY E.data DESC "

  Screen.MousePointer = vbHourglass
  
  Set rsEntrada = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rsEntrada.EOF And rsEntrada.BOF) Then
    rsEntrada.MoveFirst
  End If

  lQtde = 0
  While Not rsEntrada.EOF
      lQtde = lQtde + CLng(rsEntrada.Fields(4).Value)

      gridEntradas.AddItem rsEntrada.Fields(0).Value & vbTab & _
                          rsEntrada.Fields(4).Value & vbTab & _
                          FormataValorTexto(rsEntrada.Fields(5).Value, 2) & vbTab & _
                          rsEntrada.Fields(2).Value & vbTab & _
                          rsEntrada.Fields(3).Value & vbTab & _
                          rsEntrada.Fields(1).Value & vbTab & _
                          rsEntrada.Fields(6).Value
      rsEntrada.MoveNext
  Wend
  rsEntrada.Close
  Set rsEntrada = Nothing

  txt_qtde01.Text = lQtde
  SSTab1.Tab = 2

  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_detalharProdutoSaidas_Click()
On Error GoTo Erro
 
  Dim rsSaida As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
  Dim sTamanho As String
  Dim sCor As String
  Dim lQtde As Long
  Dim NSU_hora As String
 
  If gridProdutos.RowSel < 1 Then
    MsgBox "Selecione um produto da ABA AMARELA.", vbInformation
    SSTab1.Tab = 0
    Exit Sub
  End If

  'lbl_qtde02.Caption = gridProdutos.TextMatrix(0, 1)
  'txt_qtde02.Text = gridProdutos.TextMatrix(gridProdutos.RowSel, 1)
  txt_codProd02.Text = gridProdutos.TextMatrix(gridProdutos.RowSel, 3)
  txt_nmProd02.Text = gridProdutos.TextMatrix(gridProdutos.RowSel, 4)
  sTamanho = gridProdutos.TextMatrix(gridProdutos.RowSel, 10)
  sCor = gridProdutos.TextMatrix(gridProdutos.RowSel, 11)

  gridSaidas.Rows = 1
  gridSaidas.Row = 0

  strSQL = " Select S.Data, S.Operação, S.Cliente, C.Nome, SP.Qtde, SP.Preço, O.Nome, S.Sequência, S.NSU_hora "
  strSQL = strSQL & " From Saídas S, [Saídas - Produtos] SP, Cli_for C, [Operações Saída] O "
  
  If sTamanho <> "0" Then
      If Len(sTamanho) = 1 Then
          sTamanho = "00" & sTamanho
      ElseIf Len(sTamanho) = 2 Then
          sTamanho = "0" & sTamanho
      End If
      
      If Len(sCor) = 1 Then
          sCor = "00" & sCor
      ElseIf Len(sCor) = 2 Then
          sCor = "0" & sCor
      End If
    
      txt_tamanhoCor02 = " " & gridProdutos.TextMatrix(gridProdutos.RowSel, 5) & " - " & gridProdutos.TextMatrix(gridProdutos.RowSel, 6)
  
      strSQL = strSQL & " where SP.Código = '" & txt_codProd02.Text & sTamanho & sCor & "' AND "
  Else
      strSQL = strSQL & " where SP.Código = '" & txt_codProd02.Text & "' AND "
      
      txt_tamanhoCor02.Text = ""
  End If
  
  strSQL = strSQL & " SP.Filial = " & Combo_Filial.Text & " AND "
  strSQL = strSQL & " SP.Filial = S.Filial AND "
  strSQL = strSQL & " SP.Sequência = S.Sequência AND "
  strSQL = strSQL & " S.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " S.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
  strSQL = strSQL & " S.Cliente = C.Código and "
  strSQL = strSQL & " S.Operação = O.Código "
  strSQL = strSQL & " ORDER BY S.data DESC "

  Screen.MousePointer = vbHourglass
  
  Set rsSaida = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rsSaida.EOF And rsSaida.BOF) Then
    rsSaida.MoveFirst
  End If

  lQtde = 0
  While Not rsSaida.EOF
  
      lQtde = lQtde + CLng(rsSaida.Fields(4).Value)
      
      NSU_hora = ""
      If Not IsNull(rsSaida.Fields(8).Value) Then
          NSU_hora = rsSaida.Fields(0).Value & " " & rsSaida.Fields(8).Value
      Else
          NSU_hora = rsSaida.Fields(0).Value
      End If
  
      gridSaidas.AddItem vbTab & NSU_hora & vbTab & _
                          rsSaida.Fields(4).Value & vbTab & _
                          FormataValorTexto(rsSaida.Fields(5).Value, 2) & vbTab & _
                          rsSaida.Fields(2).Value & vbTab & _
                          rsSaida.Fields(3).Value & vbTab & _
                          rsSaida.Fields(1).Value & vbTab & _
                          rsSaida.Fields(6).Value & vbTab & _
                          rsSaida.Fields(7).Value
      rsSaida.MoveNext
  Wend
  rsSaida.Close
  Set rsSaida = Nothing

  txt_qtde02.Text = lQtde
  SSTab1.Tab = 1

  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_DRE_excluir_Click()
On Error GoTo Erro
  Dim nResponse As Variant

  If grid_DRE.RowSel < 1 Then
    MsgBox "Selecione um D.R.E. na grade.", vbInformation
    Exit Sub
  End If
  
  nResponse = MsgBox("Deseja realmente excluir este D.R.E.?", vbYesNo, "Atenção")
  If nResponse = vbNo Then
      Exit Sub
  End If
  
  db.Execute "Delete From DRE_quick where CodigoDRE = " & grid_DRE.TextMatrix(grid_DRE.RowSel, 1)
  
  MsgBox "D.R.E. excluído com sucesso", vbInformation, "Sucesso"
  
  cmd_DRE_pesquisar_Click

  Exit Sub
Erro:
  MsgBox "Erro ao realizar exclusão...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_DRE_finalizarCalculo_Click()
On Error GoTo Erro
  Dim sStrSql As String
  Dim d_DRE_receitaBruta As Double
  Dim d_subtrai_DRE_Devolucoes As Double
  Dim d_subtrai_DRE_ImpostosSobreVendas As Double
  Dim d_subtrai_DRE_ImpostosSobreVendasPer As Double
  Dim d_igual_DRE_receitaOperacionalLiquida As Double
  Dim d_igual_DRE_lucroBruto As Double
  Dim d_subtrai_DRE_despesasAdministrativas As Double
  Dim d_subtrai_DRE_despesasComerciais As Double
  Dim d_subtrai_DRE_despesasDepreciacao As Double
  Dim d_subtrai_DRE_despesasFinanceiras As Double
  Dim d_adiciona_DRE_receitasFinanceiras As Double
  Dim d_igual_DRE_lucroPrejuizoOperacional As Double
  Dim d_subtrai_DRE_despesasNaoOperacionais As Double
  Dim d_adiciona_DRE_receitasNaoOperacionais As Double
  Dim d_igual_DRE_LAIR As Double
  Dim d_subtrai_DRE_provisaoIR As Double
  Dim d_subtrai_DRE_provisaoCSLL As Double
  Dim d_igual_DRE_LucroLiquido As Double

  If Trim(txt_DRE_receitaBruta.Text) = "" Or Trim(txt_DRE_receitaBruta.Text) = "0,00" Then
      MsgBox "Não existe receita bruta até o momento para este Ano/Mês.", vbInformation, "Atenção"
      cmd_DRE_salvar.Enabled = True
      Exit Sub
  End If
  
  If cmb_ano.Text = "" Then
      DisplayMsg "Escolha o ANO."
      cmb_ano.SetFocus
      Exit Sub
  End If
  
  If cmb_mes.Text = "" Then
      DisplayMsg "Escolha o MÊS."
      cmb_mes.SetFocus
      Exit Sub
  End If
  
  If cmb_DRE_tipoAnexo.Text = "" Then
      MsgBox "Informe o seu tipo de empresa", vbInformation, "Atenção"
      cmb_DRE_tipoAnexo.SetFocus
      Exit Sub
  End If
  
  If txt_DRE_receitaBruta.Text = "" Then
      txt_DRE_receitaBruta.Text = "0,00"
  End If

  If txt_DRE_Devolucoes.Text = "" Then
      txt_DRE_Devolucoes.Text = "0,00"
  End If

  If txt_DRE_ImpostosSobreVendas.Text = "" Then
      txt_DRE_ImpostosSobreVendas.Text = "0,00"
  End If

  d_DRE_receitaBruta = CCur(txt_DRE_receitaBruta.Text)
  d_subtrai_DRE_Devolucoes = CCur(txt_DRE_Devolucoes.Text)
  
  Screen.MousePointer = vbHourglass
  
  
  If s_DRE_situacaoTributaria = "1" Then
      ' SIMPLES NACIONAL
      
      ' ===============================================================================================
      ' Buscar os D.R.E. dos últimos 12 meses para obter o percentual de imposto sobre este mês em questão
      ' Ex: Se o usuário esta calculando o D.R.E. de Março/2020 então buscar os outros 12 meses para trás (Fev/2020 até Mar/2019)
      Dim rsDRE_quick As Recordset
      Dim rsDRE_anexos As Recordset
      Dim iContadorMeses As Integer
      Dim d_DRE_Aliquota As Double
      Dim d_DRE_Redutor As Double
      Dim iAno As Integer
      Dim iMes As Integer
      Dim iAnoParametro As Integer
      Dim iMesParametro As Integer
      Dim sValorAux As String
      
      iAnoParametro = CInt(cmb_ano.Text)
      iMesParametro = cmb_mes.ListIndex + 1
      
      sStrSql = "Select * from DRE_quick "
      sStrSql = sStrSql & " Where DataAno >= " & CInt(cmb_ano.Text) - 1
      sStrSql = sStrSql & " AND DataAno <= " & CInt(cmb_ano.Text)
      sStrSql = sStrSql & " AND Filial = " & Combo_Filial.Text
    
      Set rsDRE_quick = db.OpenRecordset(sStrSql, dbOpenDynaset, dbReadOnly)
    
      iContadorMeses = 0
      While Not rsDRE_quick.EOF
          iAno = rsDRE_quick.Fields("DataANO").Value
          iMes = rsDRE_quick.Fields("DataMES").Value
          If (iAno = iAnoParametro And iMes < iMesParametro) Or (iAno = (iAnoParametro - 1) And iMes >= iMesParametro) Then
              
              iContadorMeses = iContadorMeses + 1
              d_subtrai_DRE_ImpostosSobreVendas = d_subtrai_DRE_ImpostosSobreVendas + (rsDRE_quick.Fields("ReceitaBruta").Value - rsDRE_quick.Fields("Devolucoes").Value)
          
          End If
          rsDRE_quick.MoveNext
      Wend
      rsDRE_quick.Close
      Set rsDRE_quick = Nothing
    
      If d_subtrai_DRE_ImpostosSobreVendas > 1 Then
          
          sStrSql = "Select Aliquota, ValorRedutor from DRE_anexos "
          
          If cmb_DRE_tipoAnexo.Text = "COMÉRCIO" Then
              sStrSql = sStrSql & " Where CodigoAnexo = 1 " 'ANEXO I
          ElseIf cmb_DRE_tipoAnexo.Text = "INDUSTRIA/FÁBRICA" Then
              sStrSql = sStrSql & " Where CodigoAnexo = 2 " 'ANEXO II
          End If
          
          sValorAux = CStr(d_subtrai_DRE_ImpostosSobreVendas)
          sValorAux = Replace(sValorAux, ",", ".")
          
          sStrSql = sStrSql & " AND (ValorDe < " & sValorAux
          sStrSql = sStrSql & " AND ValorAte > " & sValorAux & ")"
    
          ' Imaginar uma loja de roupas(inserida no AnexoI–Comércio), cujo faturamento
          ' em janeiro/2019 tenha sido R$ 50.000,00
          ' e que a receita bruta acumulada nos 12 meses anteriores (RBA12) tenha sido de R$ 320.000,00.
          ' A alíquota que a empresa pagará sobre o faturamento de janeiro/2019 será
          ' calculada seguindo a fórmula:
          '     [(R$ 320.000,00 x 7,30%) – 5.940,00] / R$ 320.000,00 = 5,444%
          ' Neste exemplo, a alíquota efetiva é de 5,444%, conforme o Anexo I (tabela do DB DRE_anexos),
          ' para empresas que faturam anualmente entre R$ 180.000,01 e R$ 360.000,00.
    
          Set rsDRE_anexos = db.OpenRecordset(sStrSql, dbOpenDynaset, dbReadOnly)
          
          If Not IsNull(rsDRE_anexos.Fields("Aliquota").Value) Then
              d_DRE_Redutor = rsDRE_anexos.Fields("ValorRedutor").Value
              d_DRE_Aliquota = rsDRE_anexos.Fields("Aliquota").Value
              
              d_subtrai_DRE_ImpostosSobreVendasPer = ((d_subtrai_DRE_ImpostosSobreVendas * (d_DRE_Aliquota / 100)) - d_DRE_Redutor) / d_subtrai_DRE_ImpostosSobreVendas
          End If
          rsDRE_anexos.Close
          Set rsDRE_anexos = Nothing
      End If
    
      d_subtrai_DRE_ImpostosSobreVendas = ((d_DRE_receitaBruta - d_subtrai_DRE_Devolucoes) * d_subtrai_DRE_ImpostosSobreVendasPer)
      txt_DRE_ImpostosSobreVendas.Text = Format(d_subtrai_DRE_ImpostosSobreVendas, FORMAT_VALUE)
      ' ===============================================================================================
  Else
      ' LUCRO REAL
      
      ' PIS/COFINS
      ' Sobre o Faturamento de 50.000,00 de venda de mercadorias e 50.000,00 sobre venda de serviços,
      ' será tributado em R$ 1.650,00 de PIS (100.000,00 x 1,65%) e R$ 7.600,00 de COFINS (100.000,00 x 7,60%).

      ' IRPJ/CSLL
      ' Para calcular o IRPJ deve chegar ao valor do lucro para aplicar o percentual de 15% de IRPJ e CSLL de 9%.
      
      ' Receita de Vendas – Impostos diretos    – Custos    – Despesas  = Lucro
      ' 100.000,00        – 1.650,00 – 7.600,00 – 50.000,00 – 20.000,00 = 20.750,00

      ' IRPJ = 20.750,00 x 15% = 3.112,50
      ' CSLL = 20.750,00 x 9% = 1.867,50

      ' Nota: Cálculos acima realizados para fins didáticos, porém não está sendo considerado
      ' os créditos de Pis/Pasep e Cofins não cumulativo.
      ' O IRPJ adicional deve se verificar se ultrapassa o limite de 20.000,00 ao mês
      ' ou 60.000,00 no trimestre e o que exceder este valor deve ser aplicado à alíquota de 10%.

      ' IRPJ Adicional = 20.750,00 – 20.000,00 = 750,00 x 10% = 75,00

      d_subtrai_DRE_ImpostosSobreVendas = ((d_DRE_receitaBruta - d_subtrai_DRE_Devolucoes) * CDbl(s_DRE_PIS) / 100) + ((d_DRE_receitaBruta - d_subtrai_DRE_Devolucoes) * CDbl(s_DRE_COFINS) / 100)
      txt_DRE_ImpostosSobreVendas.Text = Format(d_subtrai_DRE_ImpostosSobreVendas, FORMAT_VALUE)
  End If


  d_igual_DRE_receitaOperacionalLiquida = d_DRE_receitaBruta - d_subtrai_DRE_Devolucoes - d_subtrai_DRE_ImpostosSobreVendas

  txt_DRE_receitaOperacionalLiquida.Text = Format(d_igual_DRE_receitaOperacionalLiquida, FORMAT_VALUE)
  d_igual_DRE_lucroBruto = d_igual_DRE_receitaOperacionalLiquida - CCur(txt_DRE_cmv.Text)
  txt_DRE_lucroBruto.Text = Format(d_igual_DRE_lucroBruto, FORMAT_VALUE)

  If txt_DRE_despesasAdministrativas.Text = "" Then
      d_subtrai_DRE_despesasAdministrativas = 0
  Else
      d_subtrai_DRE_despesasAdministrativas = CCur(txt_DRE_despesasAdministrativas.Text)
  End If

  If txt_DRE_despesasComerciais.Text = "" Then
      d_subtrai_DRE_despesasComerciais = 0
  Else
      d_subtrai_DRE_despesasComerciais = CCur(txt_DRE_despesasComerciais.Text)
  End If

  If txt_DRE_despesasDepreciacao.Text = "" Then
      d_subtrai_DRE_despesasDepreciacao = 0
  Else
      d_subtrai_DRE_despesasDepreciacao = CCur(txt_DRE_despesasDepreciacao.Text)
  End If

  If txt_DRE_despesasFinanceiras.Text = "" Then
      d_subtrai_DRE_despesasFinanceiras = 0
  Else
      d_subtrai_DRE_despesasFinanceiras = CCur(txt_DRE_despesasFinanceiras.Text)
  End If

  If txt_DRE_receitasFinanceiras.Text = "" Then
      d_adiciona_DRE_receitasFinanceiras = 0
  Else
      d_adiciona_DRE_receitasFinanceiras = CCur(txt_DRE_receitasFinanceiras.Text)
  End If
  
  d_igual_DRE_lucroPrejuizoOperacional = d_igual_DRE_lucroBruto - d_subtrai_DRE_despesasAdministrativas - d_subtrai_DRE_despesasComerciais - d_subtrai_DRE_despesasDepreciacao - d_subtrai_DRE_despesasFinanceiras + d_adiciona_DRE_receitasFinanceiras
  txt_DRE_lucroPrejuizoOperacional.Text = Format(d_igual_DRE_lucroPrejuizoOperacional, FORMAT_VALUE)
  
  If txt_DRE_despesasNaoOperacionais.Text = "" Then
      d_subtrai_DRE_despesasNaoOperacionais = 0
  Else
      d_subtrai_DRE_despesasNaoOperacionais = CCur(txt_DRE_despesasNaoOperacionais.Text)
  End If

  If txt_DRE_receitasNaoOperacionais.Text = "" Then
      d_adiciona_DRE_receitasNaoOperacionais = 0
  Else
      d_adiciona_DRE_receitasNaoOperacionais = CCur(txt_DRE_receitasNaoOperacionais.Text)
  End If

  d_igual_DRE_LAIR = d_igual_DRE_lucroPrejuizoOperacional - d_subtrai_DRE_despesasNaoOperacionais + d_adiciona_DRE_receitasNaoOperacionais
  txt_DRE_LAIR.Text = Format(d_igual_DRE_LAIR, FORMAT_VALUE)
  
  
  d_subtrai_DRE_provisaoIR = (CCur(lbl_DRE_Percentual_IR.Caption) * d_igual_DRE_LAIR) / 100
  If d_subtrai_DRE_provisaoIR < 0 Then
      d_subtrai_DRE_provisaoIR = d_subtrai_DRE_provisaoIR * -1
  End If
  txt_DRE_provisaoIR.Text = Format(d_subtrai_DRE_provisaoIR, FORMAT_VALUE)
  
  d_subtrai_DRE_provisaoCSLL = (CCur(lbl_DRE_Percentual_CSLL.Caption) * d_igual_DRE_LAIR) / 100
  If d_subtrai_DRE_provisaoCSLL < 0 Then
      d_subtrai_DRE_provisaoCSLL = d_subtrai_DRE_provisaoCSLL * -1
  End If
  txt_DRE_provisaoCSLL.Text = Format(d_subtrai_DRE_provisaoCSLL, FORMAT_VALUE)
  
  txt_DRE_LucroLiquido.Text = Format(d_igual_DRE_LAIR - d_subtrai_DRE_provisaoIR - d_subtrai_DRE_provisaoCSLL, FORMAT_VALUE)

  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_DRE_help_Click()
    Dim strfile As String
    Dim objHelp As clsGeral
    Set objHelp = New clsGeral
    strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
    'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
    'Call objHelp.Show(strfile, "QuickStore10Help")
    Call objHelp.Show(strfile, "QuickStore10Help", 10058)
    Set objHelp = Nothing
End Sub

Private Sub cmd_DRE_imprimir_Click()
On Error GoTo Erro

  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome As String
  Dim strNomeLPT As String
  Dim strPortaLPT As String
  Dim intX As Integer
  Dim i As Integer
  
  strNome = "REL"
  strNomeLPT = "NOME IMPRESSORA REL"
  strPortaLPT = "PORTA IMPRESSORA REL"

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

  Dim nRow As Long
  Dim sLinha As String
  Dim sAUX As String
  
  If grid_DRE.Rows <= 1 Then
      MsgBox "Carregue o D.R.E. na grade.", vbInformation, "Atenção"
      Exit Sub
  End If
  nRow = grid_DRE.Rows
  
  Printer.Font = "LUCIDA CONSOLE"
  
  Printer.Print ""
  sLinha = "                                                                   Quick Store 10 - Soluções Comerciais inteligentes"
  
  Printer.Print ""

  sLinha = "   Relatório de D.R.E. - Demonstrativo do Resultado do Exercício"
  Printer.Print sLinha
  
  Printer.Print ""
  
  sLinha = "   " & Combo_Filial.Text & " - " & Nome_Filial.Caption
  Printer.Print sLinha
 
  ' ************************** ATENÇÃO ***********************************
  ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
  ' De preferência com o mesmo nome da impressora !!!
  
  With grid_DRE
      For nRow = 1 To .Rows - 1
 
          Printer.Print ""
          Printer.Print "   _________________________________________________________________________________________________________________"
          Printer.Print ""
          
          sAUX = .TextMatrix(nRow, 1)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   ID                            : " & sAUX
          
          sAUX = .TextMatrix(nRow, 3)
          Printer.Print "   Data Início                   :      " & sAUX
      
          sAUX = .TextMatrix(nRow, 4)
          Printer.Print "   Data Fim                      :      " & sAUX
      
          sAUX = .TextMatrix(nRow, 5)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Receita Bruta                 : " & sAUX
      
          sAUX = .TextMatrix(nRow, 6)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Devoluções                    : " & sAUX
          
          sAUX = .TextMatrix(nRow, 7)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Imposto s/ Vendas             : " & sAUX
      
          sAUX = .TextMatrix(nRow, 8)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Receita Op.Líquida            : " & sAUX
      
          sAUX = .TextMatrix(nRow, 9)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   CMV                           : " & sAUX
      
          sAUX = .TextMatrix(nRow, 10)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Lucro Bruto                   : " & sAUX

          sAUX = .TextMatrix(nRow, 11)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Administrativas      : " & sAUX
      
          sAUX = .TextMatrix(nRow, 12)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Comerciais           : " & sAUX
      
          sAUX = .TextMatrix(nRow, 13)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Depreciação          : " & sAUX
      
          sAUX = .TextMatrix(nRow, 14)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Financeiras          : " & sAUX
      
          sAUX = .TextMatrix(nRow, 15)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Receitas Financeiras          : " & sAUX
      
          sAUX = .TextMatrix(nRow, 16)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Lucro/Prejuízo Operacional    : " & sAUX
      
          sAUX = .TextMatrix(nRow, 17)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Não Operacionais     : " & sAUX
      
          sAUX = .TextMatrix(nRow, 18)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Receitas Não Operacionais     : " & sAUX
      
          sAUX = .TextMatrix(nRow, 19)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   LAIR                          : " & sAUX
          
          sAUX = .TextMatrix(nRow, 20)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Provisão IR                   : " & sAUX
      
          sAUX = .TextMatrix(nRow, 21)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Provisão CSLL                 : " & sAUX
      
          sAUX = .TextMatrix(nRow, 22)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Lucro Líquido                 : " & sAUX
          
          sAUX = .TextMatrix(nRow, 23)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Data Criação                  : " & sAUX
          
          sAUX = .TextMatrix(nRow, 24)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Usuário                       : " & sAUX
          
          sAUX = .TextMatrix(nRow, 25)
          If Len(sAUX) > 70 Then
              sLinha = "   Observações                   : " & Mid(sAUX, 1, 70)
              Printer.Print sLinha
              If Len(sAUX) > 140 Then
                  sLinha = "                                   " & Mid(sAUX, 71, 70)
                  Printer.Print sLinha
                  If Len(sAUX) > 210 Then
                      sLinha = "                                   " & Mid(sAUX, 141, 70)
                      Printer.Print sLinha
                      Printer.Print "                                   " & Mid(sAUX, 211, Len(sAUX) - 210)
                  Else
                      Printer.Print "                                   " & Mid(sAUX, 141, Len(sAUX) - 140)
                  End If
              Else
                  Printer.Print "                                   " & Mid(sAUX, 71, Len(sAUX) - 70)
              End If
              
          Else
              Printer.Print "   Observações                   : " & sAUX
          End If
      Next nRow
  End With
  
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""

  Printer.EndDoc

  Exit Sub
Erro:
  MsgBox "Erro ao realizar impressão do D.R.E...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmd_DRE_imprimirUM_Click()
On Error GoTo Erro

  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome As String
  Dim strNomeLPT As String
  Dim strPortaLPT As String
  Dim intX As Integer
  Dim i As Integer
  
  strNome = "REL"
  strNomeLPT = "NOME IMPRESSORA REL"
  strPortaLPT = "PORTA IMPRESSORA REL"

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

  Dim nRow As Long
  Dim sLinha As String
  Dim sAUX As String
  
  If grid_DRE.Rows <= 1 Then
      MsgBox "Carregue o D.R.E. na grade.", vbInformation, "Atenção"
      Exit Sub
  End If
  nRow = grid_DRE.Rows
  
  Printer.Font = "LUCIDA CONSOLE"
  
  Printer.Print ""
  sLinha = "                                                                   Quick Store 10 - Soluções Comerciais inteligentes"
  
  Printer.Print ""

  sLinha = "   Relatório de D.R.E. - Demonstrativo do Resultado do Exercício"
  Printer.Print sLinha
  
  Printer.Print ""
  
  sLinha = "   " & Combo_Filial.Text & " - " & Nome_Filial.Caption
  Printer.Print sLinha
 
  ' ************************** ATENÇÃO ***********************************
  ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
  ' De preferência com o mesmo nome da impressora !!!
  
  If grid_DRE.RowSel > 0 Then
      With grid_DRE
          sAUX = .TextMatrix(.RowSel, 1)
          
          Printer.Print ""
          Printer.Print "   _________________________________________________________________________________________________________________"
          Printer.Print ""
          
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   ID                            : " & sAUX
                  
          sAUX = .TextMatrix(.RowSel, 3)
          Printer.Print "   Data Início                   :      " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 4)
          Printer.Print "   Data Fim                      :      " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 5)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Receita Bruta                 : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 6)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Devoluções                    : " & sAUX
                  
          sAUX = .TextMatrix(.RowSel, 7)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Imposto s/ Vendas             : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 8)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Receita Op.Líquida            : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 9)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   CMV                           : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 10)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Lucro Bruto                   : " & sAUX
        
          sAUX = .TextMatrix(.RowSel, 11)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Administrativas      : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 12)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Comerciais           : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 13)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Depreciação          : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 14)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Financeiras          : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 15)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Receitas Financeiras          : " & sAUX
              
          sAUX = .TextMatrix(.RowSel, 16)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Lucro/Prejuízo Operacional    : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 17)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Despesas Não Operacionais     : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 18)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Receitas Não Operacionais     : " & sAUX
              
          sAUX = .TextMatrix(.RowSel, 19)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   LAIR                          : " & sAUX
          
          sAUX = .TextMatrix(.RowSel, 20)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Provisão IR                   : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 21)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Provisão CSLL                 : " & sAUX
      
          sAUX = .TextMatrix(.RowSel, 22)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Lucro Líquido                 : " & sAUX
          
          sAUX = .TextMatrix(.RowSel, 23)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Data Criação                  : " & sAUX
                  
          sAUX = .TextMatrix(.RowSel, 24)
          If Len(sAUX) < 15 Then
            For i = Len(sAUX) To 14
                sAUX = " " & sAUX
            Next
          End If
          Printer.Print "   Usuário                       : " & sAUX
          
          sAUX = .TextMatrix(.RowSel, 25)
          If Len(sAUX) > 70 Then
              sLinha = "   Observações                   : " & Mid(sAUX, 1, 70)
              Printer.Print sLinha
              If Len(sAUX) > 140 Then
                  sLinha = "                                   " & Mid(sAUX, 71, 70)
                  Printer.Print sLinha
                  If Len(sAUX) > 210 Then
                      sLinha = "                                   " & Mid(sAUX, 141, 70)
                      Printer.Print sLinha
                      Printer.Print "                                   " & Mid(sAUX, 211, Len(sAUX) - 210)
                  Else
                      Printer.Print "                                   " & Mid(sAUX, 141, Len(sAUX) - 140)
                  End If
              Else
                  Printer.Print "                                   " & Mid(sAUX, 71, Len(sAUX) - 70)
              End If
              
          Else
              Printer.Print "   Observações                   : " & sAUX
          End If
                  
      End With
  End If
  
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""

  Printer.EndDoc

  Exit Sub
Erro:
  MsgBox "Erro ao realizar impressão do D.R.E...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmd_DRE_limparTela_Click()
On Error GoTo Erro

  txt_DRE_receitaBruta.Text = ""
  txt_DRE_Devolucoes.Text = ""
  txt_DRE_ImpostosSobreVendas.Text = ""
  txt_DRE_receitaOperacionalLiquida.Text = ""
  txt_DRE_cmv.Text = ""
  txt_DRE_lucroBruto.Text = ""
  txt_DRE_despesasAdministrativas.Text = ""
  txt_DRE_despesasComerciais.Text = ""
  txt_DRE_despesasDepreciacao.Text = ""
  txt_DRE_despesasFinanceiras.Text = ""
  txt_DRE_receitasFinanceiras.Text = ""
  txt_DRE_lucroPrejuizoOperacional.Text = ""
  txt_DRE_despesasNaoOperacionais.Text = ""
  txt_DRE_receitasNaoOperacionais.Text = ""
  txt_DRE_LAIR.Text = ""
  txt_DRE_provisaoIR.Text = ""
  txt_DRE_provisaoCSLL.Text = ""
  txt_DRE_LucroLiquido.Text = ""
  txt_DRE_Observacoes.Text = ""
  cmb_ano.ListIndex = -1
  cmb_mes.ListIndex = -1
  cmb_DRE_tipoAnexo.ListIndex = -1
  
  Dim i As Integer
  If contador_arrayCentroCustos > 0 Then
      For i = 0 To contador_arrayCentroCustos - 1
          lst_centroCusto.Selected(i) = False
      Next
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro ao realizar limpeza da tela D.R.E...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_DRE_pesquisar_Click()
On Error GoTo Erro
  Dim rsDRE As Recordset
  Dim strSQL As String
  Dim sMesDRE As String
 
  If Nome_Filial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    Combo_Filial.SetFocus
    Exit Sub
  End If
  
  grid_DRE.Rows = 1
  
  strSQL = " Select * from DRE_Quick "
  strSQL = strSQL & " where Filial = " & Combo_Filial.Text
  
  If cmb_DRE_AnoPesquisa.Text <> "" Then
      strSQL = strSQL & " AND DataANO = " & cmb_DRE_AnoPesquisa.Text
  End If
  strSQL = strSQL & " order by DataANO, DataMES "
'  strSQL = strSQL & " order by CodigoDRE "

  Set rsDRE = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  While Not rsDRE.EOF
  
      If rsDRE.Fields("DataMES").Value = 1 Then
          sMesDRE = "Janeiro"
      ElseIf rsDRE.Fields("DataMES").Value = 2 Then
          sMesDRE = "Fevereiro"
      ElseIf rsDRE.Fields("DataMES").Value = 3 Then
          sMesDRE = "Março"
      ElseIf rsDRE.Fields("DataMES").Value = 4 Then
          sMesDRE = "Abril"
      ElseIf rsDRE.Fields("DataMES").Value = 5 Then
          sMesDRE = "Maio"
      ElseIf rsDRE.Fields("DataMES").Value = 6 Then
          sMesDRE = "Junho"
      ElseIf rsDRE.Fields("DataMES").Value = 7 Then
          sMesDRE = "Julho"
      ElseIf rsDRE.Fields("DataMES").Value = 8 Then
          sMesDRE = "Agosto"
      ElseIf rsDRE.Fields("DataMES").Value = 9 Then
          sMesDRE = "Setembro"
      ElseIf rsDRE.Fields("DataMES").Value = 10 Then
          sMesDRE = "Outubro"
      ElseIf rsDRE.Fields("DataMES").Value = 11 Then
          sMesDRE = "Novembro"
      ElseIf rsDRE.Fields("DataMES").Value = 12 Then
          sMesDRE = "Dezembro"
      End If
      
      grid_DRE.AddItem vbTab & rsDRE.Fields("CodigoDRE").Value & vbTab & _
          rsDRE.Fields("Filial").Value & vbTab & _
          rsDRE.Fields("DataANO").Value & vbTab & _
          sMesDRE & vbTab & _
          FormataValorTexto(rsDRE.Fields("ReceitaBruta").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("Devolucoes").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("ImpostoSobreVendas").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("ReceitaOperacionalLiquida").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("CMV").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("LucroBruto").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("DespesasAdministrativas").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("DespesasComerciais").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("DespesasDepreciacao").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("DespesasFinanceiras").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("ReceitasFinanceiras").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("LucroPrejuizoOperacional").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("DespesasNaoOperacionais").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("ReceitasNaoOperacionais").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("LAIR").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("ProvisaoIR").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("ProvisaoCSLL").Value, 2) & vbTab & _
          FormataValorTexto(rsDRE.Fields("LucroLiquido").Value, 2) & vbTab & _
          rsDRE.Fields("dataCriacao").Value & vbTab & _
          rsDRE.Fields("usuario").Value & vbTab & _
          rsDRE.Fields("Obs").Value
    
    rsDRE.MoveNext
  Wend
  rsDRE.Close
  Set rsDRE = Nothing

  Exit Sub
Erro:

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_DRE_salvar_Click()
On Error GoTo Erro
  Dim strSQL As String
  
  cmd_DRE_salvar.Enabled = False
  
  If Trim(txt_DRE_receitaBruta.Text) = "" Or Trim(txt_DRE_receitaBruta.Text) = "0,00" Then
      MsgBox "Realize o cálculo do D.R.E. antes de salvá-lo", vbInformation, "Atenção"
      cmd_DRE_salvar.Enabled = True
      Exit Sub
  End If
  
  If Trim(txt_DRE_LucroLiquido.Text) = "" Or Trim(txt_DRE_LucroLiquido.Text) = "0,00" Then
      MsgBox "Realize o cálculo do D.R.E. antes de salvá-lo", vbInformation, "Atenção"
      cmd_DRE_salvar.Enabled = True
      Exit Sub
  End If
  
  If cmb_ano.Text = "" Then
      DisplayMsg "Escolha o ANO."
      cmd_DRE_salvar.Enabled = True
      cmb_ano.SetFocus
      Exit Sub
  End If
  
  If cmb_mes.Text = "" Then
      DisplayMsg "Escolha o MÊS."
      cmd_DRE_salvar.Enabled = True
      cmb_mes.SetFocus
      Exit Sub
  End If
  
  If cmb_DRE_tipoAnexo.Text = "" Then
      MsgBox "Informe o seu tipo de empresa", vbInformation, "Atenção"
      cmd_DRE_salvar.Enabled = True
      cmb_DRE_tipoAnexo.SetFocus
      Exit Sub
  End If
   
  If Trim(txt_DRE_despesasAdministrativas.Text) = "" Then
      txt_DRE_despesasAdministrativas.Text = "0,00"
  End If
  
  If Trim(txt_DRE_despesasComerciais.Text) = "" Then
      txt_DRE_despesasComerciais.Text = "0,00"
  End If
  
  If Trim(txt_DRE_despesasDepreciacao.Text) = "" Then
      txt_DRE_despesasDepreciacao.Text = "0,00"
  End If
  
  If Trim(txt_DRE_despesasFinanceiras.Text) = "" Then
      txt_DRE_despesasFinanceiras.Text = "0,00"
  End If
  
  If Trim(txt_DRE_receitasFinanceiras.Text) = "" Then
      txt_DRE_receitasFinanceiras.Text = "0,00"
  End If

  If Trim(txt_DRE_despesasNaoOperacionais.Text) = "" Then
      txt_DRE_despesasNaoOperacionais.Text = "0,00"
  End If
  
  If Trim(txt_DRE_receitasNaoOperacionais.Text) = "" Then
      txt_DRE_receitasNaoOperacionais.Text = "0,00"
  End If
  
  If Trim(txt_DRE_provisaoIR.Text) = "" Then
      txt_DRE_provisaoIR.Text = "0,00"
  End If
  
  If Trim(txt_DRE_provisaoCSLL.Text) = "" Then
      txt_DRE_provisaoCSLL.Text = "0,00"
  End If
  
  If Len(Trim(txt_DRE_Observacoes.Text)) = 0 Then
      txt_DRE_Observacoes.Text = " "
  End If
  
  ' -------------------------------------------
  ' Verificar se já existe D.R.E. deste ANO/MES
  Dim rsDRE_pesquisa As Recordset
 
  strSQL = "Select * from DRE_quick "
  strSQL = strSQL & " Where Filial = " & Combo_Filial.Text
  strSQL = strSQL & " AND DataANO = " & cmb_ano.Text
  strSQL = strSQL & " AND DataMES = " & cmb_mes.ListIndex + 1
  
  Set rsDRE_pesquisa = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If Not (rsDRE_pesquisa.EOF And rsDRE_pesquisa.BOF) Then
      rsDRE_pesquisa.Close
      Set rsDRE_pesquisa = Nothing
      MsgBox "Já existe um D.R.E. criado para este ANO/Mês!", vbInformation, "Atenção"
      cmd_DRE_salvar.Enabled = True
      Exit Sub
  End If
  rsDRE_pesquisa.Close
  Set rsDRE_pesquisa = Nothing
  ' -------------------------------------------
  

  strSQL = "Insert Into DRE_Quick "
  strSQL = strSQL & "(Filial, Usuario, DataANO, DataMES, DataCriacao, Obs, ReceitaBruta, Devolucoes, ImpostoSobreVendas, "
  strSQL = strSQL & "ReceitaOperacionalLiquida, CMV, LucroBruto, DespesasAdministrativas, DespesasComerciais, "
  strSQL = strSQL & "DespesasDepreciacao, DespesasFinanceiras, ReceitasFinanceiras, LucroPrejuizoOperacional, "
  strSQL = strSQL & "DespesasNaoOperacionais, ReceitasNaoOperacionais, LAIR, ProvisaoIR, ProvisaoCSLL, LucroLiquido) "
  strSQL = strSQL & " VALUES(" & Combo_Filial.Text & "," & gnUserCode & ","
'  strSQL = strSQL & " #" & Mid(Data_Ini.Text, 4, 3) & Mid(Data_Ini.Text, 1, 3) & Mid(Data_Ini.Text, 7, 4) & "#,"
'  strSQL = strSQL & " #" & Mid(Data_Fim.Text, 4, 3) & Mid(Data_Fim.Text, 1, 3) & Mid(Data_Fim.Text, 7, 4) & "#,"
  strSQL = strSQL & cmb_ano.Text & "," & cmb_mes.ListIndex + 1 & ","
  strSQL = strSQL & " #" & Month(Now) & "/" & Day(Now) & "/" & Year(Now) & "#,"
  strSQL = strSQL & " '" & txt_DRE_Observacoes.Text & "', "
  strSQL = strSQL & Replace(CCur(txt_DRE_receitaBruta.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_Devolucoes.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_ImpostosSobreVendas.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_receitaOperacionalLiquida.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_cmv.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_lucroBruto.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_despesasAdministrativas.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_despesasComerciais.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_despesasDepreciacao.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_despesasFinanceiras.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_receitasFinanceiras.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_lucroPrejuizoOperacional.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_despesasNaoOperacionais.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_receitasNaoOperacionais.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_LAIR.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_provisaoIR.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_provisaoCSLL.Text), ",", ".") & ", "
  strSQL = strSQL & Replace(CCur(txt_DRE_LucroLiquido.Text), ",", ".") & ") "
  
  db.Execute strSQL

  MsgBox "D.R.E. salvo com sucesso!", vbInformation, "Sucesso"
  cmd_DRE_salvar.Enabled = True

  Exit Sub
Erro:
  cmd_DRE_salvar.Enabled = True
  
  MsgBox "Erro ao salvar o D.R.E...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_pesquisar_Click()
On Error GoTo Erro
 
  Dim rsEstoque As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
  Dim sTamanho As String
  Dim sCor As String
 
  If Nome_Filial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    Combo_Filial.SetFocus
    Exit Sub
  End If
 
  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Fim.SetFocus
    Exit Sub
  End If
   
  gridProdutos.Rows = 1
  gridProdutos.Row = 0
  
  If opt1.Value = True Then
      If cmb_numProdutos.Text = "TODOS" Then
        strSQL = "SELECT Sum(E.Vendas)-Sum(E.Devolução) AS SomaDeVendas, Sum(E.[Valor Vendas])-Sum(E.[Valor Devolução]) as SomaPreco,  "
      Else
        strSQL = "SELECT top " & cmb_numProdutos.Text & " Sum(E.Vendas)-Sum(E.Devolução) AS SomaDeVendas, Sum(E.[Valor Vendas])-Sum(E.[Valor Devolução]) as SomaPreco, "
      End If

      gridProdutos.Row = 0
      gridProdutos.TextMatrix(0, 1) = "Vendas"
      gridProdutos.TextMatrix(0, 2) = "R$ faturado"
  Else
      If cmb_numProdutos.Text = "TODOS" Then
        strSQL = "SELECT Sum(E.[Valor Vendas])-Sum(E.[Valor Devolução]) as SomaPreco, Sum(E.Vendas)-Sum(E.Devolução) AS SomaDeVendas, "
      Else
        strSQL = "SELECT top " & cmb_numProdutos.Text & " Sum(E.[Valor Vendas])-Sum(E.[Valor Devolução]) as SomaPreco, Sum(E.Vendas)-Sum(E.Devolução) AS SomaDeVendas, "
      End If

      gridProdutos.Row = 0
      gridProdutos.TextMatrix(0, 1) = "R$ faturado"
      gridProdutos.TextMatrix(0, 2) = "Vendas"
  End If

  strSQL = strSQL & " E.Produto, P.Nome, P.[Estoque Ideal], P.[Estoque Mínimo], EF.[Estoque Atual], E.Tamanho, E.Cor "

  strSQL = strSQL & " From Estoque E, Produtos P, [Estoque Final] EF "
  strSQL = strSQL & " where E.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " E.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
  
  If Trim(txt_parteCodigoProduto.Text) <> "" Then
      strSQL = strSQL & " P.Código like '*" & Trim(txt_parteCodigoProduto.Text) & "*' and "
  Else
      If Trim(txt_parteNomeProduto.Text) <> "" Then
          strSQL = strSQL & " P.Nome like '*" & Trim(txt_parteNomeProduto.Text) & "*' and "
      End If
  End If
  
  strSQL = strSQL & " E.Produto=P.Código and "
 
  If cboClasse.Text <> "" Then
      strSQL = strSQL & " P.Classe=" & cboClasse.Text & " and "
  End If
  
  If cboSubClasse.Text <> "" Then
      strSQL = strSQL & " P.[Sub Classe]=" & cboSubClasse.Text & " and "
  End If
  
  strSQL = strSQL & " E.Filial=" & Combo_Filial.Text & " and "
  strSQL = strSQL & " E.Filial=EF.Filial and "
  strSQL = strSQL & " E.Produto=EF.Produto and "
  strSQL = strSQL & " E.Tamanho=EF.Tamanho and "
  strSQL = strSQL & " E.Cor=EF.Cor "
  
  
  strSQL = strSQL & " GROUP BY E.Produto, P.Nome, P.[Estoque Ideal], P.[Estoque Mínimo], EF.[Estoque Atual], E.Tamanho, E.cor "
  strSQL = strSQL & " ORDER BY 1 DESC "

  Screen.MousePointer = vbHourglass
  
  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  lngContadorRegGrid = 1
  
  If Not (rsEstoque.EOF And rsEstoque.BOF) Then
    rsEstoque.MoveFirst
  End If
  
  While Not rsEstoque.EOF
  
      If rsEstoque.Fields(7).Value <> 0 Then
          sTamanho = AchaTamanho(rsEstoque.Fields(7).Value)
      Else
          sTamanho = ""
      End If

      If rsEstoque.Fields(8).Value <> 0 Then
          sCor = AchaCor(rsEstoque.Fields(8).Value)
      Else
          sCor = ""
      End If

      If opt1.Value = True And rsEstoque.Fields(0).Value <> 0 Then
          gridProdutos.AddItem lngContadorRegGrid & vbTab & rsEstoque.Fields(0).Value & vbTab & _
                          FormataValorTexto(rsEstoque.Fields(1).Value, 2) & vbTab & _
                          rsEstoque.Fields(2).Value & vbTab & _
                          rsEstoque.Fields(3).Value & vbTab & _
                          sTamanho & vbTab & _
                          sCor & vbTab & _
                          rsEstoque.Fields(6).Value & vbTab & _
                          rsEstoque.Fields(5).Value & vbTab & _
                          rsEstoque.Fields(4).Value & vbTab & _
                          rsEstoque.Fields(7).Value & vbTab & _
                          rsEstoque.Fields(8).Value
      Else
          If rsEstoque.Fields(1).Value <> 0 Then
              gridProdutos.AddItem lngContadorRegGrid & vbTab & FormataValorTexto(rsEstoque.Fields(0).Value, 2) & vbTab & _
                          rsEstoque.Fields(1).Value & vbTab & _
                          rsEstoque.Fields(2).Value & vbTab & _
                          rsEstoque.Fields(3).Value & vbTab & _
                          sTamanho & vbTab & _
                          sCor & vbTab & _
                          rsEstoque.Fields(6).Value & vbTab & _
                          rsEstoque.Fields(5).Value & vbTab & _
                          rsEstoque.Fields(4).Value & vbTab & _
                          rsEstoque.Fields(7).Value & vbTab & _
                          rsEstoque.Fields(8).Value
          End If
      End If
     
      rsEstoque.MoveNext
      lngContadorRegGrid = lngContadorRegGrid + 1
      
      If cmb_numProdutos.Text <> "TODOS" Then
          If cmb_numProdutos.Text = lngContadorRegGrid Then
              rsEstoque.MoveLast
          End If
      Else
      
      End If
  Wend
  rsEstoque.Close
  Set rsEstoque = Nothing
  
  SSTab1.Tab = 0
  
  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

'Formata o valor de acordo com o número de casas decimais e substitui separador decimal por ponto
Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
  
  If lngCasasDecimais = 2 Then
      If Len(FormataValorTexto) = 7 Then  ' 9999.99     = 9.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 6)
      ElseIf Len(FormataValorTexto) = 8 Then ' 99999.99    = 99.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 6)
      ElseIf Len(FormataValorTexto) = 9 Then ' 999999.99   = 999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 6)
      ElseIf Len(FormataValorTexto) = 10 Then ' 9999999.99   = 9.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 3) + "." + Mid(FormataValorTexto, 5, 6)
      ElseIf Len(FormataValorTexto) = 11 Then ' 99999999.99   = 99.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 3) + "." + Mid(FormataValorTexto, 6, 6)
      ElseIf Len(FormataValorTexto) = 12 Then ' 999999999.99   = 999.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 3) + "." + Mid(FormataValorTexto, 7, 6)
      End If
  End If
  
End Function

Private Sub Combo_Filial_LostFocus()
  Dim i As Integer
  Nome_Filial.Caption = ""
  
  If Combo_Filial.Text <> "" Then
      For i = 0 To iConta
        If Combo_Filial.Text = arrayFiliais(i, 0) Then
          Nome_Filial.Caption = arrayFiliais(i, 1)
          Exit For
        End If
      Next
  End If
End Sub

Private Sub Command1_Click()
    frmRelPagar4.Show
End Sub

Private Sub Command2_Click()
    frmRelPagas1.Show
End Sub

Private Sub Data_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
  End Select
End Sub

Private Sub Data_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
  End Select
End Sub

Private Sub Form_Load()
On Error GoTo Erro
  Dim rsParametros As Recordset
  Dim rsCentroCusto As Recordset
  Dim rsTamanho As Recordset
  Dim rsCor As Recordset
  Dim iContador As Integer
  
  Set rsCentroCusto = db.OpenRecordset("select Código, Nome from [Centros de Custo] where Ativo = -1 Order by 2", dbOpenDynaset)
  If Not (rsCentroCusto.EOF And rsCentroCusto.BOF) Then
      rsCentroCusto.MoveLast
      rsCentroCusto.MoveFirst
      
      ReDim arrayCentroCustos(rsCentroCusto.RecordCount, 2)
      contador_arrayCentroCustos = rsCentroCusto.RecordCount
      iContador = 0
      While Not rsCentroCusto.EOF
          lst_centroCusto.AddItem Trim(rsCentroCusto.Fields(1).Value), iContador
          arrayCentroCustos(iContador, 0) = rsCentroCusto.Fields(0).Value
          arrayCentroCustos(iContador, 1) = rsCentroCusto.Fields(1).Value
          iContador = iContador + 1
          rsCentroCusto.MoveNext
      Wend
  End If
  rsCentroCusto.Close
  Set rsCentroCusto = Nothing
  
  iContador = 0
  Set rsTamanho = db.OpenRecordset("select Código, Nome from Tamanhos ", dbOpenDynaset)
  If Not (rsTamanho.EOF And rsTamanho.BOF) Then
      rsTamanho.MoveLast
      rsTamanho.MoveFirst
      
      ReDim arrayTamanhos(rsTamanho.RecordCount, 2)
      contador_arrayTamanhos = rsTamanho.RecordCount
      While Not rsTamanho.EOF
          arrayTamanhos(iContador, 0) = rsTamanho.Fields(0).Value
          arrayTamanhos(iContador, 1) = rsTamanho.Fields(1).Value
          iContador = iContador + 1
          rsTamanho.MoveNext
      Wend
  End If
  rsTamanho.Close
  Set rsTamanho = Nothing
  
  iContador = 0
  Set rsCor = db.OpenRecordset("select Código, Nome from Cores ", dbOpenDynaset)
  If Not (rsCor.EOF And rsCor.BOF) Then
      rsCor.MoveLast
      rsCor.MoveFirst
      
      ReDim arrayCores(rsCor.RecordCount, 2)
      contador_arrayCores = rsCor.RecordCount
      While Not rsCor.EOF
          arrayCores(iContador, 0) = rsCor.Fields(0).Value
          arrayCores(iContador, 1) = rsCor.Fields(1).Value
          iContador = iContador + 1
          rsCor.MoveNext
      Wend
  End If
  rsCor.Close
  Set rsCor = Nothing
  
  
  Set rsParametros = db.OpenRecordset("select Filial, Nome, IRRF, CSLL, CodigoRegimeTributario, PIS, COFINS from [Parâmetros Filial]", dbOpenDynaset)
  
  iConta = 0
  While Not rsParametros.EOF
      arrayFiliais(iConta, 0) = rsParametros.Fields(0).Value
      arrayFiliais(iConta, 1) = rsParametros.Fields(1).Value
      arrayFiliais(iConta, 2) = rsParametros.Fields(2).Value
      arrayFiliais(iConta, 3) = rsParametros.Fields(3).Value
      arrayFiliais(iConta, 4) = rsParametros.Fields(4).Value
      arrayFiliais(iConta, 5) = rsParametros.Fields(5).Value
      arrayFiliais(iConta, 6) = rsParametros.Fields(6).Value
      Combo_Filial.AddItem rsParametros.Fields(0).Value, iConta
      iConta = iConta + 1
      rsParametros.MoveNext
  Wend
  rsParametros.Close
  Set rsParametros = Nothing

  datClasses.DatabaseName = gsQuickDBFileName
  datSubClasses.DatabaseName = gsQuickDBFileName
  
  gridProdutos.ColWidth(0) = 450
  gridProdutos.ColWidth(1) = 1300
  gridProdutos.ColWidth(2) = 1300
  gridProdutos.ColWidth(3) = 2300
  gridProdutos.ColWidth(4) = 4900
  gridProdutos.ColWidth(5) = 900
  gridProdutos.ColWidth(6) = 1800
  gridProdutos.ColWidth(7) = 800
  gridProdutos.ColWidth(8) = 800
  gridProdutos.ColWidth(9) = 800
  gridProdutos.ColWidth(10) = 0
  gridProdutos.ColWidth(11) = 0
  
  gridProdutos.Row = 0
  gridProdutos.TextMatrix(0, 1) = "Vendas"
  gridProdutos.TextMatrix(0, 2) = "R$ faturado"
  gridProdutos.TextMatrix(0, 3) = "Código"
  gridProdutos.TextMatrix(0, 4) = "Nome"
  gridProdutos.TextMatrix(0, 5) = "Tamanho"
  gridProdutos.TextMatrix(0, 6) = "Cor"
  gridProdutos.TextMatrix(0, 7) = "Estoque"
  gridProdutos.TextMatrix(0, 8) = "Mínimo"
  gridProdutos.TextMatrix(0, 9) = "Ideal"
  
  gridEntradas.ColWidth(0) = 1200
  gridEntradas.ColWidth(1) = 1200
  gridEntradas.ColWidth(2) = 1300
  gridEntradas.ColWidth(3) = 1200
  gridEntradas.ColWidth(4) = 4500
  gridEntradas.ColWidth(5) = 1200
  gridEntradas.ColWidth(6) = 4500
  
  gridEntradas.Row = 0
  gridEntradas.TextMatrix(0, 0) = "Data entrada"
  gridEntradas.TextMatrix(0, 1) = "Qtde"
  gridEntradas.TextMatrix(0, 2) = "Valor Unidade"
  gridEntradas.TextMatrix(0, 3) = "Forn/Cli"
  gridEntradas.TextMatrix(0, 4) = "Nome"
  gridEntradas.TextMatrix(0, 5) = "Oper"
  gridEntradas.TextMatrix(0, 6) = "Nome"
  
  gridSaidas.ColWidth(0) = 0
  gridSaidas.ColWidth(1) = 1900
  gridSaidas.ColWidth(2) = 1200
  gridSaidas.ColWidth(3) = 1300
  gridSaidas.ColWidth(4) = 1200
  gridSaidas.ColWidth(5) = 4000
  gridSaidas.ColWidth(6) = 600
  gridSaidas.ColWidth(7) = 4100
  gridSaidas.ColWidth(8) = 1000
  
  gridSaidas.Row = 0
  gridSaidas.TextMatrix(0, 0) = ""
  gridSaidas.TextMatrix(0, 1) = "Data saída"
  gridSaidas.TextMatrix(0, 2) = "Qtde"
  gridSaidas.TextMatrix(0, 3) = "Valor Unidade"
  gridSaidas.TextMatrix(0, 4) = "Forn/Cli"
  gridSaidas.TextMatrix(0, 5) = "Nome"
  gridSaidas.TextMatrix(0, 6) = "Oper"
  gridSaidas.TextMatrix(0, 7) = "Nome"
  gridSaidas.TextMatrix(0, 8) = "Sequência"
  
  
  gridProdutos_pareto.ColWidth(0) = 450
  gridProdutos_pareto.ColWidth(1) = 1300
  gridProdutos_pareto.ColWidth(2) = 790
  gridProdutos_pareto.ColWidth(3) = 2170
  gridProdutos_pareto.ColWidth(4) = 4750
  gridProdutos_pareto.ColWidth(5) = 800
  gridProdutos_pareto.ColWidth(6) = 1790
  gridProdutos_pareto.ColWidth(7) = 520
  gridProdutos_pareto.ColWidth(8) = 520
  gridProdutos_pareto.ColWidth(9) = 520
  
  gridProdutos_pareto.Row = 0
  gridProdutos_pareto.TextMatrix(0, 1) = "R$ faturado"
  gridProdutos_pareto.TextMatrix(0, 2) = "Vendas"
  gridProdutos_pareto.TextMatrix(0, 3) = "Código"
  gridProdutos_pareto.TextMatrix(0, 4) = "Nome"
  gridProdutos_pareto.TextMatrix(0, 5) = "Tamanho"
  gridProdutos_pareto.TextMatrix(0, 6) = "Cor"
  gridProdutos_pareto.TextMatrix(0, 7) = "% Ind"
  gridProdutos_pareto.TextMatrix(0, 8) = "% Ac"
  gridProdutos_pareto.TextMatrix(0, 9) = "Clas"
  
  
  ' Grade D.R.E.
  grid_DRE.ColWidth(0) = 0
  grid_DRE.ColWidth(1) = 600
  grid_DRE.ColWidth(2) = 400
  grid_DRE.ColWidth(3) = 1000
  grid_DRE.ColWidth(4) = 1000
  grid_DRE.ColWidth(5) = 1300
  grid_DRE.ColWidth(6) = 1300
  grid_DRE.ColWidth(7) = 1500
  grid_DRE.ColWidth(8) = 1500
  grid_DRE.ColWidth(9) = 1300
  grid_DRE.ColWidth(10) = 1300
  grid_DRE.ColWidth(11) = 1900
  grid_DRE.ColWidth(12) = 1600
  grid_DRE.ColWidth(13) = 1700
  grid_DRE.ColWidth(14) = 1700
  grid_DRE.ColWidth(15) = 1700
  grid_DRE.ColWidth(16) = 2000
  grid_DRE.ColWidth(17) = 2100
  grid_DRE.ColWidth(18) = 2000
  grid_DRE.ColWidth(19) = 1400
  grid_DRE.ColWidth(20) = 1400
  grid_DRE.ColWidth(21) = 1400
  grid_DRE.ColWidth(22) = 1400
  grid_DRE.ColWidth(23) = 1000
  grid_DRE.ColWidth(24) = 650
  grid_DRE.ColWidth(25) = 16000
  
  grid_DRE.Row = 0
  grid_DRE.TextMatrix(0, 1) = "ID"
  grid_DRE.TextMatrix(0, 2) = "Filial"
  grid_DRE.TextMatrix(0, 3) = "Ano"
  grid_DRE.TextMatrix(0, 4) = "Mês"
  grid_DRE.TextMatrix(0, 5) = "Receita Bruta"
  grid_DRE.TextMatrix(0, 6) = "Devoluções"
  grid_DRE.TextMatrix(0, 7) = "Imposto s/ Vendas"
  grid_DRE.TextMatrix(0, 8) = "Receita Op.Líquida"
  grid_DRE.TextMatrix(0, 9) = "CMV"
  grid_DRE.TextMatrix(0, 10) = "Lucro Bruto"
  grid_DRE.TextMatrix(0, 11) = "Despesas Administrativas"
  grid_DRE.TextMatrix(0, 12) = "Despesas Comerciais"
  grid_DRE.TextMatrix(0, 13) = "Despesas Depreciação"
  grid_DRE.TextMatrix(0, 14) = "Despesas Financeiras"
  grid_DRE.TextMatrix(0, 15) = "Receitas Financeiras"
  grid_DRE.TextMatrix(0, 16) = "Lucro/Prejuizo Operacional"
  grid_DRE.TextMatrix(0, 17) = "Despesas Não Operacionais"
  grid_DRE.TextMatrix(0, 18) = "Receitas Não Operacionais"
  grid_DRE.TextMatrix(0, 19) = "LAIR"
  grid_DRE.TextMatrix(0, 20) = "Provisao IR"
  grid_DRE.TextMatrix(0, 21) = "Provisao CSLL"
  grid_DRE.TextMatrix(0, 22) = "Lucro Líquido"
  grid_DRE.TextMatrix(0, 23) = "Data Criação"
  grid_DRE.TextMatrix(0, 24) = "Usuário"
  grid_DRE.TextMatrix(0, 25) = "Observações"
  '
  
  Data_Fim.Text = Format(Now, "dd/mm/yyyy")
  Data_Ini.Text = Format(Now, "dd/mm/yyyy")
  
  cmb_numProdutos.ListIndex = 5
  SSTab1.Tab = 0
  
  Combo_Filial.Text = gnCodFilial
  Combo_Filial_LostFocus
  
  Exit Sub
Erro:
  MsgBox "Erro na abertura da tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub gridProdutos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  gridProdutos.Redraw = False
End Sub

Private Sub gridProdutos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  gridProdutos.RowSel = gridProdutos.Row
  gridProdutos.Redraw = True
End Sub
