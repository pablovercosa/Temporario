VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelSaidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " SAÍDAS - Movimentação de Saídas, Entradas x NFe/NFCe"
   ClientHeight    =   8445
   ClientLeft      =   1650
   ClientTop       =   2520
   ClientWidth     =   14310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1570
   Icon            =   "RelSaidas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8445
   ScaleWidth      =   14310
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
      Left            =   -1980
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0' ORDER BY Nome"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmd_imprimir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir Grade"
      Height          =   375
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   7830
      Width           =   1935
   End
   Begin VB.TextBox txt_totalRegistros 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   375
      Left            =   10770
      TabIndex        =   58
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   7800
      Width           =   1185
   End
   Begin VB.Frame Frame9 
      Caption         =   "Tipo Nota Fiscal"
      Height          =   945
      Left            =   11460
      TabIndex        =   53
      Top             =   2370
      Visible         =   0   'False
      Width           =   2805
      Begin VB.OptionButton opt_tpNFCe 
         Appearance      =   0  'Flat
         Caption         =   "NFCe"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1230
         TabIndex        =   55
         Top             =   405
         Width           =   1035
      End
      Begin VB.OptionButton opt_tpNFe 
         Appearance      =   0  'Flat
         Caption         =   "NFe"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   330
         TabIndex        =   54
         Top             =   420
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.TextBox txt_totalNFeCanc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   375
      Left            =   6030
      TabIndex        =   51
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   7785
      Width           =   1815
   End
   Begin VB.TextBox txt_totalNFe 
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
      Height          =   375
      Left            =   2460
      TabIndex        =   49
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   7785
      Width           =   1815
   End
   Begin VB.CommandButton cmdPesquisar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   495
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3510
      Visible         =   0   'False
      Width           =   14235
   End
   Begin VB.Frame Frame8 
      Caption         =   "Tipo do Relatório"
      Height          =   525
      Left            =   60
      TabIndex        =   44
      Top             =   30
      Width           =   14205
      Begin VB.OptionButton optTipoRel2 
         Appearance      =   0  'Flat
         Caption         =   "Saídas x Nota Fiscal"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8130
         TabIndex        =   46
         Top             =   150
         Width           =   1905
      End
      Begin VB.OptionButton optTipoRel1 
         Appearance      =   0  'Flat
         Caption         =   "Relatório detalhado"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3540
         TabIndex        =   45
         Top             =   195
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.CheckBox chkNFCe 
      Appearance      =   0  'Flat
      Caption         =   "Somente NFCe"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7950
      TabIndex        =   43
      Top             =   2085
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CheckBox chkNota 
      Appearance      =   0  'Flat
      Caption         =   "Somente NFe"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7950
      TabIndex        =   42
      Top             =   1785
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CheckBox chkCupom 
      Caption         =   "Somente com Cupom Fiscal"
      Height          =   315
      Left            =   8730
      TabIndex        =   41
      Top             =   1770
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame7 
      Caption         =   "Efetivadas ou Não"
      Height          =   615
      Left            =   9660
      TabIndex        =   37
      Top             =   1710
      Width           =   4605
      Begin VB.OptionButton O_NaoEfetivadas 
         Appearance      =   0  'Flat
         Caption         =   "Não Efetivadas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   40
         Top             =   270
         Width           =   1455
      End
      Begin VB.OptionButton O_Efetivadas 
         Appearance      =   0  'Flat
         Caption         =   "Efetivadas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1470
         TabIndex        =   39
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton O_Todos_efetivadas 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Data dtaCliente 
      Caption         =   "dtaCliente"
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
      Left            =   -1500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cli_For"
      Top             =   7710
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data dtaVendedor 
      Caption         =   "dtaVendedor"
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
      Left            =   -1500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   7410
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.TextBox txtSequencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   870
      TabIndex        =   4
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   "Período"
      Height          =   615
      Left            =   2190
      TabIndex        =   29
      Top             =   1725
      Width           =   5595
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
         Height          =   420
         Left            =   2280
         Picture         =   "RelSaidas.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   150
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
         Height          =   420
         Left            =   5025
         Picture         =   "RelSaidas.frx":4F23C
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   150
         Width           =   465
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   990
         TabIndex        =   5
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   210
         Width           =   1260
         _ExtentX        =   2223
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3750
         TabIndex        =   6
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   210
         Width           =   1260
         _ExtentX        =   2223
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2940
         TabIndex        =   31
         Top             =   255
         Width           =   720
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   30
         Top             =   255
         Width           =   795
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Notas"
      Height          =   945
      Left            =   2400
      TabIndex        =   28
      Top             =   2370
      Width           =   4785
      Begin VB.ComboBox cbo_ordenar 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "RelSaidas.frx":4FB1E
         Left            =   3810
         List            =   "RelSaidas.frx":4FB40
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   510
         Width           =   2805
      End
      Begin VB.CheckBox chk_somenteNF_parcelada 
         Appearance      =   0  'Flat
         Caption         =   "Somente Saídas c/ Recebimento Parcelado"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1920
         TabIndex        =   62
         Top             =   180
         Width           =   2775
      End
      Begin VB.OptionButton O_Nota_Canc 
         Appearance      =   0  'Flat
         Caption         =   "Canceladas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton O_Nota_N_Canc 
         Appearance      =   0  'Flat
         Caption         =   "Não Canceladas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   420
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar visualização por"
         Height          =   195
         Left            =   1950
         TabIndex        =   67
         Top             =   570
         Width           =   1785
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Orçamentos"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   11460
      TabIndex        =   27
      Top             =   2370
      Visible         =   0   'False
      Width           =   2805
      Begin VB.OptionButton O_Sem_Nota 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Cancelados"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   510
         Width           =   1215
      End
      Begin VB.OptionButton O_Com_Nota 
         Caption         =   "Impresso"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   510
         Width           =   1155
      End
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Geral"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Data Data2 
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
      Left            =   -1500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Op_Saída"
      Top             =   7110
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opção"
      Height          =   945
      Left            =   60
      TabIndex        =   26
      Top             =   2370
      Width           =   2295
      Begin VB.OptionButton com_produtos 
         Appearance      =   0  'Flat
         Caption         =   "Imprimir produtos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   1845
      End
      Begin VB.OptionButton sem_produtos 
         Appearance      =   0  'Flat
         Caption         =   "Não imprimir produtos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   945
      Left            =   8910
      TabIndex        =   25
      Top             =   2370
      Width           =   2505
      Begin VB.OptionButton O_Completo 
         Appearance      =   0  'Flat
         Caption         =   "Completo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   390
         TabIndex        =   8
         Top             =   540
         Width           =   1125
      End
      Begin VB.OptionButton O_Resumido 
         Appearance      =   0  'Flat
         Caption         =   "Resumido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   390
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   945
      Left            =   7230
      TabIndex        =   24
      Top             =   2370
      Width           =   1635
      Begin VB.OptionButton O_Impressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   540
         Width           =   1245
      End
      Begin VB.OptionButton O_vídeo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   270
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gerar Relatório"
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   14235
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
      Left            =   -1455
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   6810
      Visible         =   0   'False
      Width           =   1770
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   13980
      Top             =   1530
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Operação 
      Bindings        =   "RelSaidas.frx":4FBB9
      DataSource      =   "Data2"
      Height          =   345
      Left            =   8070
      TabIndex        =   3
      Top             =   960
      Width           =   1050
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
      BevelColorFrame =   0
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      BackColorEven   =   -2147483633
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9208
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2037
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelSaidas.frx":4FBCD
      DataSource      =   "Data1"
      Height          =   345
      Left            =   780
      TabIndex        =   0
      Top             =   585
      Width           =   1050
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
      BevelColorFrame =   0
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      BackColorEven   =   -2147483633
      BackColorOdd    =   16777152
      Columns(0).Width=   3200
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "RelSaidas.frx":4FBE1
      DataSource      =   "dtaVendedor"
      Height          =   345
      Left            =   8070
      TabIndex        =   1
      Top             =   585
      Width           =   1050
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
      BevelColorFrame =   0
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      BackColorEven   =   -2147483633
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9208
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2037
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "RelSaidas.frx":4FBFB
      DataSource      =   "dtaCliente"
      Height          =   345
      Left            =   780
      TabIndex        =   2
      Top             =   960
      Width           =   1050
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
      BevelColorFrame =   0
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      BackColorEven   =   -2147483633
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9208
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2037
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBGrid grdMovimento 
      Height          =   3885
      Left            =   30
      TabIndex        =   48
      Top             =   3885
      Width           =   14295
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   12
      CheckBox3D      =   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      MaxSelectedRows =   5
      ForeColorEven   =   0
      BackColorEven   =   15724527
      BackColorOdd    =   12648384
      RowHeight       =   423
      ExtraHeight     =   238
      Columns.Count   =   12
      Columns(0).Width=   1879
      Columns(0).Caption=   "Data"
      Columns(0).Name =   "Data"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1852
      Columns(1).Caption=   "Vencimento"
      Columns(1).Name =   "DtVencimento"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1746
      Columns(2).Caption=   "Sequência"
      Columns(2).Name =   "Sequencia"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2117
      Columns(3).Caption=   "Código"
      Columns(3).Name =   "CodigoClienteFornecedor"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   6112
      Columns(4).Caption=   "Nome Cliente/Fornecedor"
      Columns(4).Name =   "Nome Cliente/Fornecedor"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   609
      Columns(5).Caption=   "Sr"
      Columns(5).Name =   "Serie"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1773
      Columns(6).Caption=   "Nota Fiscal"
      Columns(6).Name =   "NotaFiscal"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2037
      Columns(7).Caption=   "Total"
      Columns(7).Name =   "Total"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3201
      Columns(8).Caption=   "Status"
      Columns(8).Name =   "Status"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   4895
      Columns(9).Caption=   "ChaveAcesso"
      Columns(9).Name =   "ChaveAcesso"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Caption=   "ProtocoloAutorização"
      Columns(10).Name=   "ProtocoloAutorização"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3545
      Columns(11).Caption=   "ProtocoloCancelamento"
      Columns(11).Name=   "ProtocoloCancelamento"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      _ExtentX        =   25215
      _ExtentY        =   6853
      _StockProps     =   79
      ForeColor       =   0
      BackColor       =   -2147483648
      Enabled         =   0   'False
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
   Begin SSDataWidgets_B.SSDBCombo cboProduto 
      Bindings        =   "RelSaidas.frx":4FC14
      Height          =   315
      Left            =   780
      TabIndex        =   64
      Top             =   1365
      Width           =   2175
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
      BevelColorFrame =   0
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      ForeColorEven   =   0
      BackColorEven   =   -2147483633
      BackColorOdd    =   16777152
      Columns(0).Width=   3200
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      DataFieldToDisplay=   "Nome"
   End
   Begin VB.Label lbl_NomeProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3000
      TabIndex        =   66
      Top             =   1350
      Width           =   6120
   End
   Begin VB.Label lblProduto 
      Caption         =   "Produto"
      Height          =   255
      Left            =   60
      TabIndex        =   65
      Top             =   1395
      Width           =   645
   End
   Begin VB.Label lbl_obs_visaoSaidaNFCe 
      Caption         =   "Obs: NFCe com Status igual a 'Verificar Status' não  soma nestes totais R$"
      Height          =   255
      Left            =   30
      TabIndex        =   60
      Top             =   8190
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Nº total registros"
      Height          =   195
      Left            =   9420
      TabIndex        =   59
      Top             =   7875
      Width           =   1230
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Total R$ - Cancelada"
      Height          =   195
      Left            =   4410
      TabIndex        =   52
      Top             =   7860
      Width           =   1500
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total R$ Solicitado autorização"
      Height          =   195
      Left            =   30
      TabIndex        =   50
      Top             =   7860
      Width           =   2205
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1845
      TabIndex        =   36
      Top             =   960
      Width           =   5130
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Left            =   60
      TabIndex        =   35
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor"
      Height          =   195
      Left            =   7230
      TabIndex        =   34
      Top             =   645
      Width           =   690
   End
   Begin VB.Label Nome_Vendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9135
      TabIndex        =   33
      Top             =   585
      Width           =   5130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sequência"
      Height          =   195
      Left            =   60
      TabIndex        =   32
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Nome_Operação 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9135
      TabIndex        =   23
      Top             =   960
      Width           =   5130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Operação"
      Height          =   195
      Left            =   7230
      TabIndex        =   22
      Top             =   1020
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      Height          =   195
      Left            =   60
      TabIndex        =   21
      Top             =   645
      Width           =   300
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1845
      TabIndex        =   20
      Top             =   585
      Width           =   5130
   End
End
Attribute VB_Name = "frmRelSaidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsParametros As Recordset
Private rsOp_Saída As Recordset
Private rsSaidas As Recordset

'16/10/2007 - Anderson
'Implementação do filtro vendedor
'Solicitado por Agrotama
Private rsVendedor As Recordset

'08/11/2007 - Celso
'Implementação do filtro cliente
'Solicitado por Litoral Materiais de Construção.
Private rsCliente As Recordset

'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long



Private Sub B_Imprime_Click()
  Dim Erro As Integer
  Dim Str1 As String, Str_Data1 As String, Str_Data2 As String
  Dim Str_Rel As String
  Dim Aux_Data As Date
  Dim Aux_Seq As Long
  Dim Aux_Str As String
  
  '04/09/2003 - mpdea
  'Total de descontos
  Dim dblTotalDesconto As Double
  
  Call StatusMsg("")
  
  Rem Verifica empresa
  If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
    DisplayMsg "Escolha a empresa."
    Combo_Filial.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(Combo_Filial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If

  '08/03/2007 - Anderson
  'Verifica Sequência
  If txtSequencia.Text = "" Then
    txtSequencia.Text = 0
  End If
  
  '16/10/2007 - Anderson
  'Verifica Vendedor
  If Combo_Vendedor.Text = "" Then
    Combo_Vendedor.Text = 0
  End If
  
  '08/11/2007 - Celso
  'Verifica Cliente
  If Combo_Cliente.Text = "" Then
    Combo_Cliente.Text = 0
  End If
  
  
  Rem Verifica Data
  Erro = False
  If IsNull(Data_Ini.Text) Then Erro = True
  If Not Erro Then If Not IsDate(Data_Ini.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data incorreta, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
 
  Rem Verifica Data Final
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

  
  '02/09/2003 - mpdea
  'Status
  Call StatusMsg("Aguarde...")
  
  
  Str_Rel = "DELETE * FROM Saídas"
  dbTemp.Execute Str_Rel
  
  
  rsSaidas.Index = "Data"
  Aux_Data = CDate(Data_Ini.Text)
  Aux_Seq = 0
 
 
Lp1:
  rsSaidas.Seek ">", Val(Combo_Filial.Text), Aux_Data, Aux_Seq
  If rsSaidas.NoMatch Then GoTo Imprime
  If rsSaidas("Filial") <> Val(Combo_Filial.Text) Then GoTo Imprime
  If rsSaidas("Data") > CDate(Data_Fim.Text) Then GoTo Imprime
  
  
  Aux_Data = rsSaidas("Data")
  Aux_Seq = rsSaidas("Sequência")
  
  '08/03/2007 - Anderson
  'Filtra por sequência
  If txtSequencia.Text = CStr(Aux_Seq) Or txtSequencia = "0" Then
    
    If chkNFCe.Value = 1 And rsSaidas("NFCe") = 0 Then GoTo Lp1
    
    If chkNFCe.Value = 1 And IsNull(rsSaidas("NFCe")) Then GoTo Lp1
  
    If chkNota.Value = 1 And rsSaidas("Nota Impressa") = 0 Then GoTo Lp1
    
    If chkCupom.Value = 1 And rsSaidas("Observações") <> "Venda Fiscal" Then GoTo Lp1
    
    If O_Efetivadas.Value = True And rsSaidas("Movimentação Desfeita") = False Then GoTo Lp1
  
    If rsSaidas("Nota Cancelada") = True And O_Nota_N_Canc.Value = True Then GoTo Lp1
    If rsSaidas("Nota Cancelada") = False And O_Nota_Canc.Value = True Then GoTo Lp1
    
    If Nome_Operação.Caption <> "" Then
      If rsSaidas("Operação") <> Val(Combo_Operação.Text) Then GoTo Lp1
    End If
    
    '16/10/2007 - Anderson
    'Implementação do filtro
    If Nome_Vendedor.Caption <> "" Then
      If rsSaidas("Digitador") <> Val(Combo_Vendedor.Text) Then GoTo Lp1
    End If
    
    '08/11/2007 - Celso
    'Implementação do filtro cliente
    If Nome_Cliente.Caption <> "" Then
      If rsSaidas("Cliente") <> Val(Combo_Cliente.Text) Then GoTo Lp1
    End If
    
    'colocar código para trazer somente cupons fiscais
    
     '02/09/2003 - mpdea
     'Comentado devido a perda de performance
  '  Call StatusMsg("Aguarde, verificando movimentação " & str(Aux_Seq))
  
    Grava_Temp_Saídas Val(Combo_Filial.Text), Aux_Seq, cboProduto.Text
    
    '04/09/2003 - mpdea
    'Total de descontos
    dblTotalDesconto = dblTotalDesconto + _
                       CDbl(rsSaidas.Fields("Desconto").Value) + _
                       CDbl("0" & rsSaidas.Fields("DescontoSubTotal").Value)
  End If
  
  GoTo Lp1

Imprime:
  Rem  Nome do BD
   With Rel1
     .DataFiles(0) = gsTempDBFileName
     .DataFiles(1) = gsQuickDBFileName
   End With
  
  Rem Saída
  If O_Vídeo = True Then Rel1.Destination = 0
  If O_Vídeo = False Then Rel1.Destination = 1
  
  Rem Nome do arquivo .rpt
  If O_Resumido.Value = True Then
    Str1 = gsReportPath & "Saida2.RPT"
  Else
    Str1 = gsReportPath & "Saida1.RPT"
  End If

  Rel1.ReportFileName = Str1
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel1
  
  Rem Seleção
  Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
  Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
  
  
  Str_Rel = "{Saídas.Data} >="
  Str_Rel = Str_Rel + Str_Data1
  Str_Rel = Str_Rel + " And {Saídas.Data} <=" + Str_Data2
  
  
  If Nome_Operação.Caption <> "" Then
    Str_Rel = Str_Rel + " And {Saídas.Cód Operação} = " + Combo_Operação.Text
  End If
  
  If O_Sem_Nota.Value = True Then
    Str_Rel = Str_Rel + " And {Saídas.Nota} = 0"
  End If
  
  If O_Com_Nota.Value = True Then
    Str_Rel = Str_Rel + " And {Saídas.Nota} <> 0"
  End If
  
  If O_Nota_N_Canc.Value = True Then
    Str_Rel = Str_Rel + " And {Saídas.Nota Cancelada} = False"
  End If
 
  If O_Nota_Canc.Value = True Then
    Str_Rel = Str_Rel + " And {Saídas.Nota Cancelada} = True"
  End If
  
  If O_Efetivadas.Value = True Then
    Str_Rel = Str_Rel + " And {Saídas.Efetivada} = True"
  End If
  
  If O_NaoEfetivadas.Value = True Then
    Str_Rel = Str_Rel + " And {Saídas.Efetivada} = False"
  End If
  
  
  Rel1.SelectionFormula = Str_Rel
  
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  
  Rel1.Formulas(0) = Str_Rel
  
  Str_Rel = "filial = '"
  Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
  Rel1.Formulas(1) = Str_Rel
  
  Rem data inicial
  Str_Rel = "data_ini = '"
  Str_Rel = Str_Rel + Data_Ini.Text + "'"
  Rel1.Formulas(2) = Str_Rel
  
  Rem data final
  Str_Rel = "data_fim = '"
  Str_Rel = Str_Rel + Data_Fim.Text + "'"
  Rel1.Formulas(3) = Str_Rel
  
  If sem_produtos.Value = True Then Str_Rel = "emite_produtos = 'NAO'"
  If com_produtos.Value = True Then Str_Rel = "emite_produtos = 'SIM'"
  
  Rel1.Formulas(4) = Str_Rel
 
 
  '04/09/2003 - mpdea
  'Total de descontos
  Rel1.Formulas(5) = "TotalGeralDescontosComSubTotal = " & Replace(dblTotalDesconto, gsCurrencyDecimal, ".")
 

  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass

  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  
 
  Rel1.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault

End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
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

  Dim bm As Variant
  Dim nRow As Long

  Dim sLinha As String
  Dim sValor As String
  Dim sNFe As String
  Dim sCodCli As String
  Dim sNomeCli As String
  Dim sStatus As String
  
  Printer.Font = "LUCIDA CONSOLE"
  
  Printer.Print ""
  sLinha = "                                                                   Quick Store 10 - Soluções Comerciais inteligentes"
  
  Printer.Print sLinha
  sLinha = "   " & gsNomeFilial
  Printer.Print sLinha
  
  Printer.Print ""
  
  sLinha = "   Relatório de Saídas por Período      De: " & Data_Ini & " até " & Data_Fim
  Printer.Print sLinha

  Printer.Print ""

  sLinha = "   Data Mov.   Vencimento  Valor              NFe        Status                 Cliente"
  Printer.Print sLinha

  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""
  
  With grdMovimento
      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)

          ' ************************** ATENÇÃO ***********************************
          ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
          ' De preferência com o mesmo nome da impressora !!!

          sValor = .Columns("total").CellValue(bm)
          If Len(sValor) < 14 Then
            For i = Len(sValor) To 13
                sValor = " " & sValor
            Next
          End If

          sNFe = .Columns("NotaFiscal").CellValue(bm)
          If Len(sNFe) < 9 Then
            For i = Len(sNFe) To 8
                sNFe = " " & sNFe
            Next
          End If

          sCodCli = .Columns("Código").CellValue(bm)
          If Len(sCodCli) < 12 Then
            For i = Len(sCodCli) To 11
                sCodCli = " " & sCodCli
            Next
          End If
          
          sNomeCli = .Columns("Nome Cliente/Fornecedor").CellValue(bm)
          If Len(sNomeCli) > 27 Then
              sNomeCli = Mid(sNomeCli, 1, 27)
          End If
          
          sStatus = .Columns("Status").CellValue(bm) & "  "
          If Len(sStatus) > 21 Then
              sStatus = Mid(sStatus, 1, 21)
          Else
            For i = Len(sStatus) To 20
                sStatus = sStatus & " "
            Next
          End If

          sLinha = .Columns("Data").CellValue(bm) & "  "
          If .Columns("DtVencimento").CellValue(bm) <> "" Then
            sLinha = sLinha & .Columns("DtVencimento").CellValue(bm) & "  R$"
          Else
            sLinha = sLinha & "            R$"
          End If
          sLinha = sLinha & sValor
          sLinha = sLinha & "  " & sNFe
          sLinha = sLinha & "  " & sStatus
          sLinha = sLinha & "  " & sCodCli
          sLinha = sLinha & "  " & sNomeCli
          
'          If .Columns("Status").CellValue(bm) = "Solicitada/Autorizada" Then
              Printer.Print "   " & sLinha
              Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
'          End If
      Next nRow
  End With
  
  Printer.Print ""

  sValor = txt_totalNFe.Text
  If Len(sValor) < 14 Then
    For i = Len(sValor) To 13
        sValor = " " & sValor
    Next
  End If

  sLinha = "   TOTAL       R$" & sValor
  Printer.Print sLinha
  Printer.Print ""
  Printer.Print "   Observação:"
  Printer.Print "              NFCe com situação 'Verificar Status' não esta somado nestes totais."
  Printer.Print "              Sugestão é que busque informações para saber o que ocorreu e qual é a situação atual de cada NFCe."



  Printer.EndDoc
  
  Exit Sub
Erro:
    MsgBox "Erro na impressão da grade " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo ErroC
  Dim strSQL As String
  Dim rsPesqTIPO2 As Recordset
  Dim dAutorizadas As Double
  Dim dCanceladas As Double
  Dim strStatus As String
  Dim Erro As Boolean
  Dim lTotal As Long
  Dim sDtVencimento As String
  Dim sValor As String
  
  With grdMovimento
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  dAutorizadas = 0
  dCanceladas = 0
  lTotal = 0
  txt_totalNFe.Text = ""
  txt_totalNFeCanc.Text = ""
  txt_totalRegistros.Text = ""

  Rem Verifica Data
  Erro = False
  If IsNull(Data_Ini.Text) Then Erro = True
  If Not Erro Then If Not IsDate(Data_Ini.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data incorreta, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
 
  Rem Verifica Data Final
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

  
  If opt_tpNFe.Value = True Then
        '************************
        'NFe
  
        If chk_somenteNF_parcelada.Value = vbChecked Then
            
            'Condição para apenas saídas PARCELADAS
            
            strSQL = "Select S.Data, S.Sequência, S.Cliente, S.Total, N.ChaveAcesso, N.Serie, N.Numero, "
            strSQL = strSQL & " N.Status, N.ProtocoloAutorizacao,N.ProtocoloCancelamento, C.Nome, R.Vencimento, "
            strSQL = strSQL & " R.Valor, S.Serviços "
            strSQL = strSQL & " from Saídas S, NFe N, Cli_for C, [Contas a Receber] R "
            strSQL = strSQL & " WHERE (S.Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
            strSQL = strSQL & " AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
          
            If Combo_Filial.Text <> "" And Combo_Filial.Text <> "0" Then
                strSQL = strSQL & "AND S.Filial = " & Combo_Filial.Text
            End If
            
            ' condição do parcelamento
            strSQL = strSQL & " AND S.Filial = R.Filial "
            strSQL = strSQL & " AND S.Sequência = R.Sequência "
            strSQL = strSQL & " AND R.[Data Emissão] <> R.Vencimento "
            
            If txtSequencia.Text = "" Then
                If Combo_Vendedor.Text <> "" And Combo_Vendedor.Text <> "0" Then
                    strSQL = strSQL & "AND S.Digitador = " & Combo_Vendedor.Text
                End If
              
                If Combo_Operação.Text <> "" And Combo_Operação.Text <> "0" Then
                    strSQL = strSQL & "AND S.Operação = " & Combo_Operação.Text
                End If
              
                If Combo_Cliente.Text <> "" And Combo_Cliente.Text <> "0" Then
                    strSQL = strSQL & "AND S.Cliente = " & Combo_Cliente.Text
                End If
            End If
            
            strSQL = strSQL & " AND S.Cliente = C.Código"
        
            If txtSequencia.Text <> "" Then
                strSQL = strSQL & " AND S.Sequência = " & txtSequencia.Text
            End If
            strSQL = strSQL & " AND S.Sequência = N.Sequencia "
          
            strSQL = strSQL & " AND N.Modelo = '55' "
            If O_Nota_N_Canc.Value = True Then
                'Somente AUTORIZADAS
                ' 100 é autorizada
                ' 101 é cancelada
                strSQL = strSQL & " AND N.Status <> 101 "
            ElseIf O_Nota_Canc.Value = True Then
                'Somente CANCELADAS
                strSQL = strSQL & " AND N.Status = 101 "
    '''        Else
    '''            'Todas = AUTORIZADAS E CANCELADAS
    '''            strSQL = strSQL & " AND N.Status in(100,101) "
            End If
          
            'strSQL = strSQL & " ORDER BY S.Sequência DESC"
            
            'Ordenar por:
            'Sequência
            'Sequência Desc
            'Data
            'Data Desc
            'Nota Fiscal
            'Nota Fiscal DESC
            'Valor
            'Valor Desc
            'Cliente
            'Cliente Desc
            If cbo_ordenar.Text = "" Or cbo_ordenar.Text = "Sequência" Then
                strSQL = strSQL & " ORDER BY S.Sequência"
            ElseIf cbo_ordenar.Text = "Sequência DESC" Then
                strSQL = strSQL & " ORDER BY S.Sequência DESC"
            ElseIf cbo_ordenar.Text = "Data" Then
                strSQL = strSQL & " ORDER BY S.Data"
            ElseIf cbo_ordenar.Text = "Data DESC" Then
                strSQL = strSQL & " ORDER BY S.Data DESC"
            ElseIf cbo_ordenar.Text = "Nota Fiscal" Then
                strSQL = strSQL & " ORDER BY N.Numero"
            ElseIf cbo_ordenar.Text = "Nota Fiscal DESC" Then
                strSQL = strSQL & " ORDER BY N.Numero DESC"
            ElseIf cbo_ordenar.Text = "Valor" Then
                strSQL = strSQL & " ORDER BY S.Total"
            ElseIf cbo_ordenar.Text = "Valor DESC" Then
                strSQL = strSQL & " ORDER BY S.Total DESC"
            ElseIf cbo_ordenar.Text = "Cliente" Then
                strSQL = strSQL & " ORDER BY C.Nome"
            ElseIf cbo_ordenar.Text = "Cliente DESC" Then
                strSQL = strSQL & " ORDER BY C.Nome DESC"
            End If
        Else
            strSQL = "Select S.Data, S.Sequência, S.Cliente, S.Total, N.ChaveAcesso, N.Serie, N.Numero, N.Status, "
            strSQL = strSQL & " N.ProtocoloAutorizacao,N.ProtocoloCancelamento, C.Nome, S.Serviços "
            strSQL = strSQL & " from Saídas S, NFe N, Cli_for C "
            strSQL = strSQL & " WHERE (S.Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
            strSQL = strSQL & " AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
          
            If Combo_Filial.Text <> "" And Combo_Filial.Text <> "0" Then
                strSQL = strSQL & "AND S.Filial = " & Combo_Filial.Text
            End If
            
            If txtSequencia.Text = "" Or txtSequencia.Text = "0" Then
                If Combo_Vendedor.Text <> "" And Combo_Vendedor.Text <> "0" Then
                    strSQL = strSQL & "AND S.Digitador = " & Combo_Vendedor.Text
                End If
              
                If Combo_Operação.Text <> "" And Combo_Operação.Text <> "0" Then
                    strSQL = strSQL & "AND S.Operação = " & Combo_Operação.Text
                End If
              
                If Combo_Cliente.Text <> "" And Combo_Cliente.Text <> "0" Then
                    strSQL = strSQL & "AND S.Cliente = " & Combo_Cliente.Text
                End If
            End If
            
            strSQL = strSQL & " AND S.Cliente = C.Código"
        
            If txtSequencia.Text <> "" And txtSequencia.Text <> "0" Then
                strSQL = strSQL & " AND S.Sequência = " & txtSequencia.Text
            End If
            strSQL = strSQL & " AND S.Sequência = N.Sequencia "
            strSQL = strSQL & " AND S.Filial = N.Filial "
          
            strSQL = strSQL & " AND N.Modelo = '55' "
            If O_Nota_N_Canc.Value = True Then
                'Somente AUTORIZADAS
                ' 100 é autorizada
                ' 101 é cancelada
                strSQL = strSQL & " AND N.Status <> 101 "
            ElseIf O_Nota_Canc.Value = True Then
                'Somente CANCELADAS
                strSQL = strSQL & " AND N.Status = 101 "
    '''        Else
    '''            'Todas = AUTORIZADAS E CANCELADAS
    '''            strSQL = strSQL & " AND N.Status in(100,101) "
            End If
          
            'strSQL = strSQL & " ORDER BY S.Sequência DESC"
            
            'Ordenar por:
            'Sequência
            'Sequência Desc
            'Data
            'Data Desc
            'Nota Fiscal
            'Nota Fiscal DESC
            'Valor
            'Valor Desc
            'Cliente
            'Cliente Desc
            If cbo_ordenar.Text = "" Or cbo_ordenar.Text = "Sequência" Then
                strSQL = strSQL & " ORDER BY S.Sequência"
            ElseIf cbo_ordenar.Text = "Sequência DESC" Then
                strSQL = strSQL & " ORDER BY S.Sequência DESC"
            ElseIf cbo_ordenar.Text = "Data" Then
                strSQL = strSQL & " ORDER BY S.Data"
            ElseIf cbo_ordenar.Text = "Data DESC" Then
                strSQL = strSQL & " ORDER BY S.Data DESC"
            ElseIf cbo_ordenar.Text = "Nota Fiscal" Then
                strSQL = strSQL & " ORDER BY N.Numero"
            ElseIf cbo_ordenar.Text = "Nota Fiscal DESC" Then
                strSQL = strSQL & " ORDER BY N.Numero DESC"
            ElseIf cbo_ordenar.Text = "Valor" Then
                strSQL = strSQL & " ORDER BY S.Total"
            ElseIf cbo_ordenar.Text = "Valor DESC" Then
                strSQL = strSQL & " ORDER BY S.Total DESC"
            ElseIf cbo_ordenar.Text = "Cliente" Then
                strSQL = strSQL & " ORDER BY C.Nome"
            ElseIf cbo_ordenar.Text = "Cliente DESC" Then
                strSQL = strSQL & " ORDER BY C.Nome DESC"
            End If
        End If
        
        Set rsPesqTIPO2 = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
        With rsPesqTIPO2
          If Not (.BOF And .EOF) Then
            Do Until .EOF
              
              If chk_somenteNF_parcelada.Value = vbChecked Then
                  sDtVencimento = .Fields("Vencimento").Value
                  sValor = Format(.Fields("Valor").Value - .Fields("Serviços").Value, FORMAT_VALUE)
              Else
                  sDtVencimento = ""
                  sValor = Format(.Fields("Total").Value - .Fields("Serviços").Value, FORMAT_VALUE)
              End If
              
              
              If rsPesqTIPO2.Fields("Status").Value = 101 Then
                strStatus = "Cancelada"
                
                
                If O_Nota_Canc.Value = True Or Option1.Value = True Then
                    'dCanceladas = dCanceladas + CDbl(rsPesqTIPO2.Fields("Total").Value)
                    dCanceladas = dCanceladas + CDbl(sValor) - .Fields("Serviços").Value
                End If
              ElseIf rsPesqTIPO2.Fields("Status").Value = 100 Then
                strStatus = "Solicitada/Autorizada"
                
                If O_Nota_N_Canc.Value = True Or Option1.Value = True Then
                    'dAutorizadas = dAutorizadas + CDbl(rsPesqTIPO2.Fields("Total").Value)
                    dAutorizadas = dAutorizadas + CDbl(sValor) - .Fields("Serviços").Value
                End If
              Else
                strStatus = "Verificar Status"
              End If
              
              
              'Adiciona registro
              grdMovimento.AddItem .Fields("Data").Value & vbTab & _
                                 sDtVencimento & vbTab & _
                                 .Fields("Sequência").Value & vbTab & _
                                 .Fields("Cliente").Value & vbTab & _
                                 .Fields("Nome").Value & vbTab & _
                                 .Fields("Serie").Value & vbTab & _
                                 .Fields("Numero").Value & vbTab & _
                                 sValor & vbTab & _
                                 strStatus & vbTab & _
                                 .Fields("ChaveAcesso") & vbTab & _
                                 .Fields("ProtocoloAutorizacao") & vbTab & _
                                 .Fields("ProtocoloCancelamento")

                                
              .MoveNext
              lTotal = lTotal + 1
            Loop
          End If
          .Close
        End With
        Set rsPesqTIPO2 = Nothing
  Else
        '*******************************
        'NFCe....É CUPOM FISCAL

        If chk_somenteNF_parcelada.Value = vbChecked Then
            
            'Condição para apenas saídas PARCELADAS
        
            strSQL = "Select S.Data, S.Sequência, S.Cliente, S.Total, S.ChaveNFCe, S.SerieNF, S.NFCe, S.RetNFCe, C.Nome, R.Vencimento, R.Valor, S.Serviços "
            strSQL = strSQL & " from Saídas S, Cli_for C, [Contas a Receber] R  "
            strSQL = strSQL & " WHERE (S.Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
            strSQL = strSQL & " AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
            strSQL = strSQL & " AND S.NFCe > 0 "
          
            If Combo_Filial.Text <> "" And Combo_Filial.Text <> "0" Then
                strSQL = strSQL & "AND S.Filial = " & Combo_Filial.Text
            End If
            
            ' condição do parcelamento
            strSQL = strSQL & "AND S.Filial = R.Filial "
            strSQL = strSQL & "AND S.Sequência = R.Sequência "
            strSQL = strSQL & "AND R.[Data Emissão] <> R.Vencimento "
            
            If txtSequencia.Text = "" Then
                If Combo_Vendedor.Text <> "" And Combo_Vendedor.Text <> "0" Then
                    strSQL = strSQL & "AND S.Digitador = " & Combo_Vendedor.Text
                End If
              
                If Combo_Operação.Text <> "" And Combo_Operação.Text <> "0" Then
                    strSQL = strSQL & "AND S.Operação = " & Combo_Operação.Text
                End If
              
                If Combo_Cliente.Text <> "" And Combo_Cliente.Text <> "0" Then
                    strSQL = strSQL & "AND S.Cliente = " & Combo_Cliente.Text
                End If
            End If
            
            strSQL = strSQL & " AND S.Cliente = C.Código"
            
            If txtSequencia.Text <> "" Then
                strSQL = strSQL & " AND S.Sequência = " & txtSequencia.Text
            End If
          
            'strSQL = strSQL & " ORDER BY S.Sequência DESC"
            'Ordenar por:
            'Sequência
            'Sequência Desc
            'Data
            'Data Desc
            'Nota Fiscal
            'Nota Fiscal DESC
            'Valor
            'Valor Desc
            'Cliente
            'Cliente Desc
            If cbo_ordenar.Text = "" Or cbo_ordenar.Text = "Sequência" Then
                strSQL = strSQL & " ORDER BY S.Sequência"
            ElseIf cbo_ordenar.Text = "Sequência DESC" Then
                strSQL = strSQL & " ORDER BY S.Sequência DESC"
            ElseIf cbo_ordenar.Text = "Data" Then
                strSQL = strSQL & " ORDER BY S.Data"
            ElseIf cbo_ordenar.Text = "Data DESC" Then
                strSQL = strSQL & " ORDER BY S.Data DESC"
            ElseIf cbo_ordenar.Text = "Nota Fiscal" Then
                strSQL = strSQL & " ORDER BY N.Numero"
            ElseIf cbo_ordenar.Text = "Nota Fiscal DESC" Then
                strSQL = strSQL & " ORDER BY N.Numero DESC"
            ElseIf cbo_ordenar.Text = "Valor" Then
                strSQL = strSQL & " ORDER BY S.Total"
            ElseIf cbo_ordenar.Text = "Valor DESC" Then
                strSQL = strSQL & " ORDER BY S.Total DESC"
            ElseIf cbo_ordenar.Text = "Cliente" Then
                strSQL = strSQL & " ORDER BY C.Nome"
            ElseIf cbo_ordenar.Text = "Cliente DESC" Then
                strSQL = strSQL & " ORDER BY C.Nome DESC"
            End If
        
        Else
            strSQL = "Select S.Data, S.Sequência, S.Cliente, S.Total, S.ChaveNFCe, S.SerieNF, S.NFCe, "
            strSQL = strSQL & " S.RetNFCe, C.Nome, S.Serviços "
            strSQL = strSQL & " from Saídas S, Cli_for C "
            strSQL = strSQL & " WHERE (S.Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
            strSQL = strSQL & " AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
            strSQL = strSQL & " AND S.NFCe > 0 "
          
            If Combo_Filial.Text <> "" And Combo_Filial.Text <> "0" Then
                strSQL = strSQL & "AND S.Filial = " & Combo_Filial.Text
            End If
            
            If txtSequencia.Text = "" Then
                If Combo_Vendedor.Text <> "" And Combo_Vendedor.Text <> "0" Then
                    strSQL = strSQL & "AND S.Digitador = " & Combo_Vendedor.Text
                End If
              
                If Combo_Operação.Text <> "" And Combo_Operação.Text <> "0" Then
                    strSQL = strSQL & "AND S.Operação = " & Combo_Operação.Text
                End If
              
                If Combo_Cliente.Text <> "" And Combo_Cliente.Text <> "0" Then
                    strSQL = strSQL & "AND S.Cliente = " & Combo_Cliente.Text
                End If
            End If
            
            strSQL = strSQL & " AND S.Cliente = C.Código"
            
            If txtSequencia.Text <> "" Then
                strSQL = strSQL & " AND S.Sequência = " & txtSequencia.Text
            End If
          
            'strSQL = strSQL & " ORDER BY S.Sequência DESC"
            'Ordenar por:
            'Sequência
            'Sequência Desc
            'Data
            'Data Desc
            'Nota Fiscal
            'Nota Fiscal DESC
            'Valor
            'Valor Desc
            'Cliente
            'Cliente Desc
            If cbo_ordenar.Text = "" Or cbo_ordenar.Text = "Sequência" Then
                strSQL = strSQL & " ORDER BY S.Sequência"
            ElseIf cbo_ordenar.Text = "Sequência DESC" Then
                strSQL = strSQL & " ORDER BY S.Sequência DESC"
            ElseIf cbo_ordenar.Text = "Data" Then
                strSQL = strSQL & " ORDER BY S.Data"
            ElseIf cbo_ordenar.Text = "Data DESC" Then
                strSQL = strSQL & " ORDER BY S.Data DESC"
            ElseIf cbo_ordenar.Text = "Nota Fiscal" Then
                strSQL = strSQL & " ORDER BY N.Numero"
            ElseIf cbo_ordenar.Text = "Nota Fiscal DESC" Then
                strSQL = strSQL & " ORDER BY N.Numero DESC"
            ElseIf cbo_ordenar.Text = "Valor" Then
                strSQL = strSQL & " ORDER BY S.Total"
            ElseIf cbo_ordenar.Text = "Valor DESC" Then
                strSQL = strSQL & " ORDER BY S.Total DESC"
            ElseIf cbo_ordenar.Text = "Cliente" Then
                strSQL = strSQL & " ORDER BY C.Nome"
            ElseIf cbo_ordenar.Text = "Cliente DESC" Then
                strSQL = strSQL & " ORDER BY C.Nome DESC"
            End If
        End If
        
        Set rsPesqTIPO2 = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
        With rsPesqTIPO2
          If Not (.BOF And .EOF) Then
            Do Until .EOF

              If chk_somenteNF_parcelada.Value = vbChecked Then
                  sDtVencimento = .Fields("Vencimento").Value
                  sValor = Format(.Fields("Valor").Value - .Fields("Serviços").Value, FORMAT_VALUE)
              Else
                  sDtVencimento = ""
                  sValor = Format(.Fields("Total").Value - .Fields("Serviços").Value, FORMAT_VALUE)
              End If
              
              If InStr(rsPesqTIPO2("retNFCe").Value, "<detalheCancelamento>135") Then
                strStatus = "Cancelada"
                
                If O_Nota_Canc.Value = True Or Option1.Value = True Then
                    'dCanceladas = dCanceladas + CDbl(rsPesqTIPO2.Fields("Total").Value)
                    dCanceladas = dCanceladas + CDbl(sValor) - .Fields("Serviços").Value
                End If
              ElseIf InStr(rsPesqTIPO2("retNFCe").Value, "<statusAutorizacao>OK</statusAutorizacao>") _
                  Or InStr(rsPesqTIPO2("retNFCe").Value, "<statusAutorizacao>Pendente</statusAutorizacao>") _
                  Or InStr(rsPesqTIPO2("retNFCe").Value, "<detalheAutorizacao>100") Then
                strStatus = "Solicitada/Autorizada"
                
                If O_Nota_N_Canc.Value = True Or Option1.Value = True Then
                    'dAutorizadas = dAutorizadas + CDbl(rsPesqTIPO2.Fields("Total").Value)
                    dAutorizadas = dAutorizadas + CDbl(sValor) - .Fields("Serviços").Value
                End If
              Else
                strStatus = "Verificar Status"
              End If

              
              If O_Nota_Canc.Value = True Then
                  If InStr(rsPesqTIPO2("retNFCe").Value, "<detalheCancelamento>135") Then
                      'Adiciona registro
                      grdMovimento.AddItem .Fields("Data").Value & vbTab & _
                                sDtVencimento & vbTab & _
                                .Fields("Sequência").Value & vbTab & _
                                .Fields("Cliente").Value & vbTab & _
                                .Fields("Nome").Value & vbTab & _
                                .Fields("SerieNF").Value & vbTab & _
                                .Fields("NFCe").Value & vbTab & _
                                sValor & vbTab & _
                                strStatus & vbTab & _
                                "" & vbTab & _
                                "" & vbTab & _
                                ""
                                                                
                    End If
              ElseIf O_Nota_N_Canc.Value = True Then
                  If (InStr(rsPesqTIPO2("retNFCe").Value, "<statusAutorizacao>OK</statusAutorizacao>") _
                      Or InStr(rsPesqTIPO2("retNFCe").Value, "<statusAutorizacao>Pendente</statusAutorizacao>") _
                      Or InStr(rsPesqTIPO2("retNFCe").Value, "<detalheAutorizacao>100")) _
                      And InStr(rsPesqTIPO2("retNFCe").Value, "<detalheCancelamento>135") < 1 Then
                      'Adiciona registro
                      grdMovimento.AddItem .Fields("Data").Value & vbTab & _
                                sDtVencimento & vbTab & _
                                .Fields("Sequência").Value & vbTab & _
                                .Fields("Cliente").Value & vbTab & _
                                .Fields("Nome").Value & vbTab & _
                                .Fields("SerieNF").Value & vbTab & _
                                .Fields("NFCe").Value & vbTab & _
                                sValor & vbTab & _
                                strStatus & vbTab & _
                                "" & vbTab & _
                                "" & vbTab & _
                                ""
                                
                  
                  End If
              Else
                      'Adiciona registro
                      grdMovimento.AddItem .Fields("Data").Value & vbTab & _
                                sDtVencimento & vbTab & _
                                .Fields("Sequência").Value & vbTab & _
                                .Fields("Cliente").Value & vbTab & _
                                .Fields("Nome").Value & vbTab & _
                                .Fields("SerieNF").Value & vbTab & _
                                .Fields("NFCe").Value & vbTab & _
                                sValor & vbTab & _
                                strStatus & vbTab & _
                                "" & vbTab & _
                                "" & vbTab & _
                                ""

              
              End If
            
              .MoveNext
              lTotal = lTotal + 1
            Loop
          End If
          .Close
        End With
        Set rsPesqTIPO2 = Nothing
  
  End If
  
  txt_totalNFe.Text = FormatNumber(dAutorizadas, 2)
  txt_totalNFeCanc.Text = FormatNumber(dCanceladas, 2)
  txt_totalRegistros.Text = lTotal

  Exit Sub
ErroC:
  MsgBox "Erro ao tentar realizar a pesquisa - Cod " & Err.Number & " " & Err.Description, vbInformation, "Erro"
End Sub

Private Sub Combo_Filial_CloseUp()
  Combo_Filial.Text = Combo_Filial.Columns(1).Text
  Combo_Filial_LostFocus
End Sub

Private Sub Combo_Filial_LostFocus()
  Nome_Empresa.Caption = ""
  If IsNull(Combo_Filial.Text) Then Exit Sub
  If Not IsNumeric(Combo_Filial.Text) Then Exit Sub
  If Val(Combo_Filial.Text) > 99 Then Exit Sub
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Filial.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")
End Sub


Private Sub Combo_Operação_CloseUp()
  Combo_Operação.Text = Combo_Operação.Columns(1).Text
  Combo_Operação_LostFocus
End Sub

Private Sub Combo_Operação_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Operação_LostFocus()
  Call StatusMsg("")
  Nome_Operação.Caption = ""
  If IsNull(Combo_Operação.Text) Then Exit Sub
  If Not IsNumeric(Combo_Operação.Text) Then Exit Sub
  If Val(Combo_Operação.Text) > 999 Then Exit Sub
  rsOp_Saída.Index = "Código"
  rsOp_Saída.Seek "=", Val(Combo_Operação.Text)
  If rsOp_Saída.NoMatch Then Exit Sub
  Nome_Operação.Caption = rsOp_Saída("Nome")
End Sub


Private Sub Data_Ini_LostFocus()
  Data_Ini.Text = Ajusta_Data(Data_Ini.Text)
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

Private Sub Data_Fim_LostFocus()
  Data_Fim.Text = Ajusta_Data(Data_Fim.Text)
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


Private Sub cboProduto_CloseUp()
  cboProduto.Text = cboProduto.Columns(0).Text
  cboProduto_LostFocus
End Sub

Private Sub cboProduto_LostFocus()
  Dim rstProdutos As Recordset
  
  lbl_NomeProduto.Caption = ""
  
  Set rstProdutos = db.OpenRecordset("SELECT Código, Nome FROM Produtos WHERE Código = '" & cboProduto.Text & "' AND Código <> '0' ", dbOpenSnapshot)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      lbl_NomeProduto.Caption = .Fields("Nome") & ""
      cboProduto.Text = .Fields("Código") & ""
    End If
    
    If Not rstProdutos Is Nothing Then .Close
    Set rstProdutos = Nothing
  End With
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  '26/08/2013 - Jean
  'Habilitado filtro de nota fiscal para Disk Embalagens
  If CheckSerialCaseMod("QS73520-469") Then
      Frame4.Visible = True
  End If
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  '16/10/2007 - Anderson
  'Customização de relatório para Agrotama
  dtaVendedor.DatabaseName = gsQuickDBFileName
  '08/11/2007 - Celso
  'Customização de relatório para filtrar cliente
  dtaCliente.DatabaseName = gsQuickDBFileName
  
  datProdutos.DatabaseName = gsQuickDBFileName
  
  Data_Fim.Text = Format(Date, "dd/mm/yyyy")
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsOp_Saída = db.OpenRecordset("Operações Saída", , dbReadOnly)
  Set rsSaidas = db.OpenRecordset("Saídas", , dbReadOnly)
  
  '16/10/2007 - Anderson
  'Implementação do filtro vendedor
  'Solicitado por Agrotama
  Set rsVendedor = db.OpenRecordset("Funcionários", , dbReadOnly)
  
  '08/11/2007 - Celso
  'Implementação do filtro cliente
  'Solicitado por Litoral Materiais de Construção.
  Set rsCliente = db.OpenRecordset("Cli_For", , dbReadOnly)

  Combo_Filial.Text = gnCodFilial
  
  optTipoRel2.Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsOp_Saída.Close
  rsSaidas.Close
  '16/10/2007 - Anderson
  'Implementação do filtro vendedor
  'Solicitado por Agrotama
  rsVendedor.Close
  
  '08/11/2007 - Celso
  'Implementação do filtro cliente
  'Solicitado por Litoral Materiais de Construção.
  rsCliente.Close
  Set rsCliente = Nothing
  
  Set rsParametros = Nothing
  Set rsOp_Saída = Nothing
  Set rsSaidas = Nothing
  '16/10/2007 - Anderson
  'Implementação do filtro vendedor
  'Solicitado por Agrotama
  Set rsVendedor = Nothing
End Sub

'16/10/2007 - Anderson
'Implementação do filtro vendedor
'Solicitado por Agrotama
Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(1).Text
  Combo_Vendedor_LostFocus
End Sub

'16/10/2007 - Anderson
'Implementação do filtro vendedor
'Solicitado por Agrotama
Private Sub Combo_Vendedor_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

'16/10/2007 - Anderson
'Implementação do filtro vendedor
'Solicitado por Agrotama
Private Sub Combo_Vendedor_LostFocus()
  Call StatusMsg("")
  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) > 999 Then Exit Sub
  rsVendedor.Index = "Código"
  rsVendedor.Seek "=", Val(Combo_Vendedor.Text)
  If rsVendedor.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsVendedor("Nome")
End Sub

'08/11/2007 - Celso
'Implementação do filtro cliente
'Solicitado por Litoral Materiais de Construção.
Private Sub Combo_Cliente_CloseUp()
  Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
  Combo_Cliente_LostFocus
End Sub

'08/11/2007 - Celso
'Implementação do filtro cliente
'Solicitado por Litoral Materiais de Construção.
Private Sub Combo_Cliente_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

'08/11/2007 - Celso
'Implementação do filtro cliente
'Solicitado por Litoral Materiais de Construção.
Private Sub Combo_Cliente_LostFocus()
  Call StatusMsg("")
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 0 Or Val(Combo_Cliente.Text) > 99999999 Then Exit Sub
  rsCliente.Index = "Código"
  rsCliente.Seek "=", Val(Combo_Cliente.Text)
  If rsCliente.NoMatch Then Exit Sub
  Nome_Cliente.Caption = rsCliente("Nome")
End Sub

Private Sub optTipoRel1_Click()
  If optTipoRel1.Value = True Then
      Frame1.Visible = True
      Frame2.Visible = True
      Frame3.Visible = True
      Frame3.Width = 7085
      Frame7.Visible = True
      B_Imprime.Visible = True
      cmdPesquisar.Visible = False
      grdMovimento.Enabled = False
      Frame9.Visible = False
      Frame5.Visible = False
      lbl_obs_visaoSaidaNFCe.Visible = False
      
      With grdMovimento
        .Redraw = False
        .RemoveAll
        .Redraw = True
      End With
      cboProduto.Enabled = True
  Else
      Frame1.Visible = False
      Frame2.Visible = False
      Frame3.Visible = False
      Frame7.Visible = False
      B_Imprime.Visible = False
      cmdPesquisar.Visible = True
      grdMovimento.Enabled = True
      Frame9.Visible = True
      lbl_obs_visaoSaidaNFCe.Visible = True
      cboProduto.Enabled = False
      Frame5.Visible = True
  End If
End Sub

Private Sub optTipoRel2_Click()
  If optTipoRel2.Value = True Then
      Frame1.Visible = False
      Frame2.Visible = False
      Frame3.Visible = False
      Frame7.Visible = False
      B_Imprime.Visible = False
      
      Frame5.Visible = True
      Frame5.Left = 60
      Frame5.Width = 7085
      
      cmdPesquisar.Visible = True
      cmdPesquisar.Top = 3360
      grdMovimento.Enabled = True
      
      Frame9.Visible = True
      Frame9.Left = 7230
      Frame9.Width = 7030
      
      lbl_obs_visaoSaidaNFCe.Visible = True
      cboProduto.Text = ""
      lbl_NomeProduto.Caption = ""
      cboProduto.Enabled = False
  Else
      Frame5.Visible = False
      Frame9.Visible = False
      
      Frame1.Visible = True
      Frame2.Visible = True
      Frame3.Visible = True
      Frame3.Width = 7085
      
      Frame7.Visible = True
      B_Imprime.Visible = True
      cmdPesquisar.Visible = False
      grdMovimento.Enabled = False
      lbl_obs_visaoSaidaNFCe.Visible = False
      cboProduto.Enabled = True
  End If
End Sub

