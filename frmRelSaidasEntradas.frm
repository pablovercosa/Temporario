VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRelSaidasEntradas 
   BackColor       =   &H00FFA324&
   Caption         =   " Piloto - Saídas e Entradas no período"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmRelSaidasEntradas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   15480
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txt_CodOperacoesVenda 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   10560
      TabIndex        =   65
      Top             =   780
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CheckBox chk_desconsiderar_itensDoProdutoDevolvido 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Desconsiderar itens de produto devolvidos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9660
      TabIndex        =   63
      Top             =   1080
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Frame frm_vendasSeparadasPorDia 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8505
      Left            =   -1140
      TabIndex        =   60
      Top             =   8310
      Visible         =   0   'False
      Width           =   16755
      Begin VB.CommandButton cmd_imprimirGradeVendasSeparadasPorDia 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   13650
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   6390
         Width           =   1665
      End
      Begin MSFlexGridLib.MSFlexGrid gridVendasSeparadasPorDia 
         Height          =   6255
         Left            =   0
         TabIndex        =   61
         Top             =   60
         Width           =   15315
         _ExtentX        =   27014
         _ExtentY        =   11033
         _Version        =   393216
         Rows            =   1
         Cols            =   9
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
   End
   Begin VB.CheckBox chk_vendasSeparadasPorDia 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Separados por dia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8610
      TabIndex        =   59
      Top             =   1080
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CheckBox chk_produtoGradeAgrupado 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Produtos com grade agrupados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9660
      TabIndex        =   57
      Top             =   1350
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Data Data4 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   -1560
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Classes"
      Top             =   5910
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Data Data5 
      Appearance      =   0  'Flat
      Caption         =   "Data5"
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
      Height          =   270
      Left            =   -1350
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sub Classes"
      Top             =   5580
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton cmd_limparTela 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
      Caption         =   "Limpar a Tela"
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
      Left            =   14130
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   90
      Width           =   1305
   End
   Begin VB.ComboBox cmb_classificacaoSaidas 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmRelSaidasEntradas.frx":4E95A
      Left            =   7500
      List            =   "frmRelSaidasEntradas.frx":4E973
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   413
      Width           =   2115
   End
   Begin VB.CheckBox chk_visaoConsolidada 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Visão Unificada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6300
      TabIndex        =   29
      Top             =   810
      Width           =   1515
   End
   Begin VB.CheckBox chk_visaoProdutos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Visão Produtos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6300
      TabIndex        =   28
      Top             =   1320
      Width           =   1515
   End
   Begin VB.ComboBox cmb_tipoOperacao 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmRelSaidasEntradas.frx":4E9E0
      Left            =   7500
      List            =   "frmRelSaidasEntradas.frx":4E9ED
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   53
      Width           =   2115
   End
   Begin VB.CheckBox chk_formasPagamentoRecebimento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Formas de Pagamento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6300
      TabIndex        =   25
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmd_abreSequencia 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Detalhar Sequência"
      Height          =   375
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7590
      Width           =   1665
   End
   Begin VB.ComboBox cbo_ordenacao 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      ItemData        =   "frmRelSaidasEntradas.frx":4EA05
      Left            =   30
      List            =   "frmRelSaidasEntradas.frx":4EA24
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   7830
      Width           =   3375
   End
   Begin VB.CommandButton cmd_imprimirOperDet 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8040
      Width           =   1665
   End
   Begin VB.CommandButton cmd_imprimirOper 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3990
      Width           =   1665
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1230
      Width           =   1665
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
      Left            =   -1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   7020
      Visible         =   0   'False
      Width           =   1770
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
      Left            =   -1590
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cli_For"
      Top             =   6630
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
      Left            =   -1590
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1800
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
      Height          =   390
      Left            =   5760
      Picture         =   "frmRelSaidasEntradas.frx":4EA9F
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   420
      Width           =   465
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
      Height          =   390
      Left            =   2385
      Picture         =   "frmRelSaidasEntradas.frx":4F381
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   420
      Width           =   465
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "frmRelSaidasEntradas.frx":4FC63
      DataSource      =   "Data1"
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   30
      Width           =   735
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
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "frmRelSaidasEntradas.frx":4FC77
      DataSource      =   "dtaVendedor"
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   855
      Width           =   1395
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
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "frmRelSaidasEntradas.frx":4FC91
      DataSource      =   "dtaCliente"
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   1230
      Width           =   1395
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
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin MSMask.MaskEdBox Data_Ini 
      Height          =   285
      Left            =   960
      TabIndex        =   11
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   473
      Width           =   1365
      _ExtentX        =   2408
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
      Left            =   4365
      TabIndex        =   12
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   473
      Width           =   1365
      _ExtentX        =   2408
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
   Begin MSFlexGridLib.MSFlexGrid gridOperacoes 
      Height          =   2325
      Left            =   30
      TabIndex        =   15
      Top             =   1620
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   4101
      _Version        =   393216
      Rows            =   1
      Cols            =   6
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
   Begin MSFlexGridLib.MSFlexGrid gridOperacoesDetalhe 
      Height          =   2775
      Left            =   30
      TabIndex        =   18
      Top             =   4710
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454016
      BackColorSel    =   12648384
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
   Begin MSFlexGridLib.MSFlexGrid gridOperacoesProdutos 
      Height          =   825
      Left            =   13830
      TabIndex        =   47
      Top             =   4440
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   1455
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454016
      BackColorSel    =   12648384
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "frmRelSaidasEntradas.frx":4FCAA
      DataSource      =   "Data4"
      Height          =   345
      Left            =   10560
      TabIndex        =   51
      Top             =   38
      Visible         =   0   'False
      Width           =   825
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
      Columns(0).Width=   8229
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3493
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   1455
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_SubClasse 
      Bindings        =   "frmRelSaidasEntradas.frx":4FCBE
      DataSource      =   "Data5"
      Height          =   345
      Left            =   10560
      TabIndex        =   54
      Top             =   398
      Visible         =   0   'False
      Width           =   825
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
      Columns(0).Width=   8229
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3493
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   1455
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
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
   End
   Begin VB.ComboBox cmb_classificacaoEntradas 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmRelSaidasEntradas.frx":4FCD2
      Left            =   7500
      List            =   "frmRelSaidasEntradas.frx":4FCDF
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   510
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label lbl_CodOperacoesVendaEXEMPLO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Ex: 500,600"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   14160
      TabIndex        =   66
      Top             =   810
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl_CodOperacoesVenda 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Apenas estes cód.Op.Vendas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8100
      TabIndex        =   64
      Top             =   810
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lbl_avisoVisaoMovProdutosComGradeAgrupados 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   $"frmRelSaidasEntradas.frx":4FCF7
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1260
      Left            =   -780
      TabIndex        =   58
      Top             =   8340
      Visible         =   0   'False
      Width           =   6000
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_subClasse 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Sub-Classe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9660
      TabIndex        =   56
      Top             =   458
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Nome_SubClasse 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11415
      TabIndex        =   55
      Top             =   413
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Label lbl_classe 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9660
      TabIndex        =   53
      Top             =   98
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Nome_Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11415
      TabIndex        =   52
      Top             =   53
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Label lbl_outrosTit 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Cheque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   10260
      TabIndex        =   50
      Top             =   7590
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lbl_outros 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10260
      TabIndex        =   49
      Top             =   7830
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lbl_avisoCalculos 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "*Transações 'NÃO EFETIVADAS' e 'DESFEITAS' foram desconsideradas dos cálculos detalhados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3510
      TabIndex        =   48
      Top             =   8220
      Visible         =   0   'False
      Width           =   8550
   End
   Begin VB.Label lbl_valorEfetivadoNaoRecebido 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11940
      TabIndex        =   46
      Top             =   7830
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lbl_ValorNaoRecTit 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Valor não recebido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   11940
      TabIndex        =   45
      Top             =   7590
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label lbl_PrazoTit 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Prazo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8580
      TabIndex        =   44
      Top             =   7590
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lbl_prazo 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   43
      Top             =   7830
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lbl_totaEntradas 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11640
      TabIndex        =   41
      Top             =   4350
      Width           =   1605
   End
   Begin VB.Label lbl_totalSaidasDesf 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13380
      TabIndex        =   40
      Top             =   3990
      Width           =   1605
   End
   Begin VB.Label lbl_totalSaidas 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11640
      TabIndex        =   39
      Top             =   3990
      Width           =   1605
   End
   Begin VB.Label lbl_cartao 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5220
      TabIndex        =   38
      Top             =   7830
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lbl_vale 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6900
      TabIndex        =   37
      Top             =   7830
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lbl_dinheiro 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3510
      TabIndex        =   36
      Top             =   7830
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lbl_cartaoTit 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Cartão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5220
      TabIndex        =   35
      Top             =   7590
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbl_valeTit 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Vale"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6900
      TabIndex        =   34
      Top             =   7590
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbl_dinheiroTit 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Dinheiro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3510
      TabIndex        =   33
      Top             =   7590
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lbl_classificacao 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Classificação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6300
      TabIndex        =   31
      Top             =   465
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Tipo Operação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6300
      TabIndex        =   27
      Top             =   105
      Width           =   1200
   End
   Begin VB.Label lbl_ordenarPor 
      BackColor       =   &H00FFA324&
      Caption         =   "Ordenar por"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   30
      TabIndex        =   23
      Top             =   7590
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Saídas "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   10920
      TabIndex        =   17
      Top             =   4020
      Width           =   645
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Entradas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   10725
      TabIndex        =   16
      Top             =   4410
      Width           =   810
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Data Inicial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      TabIndex        =   14
      Top             =   510
      Width           =   870
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Data Final"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3555
      TabIndex        =   13
      Top             =   510
      Width           =   780
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1725
      TabIndex        =   8
      Top             =   30
      Width           =   4500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      TabIndex        =   7
      Top             =   90
      Width           =   300
   End
   Begin VB.Label Nome_Vendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2385
      TabIndex        =   6
      Top             =   855
      Width           =   3840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      TabIndex        =   5
      Top             =   900
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFA324&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      TabIndex        =   4
      Top             =   1275
      Width           =   555
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2385
      TabIndex        =   3
      Top             =   1230
      Width           =   3840
   End
End
Attribute VB_Name = "frmRelSaidasEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrFiliais() As Variant
Private numFiliais As Integer
Private arrCheques() As Variant
Private numArrCheques As Double
Private arrOperacao(100, 5) As Variant
Private rsParametros As Recordset
Private rsVendedor As Recordset
Private rsCliente As Recordset
Dim rsClasses As Recordset
Dim rsSubclasses As Recordset

Private arrProdutosDevolvidos As Variant
Private numArrProdutosDevolvidos As Long

Private Function AcharProdutoDevolvido(pProduto As String) As Double
On Error GoTo Erro
  Dim i As Double

  For i = 0 To numArrProdutosDevolvidos - 1
      If arrProdutosDevolvidos(i, 0) = pProduto Then
          AcharProdutoDevolvido = arrProdutosDevolvidos(i, 1)
          Exit Function
      End If
  Next
  
  AcharProdutoDevolvido = 0
  Exit Function
Erro:
  MsgBox "Erro na rotina de AcharProdutoDevolvido " & Err.Description, vbInformation, "Atenção"

End Function


Private Function AcharMovimentoCheques(pSequencia As Double) As String
On Error GoTo Erro
  Dim i As Double

  For i = 0 To numArrCheques - 1
      If arrCheques(i, 0) = pSequencia Then
          AcharMovimentoCheques = arrCheques(i, 1)
          Exit Function
      End If
  Next
  
  AcharMovimentoCheques = "0,00"
  Exit Function
Erro:
  MsgBox "Erro na rotina de AcharMovimentoCheques " & Err.Description, vbInformation, "Atenção"

End Function


Private Function AcharMovimentoCheques2(pFilial As Integer, pSequencia As Double) As String
On Error GoTo Erro
  Dim i As Double

  For i = 0 To numArrCheques - 1
      If arrCheques(i, 1) = pFilial And arrCheques(i, 2) = pSequencia Then
          AcharMovimentoCheques2 = arrCheques(i, 3)
          Exit Function
      End If
  Next
  
  AcharMovimentoCheques2 = "0,00"
  Exit Function
Erro:
  MsgBox "Erro na rotina de AcharMovimentoCheques2 " & Err.Description, vbInformation, "Atenção"

End Function



Private Sub MontaArray(pEfet_ou_Desf As Integer, pOperacao As Integer, pNome As String, pTipo As String, pValorEfetivada As Double, pValorDesfeita As Double)
On Error GoTo Erro
  Dim indice As Integer
  
  ' pEfet_ou_Desf:
  ' 1 = efetivada
  ' 2 = desfeita
  
  For indice = 0 To 99
    If arrOperacao(indice, 0) = "" Then
        arrOperacao(indice, 0) = pOperacao
        arrOperacao(indice, 1) = pNome
        arrOperacao(indice, 2) = pTipo
        arrOperacao(indice, 3) = pValorEfetivada
        arrOperacao(indice, 4) = pValorDesfeita
        
        Exit For
        
    ElseIf arrOperacao(indice, 0) = pOperacao Then
        If pEfet_ou_Desf = 1 Then
            arrOperacao(indice, 3) = pValorEfetivada
        Else
            arrOperacao(indice, 4) = pValorDesfeita
        End If
        
        Exit For
    End If
  Next

  Exit Sub
Erro:
  MsgBox "Erro na montagem do array " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub LimpaArray()
On Error GoTo Erro
  Dim indice As Integer
  
  For indice = 0 To 99
      arrOperacao(indice, 0) = ""
      arrOperacao(indice, 1) = ""
      arrOperacao(indice, 2) = ""
      arrOperacao(indice, 3) = ""
      arrOperacao(indice, 4) = ""
  Next

  Exit Sub
Erro:
  MsgBox "Erro na limpeza do array " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cbo_ordenacao_Click()

    If chk_visaoConsolidada.Value = vbChecked Then
        DetalharOperacaoAgrupadaDeOperacoes
    Else
        DetalharOperacao
    End If
End Sub

Private Sub chk_formasPagamentoRecebimento_Click()
    If chk_visaoProdutos.Value = vbChecked Then
        chk_formasPagamentoRecebimento.Value = vbUnchecked
        chk_vendasSeparadasPorDia.Value = vbUnchecked
        chk_vendasSeparadasPorDia.Visible = False
    ElseIf chk_formasPagamentoRecebimento.Value = vbChecked And chk_visaoProdutos.Value = vbUnchecked Then
        lbl_dinheiroTit.Visible = True
        lbl_dinheiro.Visible = True
        lbl_cartaoTit.Visible = True
        lbl_cartao.Visible = True
        lbl_valeTit.Visible = True
        lbl_vale.Visible = True
        lbl_PrazoTit.Visible = True
        lbl_prazo.Visible = True
        lbl_ValorNaoRecTit.Visible = True
        lbl_valorEfetivadoNaoRecebido.Visible = True
        lbl_avisoCalculos.Visible = True
        lbl_outros.Visible = True
        lbl_outrosTit.Visible = True
        
        chk_vendasSeparadasPorDia.Value = vbUnchecked
        chk_vendasSeparadasPorDia.Visible = True
    Else
        lbl_dinheiroTit.Visible = False
        lbl_dinheiro.Visible = False
        lbl_cartaoTit.Visible = False
        lbl_cartao.Visible = False
        lbl_valeTit.Visible = False
        lbl_vale.Visible = False
        lbl_PrazoTit.Visible = False
        lbl_prazo.Visible = False
        lbl_ValorNaoRecTit.Visible = False
        lbl_valorEfetivadoNaoRecebido.Visible = False
        lbl_avisoCalculos.Visible = False
        lbl_outros.Visible = False
        lbl_outrosTit.Visible = False
    
        gridOperacoesDetalhe.Cols = 12
        chk_vendasSeparadasPorDia.Value = vbUnchecked
        chk_vendasSeparadasPorDia.Visible = False
    End If
End Sub

Private Sub chk_produtoGradeAgrupado_Click()
    If chk_produtoGradeAgrupado.Value = vbChecked Then
        lbl_avisoVisaoMovProdutosComGradeAgrupados.Top = 7590
        lbl_avisoVisaoMovProdutosComGradeAgrupados.Left = 60
        lbl_avisoVisaoMovProdutosComGradeAgrupados.Height = 510
        lbl_avisoVisaoMovProdutosComGradeAgrupados.Width = 12560
        lbl_avisoVisaoMovProdutosComGradeAgrupados.Visible = True
    Else
        lbl_avisoVisaoMovProdutosComGradeAgrupados.Visible = False
    End If
End Sub

Private Sub chk_vendasSeparadasPorDia_Click()
    If chk_vendasSeparadasPorDia.Value = vbChecked Then
        frm_vendasSeparadasPorDia.Visible = True
        frm_vendasSeparadasPorDia.Top = 1620
        frm_vendasSeparadasPorDia.Left = 30
        
        ' Setar condições automaticamente
        cmb_tipoOperacao.Text = "Saídas"
        cmb_classificacaoSaidas.Text = "Venda"
        
        If chk_visaoConsolidada.Value = vbUnchecked Then
            chk_visaoConsolidada.Value = vbChecked
        End If
        
    Else
        frm_vendasSeparadasPorDia.Visible = False
    End If
    
    gridVendasSeparadasPorDia.Rows = 1
End Sub

Private Sub chk_visaoConsolidada_Click()
    cmb_tipoOperacao.ListIndex = 1
End Sub

Private Sub chk_visaoProdutos_Click()
    If chk_visaoProdutos.Value = vbChecked Then
        lbl_ordenarPor.Visible = False
        cbo_ordenacao.Visible = False
        gridOperacoesDetalhe.Visible = False
        
        gridOperacoesProdutos.Top = 4710
        gridOperacoesProdutos.Left = 30
        gridOperacoesProdutos.Height = 2775
        gridOperacoesProdutos.Width = 15435
        gridOperacoesProdutos.Visible = True
        
        lbl_dinheiroTit.Visible = False
        lbl_dinheiro.Visible = False
        lbl_cartaoTit.Visible = False
        lbl_cartao.Visible = False
        lbl_valeTit.Visible = False
        lbl_vale.Visible = False
        lbl_PrazoTit.Visible = False
        lbl_prazo.Visible = False
        lbl_ValorNaoRecTit.Visible = False
        lbl_valorEfetivadoNaoRecebido.Visible = False
        lbl_avisoCalculos.Visible = False
        lbl_outros.Visible = False
        lbl_outrosTit.Visible = False

        chk_formasPagamentoRecebimento.Value = vbUnchecked
        
        lbl_classe.Visible = True
        lbl_subClasse.Visible = True
        Combo_Classe.Visible = True
        Combo_SubClasse.Visible = True
        Nome_Classe.Visible = True
        Nome_SubClasse.Visible = True
        chk_produtoGradeAgrupado.Visible = True
        chk_desconsiderar_itensDoProdutoDevolvido.Visible = True
        
        lbl_CodOperacoesVenda.Visible = True
        txt_CodOperacoesVenda.Text = ""
        txt_CodOperacoesVenda.Visible = True
        lbl_CodOperacoesVendaEXEMPLO.Visible = True
        
    Else
        lbl_ordenarPor.Visible = True
        cbo_ordenacao.Visible = True
        gridOperacoesDetalhe.Visible = True
        gridOperacoesProdutos.Visible = False
        
        If chk_formasPagamentoRecebimento.Value = vbChecked Then
            lbl_dinheiroTit.Visible = True
            lbl_dinheiro.Visible = True
            lbl_cartaoTit.Visible = True
            lbl_cartao.Visible = True
            lbl_valeTit.Visible = True
            lbl_vale.Visible = True
            lbl_PrazoTit.Visible = True
            lbl_prazo.Visible = True
            lbl_ValorNaoRecTit.Visible = True
            lbl_valorEfetivadoNaoRecebido.Visible = True
            lbl_avisoCalculos.Visible = True
            lbl_outros.Visible = True
            lbl_outrosTit.Visible = True

        End If
        
        lbl_classe.Visible = False
        lbl_subClasse.Visible = False
        Combo_Classe.Visible = False
        Combo_SubClasse.Visible = False
        Nome_Classe.Visible = False
        Nome_SubClasse.Visible = False
        chk_produtoGradeAgrupado.Visible = False
        chk_desconsiderar_itensDoProdutoDevolvido.Visible = False
        
        lbl_CodOperacoesVenda.Visible = False
        txt_CodOperacoesVenda.Text = ""
        txt_CodOperacoesVenda.Visible = False
        lbl_CodOperacoesVendaEXEMPLO.Visible = False
        
        
    End If
End Sub

Private Sub cmb_tipoOperacao_Click()
    If cmb_tipoOperacao.ListIndex = 1 Then
        cmb_classificacaoSaidas.Enabled = True
    ElseIf cmb_tipoOperacao.ListIndex = 2 Then
        cmb_classificacaoSaidas.Enabled = False
        cmb_classificacaoSaidas.ListIndex = -1
        cmb_tipoOperacao.ListIndex = 0
    Else
        cmb_classificacaoSaidas.Enabled = False
        cmb_classificacaoSaidas.ListIndex = -1
    End If
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub DetalharOperacaoAgrupadaDeProdutos()
On Error GoTo Erro
  Dim strSQL As String
  Dim rstOperacoesDet As Recordset
  Dim sOperacao As String
  Dim iConta As Integer
  Dim sNomeProduto As String
  Dim sTamanho As String
  Dim sCor As String
  
  strSQL = ""
  gridOperacoesProdutos.Rows = 1
  
  iConta = 0
  If gridOperacoes.Rows > 1 Then
  
    If chk_visaoConsolidada.Value = vbUnchecked Then
        If gridOperacoes.RowSel > 0 Then
            sOperacao = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 2)
        Else
            MsgBox "Selecione uma operação na grade superior", vbInformation, "Atenção"
            Exit Sub
        End If
    Else
        For iConta = 1 To gridOperacoes.Rows - 1
          
          If iConta > 1 Then
              sOperacao = sOperacao & ","
          End If
          sOperacao = sOperacao & gridOperacoes.TextMatrix(iConta, 2)
        Next
    End If
    
    If cmb_tipoOperacao.ListIndex = -1 Or cmb_tipoOperacao.ListIndex = 0 Or cmb_tipoOperacao.ListIndex = 1 Then
        'Saídas
        
        If chk_produtoGradeAgrupado.Value = vbChecked Then
            strSQL = "SELECT P.[Código sem grade] as CodigoSemGrade, SUM(P.Qtde) as Qtde "
        Else
            strSQL = "SELECT P.Código, P.[Código sem grade] as CodigoSemGrade, SUM(P.Qtde) as Qtde "
            'strSQL = "SELECT P.Código, SUM(P.Qtde) "
        End If
        
        strSQL = strSQL & " FROM Saídas S, [Saídas - Produtos] P "
        
        If Nome_Classe.Caption <> "" Or Nome_SubClasse.Caption <> "" Then
            strSQL = strSQL & " , Produtos X "
        End If
        
        strSQL = strSQL & " Where "
        
        If Trim(Combo_Filial.Text) <> "" Then
            strSQL = strSQL & " s.Filial = " & Combo_Filial.Text & " And "
        End If
        
        strSQL = strSQL & " S.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
        
        If Trim(txt_CodOperacoesVenda.Text) <> "" Then
            strSQL = strSQL & " AND S.Operação in(" & Trim(txt_CodOperacoesVenda.Text) & ") "
        Else
        
            If iConta <= 1 Then
                strSQL = strSQL & " AND S.Operação = " & sOperacao
            Else
                strSQL = strSQL & " AND S.Operação IN (" & sOperacao & ") "
            End If
        End If
        
        strSQL = strSQL & " AND S.Efetivada = TRUE"
        strSQL = strSQL & " AND S.[Movimentação Desfeita] = FALSE"
        
        If Nome_Vendedor.Caption <> "" Then
            strSQL = strSQL & " AND S.Digitador = " & Combo_Vendedor.Text
        End If
        
        If Nome_Cliente.Caption <> "" Then
            strSQL = strSQL & " AND S.Cliente = " & Combo_Cliente.Text
        End If
        
'        If Nome_Vendedor.Caption <> "" Then
'            strSQL = strSQL & " AND S.Digitador = " & Combo_Vendedor.Text
'        End If

        strSQL = strSQL & " AND S.Filial = P.Filial "
        strSQL = strSQL & " AND S.Sequência = P.Sequência "
        
        If Nome_Classe.Caption <> "" Or Nome_SubClasse.Caption <> "" Then
            strSQL = strSQL & " AND P.[Código sem grade] = X.Código "
            
            If Nome_Classe.Caption <> "" Then
                strSQL = strSQL & " AND X.Classe = " & Combo_Classe.Text
            End If
            
            If Nome_SubClasse.Caption <> "" Or Nome_SubClasse.Caption <> "" Then
                strSQL = strSQL & " AND X.[Sub Classe] = " & Combo_SubClasse.Text
            End If
        End If

        If chk_produtoGradeAgrupado.Value = vbChecked Then
            strSQL = strSQL & " Group by P.[Código sem grade]"
        Else
            '''strSQL = strSQL & " ORDER BY P.Código "
            strSQL = strSQL & " Group by P.Código, P.[Código sem grade]"
        End If
        
        If chk_desconsiderar_itensDoProdutoDevolvido.Value = vbUnchecked Then
            '============================================
            ' Tratamento para buscar os produtos devolvidos no período e descontar no movimento vendido
            ' Observar na tela e no relatório
            Dim rsOperacoesDevolucao As Recordset
            Dim strSQL_02 As String
            
            If chk_produtoGradeAgrupado.Value = vbChecked Then
                strSQL_02 = "SELECT P.[Código sem grade] as CodigoSemGrade, SUM(P.Qtde) as Qtde "
            Else
                strSQL_02 = "SELECT P.Código, P.[Código sem grade] as CodigoSemGrade, SUM(P.Qtde) as Qtde "
                'strSQL_02 = " SELECT P.Código, SUM(P.Qtde) "
            End If
            
            strSQL_02 = strSQL_02 & " FROM Entradas E, [Entradas - Produtos] P, [Operações Entrada] O "
            
            If Nome_Classe.Caption <> "" Or Nome_SubClasse.Caption <> "" Then
                strSQL_02 = strSQL_02 & " , Produtos X "
            End If
            
            strSQL_02 = strSQL_02 & " Where "
            
            If Trim(Combo_Filial.Text) <> "" Then
                strSQL_02 = strSQL_02 & " E.Filial = " & Combo_Filial.Text & " And "
            End If
            
'''            strSQL_02 = strSQL_02 & " Where E.Filial = " & Combo_Filial.Text
            strSQL_02 = strSQL_02 & " E.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  E.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
            strSQL_02 = strSQL_02 & " AND E.Operação = O.Código "
            strSQL_02 = strSQL_02 & " AND O.Tipo = 'D' "
            
    '''        If Nome_Vendedor.Caption <> "" Then
    '''            strSQL_02 = strSQL_02 & " AND E.Digitador = " & Combo_Vendedor.Text
    '''        End If
    '''
    '''        If Nome_Cliente.Caption <> "" Then
    '''            strSQL_02 = strSQL_02 & " AND E.Fornecedor = " & Combo_Cliente.Text
    '''        End If
    
            strSQL_02 = strSQL_02 & " AND E.Filial = P.Filial "
            strSQL_02 = strSQL_02 & " AND E.Sequência = P.Sequência "
            
            If Nome_Classe.Caption <> "" Or Nome_SubClasse.Caption <> "" Then
                strSQL_02 = strSQL_02 & " AND P.[Código sem grade] = X.Código "
                
                If Nome_Classe.Caption <> "" Then
                    strSQL_02 = strSQL_02 & " AND X.Classe = " & Combo_Classe.Text
                End If
                
                If Nome_SubClasse.Caption <> "" Or Nome_SubClasse.Caption <> "" Then
                    strSQL_02 = strSQL_02 & " AND X.[Sub Classe] = " & Combo_SubClasse.Text
                End If
            End If
    
            If chk_produtoGradeAgrupado.Value = vbChecked Then
                strSQL_02 = strSQL_02 & " Group by P.[Código sem grade]"
            Else
                'strSQL_02 = strSQL_02 & " Group by P.Código"
                strSQL_02 = strSQL_02 & " Group by P.Código, P.[Código sem grade]"
            End If
            
    
            Set rsOperacoesDevolucao = db.OpenRecordset(strSQL_02, dbOpenDynaset)
            
            ' Descarregar no arrayDevolucoesProdutos
            If Not (rsOperacoesDevolucao.EOF And rsOperacoesDevolucao.BOF) Then
                numArrProdutosDevolvidos = 0
                ReDim arrProdutosDevolvidos(rsOperacoesDevolucao.RecordCount, 2)
                
                While Not rsOperacoesDevolucao.EOF
                    ' CodigoProduto devolvido
                    arrProdutosDevolvidos(numArrProdutosDevolvidos, 0) = rsOperacoesDevolucao.Fields(0).Value
                    ' Quantidade devolvida
                    arrProdutosDevolvidos(numArrProdutosDevolvidos, 1) = rsOperacoesDevolucao.Fields("Qtde").Value
    
                    numArrProdutosDevolvidos = numArrProdutosDevolvidos + 1
                    rsOperacoesDevolucao.MoveNext
                Wend
            End If
            
            rsOperacoesDevolucao.Close
            Set rsOperacoesDevolucao = Nothing
            '============================================
        End If
        
        
    ElseIf cmb_tipoOperacao.ListIndex = -1 Or cmb_tipoOperacao.ListIndex = 0 Or cmb_tipoOperacao.ListIndex = 2 Then

'''        strSQL = "SELECT E.Data, E.Sequência, E.Caixa, E.Digitador, E.Fornecedor, C.Nome, E.total, E.Efetivada "
'''        strSQL = strSQL & " FROM Entradas E, Cli_For C "
'''        strSQL = strSQL & " Where E.Filial = " & Combo_Filial.Text
'''        strSQL = strSQL & " AND E.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  E.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
'''        strSQL = strSQL & " AND E.Operação = " & sOperacao
'''
'''        If Nome_Vendedor.Caption <> "" Then
'''            strSQL = strSQL & " AND E.Digitador = " & Combo_Vendedor.Text
'''        End If
'''
'''        If Nome_Cliente.Caption <> "" Then
'''            strSQL = strSQL & " AND E.Fornecedor = " & Combo_Cliente.Text
'''        End If
'''
'''        strSQL = strSQL & " AND E.Fornecedor = C.código "
'''
'''
'''        'Ordenação por:
'''        'DATA HORA SEQUÊNCIA
'''        'EFETIVADAS
'''        'DESFEITAS
'''        'VALOR CRESCENTE
'''        'VALOR DECRESCENTE
'''        'CAIXA
'''        'VENDEDOR
'''        'CLIENTE/FORNECEDOR
'''
'''        If cbo_ordenacao.Text = "" Or cbo_ordenacao.Text = "DATA HORA SEQUÊNCIA" Or cbo_ordenacao.Text = "DESFEITAS" Then
'''            strSQL = strSQL & " ORDER BY E.Data, E.Sequência "
'''        ElseIf cbo_ordenacao.Text = "EFETIVADAS" Then
'''            strSQL = strSQL & " ORDER BY E.Efetivada, E.Data, E.Sequência "
'''        ElseIf cbo_ordenacao.Text = "VALOR CRESCENTE" Then
'''            strSQL = strSQL & " ORDER BY E.total "
'''        ElseIf cbo_ordenacao.Text = "VALOR DECRESCENTE" Then
'''            strSQL = strSQL & " ORDER BY E.total DESC "
'''        ElseIf cbo_ordenacao.Text = "CAIXA" Then
'''            strSQL = strSQL & " ORDER BY E.Caixa "
'''        ElseIf cbo_ordenacao.Text = "VENDEDOR" Then
'''            strSQL = strSQL & " ORDER BY E.Digitador "
'''        ElseIf cbo_ordenacao.Text = "CLIENTE/FORNECEDOR" Then
'''            strSQL = strSQL & " ORDER BY E.Fornecedor "
'''        End If
        
    End If
    
    Set rstOperacoesDet = db.OpenRecordset(strSQL, dbOpenDynaset)
 
    Dim dQtdeDevolvidaProd As Double
    Dim dQtdeProd As Double
 
    Dim rsProdutos As Recordset
    Dim rsGrade As Recordset
    Dim rsTamanho As Recordset
    Dim rsCor As Recordset
    Dim sCodigoProdutoAux As String
 
    Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
    Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
    Set rsTamanho = db.OpenRecordset("Tamanhos", , dbReadOnly)
    Set rsCor = db.OpenRecordset("Cores", , dbReadOnly)
 
    With rstOperacoesDet
        If Not (.BOF And .EOF) Then
            .MoveFirst
  
            Do Until .EOF
                rsProdutos.Index = "Código"
                rsProdutos.Seek "=", .Fields(0).Value
                If Not rsProdutos.NoMatch Then
                    sNomeProduto = rsProdutos.Fields("Nome").Value
                    sTamanho = ""
                    sCor = ""
                Else
                  rsGrade.Index = "Código"
                  rsGrade.Seek "=", .Fields(0).Value
                  If rsGrade.NoMatch Then
                      sNomeProduto = ""
                      sTamanho = ""
                      sCor = ""
                  Else
                      rsTamanho.Index = "Código"
                      rsTamanho.Seek "=", Mid(.Fields(0).Value, Len(.Fields(0).Value) - 5, 3)
                      If Not rsTamanho.NoMatch Then
                          sTamanho = rsTamanho.Fields("Nome").Value
                      Else
                          sTamanho = ""
                      End If
                      
                      rsCor.Index = "Código"
                      rsCor.Seek "=", Mid(.Fields(0).Value, Len(.Fields(0).Value) - 2, 3)
                      If Not rsCor.NoMatch Then
                          sCor = rsCor.Fields("Nome").Value
                      Else
                          sCor = ""
                      End If
                      
                      sCodigoProdutoAux = rsGrade("Código Original")
                      rsProdutos.Index = "Código"
                      rsProdutos.Seek "=", sCodigoProdutoAux
                      If Not rsProdutos.NoMatch Then
                          sNomeProduto = rsProdutos.Fields("Nome").Value
                      Else
                          sNomeProduto = ""
                          sTamanho = ""
                          sCor = ""
                      End If
                  End If
                End If
                
                dQtdeProd = .Fields("Qtde").Value
                
                ' Verificar se houveram devoluções no período deste produto
                If chk_desconsiderar_itensDoProdutoDevolvido.Value = vbUnchecked Then
                    dQtdeDevolvidaProd = 0
                    dQtdeDevolvidaProd = AcharProdutoDevolvido(.Fields(0).Value)
                    If dQtdeDevolvidaProd > 0 Then
                        dQtdeProd = dQtdeProd - dQtdeDevolvidaProd
                    End If
                End If
                '
                
                gridOperacoesProdutos.AddItem vbTab & .Fields(0).Value & vbTab & _
                              sNomeProduto & vbTab & _
                              sTamanho & vbTab & _
                              sCor & vbTab & _
                              dQtdeProd  '.Fields("Qtde").Value
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstOperacoesDet = Nothing
  End If
  
  rsProdutos.Close
  rsGrade.Close
  rsTamanho.Close
  rsCor.Close
  
  Set rsProdutos = Nothing
  Set rsGrade = Nothing
  Set rsTamanho = Nothing
  Set rsCor = Nothing
  
  Exit Sub
Erro:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Sub

Private Sub DetalharOperacaoAgrupadaDeOperacoes()
On Error GoTo Erro
  Dim strSQL As String
  Dim rstOperacoesDet As Recordset
  Dim rstMovimentoCheques As Recordset
  Dim sOperacao As String
  Dim sTipoOperacao As String
  Dim sEfetAux As String
  Dim sDesfAux As String
  Dim dValorDinheiro As Double
  Dim dValorCartao As Double
  Dim dValorVale As Double
  Dim dValorPrazo As Double
  Dim dValorOutros As Double   ' cheque etc
  Dim dValorNaoRecebidoConfere As Double
  Dim sValorDinheiro As String
  Dim sValorCartao As String
  Dim sValorVale As String
  Dim sValorPrazo As String
  Dim sValorOutros As String
  Dim iConta As Integer
  
  strSQL = ""
  gridOperacoesDetalhe.Rows = 1
  
  lbl_dinheiro.Caption = "0,00"
  lbl_cartao.Caption = "0,00"
  lbl_vale.Caption = "0,00"
  lbl_prazo.Caption = "0,00"
  lbl_valorEfetivadoNaoRecebido.Caption = "0,00"
  lbl_outros.Caption = "0,00"

  iConta = 0
  If gridOperacoes.Rows > 1 Then
  
    For iConta = 1 To gridOperacoes.Rows - 1
      
      If iConta > 1 Then
          sOperacao = sOperacao & ","
      End If
      sOperacao = sOperacao & gridOperacoes.TextMatrix(iConta, 2)
      sTipoOperacao = gridOperacoes.TextMatrix(iConta, 1)
    Next
    
    If sTipoOperacao = "Saída" Then
    
        If chk_formasPagamentoRecebimento.Value = vbChecked Then
        
            ' **************************************
            ' Ver se recebeu algo em cheques
            strSQL = "Select Filial, Sequência, sum(Valor) from [Movimento - Cheques] "
            strSQL = strSQL & " where Bom >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# "
            strSQL = strSQL & " and Filial = " & Combo_Filial.Text
            strSQL = strSQL & " Group by  Filial, Sequência"

            Set rstMovimentoCheques = db.OpenRecordset(strSQL, dbOpenDynaset)
            If Not (rstMovimentoCheques.EOF And rstMovimentoCheques.BOF) Then
                
                Dim lContador As Long
                ReDim arrCheques(rstMovimentoCheques.RecordCount, 2)
                
                numArrCheques = rstMovimentoCheques.RecordCount
                
                lContador = 0
                While Not rstMovimentoCheques.EOF
                    arrCheques(lContador, 0) = rstMovimentoCheques.Fields(1).Value
                    arrCheques(lContador, 1) = rstMovimentoCheques.Fields(2).Value
                    
                    lContador = lContador + 1
                    rstMovimentoCheques.MoveNext
                Wend
            End If
            
            rstMovimentoCheques.Close
            Set rstMovimentoCheques = Nothing
            ' **************************************
        
            gridOperacoesDetalhe.Cols = 18
            gridOperacoesDetalhe.ColWidth(12) = 1500
            gridOperacoesDetalhe.ColWidth(13) = 1500
            gridOperacoesDetalhe.ColWidth(14) = 1500
            gridOperacoesDetalhe.ColWidth(15) = 1500
            gridOperacoesDetalhe.ColWidth(16) = 1500
            gridOperacoesDetalhe.ColWidth(17) = 1500
            
            gridOperacoesDetalhe.TextMatrix(0, 12) = "Dinheiro"
            gridOperacoesDetalhe.TextMatrix(0, 13) = "Cartão"
            gridOperacoesDetalhe.TextMatrix(0, 14) = "Vale"
            gridOperacoesDetalhe.TextMatrix(0, 15) = "Num.Cartão"
            gridOperacoesDetalhe.TextMatrix(0, 16) = "Prazo"
            gridOperacoesDetalhe.TextMatrix(0, 17) = "Outros"
  
            strSQL = "SELECT S.Data, S.NSU_Hora, S.Sequência, S.Caixa, S.Digitador, S.Operador, S.Cliente, C.Nome, S.total, S.Efetivada, S.[Movimentação Desfeita], S.[Recebe - Dinheiro], S.[Recebe - Cartão], S.[Recebe - Vale], S.[Recebe - Num Cartão], S.[Total Prazo], S.[Valor Recebido] "
        Else
            strSQL = "SELECT S.Data, S.NSU_Hora, S.Sequência, S.Caixa, S.Digitador, S.Operador, S.Cliente, C.Nome, S.total, S.Efetivada, S.[Movimentação Desfeita]"
        End If
        
        strSQL = strSQL & " FROM Saídas S, Cli_For C "
        strSQL = strSQL & " Where S.Filial = " & Combo_Filial.Text
        strSQL = strSQL & " AND S.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
        
        If iConta <= 1 Then
            strSQL = strSQL & " AND S.Operação = " & sOperacao
        Else
            strSQL = strSQL & " AND S.Operação IN (" & sOperacao & ") "
        End If
        
        If Nome_Vendedor.Caption <> "" Then
            strSQL = strSQL & " AND S.Digitador = " & Combo_Vendedor.Text
        End If
        
        If Nome_Cliente.Caption <> "" Then
            strSQL = strSQL & " AND S.Cliente = " & Combo_Cliente.Text
        End If
        
        If Nome_Vendedor.Caption <> "" Then
            strSQL = strSQL & " AND S.Digitador = " & Combo_Vendedor.Text
        End If
       
        strSQL = strSQL & " AND S.Cliente = C.código "
        
        'Ordenação por:
        'DATA HORA SEQUÊNCIA
        'EFETIVADAS
        'DESFEITAS
        'VALOR CRESCENTE
        'VALOR DECRESCENTE
        'CAIXA
        'VENDEDOR
        'CLIENTE/FORNECEDOR
        
        If cbo_ordenacao.Text = "" Or cbo_ordenacao.Text = "DATA HORA SEQUÊNCIA" Then
            strSQL = strSQL & " ORDER BY S.Data, S.NSU_Hora, S.Sequência "
        ElseIf cbo_ordenacao.Text = "EFETIVADAS" Then
            strSQL = strSQL & " ORDER BY S.Efetivada, S.Data, S.NSU_Hora, S.Sequência "
        ElseIf cbo_ordenacao.Text = "DESFEITAS" Then
            strSQL = strSQL & " ORDER BY S.[Movimentação Desfeita], S.Data, S.NSU_Hora, S.Sequência "
        ElseIf cbo_ordenacao.Text = "VALOR CRESCENTE" Then
            strSQL = strSQL & " ORDER BY S.total "
        ElseIf cbo_ordenacao.Text = "VALOR DECRESCENTE" Then
            strSQL = strSQL & " ORDER BY S.total DESC "
        ElseIf cbo_ordenacao.Text = "CAIXA" Then
            strSQL = strSQL & " ORDER BY S.Caixa "
        ElseIf cbo_ordenacao.Text = "VENDEDOR" Then
            strSQL = strSQL & " ORDER BY S.Digitador "
        ElseIf cbo_ordenacao.Text = "CLIENTE/FORNECEDOR" Then
            strSQL = strSQL & " ORDER BY S.Cliente "
        End If
    Else
'''        strSQL = "SELECT E.Data, E.Sequência, E.Caixa, E.Digitador, E.Fornecedor, C.Nome, E.total, E.Efetivada "
'''        strSQL = strSQL & " FROM Entradas E, Cli_For C "
'''        strSQL = strSQL & " Where E.Filial = " & Combo_Filial.Text
'''        strSQL = strSQL & " AND E.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  E.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
'''        strSQL = strSQL & " AND E.Operação = " & sOperacao
'''
'''        If Nome_Vendedor.Caption <> "" Then
'''            strSQL = strSQL & " AND E.Digitador = " & Combo_Vendedor.Text
'''        End If
'''
'''        If Nome_Cliente.Caption <> "" Then
'''            strSQL = strSQL & " AND E.Fornecedor = " & Combo_Cliente.Text
'''        End If
'''
'''        strSQL = strSQL & " AND E.Fornecedor = C.código "
'''
'''
'''        'Ordenação por:
'''        'DATA HORA SEQUÊNCIA
'''        'EFETIVADAS
'''        'DESFEITAS
'''        'VALOR CRESCENTE
'''        'VALOR DECRESCENTE
'''        'CAIXA
'''        'VENDEDOR
'''        'CLIENTE/FORNECEDOR
'''
'''        If cbo_ordenacao.Text = "" Or cbo_ordenacao.Text = "DATA HORA SEQUÊNCIA" Or cbo_ordenacao.Text = "DESFEITAS" Then
'''            strSQL = strSQL & " ORDER BY E.Data, E.Sequência "
'''        ElseIf cbo_ordenacao.Text = "EFETIVADAS" Then
'''            strSQL = strSQL & " ORDER BY E.Efetivada, E.Data, E.Sequência "
'''        ElseIf cbo_ordenacao.Text = "VALOR CRESCENTE" Then
'''            strSQL = strSQL & " ORDER BY E.total "
'''        ElseIf cbo_ordenacao.Text = "VALOR DECRESCENTE" Then
'''            strSQL = strSQL & " ORDER BY E.total DESC "
'''        ElseIf cbo_ordenacao.Text = "CAIXA" Then
'''            strSQL = strSQL & " ORDER BY E.Caixa "
'''        ElseIf cbo_ordenacao.Text = "VENDEDOR" Then
'''            strSQL = strSQL & " ORDER BY E.Digitador "
'''        ElseIf cbo_ordenacao.Text = "CLIENTE/FORNECEDOR" Then
'''            strSQL = strSQL & " ORDER BY E.Fornecedor "
'''        End If
        
    End If
    
 
    dValorDinheiro = 0
    dValorCartao = 0
    dValorVale = 0
    dValorPrazo = 0
    dValorNaoRecebidoConfere = 0
    dValorOutros = 0
 
    Set rstOperacoesDet = db.OpenRecordset(strSQL, dbOpenDynaset)
 
    With rstOperacoesDet
        If Not (.BOF And .EOF) Then
            .MoveFirst
    
            Do Until .EOF
                If sTipoOperacao = "Saída" Then
                    sEfetAux = ""
                    sDesfAux = ""
                    If .Fields("Efetivada").Value = True Then
                        sEfetAux = "SIM"
                    Else
                        sEfetAux = "NÃO"
                    End If
                    
                    If .Fields("Movimentação Desfeita").Value = True Then
                        sDesfAux = "SIM"
                    Else
                        sDesfAux = "NÃO"
                    End If
                    
                    If chk_formasPagamentoRecebimento.Value = vbChecked Then
                        If .Fields("Efetivada").Value = True And .Fields("Movimentação Desfeita").Value = False Then
                        
                            If Not IsNull(.Fields("Recebe - Dinheiro").Value) Then
                                dValorDinheiro = dValorDinheiro + .Fields("Recebe - Dinheiro").Value
                                sValorDinheiro = .Fields("Recebe - Dinheiro").Value
                            Else
                                sValorDinheiro = "0"
                            End If
                            If Not IsNull(.Fields("Recebe - Cartão").Value) Then
                                dValorCartao = dValorCartao + .Fields("Recebe - Cartão").Value
                                sValorCartao = .Fields("Recebe - Cartão").Value
                            Else
                                sValorCartao = "0"
                            End If
                            If Not IsNull(.Fields("Recebe - Vale").Value) Then
                                dValorVale = dValorVale + .Fields("Recebe - Vale").Value
                                sValorVale = .Fields("Recebe - Vale").Value
                            Else
                                sValorVale = "0"
                            End If
                            If Not IsNull(.Fields("Total Prazo").Value) Then
                                dValorPrazo = dValorPrazo + .Fields("Total Prazo").Value
                                sValorPrazo = .Fields("Total Prazo").Value
                            Else
                                sValorPrazo = "0"
                            End If
                            
                            sValorOutros = AcharMovimentoCheques(.Fields("Sequência").Value)
                            dValorOutros = dValorOutros + CDbl(sValorOutros)
                            
                            If Not IsNull(.Fields("Valor Recebido").Value) Then
                                If .Fields("Valor Recebido").Value = 0 Then
                                    dValorNaoRecebidoConfere = dValorNaoRecebidoConfere + .Fields("total").Value
                                End If
                            End If
                        End If
                        
                        If sValorDinheiro = "" Then
                            sValorDinheiro = "0"
                        End If
                        If sValorCartao = "" Then
                            sValorCartao = "0"
                        End If
                        If sValorVale = "" Then
                            sValorVale = "0"
                        End If
                        If sValorPrazo = "" Then
                            sValorPrazo = "0"
                        End If
                        
                        If sValorOutros = "" Then
                            sValorOutros = "0"
                        End If
                        
                        gridOperacoesDetalhe.AddItem vbTab & .Fields("Data").Value & vbTab & _
                              .Fields("NSU_Hora").Value & vbTab & _
                              .Fields("Sequência").Value & vbTab & _
                              FormataValorTexto(.Fields("total").Value, 2) & "" & vbTab & _
                              sEfetAux & vbTab & _
                              sDesfAux & vbTab & _
                              .Fields("Caixa").Value & vbTab & _
                              .Fields("Digitador").Value & vbTab & _
                              .Fields("Operador").Value & vbTab & _
                              .Fields("Cliente").Value & vbTab & _
                              .Fields("Nome").Value & vbTab & _
                              FormataValorTexto(sValorDinheiro, 2) & vbTab & _
                              FormataValorTexto(sValorCartao, 2) & vbTab & _
                              FormataValorTexto(sValorVale, 2) & vbTab & _
                              .Fields("Recebe - Num Cartão").Value & vbTab & _
                              FormataValorTexto(sValorPrazo, 2) & vbTab & _
                              FormataValorTexto(sValorOutros, 2)
                    Else
                        gridOperacoesDetalhe.AddItem vbTab & .Fields("Data").Value & vbTab & _
                              .Fields("NSU_Hora").Value & vbTab & _
                              .Fields("Sequência").Value & vbTab & _
                              FormataValorTexto(.Fields("total").Value, 2) & "" & vbTab & _
                              sEfetAux & vbTab & _
                              sDesfAux & vbTab & _
                              .Fields("Caixa").Value & vbTab & _
                              .Fields("Digitador").Value & vbTab & _
                              .Fields("Operador").Value & vbTab & _
                              .Fields("Cliente").Value & vbTab & _
                              .Fields("Nome").Value
                    End If

                Else
                
'''                    sDesfAux = "NÃO"
'''                    sEfetAux = ""
'''                    If .Fields("Efetivada").Value = True Then
'''                        sEfetAux = "SIM"
'''                    End If
'''
'''                    gridOperacoesDetalhe.AddItem vbTab & .Fields("Data").Value & vbTab & _
'''                          vbTab & _
'''                          .Fields("Sequência").Value & vbTab & _
'''                          FormataValorTexto(.Fields("total").Value, 2) & "" & vbTab & _
'''                          sEfetAux & vbTab & _
'''                          sDesfAux & vbTab & _
'''                          .Fields("Caixa").Value & vbTab & _
'''                          .Fields("Digitador").Value & vbTab & _
'''                          vbTab & _
'''                          .Fields("Fornecedor").Value & vbTab & _
'''                          .Fields("Nome").Value
                End If
                
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstOperacoesDet = Nothing
  End If
  
  lbl_dinheiro.Caption = Format(dValorDinheiro, FORMAT_VALUE)
  lbl_cartao.Caption = Format(dValorCartao, FORMAT_VALUE)
  lbl_vale.Caption = Format(dValorVale, FORMAT_VALUE)
  lbl_prazo.Caption = Format(dValorPrazo, FORMAT_VALUE)
  lbl_valorEfetivadoNaoRecebido.Caption = Format(dValorNaoRecebidoConfere, FORMAT_VALUE)
  lbl_outros.Caption = Format(dValorOutros, FORMAT_VALUE)

  Exit Sub
Erro:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
End Sub

Private Sub DetalharOperacaoAgrupadaDeOperacoes_separadasPorDia()
On Error GoTo Erro
  Dim strSQL As String
  Dim rstOperacoesDet As Recordset
  Dim rstMovimentoCheques As Recordset
  Dim sOperacao As String
  Dim sTipoOperacao As String
  Dim sEfetAux As String
  Dim sDesfAux As String
  Dim dValorDinheiro As Double
  Dim dValorCartao As Double
  Dim dValorVale As Double
  Dim dValorPrazo As Double
  Dim dValorOutros As Double   ' cheque etc
  Dim dValorNaoRecebidoConfere As Double
  Dim sValorDinheiro As String
  Dim sValorCartao As String
  Dim sValorVale As String
  Dim sValorPrazo As String
  Dim sValorOutros As String
  Dim iConta As Integer
  Dim dValorTotalVendaDia As Double
  
  Dim sDiaCalculo As String
  Dim iFilialCalculo As Integer
  Dim dValorDinheiroTot As Double
  Dim dValorCartaoTot As Double
  Dim dValorValeTot As Double
  Dim dValorPrazoTot As Double
  Dim dValorOutrosTot As Double   ' cheque etc
  Dim dValorNaoRecebidoConfereTot As Double
  Dim i As Integer
  
  strSQL = ""
  gridVendasSeparadasPorDia.Rows = 1

  ' *****************
  ' Filiais
  If Trim(Nome_Empresa.Caption) <> "" Then
      numFiliais = 1
      ReDim arrFiliais(numFiliais, 2)
      arrFiliais(0, 0) = Combo_Filial.Text
  Else
      numFiliais = rsParametros.RecordCount
      ReDim arrFiliais(numFiliais, 2)
      rsParametros.MoveFirst
      For i = 1 To rsParametros.RecordCount
          arrFiliais(i - 1, 0) = rsParametros.Fields(0)
          rsParametros.MoveNext
      Next
  End If
  ' *****************


  iConta = 0
  
  Dim rsOperacoesSaida As Recordset
  Set rsOperacoesSaida = db.OpenRecordset("Select Código From [Operações Saída] Where Tipo = 'V' ", dbOpenDynaset)
  If Not (rsOperacoesSaida.EOF And rsOperacoesSaida.BOF) Then
      rsOperacoesSaida.MoveLast
      rsOperacoesSaida.MoveFirst
      For iConta = 1 To rsOperacoesSaida.RecordCount
          If rsOperacoesSaida.RecordCount = 1 Or iConta = 1 Then
              sOperacao = rsOperacoesSaida.Fields(0)
          Else
              sOperacao = sOperacao & "," & rsOperacoesSaida.Fields(0)
          End If
          rsOperacoesSaida.MoveNext
      Next
  Else
      Exit Sub
  End If
  rsOperacoesSaida.Close
  Set rsOperacoesSaida = Nothing
        
  ' **************************************
  ' Ver se recebeu algo em cheques
  '''strSQL = "Select Bom, Filial, sum(Valor) from [Movimento - Cheques] "
  strSQL = "Select Bom, Filial, Sequência, sum(Valor) from [Movimento - Cheques] "
  strSQL = strSQL & " where Bom >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# "
  
  If Trim(Nome_Empresa.Caption) <> "" Then
      strSQL = strSQL & " and Filial = " & Combo_Filial.Text
  End If
  
  strSQL = strSQL & " Group by  Bom, Filial, Sequência "

  Set rstMovimentoCheques = db.OpenRecordset(strSQL, dbOpenDynaset)
  If Not (rstMovimentoCheques.EOF And rstMovimentoCheques.BOF) Then
      
      Dim lContador As Long
      ReDim arrCheques(rstMovimentoCheques.RecordCount, 3)
      
      numArrCheques = rstMovimentoCheques.RecordCount
      
      lContador = 0
      While Not rstMovimentoCheques.EOF
          arrCheques(lContador, 0) = rstMovimentoCheques.Fields(0).Value
          arrCheques(lContador, 1) = rstMovimentoCheques.Fields(1).Value
          arrCheques(lContador, 2) = rstMovimentoCheques.Fields(2).Value
          arrCheques(lContador, 3) = rstMovimentoCheques.Fields(3).Value
          
          lContador = lContador + 1
          rstMovimentoCheques.MoveNext
      Wend
  End If
  
  rstMovimentoCheques.Close
  Set rstMovimentoCheques = Nothing
  ' **************************************
        
  strSQL = "SELECT S.Filial, S.Data, S.Sequência, S.Caixa, S.Digitador, S.Operador, S.Cliente, C.Nome, S.total, S.Efetivada,  "
  strSQL = strSQL & " S.[Movimentação Desfeita], S.[Recebe - Dinheiro], S.[Recebe - Cartão], S.[Recebe - Vale], "
  strSQL = strSQL & " S.[Recebe - Num Cartão], S.[Total Prazo], S.[Valor Recebido] "
  strSQL = strSQL & " FROM Saídas S, Cli_For C "
  
  If Trim(Nome_Empresa.Caption) <> "" Then
      strSQL = strSQL & " Where S.Filial = " & Combo_Filial.Text & " AND "
  Else
      strSQL = strSQL & " Where "
  End If
  
  strSQL = strSQL & " S.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
  
  If iConta <= 1 Then
      strSQL = strSQL & " AND S.Operação = " & sOperacao
  Else
      strSQL = strSQL & " AND S.Operação IN (" & sOperacao & ") "
  End If
  
  If Nome_Vendedor.Caption <> "" Then
      strSQL = strSQL & " AND S.Digitador = " & Combo_Vendedor.Text
  End If
  
  If Nome_Cliente.Caption <> "" Then
      strSQL = strSQL & " AND S.Cliente = " & Combo_Cliente.Text
  End If
  
  strSQL = strSQL & " AND S.Cliente = C.código "

  strSQL = strSQL & " ORDER BY S.Data, S.Filial "
    
  dValorDinheiro = 0
  dValorCartao = 0
  dValorVale = 0
  dValorPrazo = 0
  dValorNaoRecebidoConfere = 0
  dValorOutros = 0
 
  sDiaCalculo = ""
 
  Set rstOperacoesDet = db.OpenRecordset(strSQL, dbOpenDynaset)
 
  With rstOperacoesDet
        If Not (.BOF And .EOF) Then
            .MoveFirst
    
            Do Until .EOF
            
                If sDiaCalculo = "" Then
                    sDiaCalculo = .Fields("Data").Value
                    iFilialCalculo = .Fields("Filial").Value
                End If
                
                sEfetAux = ""
                sDesfAux = ""
                If .Fields("Efetivada").Value = True Then
                    sEfetAux = "SIM"
                Else
                    sEfetAux = "NÃO"
                End If
                
                If .Fields("Movimentação Desfeita").Value = True Then
                    sDesfAux = "SIM"
                Else
                    sDesfAux = "NÃO"
                End If
                
                If .Fields("Efetivada").Value = True And .Fields("Movimentação Desfeita").Value = False Then
                    
                    If Not IsNull(.Fields("Recebe - Dinheiro").Value) Then
                        dValorDinheiro = dValorDinheiro + .Fields("Recebe - Dinheiro").Value
                        sValorDinheiro = dValorDinheiro
                    Else
                        sValorDinheiro = "0"
                    End If
                    If Not IsNull(.Fields("Recebe - Cartão").Value) Then
                        dValorCartao = dValorCartao + .Fields("Recebe - Cartão").Value
                        sValorCartao = dValorCartao
                    Else
                        sValorCartao = "0"
                    End If
                    If Not IsNull(.Fields("Recebe - Vale").Value) Then
                        dValorVale = dValorVale + .Fields("Recebe - Vale").Value
                        sValorVale = dValorVale
                    Else
                        sValorVale = "0"
                    End If
                    If Not IsNull(.Fields("Total Prazo").Value) Then
                        dValorPrazo = dValorPrazo + .Fields("Total Prazo").Value
                        sValorPrazo = dValorPrazo
                    Else
                        sValorPrazo = "0"
                    End If
                    
                    sValorOutros = AcharMovimentoCheques2(.Fields("Filial").Value, .Fields("Sequência").Value)
                    
                    dValorOutros = dValorOutros + CDbl(sValorOutros)
                    sValorOutros = dValorOutros
                    
                    If Not IsNull(.Fields("Valor Recebido").Value) Then
                        If .Fields("Valor Recebido").Value = 0 Then
                            dValorNaoRecebidoConfere = dValorNaoRecebidoConfere + .Fields("total").Value
                        End If
                    End If
                End If
                    
                If sValorDinheiro = "" Then
                    sValorDinheiro = "0"
                End If
                If sValorCartao = "" Then
                    sValorCartao = "0"
                End If
                If sValorVale = "" Then
                    sValorVale = "0"
                End If
                If sValorPrazo = "" Then
                    sValorPrazo = "0"
                End If
                
                If sValorOutros = "" Then
                    sValorOutros = "0"
                End If
                
                dValorTotalVendaDia = dValorOutros + dValorPrazo + dValorVale + dValorCartao + dValorDinheiro
                
            
'                    gridVendasSeparadasPorDia.TextMatrix(0, 0) = ""
'                    gridVendasSeparadasPorDia.TextMatrix(0, 1) = "Data/Filial"
'                    gridVendasSeparadasPorDia.TextMatrix(0, 2) = "Dinheiro"
'                    gridVendasSeparadasPorDia.TextMatrix(0, 3) = "Cartão"
'                    gridVendasSeparadasPorDia.TextMatrix(0, 4) = "Prazo"
'                    gridVendasSeparadasPorDia.TextMatrix(0, 5) = "Cheque"
'                    gridVendasSeparadasPorDia.TextMatrix(0, 6) = "Vale"
'                    gridVendasSeparadasPorDia.TextMatrix(0, 7) = "Não financeiro"
'                    gridVendasSeparadasPorDia.TextMatrix(0, 8) = "TOTAL"
                    
                .MoveNext
                
                If .EOF = True Then
                    If gridVendasSeparadasPorDia.Rows = 1 Then
                        gridVendasSeparadasPorDia.AddItem vbTab & sDiaCalculo
                    End If
                    
                    gridVendasSeparadasPorDia.AddItem vbTab & "Filial " & iFilialCalculo & " VENDAS" & vbTab & _
                          FormataValorTexto(sValorDinheiro, 2) & vbTab & _
                          FormataValorTexto(sValorCartao, 2) & vbTab & _
                          FormataValorTexto(sValorPrazo, 2) & vbTab & _
                          FormataValorTexto(sValorOutros, 2) & vbTab & _
                          FormataValorTexto(sValorVale, 2) & vbTab & _
                          FormataValorTexto(dValorNaoRecebidoConfere, 2) & vbTab & _
                          FormataValorTexto(dValorTotalVendaDia, 2)
                          
                    For i = 0 To numFiliais - 1
                        If iFilialCalculo = arrFiliais(i, 0) Then
                            arrFiliais(i, 1) = arrFiliais(i, 1) + dValorTotalVendaDia
                        End If
                    Next
                          
                    ' Imprimir as últimas linhas com:
                    ' TOTAL PERÍODO VENDAS POR FILIAL (uma linha para cada filial)
                    ' TOTAL GERAL (uma linha somada todas as filiais)
                    ' *Obs: Dev. desconsideradas dos cálculos
                    Dim dTotalGeralPeriodo As Double
                    
                    gridVendasSeparadasPorDia.AddItem vbTab & ""
                    For i = 0 To numFiliais - 1
                        gridVendasSeparadasPorDia.AddItem vbTab & "TOTAL VENDAS Filial " & arrFiliais(i, 0) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & _
                              FormataValorTexto(arrFiliais(i, 1), 2)
                        dTotalGeralPeriodo = dTotalGeralPeriodo + arrFiliais(i, 1)
                    Next
                    gridVendasSeparadasPorDia.AddItem vbTab & "TOTAL GERAL VENDAS " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & _
                              FormataValorTexto(dTotalGeralPeriodo, 2)
                    
                    gridVendasSeparadasPorDia.AddItem vbTab & "OBS: As devoluções de produtos " & vbTab & "(caso elas existam) " & vbTab & "são desconsideradas " & vbTab & "dos cálculos"
                    
                ElseIf (sDiaCalculo <> .Fields("Data").Value) Or (sDiaCalculo = .Fields("Data").Value And iFilialCalculo <> .Fields("Filial").Value) Then
                    
                    If gridVendasSeparadasPorDia.Rows = 1 Then
                        gridVendasSeparadasPorDia.AddItem vbTab & sDiaCalculo
                    End If
                    
                    gridVendasSeparadasPorDia.AddItem vbTab & "Filial " & iFilialCalculo & " VENDAS" & vbTab & _
                          FormataValorTexto(sValorDinheiro, 2) & vbTab & _
                          FormataValorTexto(sValorCartao, 2) & vbTab & _
                          FormataValorTexto(sValorPrazo, 2) & vbTab & _
                          FormataValorTexto(sValorOutros, 2) & vbTab & _
                          FormataValorTexto(sValorVale, 2) & vbTab & _
                          FormataValorTexto(dValorNaoRecebidoConfere, 2) & vbTab & _
                          FormataValorTexto(dValorTotalVendaDia, 2)

                    If sDiaCalculo <> .Fields("Data").Value Then
                        sDiaCalculo = .Fields("Data").Value
                        gridVendasSeparadasPorDia.AddItem vbTab & sDiaCalculo
                    End If

                    For i = 0 To numFiliais - 1
                        If iFilialCalculo = arrFiliais(i, 0) Then
                            arrFiliais(i, 1) = arrFiliais(i, 1) + dValorTotalVendaDia
                        End If
                    Next
          
                    'sDiaCalculo = .Fields("Data").Value
                    iFilialCalculo = .Fields("Filial").Value
                    sValorDinheiro = ""
                    dValorDinheiro = 0
                    sValorCartao = ""
                    dValorCartao = 0
                    sValorPrazo = ""
                    dValorPrazo = 0
                    sValorVale = ""
                    dValorVale = 0
                    sValorOutros = ""
                    dValorOutros = 0
                    dValorNaoRecebidoConfere = 0
                    dValorTotalVendaDia = 0
                End If
            Loop
        End If
        .Close
  End With
  Set rstOperacoesDet = Nothing

  Exit Sub
Erro:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
End Sub


Private Sub DetalharOperacao()
On Error GoTo Erro
  Dim strSQL As String
  Dim rstOperacoesDet As Recordset
  Dim rstMovimentoCheques As Recordset
  Dim sOperacao As String
  Dim sTipoOperacao As String
  Dim sEfetAux As String
  Dim sDesfAux As String
  Dim dValorDinheiro As Double
  Dim dValorCartao As Double
  Dim dValorVale As Double
  Dim dValorPrazo As Double
  Dim dValorOutros As Double   ' cheque etc
  Dim dValorNaoRecebidoConfere As Double
  Dim sValorDinheiro As String
  Dim sValorCartao As String
  Dim sValorVale As String
  Dim sValorPrazo As String
  Dim sValorOutros As String
  Dim iConta As Integer
  Dim sHora As String
  Dim sRecebeNumCartao As String
 
  
  strSQL = ""
  gridOperacoesDetalhe.Rows = 1
  
  lbl_dinheiro.Caption = "0,00"
  lbl_cartao.Caption = "0,00"
  lbl_vale.Caption = "0,00"
  lbl_prazo.Caption = "0,00"
  lbl_valorEfetivadoNaoRecebido.Caption = "0,00"
  lbl_outros.Caption = "0,00"

  If gridOperacoes.RowSel > 0 Then
  
    sOperacao = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 2)
    sTipoOperacao = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 1)
    
    If sTipoOperacao = "Saída" Then
    
        If chk_formasPagamentoRecebimento.Value = vbChecked Then
        
            ' **************************************
            ' Ver se recebeu algo em cheques
            strSQL = "Select Filial, Sequência, sum(Valor) from [Movimento - Cheques] "
            strSQL = strSQL & " where Bom >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# "
            strSQL = strSQL & " and Filial = " & Combo_Filial.Text
            strSQL = strSQL & " Group by  Filial, Sequência"

            Set rstMovimentoCheques = db.OpenRecordset(strSQL, dbOpenDynaset)
            If Not (rstMovimentoCheques.EOF And rstMovimentoCheques.BOF) Then
                
                Dim lContador As Long
                ReDim arrCheques(rstMovimentoCheques.RecordCount, 2)
                
                numArrCheques = rstMovimentoCheques.RecordCount
                
                lContador = 0
                While Not rstMovimentoCheques.EOF
                    arrCheques(lContador, 0) = rstMovimentoCheques.Fields(1).Value
                    arrCheques(lContador, 1) = rstMovimentoCheques.Fields(2).Value
                    
                    lContador = lContador + 1
                    rstMovimentoCheques.MoveNext
                Wend
            End If
            
            rstMovimentoCheques.Close
            Set rstMovimentoCheques = Nothing
            ' **************************************
        
        
            gridOperacoesDetalhe.Cols = 18
            gridOperacoesDetalhe.ColWidth(12) = 1500
            gridOperacoesDetalhe.ColWidth(13) = 1500
            gridOperacoesDetalhe.ColWidth(14) = 1500
            gridOperacoesDetalhe.ColWidth(15) = 1500
            gridOperacoesDetalhe.ColWidth(16) = 1500
            gridOperacoesDetalhe.ColWidth(17) = 1500
            
            gridOperacoesDetalhe.TextMatrix(0, 12) = "Dinheiro"
            gridOperacoesDetalhe.TextMatrix(0, 13) = "Cartão"
            gridOperacoesDetalhe.TextMatrix(0, 14) = "Vale"
            gridOperacoesDetalhe.TextMatrix(0, 15) = "Num.Cartão"
            gridOperacoesDetalhe.TextMatrix(0, 16) = "Prazo"
            gridOperacoesDetalhe.TextMatrix(0, 17) = "Outros"
  
            strSQL = "SELECT S.Data, S.NSU_Hora, S.Sequência, S.Caixa, S.Digitador, S.Operador, S.Cliente, C.Nome, S.total, S.Efetivada, S.[Movimentação Desfeita], S.[Recebe - Dinheiro], S.[Recebe - Cartão], S.[Recebe - Vale], S.[Recebe - Num Cartão], S.[Total Prazo], S.[Valor Recebido] "
        Else
            strSQL = "SELECT S.Data, S.NSU_Hora, S.Sequência, S.Caixa, S.Digitador, S.Operador, S.Cliente, C.Nome, S.total, S.Efetivada, S.[Movimentação Desfeita]"
        End If
    
        'strSQL = "SELECT S.Data, S.NSU_Hora, S.Sequência, S.Caixa, S.Digitador, S.Operador, S.Cliente, C.Nome, S.total, S.Efetivada, S.[Movimentação Desfeita]"
        strSQL = strSQL & " FROM Saídas S, Cli_For C "
        strSQL = strSQL & " Where S.Filial = " & Combo_Filial.Text
        strSQL = strSQL & " AND S.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
        strSQL = strSQL & " AND S.Operação = " & sOperacao
        
        If Nome_Vendedor.Caption <> "" Then
            strSQL = strSQL & " AND S.Digitador = " & Combo_Vendedor.Text
        End If
        
        If Nome_Cliente.Caption <> "" Then
            strSQL = strSQL & " AND S.Cliente = " & Combo_Cliente.Text
        End If
        
        If Nome_Vendedor.Caption <> "" Then
            strSQL = strSQL & " AND S.Digitador = " & Combo_Vendedor.Text
        End If
       
        strSQL = strSQL & " AND S.Cliente = C.código "
        
        'Ordenação por:
        'DATA HORA SEQUÊNCIA
        'EFETIVADAS
        'DESFEITAS
        'VALOR CRESCENTE
        'VALOR DECRESCENTE
        'CAIXA
        'VENDEDOR
        'CLIENTE/FORNECEDOR
        
        If cbo_ordenacao.Text = "" Or cbo_ordenacao.Text = "DATA HORA SEQUÊNCIA" Then
            strSQL = strSQL & " ORDER BY S.Data, S.NSU_Hora, S.Sequência "
        ElseIf cbo_ordenacao.Text = "EFETIVADAS" Then
            strSQL = strSQL & " ORDER BY S.Efetivada, S.Data, S.NSU_Hora, S.Sequência "
        ElseIf cbo_ordenacao.Text = "DESFEITAS" Then
            strSQL = strSQL & " ORDER BY S.[Movimentação Desfeita], S.Data, S.NSU_Hora, S.Sequência "
        ElseIf cbo_ordenacao.Text = "VALOR CRESCENTE" Then
            strSQL = strSQL & " ORDER BY S.total "
        ElseIf cbo_ordenacao.Text = "VALOR DECRESCENTE" Then
            strSQL = strSQL & " ORDER BY S.total DESC "
        ElseIf cbo_ordenacao.Text = "CAIXA" Then
            strSQL = strSQL & " ORDER BY S.Caixa "
        ElseIf cbo_ordenacao.Text = "VENDEDOR" Then
            strSQL = strSQL & " ORDER BY S.Digitador "
        ElseIf cbo_ordenacao.Text = "CLIENTE/FORNECEDOR" Then
            strSQL = strSQL & " ORDER BY S.Cliente "
        End If
    Else
        strSQL = "SELECT E.Data, E.Sequência, E.Caixa, E.Digitador, E.Fornecedor, C.Nome, E.total, E.Efetivada "
        strSQL = strSQL & " FROM Entradas E, Cli_For C "
        strSQL = strSQL & " Where E.Filial = " & Combo_Filial.Text
        strSQL = strSQL & " AND E.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  E.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
        strSQL = strSQL & " AND E.Operação = " & sOperacao
        
        If Nome_Vendedor.Caption <> "" Then
            strSQL = strSQL & " AND E.Digitador = " & Combo_Vendedor.Text
        End If
        
        If Nome_Cliente.Caption <> "" Then
            strSQL = strSQL & " AND E.Fornecedor = " & Combo_Cliente.Text
        End If
        
        strSQL = strSQL & " AND E.Fornecedor = C.código "

        
        'Ordenação por:
        'DATA HORA SEQUÊNCIA
        'EFETIVADAS
        'DESFEITAS
        'VALOR CRESCENTE
        'VALOR DECRESCENTE
        'CAIXA
        'VENDEDOR
        'CLIENTE/FORNECEDOR
        
        If cbo_ordenacao.Text = "" Or cbo_ordenacao.Text = "DATA HORA SEQUÊNCIA" Or cbo_ordenacao.Text = "DESFEITAS" Then
            strSQL = strSQL & " ORDER BY E.Data, E.Sequência "
        ElseIf cbo_ordenacao.Text = "EFETIVADAS" Then
            strSQL = strSQL & " ORDER BY E.Efetivada, E.Data, E.Sequência "
        ElseIf cbo_ordenacao.Text = "VALOR CRESCENTE" Then
            strSQL = strSQL & " ORDER BY E.total "
        ElseIf cbo_ordenacao.Text = "VALOR DECRESCENTE" Then
            strSQL = strSQL & " ORDER BY E.total DESC "
        ElseIf cbo_ordenacao.Text = "CAIXA" Then
            strSQL = strSQL & " ORDER BY E.Caixa "
        ElseIf cbo_ordenacao.Text = "VENDEDOR" Then
            strSQL = strSQL & " ORDER BY E.Digitador "
        ElseIf cbo_ordenacao.Text = "CLIENTE/FORNECEDOR" Then
            strSQL = strSQL & " ORDER BY E.Fornecedor "
        End If
        
    End If
    
 
    dValorDinheiro = 0
    dValorCartao = 0
    dValorVale = 0
    dValorPrazo = 0
    dValorNaoRecebidoConfere = 0
    dValorOutros = 0
 
    Set rstOperacoesDet = db.OpenRecordset(strSQL, dbOpenDynaset)
 
    With rstOperacoesDet
        If Not (.BOF And .EOF) Then
            .MoveFirst
    
            Do Until .EOF
            
                If sTipoOperacao = "Saída" Then
                    
                    If Not IsNull(.Fields("NSU_Hora").Value) Then
                        sHora = .Fields("NSU_Hora").Value
                    Else
                        sHora = ""
                    End If
                    
                    If chk_formasPagamentoRecebimento.Value = vbChecked Then
                        If Not IsNull(.Fields("Recebe - Num Cartão").Value) Then
                            sRecebeNumCartao = .Fields("Recebe - Num Cartão").Value
                        Else
                            sRecebeNumCartao = ""
                        End If
                    Else
                        sRecebeNumCartao = ""
                    End If
                    sEfetAux = ""
                    sDesfAux = ""
                    If .Fields("Efetivada").Value = True Then
                        sEfetAux = "SIM"
                    Else
                        sEfetAux = "NÃO"
                    End If
                    
                    If .Fields("Movimentação Desfeita").Value = True Then
                        sDesfAux = "SIM"
                    Else
                        sDesfAux = "NÃO"
                    End If
                    
                    If chk_formasPagamentoRecebimento.Value = vbChecked Then
                        If .Fields("Efetivada").Value = True And .Fields("Movimentação Desfeita").Value = False Then
                        
                            If Not IsNull(.Fields("Recebe - Dinheiro").Value) Then
                                dValorDinheiro = dValorDinheiro + .Fields("Recebe - Dinheiro").Value
                                sValorDinheiro = .Fields("Recebe - Dinheiro").Value
                            Else
                                sValorDinheiro = "0"
                            End If
                            If Not IsNull(.Fields("Recebe - Cartão").Value) Then
                                dValorCartao = dValorCartao + .Fields("Recebe - Cartão").Value
                                sValorCartao = .Fields("Recebe - Cartão").Value
                            Else
                                sValorCartao = "0"
                            End If
                            If Not IsNull(.Fields("Recebe - Vale").Value) Then
                                dValorVale = dValorVale + .Fields("Recebe - Vale").Value
                                sValorVale = .Fields("Recebe - Vale").Value
                            Else
                                sValorVale = "0"
                            End If
                            If Not IsNull(.Fields("Total Prazo").Value) Then
                                dValorPrazo = dValorPrazo + .Fields("Total Prazo").Value
                                sValorPrazo = .Fields("Total Prazo").Value
                            Else
                                sValorPrazo = "0"
                            End If
                            
                            sValorOutros = AcharMovimentoCheques(.Fields("Sequência").Value)
                            dValorOutros = dValorOutros + CDbl(sValorOutros)
                            
                            If Not IsNull(.Fields("Valor Recebido").Value) Then
                                If .Fields("Valor Recebido").Value = 0 Then
                                    dValorNaoRecebidoConfere = dValorNaoRecebidoConfere + .Fields("total").Value
                                End If
                            End If
                        End If
                        
                        If sValorDinheiro = "" Then
                          sValorDinheiro = "0"
                        End If
                        If sValorCartao = "" Then
                          sValorCartao = "0"
                        End If
                        If sValorVale = "" Then
                          sValorVale = "0"
                        End If
                        If sValorPrazo = "" Then
                          sValorPrazo = "0"
                        End If
                        If sValorOutros = "" Then
                          sValorOutros = "0"
                        End If
                    
                        gridOperacoesDetalhe.AddItem vbTab & .Fields("Data").Value & vbTab & _
                              sHora & vbTab & _
                              .Fields("Sequência").Value & vbTab & _
                              FormataValorTexto(.Fields("total").Value, 2) & "" & vbTab & _
                              sEfetAux & vbTab & _
                              sDesfAux & vbTab & _
                              .Fields("Caixa").Value & vbTab & _
                              .Fields("Digitador").Value & vbTab & _
                              .Fields("Operador").Value & vbTab & _
                              .Fields("Cliente").Value & vbTab & _
                              .Fields("Nome").Value & vbTab & _
                              FormataValorTexto(sValorDinheiro, 2) & vbTab & _
                              FormataValorTexto(sValorCartao, 2) & vbTab & _
                              FormataValorTexto(sValorVale, 2) & vbTab & _
                              sRecebeNumCartao & vbTab & _
                              FormataValorTexto(sValorPrazo, 2) & vbTab & _
                              FormataValorTexto(sValorOutros, 2)
                    Else
                        gridOperacoesDetalhe.AddItem vbTab & .Fields("Data").Value & vbTab & _
                              sHora & vbTab & _
                              .Fields("Sequência").Value & vbTab & _
                              FormataValorTexto(.Fields("total").Value, 2) & "" & vbTab & _
                              sEfetAux & vbTab & _
                              sDesfAux & vbTab & _
                              .Fields("Caixa").Value & vbTab & _
                              .Fields("Digitador").Value & vbTab & _
                              .Fields("Operador").Value & vbTab & _
                              .Fields("Cliente").Value & vbTab & _
                              .Fields("Nome").Value
                    End If
                
                Else
                    ' TipoOperacao ENTRADA
                    
                    sDesfAux = "NÃO"
                    sEfetAux = ""
                    If .Fields("Efetivada").Value = True Then
                        sEfetAux = "SIM"
                    End If
                    
                    gridOperacoesDetalhe.AddItem vbTab & .Fields("Data").Value & vbTab & _
                          vbTab & _
                          .Fields("Sequência").Value & vbTab & _
                          FormataValorTexto(.Fields("total").Value, 2) & "" & vbTab & _
                          sEfetAux & vbTab & _
                          sDesfAux & vbTab & _
                          .Fields("Caixa").Value & vbTab & _
                          .Fields("Digitador").Value & vbTab & _
                          vbTab & _
                          .Fields("Fornecedor").Value & vbTab & _
                          .Fields("Nome").Value
                End If
                
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstOperacoesDet = Nothing
  End If
  
  lbl_dinheiro.Caption = Format(dValorDinheiro, FORMAT_VALUE)
  lbl_cartao.Caption = Format(dValorCartao, FORMAT_VALUE)
  lbl_vale.Caption = Format(dValorVale, FORMAT_VALUE)
  lbl_prazo.Caption = Format(dValorPrazo, FORMAT_VALUE)
  lbl_valorEfetivadoNaoRecebido.Caption = Format(dValorNaoRecebidoConfere, FORMAT_VALUE)
  lbl_outros.Caption = Format(dValorOutros, FORMAT_VALUE)

  Exit Sub
Erro:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
End Sub

Private Sub cmd_imprimirGradeVendasSeparadasPorDia_Click()
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

  Dim sLinha          As String
  Dim sData_Filial    As String
  Dim sDinheiro       As String
  Dim sCartao         As String
  Dim sPrazo          As String
  Dim sCheque         As String
  Dim sVale           As String
  Dim sNaoFinanc      As String
  Dim sTOTAL          As String
  
  Printer.Font = "LUCIDA CONSOLE"
  
  Printer.Print ""
  sLinha = "                                                                   Quick Store 10 - Soluções Comerciais inteligentes"
  
  Printer.Print ""

  sLinha = "   Relatório de Saídas e Entradas por Período  De: " & Data_Ini & " até " & Data_Fim
  Printer.Print sLinha
  
  Printer.Print ""
  
  If Nome_Cliente.Caption <> "" Then
      sLinha = "   " & Combo_Cliente.Text & " - " & Nome_Cliente.Caption
      Printer.Print sLinha
  End If
 
  If Nome_Vendedor.Caption <> "" Then
      sLinha = "   " & Combo_Vendedor.Text & " - " & Nome_Vendedor.Caption
      Printer.Print sLinha
  End If

  Printer.Print ""
  
'  gridVendasSeparadasPorDia.TextMatrix(0, 0) = ""
'  gridVendasSeparadasPorDia.TextMatrix(0, 1) = "Data/Filial"
'  gridVendasSeparadasPorDia.TextMatrix(0, 2) = "Dinheiro"
'  gridVendasSeparadasPorDia.TextMatrix(0, 3) = "Cartão"
'  gridVendasSeparadasPorDia.TextMatrix(0, 4) = "Prazo"
'  gridVendasSeparadasPorDia.TextMatrix(0, 5) = "Cheque"
'  gridVendasSeparadasPorDia.TextMatrix(0, 6) = "Vale"
'  gridVendasSeparadasPorDia.TextMatrix(0, 7) = "Não financeiro"
'  gridVendasSeparadasPorDia.TextMatrix(0, 8) = "TOTAL"
  
  sLinha = "   Data/Filial          Dinheiro     Cartão       Prazo        Cheque      Vale        Não Financ. TOTAL        "
  Printer.Print sLinha
    
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""
    
  With gridVendasSeparadasPorDia
      For nRow = 1 To .Rows - 1
          ' ************************** ATENÇÃO ***********************************
          ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
          ' De preferência com o mesmo nome da impressora !!!
    
          sData_Filial = gridVendasSeparadasPorDia.TextMatrix(nRow, 1)
              
          sDinheiro = gridVendasSeparadasPorDia.TextMatrix(nRow, 2)
          If Len(sDinheiro) < 13 Then
            For i = Len(sDinheiro) To 12
                sDinheiro = " " & sDinheiro
            Next
          End If

          sCartao = gridVendasSeparadasPorDia.TextMatrix(nRow, 3)
          If Len(sCartao) < 13 Then
            For i = Len(sCartao) To 12
                sCartao = " " & sCartao
            Next
          End If

          sPrazo = gridVendasSeparadasPorDia.TextMatrix(nRow, 4)
          If Len(sPrazo) < 13 Then
            For i = Len(sPrazo) To 12
                sPrazo = " " & sPrazo
            Next
          End If
    
          sCheque = gridVendasSeparadasPorDia.TextMatrix(nRow, 5)
          If Len(sCheque) < 12 Then
            For i = Len(sCheque) To 11
                sCheque = " " & sCheque
            Next
          End If
    
          sVale = gridVendasSeparadasPorDia.TextMatrix(nRow, 6)
          If Len(sVale) < 12 Then
            For i = Len(sVale) To 11
                sVale = " " & sVale
            Next
          End If

          sNaoFinanc = gridVendasSeparadasPorDia.TextMatrix(nRow, 7)
          If Len(sNaoFinanc) < 12 Then
            For i = Len(sNaoFinanc) To 11
                sNaoFinanc = " " & sNaoFinanc
            Next
          End If
          
          sTOTAL = gridVendasSeparadasPorDia.TextMatrix(nRow, 8)
          If Len(sTOTAL) < 13 Then
            For i = Len(sTOTAL) To 12
                sTOTAL = " " & sTOTAL
            Next
          End If
          
          If InStr(1, sData_Filial, "TOTAL VENDAS") > 0 Or InStr(1, sData_Filial, "TOTAL GERAL") > 0 Then
              If Len(sData_Filial) < 80 Then
                For i = Len(sData_Filial) To 79
                    sData_Filial = sData_Filial & " "
                Next
              End If
              
              sTOTAL = gridVendasSeparadasPorDia.TextMatrix(nRow, 8)
              If Len(sTOTAL) < 23 Then
                For i = Len(sTOTAL) To 22
                    sTOTAL = " " & sTOTAL
                Next
              End If
              
              sLinha = sData_Filial
              sLinha = sLinha & sTOTAL
              
          Else
              sLinha = sData_Filial
              sLinha = sLinha & sDinheiro
              sLinha = sLinha & sCartao
              sLinha = sLinha & sPrazo
              sLinha = sLinha & sCheque
              sLinha = sLinha & sVale
              sLinha = sLinha & sNaoFinanc
              sLinha = sLinha & sTOTAL
          End If
    
          
          If InStr(1, sLinha, "OBS:") = 0 Then
              Printer.Print "   " & sLinha
          Else
              Printer.Print ""
              Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
              Printer.Print "   *OBS: As devoluções de produtos (caso elas existam) são desconsideradas dos cálculos"
          End If
          Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      Next nRow
  End With
      
  Printer.Print ""

  Printer.EndDoc
  
  Exit Sub
Erro:
    MsgBox "Erro na impressão da grade " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_imprimirOper_Click()
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
  Dim sValorEfetivadas As String
  Dim sValorDesfeitas As String
  Dim sTipoOper As String
  Dim sOperacao As String
  Dim sOperacaoNome As String
  
  Printer.Font = "LUCIDA CONSOLE"
  
  Printer.Print ""
  sLinha = "                                                                   Quick Store 10 - Soluções Comerciais inteligentes"
  
  Printer.Print ""

  sLinha = "   Relatório de Saídas e Entradas por Período  De: " & Data_Ini & " até " & Data_Fim
  Printer.Print sLinha
  
  Printer.Print ""
  
  sLinha = "   " & Combo_Filial.Text & " - " & Nome_Empresa.Caption
  Printer.Print sLinha
  
  If Nome_Cliente.Caption <> "" Then
      sLinha = "   " & Combo_Cliente.Text & " - " & Nome_Cliente.Caption
      Printer.Print sLinha
  End If
 
  If Nome_Vendedor.Caption <> "" Then
      sLinha = "   " & Combo_Vendedor.Text & " - " & Nome_Vendedor.Caption
      Printer.Print sLinha
  End If
  

  Printer.Print ""

  sLinha = "   Tipo Oper.  Operação  Nome                                          Total Efetivadas      Total Desfeitas"
  Printer.Print sLinha

  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""

  With gridOperacoes
      For nRow = 1 To .Rows - 1
          ' ************************** ATENÇÃO ***********************************
          ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
          ' De preferência com o mesmo nome da impressora !!!

          sTipoOper = gridOperacoes.TextMatrix(nRow, 1)
          If sTipoOper = "Saída" Then
                sTipoOper = sTipoOper & "     "
          Else
                ' Entrada
                sTipoOper = sTipoOper & "   "
          End If
          
          sOperacao = gridOperacoes.TextMatrix(nRow, 2)
          If Len(sOperacao) < 8 Then
            For i = Len(sOperacao) To 7
                sOperacao = " " & sOperacao
            Next
          End If
          
          sOperacaoNome = gridOperacoes.TextMatrix(nRow, 3)
          If Len(sOperacaoNome) < 44 Then
            For i = Len(sOperacaoNome) To 43
                sOperacaoNome = sOperacaoNome & " "
            Next
          Else
              sOperacaoNome = Mid(sOperacaoNome, 1, 44)
          End If

          sValorEfetivadas = gridOperacoes.TextMatrix(nRow, 4)
          If Len(sValorEfetivadas) < 16 Then
            For i = Len(sValorEfetivadas) To 15
                sValorEfetivadas = " " & sValorEfetivadas
            Next
          End If

          sValorDesfeitas = gridOperacoes.TextMatrix(nRow, 5)
          If Len(sValorDesfeitas) < 15 Then
            For i = Len(sValorDesfeitas) To 14
                sValorDesfeitas = " " & sValorDesfeitas
            Next
          End If

          sLinha = sTipoOper
          sLinha = sLinha & "  " & sOperacao
          sLinha = sLinha & "  " & sOperacaoNome
          sLinha = sLinha & "  " & sValorEfetivadas
          sLinha = sLinha & "      " & sValorDesfeitas
          
          Printer.Print "   " & sLinha
          Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      Next nRow
  End With
  
  Printer.Print ""

  sValorEfetivadas = lbl_totalSaidas.Caption
  If Len(sValorEfetivadas) < 15 Then
    For i = Len(sValorEfetivadas) To 14
        sValorEfetivadas = " " & sValorEfetivadas
    Next
  End If
  
  sValorDesfeitas = lbl_totalSaidasDesf.Caption
  If Len(sValorDesfeitas) < 15 Then
    For i = Len(sValorDesfeitas) To 14
        sValorDesfeitas = " " & sValorDesfeitas
    Next
  End If
  
  Dim sValor As String
  sValor = lbl_totaEntradas.Caption
  If Len(sValor) < 15 Then
    For i = Len(sValor) To 14
        sValor = " " & sValor
    Next
  End If

  sLinha = "   Total Saídas Efetivadas R$ " & sValorEfetivadas
  Printer.Print sLinha
  sLinha = "   Total Saídas Desfeitas  R$ " & sValorDesfeitas
  Printer.Print sLinha
  sLinha = "   Total Entradas          R$ " & sValor
  Printer.Print sLinha

  Printer.EndDoc
  
  Exit Sub
Erro:
    MsgBox "Erro na impressão da grade " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_imprimirOperDet_Click()
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
  Dim sData As String
  Dim sHora As String
  Dim sSequencia As String
  Dim sValor As String
  Dim sEfetivada As String
  Dim sDesfeita As String
  Dim sCaixa As String
  Dim sVendedor As String
  Dim sOperador As String
  Dim sCliForn As String
  Dim sValorTotalEfetivada As String
  Dim sValorTotalDesfeita As String
  Dim sOperacao As String
  Dim sTipoOperacao As String
  Dim iConta As Integer
  Dim sCodigoProduto As String
  Dim sNomeProduto As String
  Dim sTamanho As String
  Dim sCor As String
  Dim sQuantidadeProduto As String
  
  
  Printer.Font = "LUCIDA CONSOLE"
  
  Printer.Print ""
  sLinha = "                                                                   Quick Store 10 - Soluções Comerciais inteligentes"
  
  Printer.Print ""

  sLinha = "   Relatório de Saídas e Entradas por Período  De: " & Data_Ini & " até " & Data_Fim
  Printer.Print sLinha
  
  Printer.Print ""
  
  sLinha = "   " & Combo_Filial.Text & " - " & Nome_Empresa.Caption
  Printer.Print sLinha
  
  If Nome_Cliente.Caption <> "" Then
      sLinha = "   " & Combo_Cliente.Text & " - " & Nome_Cliente.Caption
      Printer.Print sLinha
  End If
 
  If Nome_Vendedor.Caption <> "" Then
      sLinha = "   " & Combo_Vendedor.Text & " - " & Nome_Vendedor.Caption
      Printer.Print sLinha
  End If

  Printer.Print ""
  
  If chk_visaoConsolidada.Value = vbChecked Then
      For iConta = 1 To gridOperacoes.Rows - 1
        
        If iConta > 1 Then
            sOperacao = sOperacao & ","
        End If
        sOperacao = sOperacao & gridOperacoes.TextMatrix(iConta, 2)
      Next
      sLinha = "   Operação      : " & sOperacao
      Printer.Print sLinha
  Else
      If gridOperacoes.RowSel > 0 Then
          sOperacao = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 2)
          sTipoOperacao = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 1)
          sValorTotalEfetivada = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 4)
          sValorTotalDesfeita = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 5)
          sLinha = "   Operação      : " & sOperacao
          Printer.Print sLinha
          sLinha = "   Tipo Operação : " & sTipoOperacao
          Printer.Print sLinha
      End If
  End If
      
  Printer.Print ""
    
  If chk_visaoProdutos.Value = vbUnchecked Then
      ' ******************************************
      ' Visão sequência de vendas/compras...
    
      sLinha = "   Data        Hora      Sequência    Valor           Efetivada  Desfeita  Caixa  Vendedor  Operador  Cli/Forn  "
      Printer.Print sLinha
    
      Printer.Print "   _________________________________________________________________________________________________________________"
      Printer.Print ""
    
      With gridOperacoesDetalhe
          For nRow = 1 To .Rows - 1
              ' ************************** ATENÇÃO ***********************************
              ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
              ' De preferência com o mesmo nome da impressora !!!
    
              sData = gridOperacoesDetalhe.TextMatrix(nRow, 1)
              sHora = gridOperacoesDetalhe.TextMatrix(nRow, 2)
              
              If Len(sHora) < 8 Then
                For i = Len(sHora) To 7
                    sHora = sHora & " "
                Next
              End If
    
              sSequencia = gridOperacoesDetalhe.TextMatrix(nRow, 3)
              If Len(sSequencia) < 10 Then
                For i = Len(sSequencia) To 9
                    sSequencia = sSequencia & " "
                Next
              End If
    
              sValor = gridOperacoesDetalhe.TextMatrix(nRow, 4)
              If Len(sValor) < 16 Then
                For i = Len(sValor) To 15
                    sValor = " " & sValor
                Next
              End If
    
              sEfetivada = gridOperacoesDetalhe.TextMatrix(nRow, 5)
              sEfetivada = sEfetivada & "      "
    
              sDesfeita = gridOperacoesDetalhe.TextMatrix(nRow, 6)
              sDesfeita = sDesfeita & "     "
    
              sCaixa = gridOperacoesDetalhe.TextMatrix(nRow, 7)
              If Len(sCaixa) < 5 Then
                For i = Len(sCaixa) To 4
                    sCaixa = " " & sCaixa
                Next
              End If
              
              sVendedor = gridOperacoesDetalhe.TextMatrix(nRow, 8)
              If Len(sVendedor) < 8 Then
                For i = Len(sVendedor) To 7
                    sVendedor = " " & sVendedor
                Next
              End If
              
              sOperador = gridOperacoesDetalhe.TextMatrix(nRow, 9)
              If Len(sOperador) < 8 Then
                For i = Len(sOperador) To 7
                    sOperador = " " & sOperador
                Next
              End If
    
              sCliForn = gridOperacoesDetalhe.TextMatrix(nRow, 10)
              If Len(sCliForn) < 10 Then
                For i = Len(sCliForn) To 9
                    sCliForn = " " & sCliForn
                Next
              End If
    
              sLinha = sData
              sLinha = sLinha & "  " & sHora
              sLinha = sLinha & "  " & sSequencia
              sLinha = sLinha & "  " & sValor
              sLinha = sLinha & "  " & sEfetivada
              sLinha = sLinha & "  " & sDesfeita
              sLinha = sLinha & "  " & sCaixa
              sLinha = sLinha & "  " & sVendedor
              sLinha = sLinha & "  " & sOperador
              sLinha = sLinha & "  " & sCliForn
              
              Printer.Print "   " & sLinha
              Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
          Next nRow
      End With
      
      Printer.Print ""
    
      If Len(sValorTotalEfetivada) < 15 Then
        For i = Len(sValorTotalEfetivada) To 14
            sValorTotalEfetivada = " " & sValorTotalEfetivada
        Next
      End If
      
      If Len(sValorTotalDesfeita) < 15 Then
        For i = Len(sValorTotalDesfeita) To 14
            sValorTotalDesfeita = " " & sValorTotalDesfeita
        Next
      End If
      
      
      If chk_visaoConsolidada.Value = vbChecked Then
          sLinha = "   Total Efetivadas R$ " & lbl_totalSaidas.Caption
          Printer.Print sLinha
          sLinha = "   Total Desfeitas  R$ " & lbl_totalSaidasDesf.Caption
          Printer.Print sLinha
      Else
          sLinha = "   Total Efetivadas R$ " & sValorTotalEfetivada
          Printer.Print sLinha
          sLinha = "   Total Desfeitas  R$ " & sValorTotalDesfeita
          Printer.Print sLinha
      End If
    
      If chk_formasPagamentoRecebimento.Value = vbChecked Then
        Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
        Printer.Print "   Dinheiro : " & lbl_dinheiro.Caption
        Printer.Print "   Cartão   : " & lbl_cartao.Caption
        Printer.Print "   Vale     : " & lbl_vale.Caption
        Printer.Print "   Prazo    : " & lbl_prazo.Caption
        Printer.Print "   Outros   : " & lbl_outros.Caption
        Printer.Print "   *Efetivado e não recebido  : " & lbl_valorEfetivadoNaoRecebido.Caption
        Printer.Print "   *Transações 'NÃO EFETIVADAS' e 'DESFEITAS' foram desconsideradas dos cálculos detalhados"
        Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      End If
  Else
      ' ******************************************
      ' Visão Produtos
      sLinha = "   Código               Nome Produto                             Tamanho         Cor             Quantidade Mov."
      Printer.Print sLinha
    
      Printer.Print "   _________________________________________________________________________________________________________________"
      Printer.Print ""
    
      With gridOperacoesProdutos
          For nRow = 1 To .Rows - 1
              ' ************************** ATENÇÃO ***********************************
              ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
              ' De preferência com o mesmo nome da impressora !!!
    
              sCodigoProduto = gridOperacoesProdutos.TextMatrix(nRow, 1)
              If Len(sCodigoProduto) < 20 Then
                For i = Len(sCodigoProduto) To 19
                    sCodigoProduto = sCodigoProduto & " "
                Next
              End If
              
              sNomeProduto = gridOperacoesProdutos.TextMatrix(nRow, 2)
              If Len(sNomeProduto) < 40 Then
                For i = Len(sNomeProduto) To 39
                    sNomeProduto = sNomeProduto & " "
                Next
              ElseIf Len(sNomeProduto) > 40 Then
                sNomeProduto = Mid(sNomeProduto, 1, 40)
              End If
    
              sTamanho = gridOperacoesProdutos.TextMatrix(nRow, 3)
              If Len(sTamanho) < 15 Then
                For i = Len(sTamanho) To 14
                    sTamanho = sTamanho & " "
                Next
              ElseIf Len(sTamanho) > 15 Then
                sTamanho = Mid(sTamanho, 1, 15)
              End If
    
              sCor = gridOperacoesProdutos.TextMatrix(nRow, 4)
              If Len(sCor) < 15 Then
                For i = Len(sCor) To 14
                    sCor = sCor & " "
                Next
              ElseIf Len(sCor) > 15 Then
                sCor = Mid(sCor, 1, 15)
              End If
              
              sQuantidadeProduto = gridOperacoesProdutos.TextMatrix(nRow, 5)
              If Len(sQuantidadeProduto) < 15 Then
                For i = Len(sQuantidadeProduto) To 14
                    sQuantidadeProduto = " " & sQuantidadeProduto
                Next
              End If
              
              sLinha = sData
              sLinha = sLinha & sCodigoProduto
              sLinha = sLinha & " " & sNomeProduto
              sLinha = sLinha & " " & sTamanho
              sLinha = sLinha & " " & sCor
              sLinha = sLinha & " " & sQuantidadeProduto
              
              Printer.Print "   " & sLinha
              Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
          Next nRow
      End With
  
      ' Listar os produtos que foram devolvidos no período e que o movimento foi estornado neste relatório
      If chk_desconsiderar_itensDoProdutoDevolvido.Value = vbUnchecked Then
          If numArrProdutosDevolvidos > 0 Then
              Dim iContador As Long
              Dim X As Long
              Dim sProdDev As String
              Printer.Print " "
              Printer.Print " "
              Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
              Printer.Print "   Atenção: "
              Printer.Print "            Para os critérios de pesquisa que você selecionou acima, abaixo está a lista dos produtos que foram "
              Printer.Print "        devolvidos por clientes no período relacionado e consequentemente o movimento foi estornado nestes produtos "
              Printer.Print "        listados acima. (Se existir)."
              Printer.Print " "
              Printer.Print "   Produto                Quantidade Devolvida"
              Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
              For iContador = 0 To numArrProdutosDevolvidos - 1
                  sProdDev = arrProdutosDevolvidos(iContador, 0)
                  If Len(sProdDev) < 20 Then
                      For X = Len(sProdDev) To 19
                          sProdDev = sProdDev & " "
                      Next
                  End If
                  Printer.Print "   " & sProdDev & "     " & arrProdutosDevolvidos(iContador, 1)
              Next
              Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
          End If
      End If
  End If

  Printer.EndDoc
  
  Exit Sub
Erro:
    MsgBox "Erro na impressão da grade " & Err.Description, vbInformation, "Atenção"
End Sub


Private Sub cmd_limparTela_Click()

  Combo_Filial_LostFocus
  Data_Ini.Text = Format(Date, "DD/MM/YYYY")
  Data_Fim.Text = Format(Date, "DD/MM/YYYY")
  Combo_Vendedor.Text = ""
  Combo_Vendedor_LostFocus
  Combo_Cliente.Text = ""
  Combo_Cliente_LostFocus
  

  cmb_classificacaoSaidas.Enabled = False
  chk_visaoConsolidada.Value = vbUnchecked
  chk_formasPagamentoRecebimento.Value = vbUnchecked
  chk_visaoProdutos.Value = vbUnchecked
  cmb_tipoOperacao.ListIndex = -1
  
  gridOperacoes.Rows = 1
  gridOperacoesDetalhe.Rows = 1
  gridOperacoesProdutos.Rows = 1
  gridOperacoesProdutos.Visible = False
  lbl_totalSaidas.Caption = "0,00"
  lbl_totalSaidasDesf.Caption = "0,00"
  lbl_totaEntradas.Caption = "0,00"
  
  cbo_ordenacao.ListIndex = -1
  
  lbl_dinheiroTit.Visible = False
  lbl_dinheiro.Visible = False
  lbl_cartaoTit.Visible = False
  lbl_cartao.Visible = False
  lbl_valeTit.Visible = False
  lbl_vale.Visible = False
  lbl_PrazoTit.Visible = False
  lbl_prazo.Visible = False
  lbl_ValorNaoRecTit.Visible = False
  lbl_valorEfetivadoNaoRecebido.Visible = False
  lbl_avisoCalculos.Visible = False
  lbl_outros.Visible = False
  lbl_outrosTit.Visible = False
  
  chk_produtoGradeAgrupado.Value = vbUnchecked
  Combo_Classe.Text = ""
  Nome_Classe.Caption = ""
  Combo_SubClasse.Text = ""
  Nome_SubClasse.Caption = ""
  
  chk_vendasSeparadasPorDia.Value = vbUnchecked
  chk_vendasSeparadasPorDia.Visible = False
  
  chk_desconsiderar_itensDoProdutoDevolvido.Value = vbUnchecked
  chk_desconsiderar_itensDoProdutoDevolvido.Visible = False
  
  frm_vendasSeparadasPorDia.Visible = False
  
  lbl_CodOperacoesVenda.Visible = False
  txt_CodOperacoesVenda.Text = ""
  txt_CodOperacoesVenda.Visible = False
  lbl_CodOperacoesVendaEXEMPLO.Visible = False
  
End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo Erro

  Dim Erro As Boolean
  
  If Nome_Empresa.Caption = "" And chk_vendasSeparadasPorDia.Value = vbUnchecked And chk_visaoProdutos.Value = vbUnchecked Then
    DisplayMsg "Selecione uma Filial."
    Combo_Filial.SetFocus
    Exit Sub
  End If

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
    DisplayMsg "Escolha um período no máximo de 90 dias"
    Data_Fim.SetFocus
    Exit Sub
  End If

  Call StatusMsg("Pesquisando operações...")
  Screen.MousePointer = vbHourglass
  Call PesquisarOperacoes
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")

  Exit Sub
Erro:
  MsgBox "Erro na pesquisa de operações " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub PesquisarOperacoes()
On Error GoTo Erro
  
  Dim sSQL As String
  Dim rsOperacoesAux As Recordset
  Dim dValorEfetivada As Double
  Dim dValorDesfeita As Double
  Dim dTotalSaidas As Double
  Dim dTotalSaidasDesf As Double
  Dim dTotalEntradas As Double
  
  gridOperacoes.Rows = 1
  gridOperacoesDetalhe.Rows = 1
  lbl_totalSaidas.Caption = Format(0, FORMAT_VALUE)
  lbl_totalSaidasDesf.Caption = Format(0, FORMAT_VALUE)
  lbl_totaEntradas.Caption = Format(0, FORMAT_VALUE)
  cbo_ordenacao.ListIndex = 0
  
  LimpaArray
  
  
  ' ==========================================================================================
  ' Tratamento Operações de Saída
  If (cmb_tipoOperacao.ListIndex = 1 Or cmb_tipoOperacao.ListIndex = 0 Or cmb_tipoOperacao.ListIndex = -1) And chk_vendasSeparadasPorDia.Value = vbUnchecked Then
      sSQL = "Select S.Operação, O.Nome, S.Efetivada, S.[Movimentação Desfeita], sum(total) "
      sSQL = sSQL & " From Saídas S, [Operações Saída] O"
      
      sSQL = sSQL & " Where "
      
      If Trim(Combo_Filial.Text) <> "" Then
          sSQL = sSQL & " S.Filial= " & Combo_Filial.Text & " AND "
      End If
      
      sSQL = sSQL & " S.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  S.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
      
      If Nome_Vendedor.Caption <> "" Then
          sSQL = sSQL & " AND S.Digitador = " & Combo_Vendedor.Text
      End If
      
      If Nome_Cliente.Caption <> "" Then
          sSQL = sSQL & " AND S.Cliente = " & Combo_Cliente.Text
      End If
     
      If Nome_Vendedor.Caption <> "" Then
          sSQL = sSQL & " AND S.Digitador = " & Combo_Vendedor.Text
      End If
      
      If Trim(txt_CodOperacoesVenda.Text) <> "" Then
          sSQL = sSQL & " AND S.Operação in(" & Trim(txt_CodOperacoesVenda.Text) & ") "
      End If
      
      sSQL = sSQL & " AND S.Operação = O.Código "
      
      'Operação de Saída - Classificação:
      'Venda
      'Transferência de Saída
      'Ajuste de Saída
      'Grátis Saída / Devolução
      'Empréstimo Saída
      'Orçamento
      If cmb_classificacaoSaidas.ListIndex > -1 Then
          If cmb_classificacaoSaidas.ListIndex = 1 Then
              sSQL = sSQL & " AND O.Tipo = 'V' "
          ElseIf cmb_classificacaoSaidas.ListIndex = 2 Then
              sSQL = sSQL & " AND O.Tipo = 'T' "
          ElseIf cmb_classificacaoSaidas.ListIndex = 3 Then
              sSQL = sSQL & " AND O.Tipo = 'A' "
          ElseIf cmb_classificacaoSaidas.ListIndex = 4 Then
              sSQL = sSQL & " AND O.Tipo = 'G' "
          ElseIf cmb_classificacaoSaidas.ListIndex = 5 Then
              sSQL = sSQL & " AND O.Tipo = 'E' "
          ElseIf cmb_classificacaoSaidas.ListIndex = 6 Then
              sSQL = sSQL & " AND O.Tipo = 'O' "
          End If
      End If
      
      sSQL = sSQL & " Group by  S.Operação, O.Nome, S.Efetivada, S.[Movimentação Desfeita]"
      
      Set rsOperacoesAux = db.OpenRecordset(sSQL, dbOpenDynaset)
     
      With rsOperacoesAux
        If Not (.BOF And .EOF) Then
          .MoveFirst
    
          Do Until .EOF
              dValorEfetivada = 0
              dValorDesfeita = 0
    
              If .Fields(2).Value = True And .Fields(3).Value = False Then
                  dValorEfetivada = .Fields(4).Value
                  MontaArray 1, .Fields("Operação").Value, .Fields("Nome").Value, "Saída", dValorEfetivada, 0
              ElseIf .Fields(2).Value = True And .Fields(3).Value = True Then
                  dValorDesfeita = .Fields(4).Value
                  MontaArray 2, .Fields("Operação").Value, .Fields("Nome").Value, "Saída", 0, dValorDesfeita
              End If
        
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsOperacoesAux = Nothing
     
      Dim indiceFor As Integer
      For indiceFor = 0 To 99
          If arrOperacao(indiceFor, 0) <> "" Then
              gridOperacoes.AddItem vbTab & arrOperacao(indiceFor, 2) & vbTab & arrOperacao(indiceFor, 0) & vbTab & _
                    arrOperacao(indiceFor, 1) & vbTab & _
                    FormataValorTexto(arrOperacao(indiceFor, 3), 2) & vbTab & _
                    FormataValorTexto(arrOperacao(indiceFor, 4), 2) & vbTab
                    
                    dTotalSaidas = dTotalSaidas + arrOperacao(indiceFor, 3)
                    dTotalSaidasDesf = dTotalSaidasDesf + arrOperacao(indiceFor, 4)
                    
          ElseIf arrOperacao(indiceFor, 0) = "" Then
              Exit For
          End If
      Next
      
      If arrOperacao(0, 0) = "" Then
          gridOperacoesDetalhe.Rows = 1
          gridOperacoesProdutos.Rows = 1
          
          Exit Sub
      Else
          If chk_visaoConsolidada.Value = vbChecked And chk_visaoProdutos.Value = vbUnchecked And chk_vendasSeparadasPorDia.Value = vbUnchecked Then
              DetalharOperacaoAgrupadaDeOperacoes
          ElseIf chk_visaoConsolidada.Value = vbChecked And chk_visaoProdutos.Value = vbChecked Then
              DetalharOperacaoAgrupadaDeProdutos
          ElseIf chk_visaoConsolidada.Value = vbChecked And chk_vendasSeparadasPorDia.Value = vbChecked And chk_visaoProdutos.Value = vbUnchecked Then
              DetalharOperacaoAgrupadaDeOperacoes_separadasPorDia
          End If
       End If

  ElseIf chk_visaoConsolidada.Value = vbChecked And chk_vendasSeparadasPorDia.Value = vbChecked And chk_visaoProdutos.Value = vbUnchecked Then
        DetalharOperacaoAgrupadaDeOperacoes_separadasPorDia
  End If
  ' ==========================================================================================
  
  ' ==========================================================================================
  ' Tratamento Operações de Entrada
  If cmb_tipoOperacao.ListIndex = 2 Or cmb_tipoOperacao.ListIndex = 0 Or cmb_tipoOperacao.ListIndex = -1 Then
      sSQL = "Select E.Operação, O.Nome, E.Efetivada,  sum(total) "
      sSQL = sSQL & " From Entradas E, [Operações Entrada] O"
      
      sSQL = sSQL & " Where "
      
      If Trim(Combo_Filial.Text) <> "" Then
          sSQL = sSQL & " E.Filial= " & Combo_Filial.Text & " AND "
      End If
      
      sSQL = sSQL & " E.Data >= #" & Format(Data_Ini, "mm/dd/yyyy") & "# AND  E.Data <= #" & Format(Data_Fim, "mm/dd/yyyy") & " 23:59:59# "
      
      If Nome_Vendedor.Caption <> "" Then
          sSQL = sSQL & " AND E.Digitador = " & Combo_Vendedor.Text
      End If
      
      If Nome_Cliente.Caption <> "" Then
          sSQL = sSQL & " AND E.Fornecedor = " & Combo_Cliente.Text
      End If
     
      sSQL = sSQL & " AND E.Operação = O.Código "
      sSQL = sSQL & " Group by  E.Operação, O.Nome, E.Efetivada "
      
      Set rsOperacoesAux = db.OpenRecordset(sSQL, dbOpenDynaset)
     
      With rsOperacoesAux
        If Not (.BOF And .EOF) Then
          .MoveFirst
    
          Do Until .EOF
              dValorEfetivada = 0
    
              ' Operação de Entrada não tem Desfeita (pois a opção desfazer EXCLUI o registro de entrada)
    
              dValorEfetivada = .Fields(3).Value
              MontaArray 1, .Fields("Operação").Value, .Fields("Nome").Value, "Entrada", dValorEfetivada, 0
       
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsOperacoesAux = Nothing
     
      For indiceFor = 0 To 99
          If arrOperacao(indiceFor, 0) <> "" And arrOperacao(indiceFor, 2) = "Entrada" Then
              gridOperacoes.AddItem vbTab & arrOperacao(indiceFor, 2) & vbTab & arrOperacao(indiceFor, 0) & vbTab & _
                    arrOperacao(indiceFor, 1) & vbTab & _
                    FormataValorTexto(arrOperacao(indiceFor, 3), 2) & vbTab & _
                    "0,00" & vbTab & _
                    FormataValorTexto(arrOperacao(indiceFor, 3), 2)
                    
                    dTotalEntradas = dTotalEntradas + arrOperacao(indiceFor, 3)
          ElseIf arrOperacao(indiceFor, 0) = "" Then
              Exit For
          End If
      Next
  End If
  ' ==========================================================================================
  
  
  lbl_totalSaidas.Caption = Format(dTotalSaidas, FORMAT_VALUE)
  lbl_totalSaidasDesf.Caption = Format(dTotalSaidasDesf, FORMAT_VALUE)
  lbl_totaEntradas.Caption = Format(dTotalEntradas, FORMAT_VALUE)
  
 
  Exit Sub
Erro:
  MsgBox "Erro na pesquisa de operações " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

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

Private Sub Combo_Classe_CloseUp()
  Combo_Classe.Text = Combo_Classe.Columns(1).Text
  Nome_Classe.Caption = Combo_Classe.Columns(0).Text
  Combo_Classe_LostFocus
End Sub

Private Sub Combo_Classe_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Classe_LostFocus()
  Call StatusMsg("")
 
  Nome_Classe.Caption = ""
  If IsNull(Combo_Classe.Text) Then Exit Sub
  If Combo_Classe.Text = "" Then Exit Sub
  If Combo_Classe.Text = "0" Then Exit Sub
  If Not IsNumeric(Combo_Classe.Text) Then Exit Sub
  
  rsClasses.Index = "Código"
   
  If rsClasses.RecordCount > 0 Then
    rsClasses.MoveFirst
  End If
  
  rsClasses.Seek "=", Combo_Classe.Text
  
  If Not rsClasses.NoMatch Then
    Nome_Classe.Caption = rsClasses("Nome")
  Else
    Nome_Classe.Caption = ""
  End If
  
End Sub

Private Sub Combo_SubClasse_CloseUp()
  Combo_SubClasse.Text = Combo_SubClasse.Columns(1).Text
  Nome_SubClasse.Caption = Combo_SubClasse.Columns(0).Text
  Combo_SubClasse_LostFocus
End Sub

Private Sub Combo_SubClasse_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_SubClasse_LostFocus()

  Call StatusMsg("")
 
  Nome_SubClasse.Caption = ""
  If IsNull(Combo_SubClasse.Text) Then Exit Sub
  If Combo_SubClasse.Text = "" Then Exit Sub
  If Combo_SubClasse.Text = "0" Then Exit Sub
  If Not IsNumeric(Combo_SubClasse.Text) Then Exit Sub

  rsSubclasses.Index = "Código"
   
  If rsSubclasses.RecordCount > 0 Then
    rsSubclasses.MoveFirst
  End If
  
  rsSubclasses.Seek "=", Combo_SubClasse.Text
  
  If Not rsSubclasses.NoMatch Then
    Nome_SubClasse.Caption = rsSubclasses("Nome")
  Else
    Nome_SubClasse.Caption = ""
  End If

End Sub

Private Sub Combo_Cliente_CloseUp()
  Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
  Combo_Cliente_LostFocus
End Sub

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

Private Sub Combo_Filial_CloseUp()
  Combo_Filial.Text = Combo_Filial.Columns(1).Text
  Combo_Filial_LostFocus
End Sub

Private Sub Combo_Filial_LostFocus()

  If Filial_Liberada <> 0 Then
    If Val(Combo_Filial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      cmdPesquisar.Enabled = False
      Exit Sub
    End If
  End If
  
  cmdPesquisar.Enabled = True
  
  Nome_Empresa.Caption = ""
  If IsNull(Combo_Filial.Text) Then Exit Sub
  If Not IsNumeric(Combo_Filial.Text) Then Exit Sub
  If Val(Combo_Filial.Text) > 99 Then Exit Sub
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Filial.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")
End Sub

Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(1).Text
  Combo_Vendedor_LostFocus
End Sub

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

Private Sub cmd_abreSequencia_Click()
On Error GoTo Erro

  Dim sTipoOperacao As String

  If gridOperacoes.RowSel > 0 Then
    sTipoOperacao = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 1)
    
    If sTipoOperacao = "Saída" Then
        If gridOperacoesDetalhe.RowSel > 0 Then
            Dim objSaidas As frmSaidas
            Set objSaidas = New frmSaidas
            
            objSaidas.txtSeq = gridOperacoesDetalhe.TextMatrix(gridOperacoesDetalhe.RowSel, 3)
            objSaidas.SearchRecord_peloNumSeq (gridOperacoesDetalhe.TextMatrix(gridOperacoesDetalhe.RowSel, 3))
            objSaidas.Show
            
            Set objSaidas = Nothing
        Else
            MsgBox "selecione uma Sequência na grade", vbInformation, "Atenção"
            Exit Sub
        End If
    Else
        If gridOperacoesDetalhe.RowSel > 0 Then
            Dim objEntradas As frmEntrada
            Set objEntradas = New frmEntrada
            
            objEntradas.txtSeq = gridOperacoesDetalhe.TextMatrix(gridOperacoesDetalhe.RowSel, 3)
            objEntradas.SearchRecord_porSequencia (gridOperacoesDetalhe.TextMatrix(gridOperacoesDetalhe.RowSel, 3))
            objEntradas.Show
            
            Set objEntradas = Nothing
        Else
            MsgBox "selecione uma Sequência na grade", vbInformation, "Atenção"
            Exit Sub
        End If
    
    End If
    
  Else
    MsgBox "selecione uma Sequência na grade", vbInformation, "Atenção"
    Exit Sub
    
  End If
    
  Exit Sub
  
Erro:
  MsgBox "Erro no detalhamento da Sequência" & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Form_Load()
On Error GoTo Erro

  Data1.DatabaseName = gsQuickDBFileName
  dtaVendedor.DatabaseName = gsQuickDBFileName
  dtaCliente.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
  
  
  Data_Ini.Text = Format(Date, "DD/MM/YYYY")
  Data_Fim.Text = Format(Date, "DD/MM/YYYY")

  gridOperacoes.ColWidth(0) = 0
  gridOperacoes.ColWidth(1) = 1700
  gridOperacoes.ColWidth(2) = 1200
  gridOperacoes.ColWidth(3) = 8500
  gridOperacoes.ColWidth(4) = 1800
  gridOperacoes.ColWidth(5) = 1800
  
  gridOperacoes.Row = 0
  gridOperacoes.TextMatrix(0, 0) = ""
  gridOperacoes.TextMatrix(0, 1) = "Tipo da Operação"
  gridOperacoes.TextMatrix(0, 2) = "Operação"
  gridOperacoes.TextMatrix(0, 3) = "Nome"
  gridOperacoes.TextMatrix(0, 4) = "Total Efetivadas"
  gridOperacoes.TextMatrix(0, 5) = "Total Desfeitas"
  
  
  gridOperacoesDetalhe.ColWidth(0) = 0
  gridOperacoesDetalhe.ColWidth(1) = 1100
  gridOperacoesDetalhe.ColWidth(2) = 950
  gridOperacoesDetalhe.ColWidth(3) = 1200
  gridOperacoesDetalhe.ColWidth(4) = 1300
  gridOperacoesDetalhe.ColWidth(5) = 1000
  gridOperacoesDetalhe.ColWidth(6) = 1000
  gridOperacoesDetalhe.ColWidth(7) = 600
  gridOperacoesDetalhe.ColWidth(8) = 900
  gridOperacoesDetalhe.ColWidth(9) = 900
  gridOperacoesDetalhe.ColWidth(10) = 1000
  gridOperacoesDetalhe.ColWidth(11) = 5000
  
  gridOperacoesDetalhe.Row = 0
  gridOperacoesDetalhe.TextMatrix(0, 0) = ""
  gridOperacoesDetalhe.TextMatrix(0, 1) = "Data"
  gridOperacoesDetalhe.TextMatrix(0, 2) = "Hora"
  gridOperacoesDetalhe.TextMatrix(0, 3) = "Sequência"
  gridOperacoesDetalhe.TextMatrix(0, 4) = "Valor"
  gridOperacoesDetalhe.TextMatrix(0, 5) = "Efetivada"
  gridOperacoesDetalhe.TextMatrix(0, 6) = "Desfeita"
  gridOperacoesDetalhe.TextMatrix(0, 7) = "Caixa"
  gridOperacoesDetalhe.TextMatrix(0, 8) = "Vendedor"
  gridOperacoesDetalhe.TextMatrix(0, 9) = "Operador"
  gridOperacoesDetalhe.TextMatrix(0, 10) = "Cli/Forn"
  gridOperacoesDetalhe.TextMatrix(0, 11) = "Nome"
  
  
  gridOperacoesProdutos.ColWidth(0) = 0
  gridOperacoesProdutos.ColWidth(1) = 2200
  gridOperacoesProdutos.ColWidth(2) = 5840
  gridOperacoesProdutos.ColWidth(3) = 2200
  gridOperacoesProdutos.ColWidth(4) = 2200
  gridOperacoesProdutos.ColWidth(5) = 2500
  
  gridOperacoesProdutos.Row = 0
  gridOperacoesProdutos.TextMatrix(0, 0) = ""
  gridOperacoesProdutos.TextMatrix(0, 1) = "Código"
  gridOperacoesProdutos.TextMatrix(0, 2) = "Nome produto"
  gridOperacoesProdutos.TextMatrix(0, 3) = "Tamanho"
  gridOperacoesProdutos.TextMatrix(0, 4) = "Cor"
  gridOperacoesProdutos.TextMatrix(0, 5) = "Quantidade Movimentada"


  gridVendasSeparadasPorDia.ColWidth(0) = 0
  gridVendasSeparadasPorDia.ColWidth(1) = 2800
  gridVendasSeparadasPorDia.ColWidth(2) = 1740
  gridVendasSeparadasPorDia.ColWidth(3) = 1740
  gridVendasSeparadasPorDia.ColWidth(4) = 1740
  gridVendasSeparadasPorDia.ColWidth(5) = 1740
  gridVendasSeparadasPorDia.ColWidth(6) = 1740
  gridVendasSeparadasPorDia.ColWidth(7) = 1740
  gridVendasSeparadasPorDia.ColWidth(8) = 1740

  gridVendasSeparadasPorDia.Row = 0
  gridVendasSeparadasPorDia.TextMatrix(0, 0) = ""
  gridVendasSeparadasPorDia.TextMatrix(0, 1) = "Data/Filial"
  gridVendasSeparadasPorDia.TextMatrix(0, 2) = "Dinheiro"
  gridVendasSeparadasPorDia.TextMatrix(0, 3) = "Cartão"
  gridVendasSeparadasPorDia.TextMatrix(0, 4) = "Prazo"
  gridVendasSeparadasPorDia.TextMatrix(0, 5) = "Cheque"
  gridVendasSeparadasPorDia.TextMatrix(0, 6) = "Vale"
  gridVendasSeparadasPorDia.TextMatrix(0, 7) = "Não financeiro"
  gridVendasSeparadasPorDia.TextMatrix(0, 8) = "TOTAL"


  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsVendedor = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsCliente = db.OpenRecordset("Cli_For", , dbReadOnly)

  Combo_Filial.Text = gnCodFilial
  Combo_Filial_LostFocus
  
  cmdPesquisar_Click

  Exit Sub
  
Erro:
  MsgBox "Erro na abertura da tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Erro

    rsClasses.Close
    Set rsClasses = Nothing
    rsSubclasses.Close
    Set rsSubclasses = Nothing
    
    Exit Sub
Erro:
  MsgBox "Erro no fechamento da tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub gridOperacoes_Click()
    If chk_visaoConsolidada.Value = vbUnchecked And chk_visaoProdutos.Value = vbUnchecked Then
        DetalharOperacao
    ElseIf chk_visaoConsolidada.Value = vbUnchecked And chk_visaoProdutos.Value = vbChecked Then
        DetalharOperacaoAgrupadaDeProdutos
    End If
End Sub
