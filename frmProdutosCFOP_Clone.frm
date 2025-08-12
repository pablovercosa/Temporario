VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProdutosCFOP_Clone 
   Caption         =   " Clonar as características do produto"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdutosCFOP_Clone.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   14070
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data4 
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
      Left            =   10200
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cli_For"
      Top             =   1290
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clonar as seguintes características"
      Height          =   1575
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   11700
      Begin VB.CheckBox chk_opcao_SubClasse 
         Appearance      =   0  'Flat
         Caption         =   "Sub Classe"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10080
         TabIndex        =   48
         Top             =   510
         Width           =   1245
      End
      Begin VB.CheckBox chk_opcao_Classe 
         Appearance      =   0  'Flat
         Caption         =   "Classe"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10080
         TabIndex        =   47
         Top             =   240
         Width           =   795
      End
      Begin VB.CheckBox chk_CodigoBeneficio 
         Appearance      =   0  'Flat
         Caption         =   "Código Benefício"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1755
      End
      Begin VB.CheckBox chk_EntradaIPI 
         Appearance      =   0  'Flat
         Caption         =   "% IPI Entrada"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6030
         TabIndex        =   16
         Top             =   1050
         Width           =   1545
      End
      Begin VB.CheckBox chk_opcao_EntradaICMSST_baseCalculo 
         Appearance      =   0  'Flat
         Caption         =   "ICMSST BaseCalculo Entrada"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6030
         TabIndex        =   15
         Top             =   780
         Width           =   2415
      End
      Begin VB.CheckBox chk_opcao_SaidaICMSST_baseCalculo 
         Appearance      =   0  'Flat
         Caption         =   "ICMSST BaseCalculo Saída"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3660
         TabIndex        =   11
         Top             =   780
         Width           =   2265
      End
      Begin VB.CheckBox chk_opcao_EntradaICMSST 
         Appearance      =   0  'Flat
         Caption         =   "ICMSST Entrada"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6030
         TabIndex        =   14
         Top             =   510
         Width           =   1545
      End
      Begin VB.CheckBox chk_opcao_SaidaICMSST 
         Appearance      =   0  'Flat
         Caption         =   "ICMSST Saída"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3660
         TabIndex        =   10
         Top             =   510
         Width           =   1365
      End
      Begin VB.CheckBox chk_opcao_EntradaICMS 
         Appearance      =   0  'Flat
         Caption         =   "ICMS Entrada"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6030
         TabIndex        =   13
         Top             =   240
         Width           =   1365
      End
      Begin VB.CheckBox chk_fornecedor 
         Appearance      =   0  'Flat
         Caption         =   "Fornecedor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8550
         TabIndex        =   17
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chk_NCM 
         Appearance      =   0  'Flat
         Caption         =   "NCM"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   705
      End
      Begin VB.CheckBox chk_situacaoTrib 
         Appearance      =   0  'Flat
         Caption         =   "Situação Tributária (Empresa Lucro Real)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   510
         Width           =   3225
      End
      Begin VB.CheckBox chk_opcao_SaidaICMS 
         Appearance      =   0  'Flat
         Caption         =   "ICMS Saída"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3660
         TabIndex        =   9
         Top             =   240
         Width           =   1245
      End
      Begin VB.CheckBox chk_SaidaIPI 
         Appearance      =   0  'Flat
         Caption         =   "% IPI Saída"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3660
         TabIndex        =   12
         Top             =   1050
         Width           =   1545
      End
      Begin VB.CheckBox chk_CFOPs_vinculados 
         Appearance      =   0  'Flat
         Caption         =   "CFOPs de Operações vinculados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   2685
      End
      Begin VB.CheckBox chk_situacaoTribPIS 
         Appearance      =   0  'Flat
         Caption         =   "Situação Tributária do PIS"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1050
         Width           =   2265
      End
      Begin SSDataWidgets_B.SSDBCombo cboCliente 
         Bindings        =   "frmProdutosCFOP_Clone.frx":4E95A
         DataSource      =   "Data4"
         Height          =   330
         Left            =   8550
         TabIndex        =   18
         Top             =   510
         Visible         =   0   'False
         Width           =   1005
         DataFieldList   =   "Nome"
         ListAutoValidate=   0   'False
         MaxDropDownItems=   16
         BevelType       =   0
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
         BorderStyle     =   0
         BevelColorFace  =   15066597
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   9075
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1746
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         Columns(2).Width=   847
         Columns(2).Caption=   "Tipo"
         Columns(2).Name =   "Tipo"
         Columns(2).CaptionAlignment=   0
         Columns(2).DataField=   "Tipo"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   4339
         Columns(3).Caption=   "Cidade"
         Columns(3).Name =   "Cidade"
         Columns(3).CaptionAlignment=   0
         Columns(3).DataField=   "Cidade"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1191
         Columns(4).Caption=   "Estado"
         Columns(4).Name =   "Estado"
         Columns(4).CaptionAlignment=   0
         Columns(4).DataField=   "Estado"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         _ExtentX        =   1773
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   0
         BackColor       =   12648447
      End
      Begin VB.Label Nome_Cliente 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   570
         Left            =   8550
         TabIndex        =   19
         Top             =   870
         Visible         =   0   'False
         Width           =   3075
      End
   End
   Begin VB.CommandButton cmd_clonarProdutos 
      BackColor       =   &H00FFA324&
      Caption         =   "Clonar características"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7380
      Width           =   2205
   End
   Begin VB.CommandButton cmd_desvinculaUM 
      BackColor       =   &H00C0FFFF&
      Caption         =   "/\"
      Height          =   285
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5190
      Width           =   1035
   End
   Begin VB.CommandButton cmd_desvinculaTODOS 
      BackColor       =   &H00C0FFFF&
      Caption         =   "//\\"
      Height          =   285
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5190
      Width           =   1035
   End
   Begin VB.CommandButton cmd_vinculaUM 
      BackColor       =   &H00FFA324&
      Caption         =   "\/"
      Height          =   285
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5190
      Width           =   1035
   End
   Begin VB.CommandButton cmd_vinculaTODOS 
      BackColor       =   &H00FFA324&
      Caption         =   "\\//"
      Height          =   285
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5190
      Width           =   1035
   End
   Begin VB.CommandButton cmd_pesquisar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar produtos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2850
      Width           =   2205
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisar produtos por"
      Height          =   1245
      Left            =   60
      TabIndex        =   20
      Top             =   2070
      Width           =   11700
      Begin VB.Data datCodigoNBM 
         Caption         =   "datCodigoNBM"
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
         Left            =   13410
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Codigo, Nome FROM AliquotasNCM ORDER BY Codigo"
         Top             =   810
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.OptionButton opt_Nada 
         Appearance      =   0  'Flat
         Caption         =   "Nenhuma das opções acima"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   930
         Value           =   -1  'True
         Width           =   2385
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
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
         Left            =   13410
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Sub_Classe"
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   13410
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Classe"
         Top             =   150
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txt_codigo_pesquisa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1920
         TabIndex        =   25
         Top             =   450
         Width           =   2625
      End
      Begin VB.OptionButton opt_fornecedor 
         Appearance      =   0  'Flat
         Caption         =   "Código Fornecedor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   705
         Width           =   1725
      End
      Begin VB.OptionButton opt_codigo 
         Appearance      =   0  'Flat
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   495
         Width           =   855
      End
      Begin VB.OptionButton opt_nome 
         Appearance      =   0  'Flat
         Caption         =   "Parte do Nome"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   270
         Width           =   1695
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
         Bindings        =   "frmProdutosCFOP_Clone.frx":4E96E
         DataSource      =   "Data1"
         Height          =   315
         Left            =   6210
         TabIndex        =   26
         Top             =   150
         Width           =   2235
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
         GroupHeaders    =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         RowHeight       =   476
         Columns.Count   =   2
         Columns(0).Width=   5450
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3545
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   3942
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Sub_Classe 
         Bindings        =   "frmProdutosCFOP_Clone.frx":4E982
         DataSource      =   "Data2"
         Height          =   315
         Left            =   6210
         TabIndex        =   27
         Top             =   510
         Width           =   2235
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
         ForeColorEven   =   0
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         Columns(0).Width=   3200
         _ExtentX        =   3942
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBCombo cboCodigoNBM 
         Bindings        =   "frmProdutosCFOP_Clone.frx":4E996
         DataSource      =   "datCodigoNBM"
         Height          =   315
         Left            =   6210
         TabIndex        =   28
         Top             =   855
         Width           =   2235
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
         ForeColorEven   =   0
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         Columns(0).Width=   3200
         _ExtentX        =   3942
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Código"
      End
      Begin VB.Label Label3 
         Caption         =   "e"
         Height          =   255
         Left            =   4950
         TabIndex        =   46
         Top             =   540
         Width           =   135
      End
      Begin VB.Label lblNomeCodigoNBM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8490
         TabIndex        =   45
         Top             =   855
         Width           =   3105
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "NCM"
         Height          =   195
         Left            =   5760
         TabIndex        =   43
         Top             =   900
         Width           =   330
      End
      Begin VB.Label Label13 
         Caption         =   "Subclasse"
         Height          =   255
         Left            =   5400
         TabIndex        =   39
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label12 
         Caption         =   "Classe"
         Height          =   255
         Left            =   5640
         TabIndex        =   38
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Nome_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8490
         TabIndex        =   37
         Top             =   150
         Width           =   3105
      End
      Begin VB.Label Nome_Sub_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8490
         TabIndex        =   34
         Top             =   510
         Width           =   3105
      End
   End
   Begin VB.TextBox txt_codigoProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
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
      Height          =   330
      Left            =   600
      TabIndex        =   1
      Top             =   52
      Width           =   2415
   End
   Begin VB.TextBox txt_nomeProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Height          =   330
      Left            =   3075
      TabIndex        =   2
      Top             =   52
      Width           =   8685
   End
   Begin MSFlexGridLib.MSFlexGrid gridPesquisa 
      Height          =   1740
      Left            =   60
      TabIndex        =   30
      Top             =   3360
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   3069
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483648
      BackColorFixed  =   12648447
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483648
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin MSFlexGridLib.MSFlexGrid gridClonar 
      Height          =   2040
      Left            =   60
      TabIndex        =   35
      Top             =   5790
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   3598
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483648
      BackColorFixed  =   16753444
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483648
      GridColor       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Label lbl_pesquisa 
      Caption         =   "Aguarde, carregando a grade..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   8400
      TabIndex        =   42
      Top             =   5310
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label lbl_codigoAlvo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   6120
      TabIndex        =   41
      Top             =   5520
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFA324&
      Caption         =   "Grade dos produtos selecionados que receberão as características do produto"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   90
      TabIndex        =   40
      Top             =   5520
      Width           =   5985
   End
   Begin VB.Label Label9 
      Caption         =   "Código"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   495
   End
End
Attribute VB_Name = "frmProdutosCFOP_Clone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public codigoProduto As String
Public nomeProduto As String
Public CodigoNCM As String
Public SituacaoTributaria As String
Public CodigoBeneficio As String
Public PercentualIPI_Saida As String
Public PercentualIPI_Entrada As String
Public PercentualICMS_Saida As String
Public PercentualICMS_Entrada As String
Public PercentualICMSST_Saida As String
Public PercentualICMSST_Entrada As String
Public ValorBC_ICMSST_Saida As String
Public ValorBC_ICMSST_Entrada As String
Public TipoSituacaoTributariaPIS As Integer

Public ClasseProduto As String
Public SubClasseProduto As String

Private arrFornecedores(30) As Long
Private arrFornecedoresNum As Integer


Dim rsCliFor As Recordset
Dim rsClasses As Recordset
Dim rsSub_Classes As Recordset
Dim rsCFOP As Recordset

Private Sub cboCliente_Click()
  cboCliente.Text = cboCliente.Columns(1).Text
End Sub

Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(1).Text
  cboCliente_LostFocus
End Sub

Private Sub cboCliente_LostFocus()
  Dim Aux As String
  Dim i As Integer
  Dim boFornecedorValido As Boolean

  Aux = cboCliente.Text
  If IsNull(Aux) Then Exit Sub
  If Aux = "" Then Exit Sub
  If Not IsNumeric(Aux) Then
      cboCliente.Text = ""
      Nome_Cliente.Caption = ""
      Exit Sub
  End If
  If Val(Aux) < 1 Then
      Nome_Cliente.Caption = ""
      Exit Sub
  End If
  If Val(Aux) > 99999999 Then
      Nome_Cliente.Caption = ""
      Exit Sub
  End If
  
  boFornecedorValido = False
  For i = 0 To arrFornecedoresNum - 1
      If arrFornecedores(i) = Aux Then
          boFornecedorValido = True
          Exit For
      End If
  Next
  
  If boFornecedorValido = True Then
      rsCliFor.Index = "Código"
      rsCliFor.Seek "=", Val(Aux)
      If rsCliFor.NoMatch Then
          DisplayMsg "Cliente incorreto."
          cboCliente.SetFocus
          Exit Sub
      End If
      
      Nome_Cliente.Caption = rsCliFor("Nome") & ""
  Else
      cboCliente.Text = ""
      Nome_Cliente.Caption = ""
  End If
End Sub

Private Sub cboCodigoNBM_CloseUp()
  cboCodigoNBM.Text = cboCodigoNBM.Columns(0).Text
  cboCodigoNBM_LostFocus
End Sub

Private Sub cboCodigoNBM_LostFocus()
  Dim rstCodigoNBM As Recordset
  
  lblNomeCodigoNBM.Caption = ""
  If Len(cboCodigoNBM.Text) <= 0 Then Exit Sub
  
  ' Pilatti Dezembro/17
  Set rstCodigoNBM = db.OpenRecordset("SELECT Codigo, Nome FROM AliquotasNCM WHERE Codigo = '" & CStr(cboCodigoNBM.Text) & "'", dbOpenSnapshot)
  
  With rstCodigoNBM
    If Not (.BOF And .EOF) Then
      lblNomeCodigoNBM.Caption = .Fields("Nome") & ""
    End If
    
    If Not rstCodigoNBM Is Nothing Then .Close
    Set rstCodigoNBM = Nothing
  End With
  
  If cboCodigoNBM.Text <> "" And lblNomeCodigoNBM.Caption = "" Then
      MsgBox "Antes de vincular este NCM " & cboCodigoNBM.Text & " neste produto você deve realizar o cadastro do NCM no sistema. O caminho é pelo menu principal, Aba 'Cadastro', opção 'Códigos NCM'.", vbInformation, "Realize o cadastro"
  End If
End Sub

Private Sub chk_fornecedor_Click()
  If chk_fornecedor.Value = vbChecked Then
      cboCliente.Visible = True
      Nome_Cliente.Visible = True
  Else
      cboCliente.Text = ""
      Nome_Cliente.Caption = ""
      cboCliente.Visible = False
      Nome_Cliente.Visible = False
  End If
End Sub

Private Sub cmd_clonarProdutos_Click()
On Error GoTo Erro

  Dim retMsg As Variant
  retMsg = MsgBox("Serão clonadas as características fiscais do produto " & txt_codigoProduto.Text & " nos produtos da última grade. Deseja realmente prosseguir?", vbYesNo, "Atenção")
  
  If retMsg = vbNo Then
      Exit Sub
  End If

  If chk_NCM.Value = False And chk_situacaoTrib.Value = False And chk_CFOPs_vinculados.Value = False And _
  chk_situacaoTribPIS.Value = False And chk_CodigoBeneficio.Value = False And chk_opcao_SaidaICMS.Value = False And _
  chk_opcao_SaidaICMSST.Value = False And chk_opcao_SaidaICMSST_baseCalculo.Value = False And chk_SaidaIPI.Value = False And _
  chk_opcao_EntradaICMS.Value = False And chk_opcao_EntradaICMSST.Value = False And chk_opcao_EntradaICMSST_baseCalculo.Value = False And _
  chk_EntradaIPI.Value = False And chk_fornecedor.Value = False And chk_opcao_Classe.Value = False And chk_opcao_SubClasse.Value = False Then
  
      MsgBox "Escolha pelo menos uma das características fiscais que deseja clonar.", vbInformation, "Atenção"
      Exit Sub
  End If
  
  If chk_fornecedor.Value = vbChecked And Nome_Cliente.Caption = "" Then
      MsgBox "Escolha um fornecedor.", vbInformation, "Atenção"
      Exit Sub
  End If

  lbl_pesquisa.Visible = True
  lbl_pesquisa.Caption = "Aguarde, salvando..."
  DoEvents

  Dim lCont As Long
  Dim bTem As Boolean
  Dim sCodigoProd As String
  Dim sSql As String
  
  Call ws.BeginTrans

  
  For lCont = 1 To gridClonar.Rows - 1
      sCodigoProd = gridClonar.TextMatrix(lCont, 1)
      
      bTem = False
      
      'Atualizar na tabela de Produtos
      If chk_NCM.Value = vbChecked Or chk_situacaoTrib.Value = vbChecked Or chk_situacaoTribPIS.Value = vbChecked Or _
          chk_CodigoBeneficio.Value = vbChecked Or chk_opcao_SaidaICMS.Value = vbChecked Or chk_opcao_SaidaICMSST.Value = vbChecked Or _
          chk_opcao_SaidaICMSST_baseCalculo.Value = vbChecked Or chk_SaidaIPI.Value = vbChecked Or chk_opcao_EntradaICMS.Value = vbChecked Or _
          chk_opcao_EntradaICMSST.Value = vbChecked Or chk_opcao_EntradaICMSST_baseCalculo.Value = vbChecked Or chk_EntradaIPI.Value = vbChecked Or _
          chk_opcao_Classe.Value = vbChecked Or chk_opcao_SubClasse.Value = vbChecked Then

          sSql = "Update Produtos "
      
          If chk_NCM.Value = vbChecked Then
              sSql = sSql & " set CodigoNBM='" & CodigoNCM & "' "
              bTem = True
          End If
      
          If chk_situacaoTrib.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              sSql = sSql & " [Situação Tributária]='" & SituacaoTributaria & "' "
              bTem = True
          End If
      
          If chk_CodigoBeneficio.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              sSql = sSql & " CodigoBeneficio='" & CodigoBeneficio & "' "
              bTem = True
          End If
          
          If chk_opcao_SaidaICMS.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(PercentualICMS_Saida) = "" Then
                  PercentualICMS_Saida = "0"
              End If
              
              sSql = sSql & " [Percentual ICM Saida]=" & Replace(CStr(PercentualICMS_Saida), ",", ".")
              bTem = True
          End If
          
          If chk_opcao_EntradaICMS.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(PercentualICMS_Entrada) = "" Then
                  PercentualICMS_Entrada = "0"
              End If
              
              sSql = sSql & " [Percentual ICM Entrada]=" & Replace(CStr(PercentualICMS_Entrada), ",", ".")
              bTem = True
          End If
      
          If chk_SaidaIPI.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(PercentualIPI_Saida) = "" Then
                  PercentualIPI_Saida = "0"
              End If
              
              sSql = sSql & " [Percentual IPI]=" & Replace(CStr(PercentualIPI_Saida), ",", ".")
              bTem = True
          End If
          
          If chk_EntradaIPI.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(PercentualIPI_Entrada) = "" Then
                  PercentualIPI_Entrada = "0"
              End If
              
              sSql = sSql & " Percentual_IPI_Entrada=" & Replace(CStr(PercentualIPI_Entrada), ",", ".")
              bTem = True
          End If

          If chk_opcao_SaidaICMSST.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(PercentualICMSST_Saida) = "" Then
                  PercentualICMSST_Saida = "0"
              End If
              
              sSql = sSql & " Percentual_ICMSST_Saida=" & Replace(CStr(PercentualICMSST_Saida), ",", ".")
              bTem = True
          End If
          
          If chk_opcao_EntradaICMSST.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(PercentualICMSST_Entrada) = "" Then
                  PercentualICMSST_Entrada = "0"
              End If
              
              sSql = sSql & " Percentual_ICMSST_Entrada=" & Replace(CStr(PercentualICMSST_Entrada), ",", ".")
              bTem = True
          End If

          If chk_opcao_SaidaICMSST_baseCalculo.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(ValorBC_ICMSST_Saida) = "" Then
                  ValorBC_ICMSST_Saida = "0"
              End If
              
              sSql = sSql & " BaseCalculoICMSST_Saida=" & Replace(CStr(ValorBC_ICMSST_Saida), ",", ".")
              bTem = True
          End If

          If chk_opcao_EntradaICMSST_baseCalculo.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(ValorBC_ICMSST_Entrada) = "" Then
                  ValorBC_ICMSST_Entrada = "0"
              End If
              
              sSql = sSql & " BaseCalculoICMSST_Entrada=" & Replace(CStr(ValorBC_ICMSST_Entrada), ",", ".")
              bTem = True
          End If
          
          If chk_opcao_Classe.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(ClasseProduto) = "" Then
                  ClasseProduto = "0"
              End If
              
              sSql = sSql & " Classe=" & Replace(CStr(ClasseProduto), ",", ".")
              bTem = True
          End If
          
          If chk_opcao_SubClasse.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              
              If Trim(SubClasseProduto) = "" Then
                  SubClasseProduto = "0"
              End If
              
              sSql = sSql & " [Sub Classe]=" & Replace(CStr(SubClasseProduto), ",", ".")
              bTem = True
          End If
          
          
          
          If chk_situacaoTribPIS.Value = vbChecked Then
              If bTem = True Then
                  sSql = sSql & ", "
              Else
                  sSql = sSql & " set "
              End If
              sSql = sSql & " [TipoSituacaoTributariaPIS]=" & TipoSituacaoTributariaPIS
          End If

          sSql = sSql & " WHERE Código='" & sCodigoProd & "' "

          db.Execute sSql
      
      End If
      
      'Atualizar na tabela de ProdutoCFOP
      If chk_CFOPs_vinculados.Value = vbChecked Then
          
          sSql = " DELETE From ProdutoCFOP WHERE CodProduto='" & sCodigoProd & "' "
          db.Execute sSql
          
          
          If Not (rsCFOP.EOF And rsCFOP.BOF) Then
              rsCFOP.MoveFirst
          End If
          
          While Not rsCFOP.EOF
              sSql = " INSERT INTO ProdutoCFOP (CodProduto, CodOperacao, CFOP, CSO) "
              sSql = sSql & " VALUES ('" & sCodigoProd & "', " & rsCFOP.Fields("CodOperacao").Value & ","
              sSql = sSql & " '" & rsCFOP.Fields("CFOP").Value & "', '" & rsCFOP.Fields("CSO").Value & "') "

              db.Execute sSql

              rsCFOP.MoveNext
          Wend
      End If
      
      ' Atualizar na tabela de Forn_Prod
      If chk_fornecedor.Value = vbChecked Then
          sSql = " INSERT INTO Forn_Prod (Produto, Fornecedor) "
          sSql = sSql & " VALUES ('" & sCodigoProd & "', " & cboCliente.Text & ")"
 
          db.Execute sSql
      End If
  Next
  
  Call ws.CommitTrans
  
  lbl_pesquisa.Visible = False
  DoEvents

  MsgBox "Clonagem das características fiscais foram realizadas com sucesso nos produtos selecionados!", vbInformation, "Sucesso"
  
  'Limpar todos os campos da tela...
  gridPesquisa.Rows = 1
  gridClonar.Rows = 1
  Combo_Classe.Text = ""
  Combo_Sub_Classe.Text = ""
  Nome_Classe.Caption = ""
  Nome_Sub_Classe.Caption = ""
  txt_codigo_pesquisa.Text = ""
  chk_NCM.Value = False
  chk_situacaoTrib.Value = False
  chk_CFOPs_vinculados.Value = False
  chk_situacaoTribPIS.Value = False
  chk_CodigoBeneficio.Value = False
  chk_opcao_SaidaICMS.Value = False
  chk_opcao_EntradaICMS.Value = False
  chk_opcao_SaidaICMSST.Value = False
  chk_opcao_EntradaICMSST.Value = False
  chk_opcao_SaidaICMSST_baseCalculo.Value = False
  chk_opcao_EntradaICMSST_baseCalculo.Value = False
  chk_SaidaIPI.Value = False
  chk_EntradaIPI.Value = False
  chk_opcao_Classe.Value = False
  chk_opcao_SubClasse.Value = False

  Exit Sub
Erro:
    Call ws.Rollback

    MsgBox "Erro na função de clonar características fiscais de um produto para outros selecionados na grade. " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_desvinculaTODOS_Click()
  gridClonar.Rows = 1

End Sub

Private Sub cmd_desvinculaUM_Click()
On Error GoTo Erro

  If gridClonar.Rows = 2 Then
      gridClonar.Rows = 1
      Exit Sub
  End If

  If gridClonar.RowSel > 0 Then
      gridClonar.RemoveItem gridClonar.RowSel
  Else
      MsgBox "Selecione um registro na grade!", vbInformation, "Atenção"
      Exit Sub
  End If

  Exit Sub
Erro:
    MsgBox "Erro na função Desvincular um produto " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cmd_pesquisar_Click()
    lbl_pesquisa.Visible = True
    lbl_pesquisa.Caption = "Aguarde, carregando a grade..."
    DoEvents

    PesquisarProdutos

    lbl_pesquisa.Visible = False
    DoEvents
End Sub

Private Sub PesquisarProdutos()
On Error GoTo ErrHandler
  
  Dim rsProdutos As Recordset
  Dim strSQL As String
  
  gridPesquisa.Rows = 1
  gridPesquisa.Row = 0

'  With gridPesquisa
'    .Redraw = False
'    .RemoveAll
'    .Redraw = True
'  End With

  If opt_Nada.Value = False And Trim(txt_codigo_pesquisa.Text) = "" Then
      MsgBox "Entre com algum dado para a pesquisa ou MARQUE a opção 'Nenhuma das opções acima'", vbInformation, "Atenção"
      txt_codigo_pesquisa.SetFocus
      Exit Sub
  End If

  strSQL = "SELECT A.Código, A.Nome, A.Classe, B.Nome, A.[Sub Classe], C.Nome "
  strSQL = strSQL & " From Produtos A, Classes B, [Sub Classes] C "
  strSQL = strSQL & " Where "

  If opt_nome.Value = True Then
      strSQL = strSQL & " A.Nome like '*" & Trim(txt_codigo_pesquisa.Text) & "*' and "
  ElseIf opt_codigo.Value = True Then
      strSQL = strSQL & " A.Código = '" & Trim(txt_codigo_pesquisa.Text) & "' and "
  ElseIf opt_fornecedor.Value = True Then
      strSQL = strSQL & " A.[Código do Fornecedor] = '" & Trim(txt_codigo_pesquisa.Text) & "' and "
  End If

  If Trim(Combo_Classe.Text) <> "" And Trim(Combo_Classe.Text) <> "0" Then
      strSQL = strSQL & " A.Classe = " & Trim(Combo_Classe.Text) & " and "
  End If

  If Trim(Combo_Sub_Classe.Text) <> "" And Trim(Combo_Sub_Classe.Text) <> "0" Then
      strSQL = strSQL & " A.[Sub Classe] = " & Trim(Combo_Sub_Classe.Text) & " and "
  End If
  
  If Trim(cboCodigoNBM.Text) <> "" And Trim(cboCodigoNBM.Text) <> "0" Then
      strSQL = strSQL & " A.CodigoNBM = '" & Trim(cboCodigoNBM.Text) & "' and "
  End If

  strSQL = strSQL & " A.Classe = B.Código and "
  strSQL = strSQL & " A.[Sub Classe] = C.Código "
  strSQL = strSQL & " ORDER BY A.Nome "
 
  Set rsProdutos = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsProdutos
    If Not (.BOF And .EOF) Then
      Do Until .EOF
          gridPesquisa.AddItem vbTab & .Fields(0).Value & vbTab & _
                      .Fields(1).Value & vbTab & _
                      .Fields(3).Value & vbTab & _
                      .Fields(5).Value
          .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsProdutos = Nothing
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao exibir registros de pesquisa. Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmd_vinculaTODOS_Click()
On Error GoTo Erro

  lbl_pesquisa.Visible = True
  lbl_pesquisa.Caption = "Aguarde, carregando a grade..."

  DoEvents

  Dim lCont As Long
  
  For lCont = 1 To gridPesquisa.Rows - 1
      If gridPesquisa.TextMatrix(lCont, 1) <> codigoProduto Then
          gridClonar.AddItem vbTab & gridPesquisa.TextMatrix(lCont, 1) & vbTab & _
                      gridPesquisa.TextMatrix(lCont, 2) & vbTab & _
                      gridPesquisa.TextMatrix(lCont, 3) & vbTab & _
                      gridPesquisa.TextMatrix(lCont, 4)
      End If
  Next
  
  lbl_pesquisa.Visible = False
  DoEvents
  
  Exit Sub
Erro:
    MsgBox "Erro na função Vincular todos os produto " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub cmd_vinculaUM_Click()
On Error GoTo Erro

  If gridPesquisa.RowSel > 0 Then
      If gridPesquisa.TextMatrix(gridPesquisa.RowSel, 1) <> codigoProduto Then
          gridClonar.AddItem vbTab & gridPesquisa.TextMatrix(gridPesquisa.RowSel, 1) & vbTab & _
                      gridPesquisa.TextMatrix(gridPesquisa.RowSel, 2) & vbTab & _
                      gridPesquisa.TextMatrix(gridPesquisa.RowSel, 3) & vbTab & _
                      gridPesquisa.TextMatrix(gridPesquisa.RowSel, 4)
      End If
  Else
      MsgBox "Selecione um registro na grade!", vbInformation, "Atenção"
      Exit Sub
  End If

  Exit Sub
Erro:
    MsgBox "Erro na função Vincular um produto " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub Form_Load()
On Error GoTo Erro
  Dim strSQL As String

  Me.Caption = " Clonar as características do produto " & codigoProduto

  If Not IsNull(codigoProduto) And codigoProduto <> "" Then
    txt_codigoProduto.Text = codigoProduto
    txt_nomeProduto.Text = nomeProduto
    lbl_codigoAlvo.Caption = codigoProduto
  End If

  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  datCodigoNBM.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName

  strSQL = "SELECT C.Nome, C.Código, C.Tipo, C.Cidade, C.Estado "
  strSQL = strSQL & " From Cli_For C, Forn_Prod F"
  strSQL = strSQL & " Where F.Produto = '" & txt_codigoProduto.Text & "'"
  strSQL = strSQL & " And F.Fornecedor = C.Código"
  strSQL = strSQL & " ORDER BY Nome"
  Data4.RecordSource = strSQL
  Data4.Refresh
  
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  
  Dim rsCliForAux As Recordset
  Set rsCliForAux = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  If rsCliForAux.EOF And rsCliForAux.BOF Then
      chk_fornecedor.Visible = False
  Else
      rsCliForAux.MoveLast
      rsCliForAux.MoveFirst
      arrFornecedoresNum = 0
      While Not rsCliForAux.EOF
          arrFornecedores(arrFornecedoresNum) = rsCliForAux.Fields(1).Value
          arrFornecedoresNum = arrFornecedoresNum + 1
          rsCliForAux.MoveNext
      Wend
  End If
  rsCliForAux.Close
  Set rsCliForAux = Nothing

  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)

  gridPesquisa.ColWidth(0) = 0
  gridPesquisa.ColWidth(1) = 2200
  gridPesquisa.ColWidth(2) = 4200
  gridPesquisa.ColWidth(3) = 2440
  gridPesquisa.ColWidth(4) = 2440
  
  gridPesquisa.Row = 0
  gridPesquisa.TextMatrix(0, 1) = "Código"
  gridPesquisa.TextMatrix(0, 2) = "Nome"
  gridPesquisa.TextMatrix(0, 3) = "Classe"
  gridPesquisa.TextMatrix(0, 4) = "Sub-Classe"

  gridClonar.ColWidth(0) = 0
  gridClonar.ColWidth(1) = 2200
  gridClonar.ColWidth(2) = 4200
  gridClonar.ColWidth(3) = 2440
  gridClonar.ColWidth(4) = 2440
  
  gridClonar.Row = 0
  gridClonar.TextMatrix(0, 1) = "Código"
  gridClonar.TextMatrix(0, 2) = "Nome"
  gridClonar.TextMatrix(0, 3) = "Classe"
  gridClonar.TextMatrix(0, 4) = "Sub-Classe"
  
  strSQL = "SELECT CodOperacao, CFOP, CSO FROM ProdutoCFOP "
  strSQL = strSQL & " where CodProduto = '" & codigoProduto & "' "
  
  Set rsCFOP = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  Exit Sub
Erro:

  MsgBox "Erro ao realizar carga da tela...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
  
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

Private Sub Combo_Sub_Classe_CloseUp()
  Combo_Sub_Classe.Text = Combo_Sub_Classe.Columns(1).Text
  Combo_Sub_Classe_LostFocus
End Sub

Private Sub Combo_Sub_Classe_LostFocus()
  Dim Aux As Variant
  
  Nome_Sub_Classe.Caption = ""
  Aux = Combo_Sub_Classe.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsSub_Classes.Index = "Código"
  rsSub_Classes.Seek "=", Val(Aux)
  If rsSub_Classes.NoMatch Then Exit Sub
  
  If Not IsNull(rsSub_Classes("Nome")) Then
      Nome_Sub_Classe.Caption = rsSub_Classes("Nome")
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsClasses.Close
  rsSub_Classes.Close
  rsCFOP.Close
  rsCliFor.Close
  
  Set rsClasses = Nothing
  Set rsSub_Classes = Nothing
  Set rsCFOP = Nothing
  Set rsCliFor = Nothing
End Sub

Private Sub gridPesquisa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridPesquisa.Redraw = False
End Sub

Private Sub gridPesquisa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridPesquisa.RowSel = gridPesquisa.Row
  gridPesquisa.Redraw = True
End Sub
