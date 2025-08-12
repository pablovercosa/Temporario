VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmProdutosTributos 
   Caption         =   " Tributos do produto"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14520
   Icon            =   "frmProdutosTributos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox pRedBC_ICM 
      Height          =   315
      Left            =   12480
      TabIndex        =   43
      Top             =   4470
      Width           =   795
      _ExtentX        =   1397
      _ExtentY        =   550
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "##0.00"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmd_salvar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salvar"
      Height          =   430
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5070
      Width           =   14355
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   14355
      Begin VB.Label lbl_codigoProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   750
         TabIndex        =   28
         Top             =   180
         Width           =   2325
      End
      Begin VB.Label lbl_nomeProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   180
         Width           =   11085
      End
      Begin VB.Label Label9 
         Caption         =   "Produto"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame Frame10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Left            =   90
      TabIndex        =   0
      Top             =   780
      Width           =   14355
      Begin VB.ComboBox cmb_SituacaoTributariaEntrada 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmProdutosTributos.frx":4E95A
         Left            =   11010
         List            =   "frmProdutosTributos.frx":4E97F
         TabIndex        =   41
         Top             =   900
         Width           =   1095
      End
      Begin VB.Data datCodigoNBM 
         Caption         =   "datCodigoNBM"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10020
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Codigo, Nome, CEST FROM AliquotasNCM ORDER BY Codigo"
         Top             =   300
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.ComboBox cmb_codigoBeneficio 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmProdutosTributos.frx":4E9BA
         Left            =   3090
         List            =   "frmProdutosTributos.frx":4E9BC
         TabIndex        =   21
         Top             =   3720
         Width           =   5955
      End
      Begin VB.ComboBox cmb_situacaoTributariaDoPIS 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmProdutosTributos.frx":4E9BE
         Left            =   3090
         List            =   "frmProdutosTributos.frx":4EA28
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3240
         Width           =   11115
      End
      Begin VB.ComboBox cmb_SituacaoTributaria 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmProdutosTributos.frx":4F3B0
         Left            =   3090
         List            =   "frmProdutosTributos.frx":4F3D5
         TabIndex        =   8
         Top             =   900
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Perc_IPI_Saida 
         Height          =   315
         Left            =   3090
         TabIndex        =   1
         Top             =   2790
         Width           =   795
         _ExtentX        =   1397
         _ExtentY        =   550
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Perc_ICMS_Saida 
         Height          =   315
         Left            =   3090
         TabIndex        =   10
         Top             =   1530
         Width           =   795
         _ExtentX        =   1397
         _ExtentY        =   550
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Perc_ICMS_Entrada 
         Height          =   315
         Left            =   11010
         TabIndex        =   11
         ToolTipText     =   "Insira aqui a Aliquota do ICMS que consta na NFe de Compra"
         Top             =   1530
         Width           =   795
         _ExtentX        =   1397
         _ExtentY        =   550
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor_baseCalculo_ICMSST_Saida 
         Height          =   315
         Left            =   3090
         TabIndex        =   12
         Top             =   2370
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   550
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Perc_ICMSST_Saida 
         Height          =   315
         Left            =   3090
         TabIndex        =   18
         Top             =   1950
         Width           =   795
         _ExtentX        =   1397
         _ExtentY        =   550
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Perc_IPI_Entrada 
         Height          =   315
         Left            =   11010
         TabIndex        =   25
         Top             =   2790
         Width           =   795
         _ExtentX        =   1397
         _ExtentY        =   550
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00"
         PromptChar      =   "_"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCodigoNBM 
         Bindings        =   "frmProdutosTributos.frx":4F410
         DataSource      =   "datCodigoNBM"
         Height          =   315
         Left            =   690
         TabIndex        =   29
         Top             =   210
         Width           =   2355
         DataFieldList   =   "Codigo"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorEven   =   0
         BackColorEven   =   14737632
         BackColorOdd    =   12648384
         Columns(0).Width=   3200
         _ExtentX        =   4154
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Codigo"
      End
      Begin MSMask.MaskEdBox Valor_baseCalculo_ICMSST_Entrada 
         Height          =   315
         Left            =   11010
         TabIndex        =   32
         Top             =   2370
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   550
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Perc_ICMSST_Entrada 
         Height          =   315
         Left            =   11010
         TabIndex        =   33
         Top             =   1950
         Width           =   795
         _ExtentX        =   1397
         _ExtentY        =   550
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "%"
         Height          =   255
         Left            =   13320
         TabIndex        =   44
         Top             =   3780
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Porcentagem Redução Base Cálculo ICMS"
         Height          =   315
         Left            =   9240
         TabIndex        =   42
         Top             =   3780
         Width           =   3105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Origem + Situação Tributária Entrada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7410
         TabIndex        =   40
         Top             =   960
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ex: 010                                     1º caracter é a origem                 2ºe3º é a Tributação pelo ICMS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4260
         TabIndex        =   39
         Top             =   900
         Width           =   2370
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl_perc5 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   11940
         TabIndex        =   38
         Top             =   2820
         Width           =   285
      End
      Begin VB.Label lbl_perc4 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4020
         TabIndex        =   37
         Top             =   2820
         Width           =   285
      End
      Begin VB.Label lbl_tit_ValprBC_ICMSST_Entrada 
         AutoSize        =   -1  'True
         Caption         =   "Valor Base de Cálculo ICMS ST Entrada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7410
         TabIndex        =   36
         Top             =   2430
         Width           =   2790
      End
      Begin VB.Label Label3 
         Caption         =   "ICMSST de Entrada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7410
         TabIndex        =   35
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label lbl_perc6 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   11940
         TabIndex        =   34
         Top             =   1980
         Width           =   285
      End
      Begin VB.Label lblNomeCodigoNBM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   3090
         TabIndex        =   30
         Top             =   210
         Width           =   9105
      End
      Begin VB.Label lbl_tit_IPIEntrada 
         Caption         =   "IPI Entrada ou p/ Devolução ao Fornecedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7410
         TabIndex        =   26
         Top             =   2820
         Width           =   3345
      End
      Begin VB.Label Label6 
         Caption         =   "Situação Tributária do PIS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   3300
         Width           =   1995
      End
      Begin VB.Label lbl_tit_IPISaida 
         Caption         =   "IPI Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   2820
         Width           =   765
      End
      Begin VB.Label lblCodBeneficio 
         Caption         =   "Código do Benefício"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   3780
         Width           =   1695
      End
      Begin VB.Label lbl_perc3 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4020
         TabIndex        =   20
         Top             =   1980
         Width           =   285
      End
      Begin VB.Label lbl_tit_ICMSST_Saida 
         Caption         =   "ICMSST de Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label lbl_perc2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   11940
         TabIndex        =   17
         Top             =   1560
         Width           =   285
      End
      Begin VB.Label lbl_perc1 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4020
         TabIndex        =   16
         Top             =   1560
         Width           =   285
      End
      Begin VB.Label lbl_tit_ValprBC_ICMSST_Saida 
         AutoSize        =   -1  'True
         Caption         =   "Valor Base de Cálculo ICMS ST Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   2430
         Width           =   2610
      End
      Begin VB.Label lbl_tit_ICMSSaida 
         Caption         =   "ICMS de Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label lbl_tit_ICMSEntrada 
         Caption         =   "ICMS de Entrada ou p/ Devolução ao Fornecedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7410
         TabIndex        =   13
         Top             =   1560
         Width           =   3555
      End
      Begin VB.Label lbl_CodigoCEST 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   12810
         TabIndex        =   7
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label lbl_CEST 
         Caption         =   "CEST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12270
         TabIndex        =   6
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lbl_tit_NCM 
         Caption         =   "NCM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lbl_tit_SituacaoTributaria 
         AutoSize        =   -1  'True
         Caption         =   "Origem + Situação Tributária Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   960
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmProdutosTributos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodigoProduto As String
Public nomeProduto As String
Public NCM As String
Public SituacaoTributaria As String
Public SituacaoTributariaEntrada As String
Public ICMS_Saida As String
Public ICMS_EntradaDev As String
Public ICMSST_Saida As String
Public ICMSST_SaidaBaseCalculo As String
Public ICMSST_Entrada As String
Public ICMSST_EntradaBaseCalculo As String
Public Perc_IPISaida As String
Public Perc_IPIEntrada As String
Public Perc_RedBC_ICM As String
Public SituacaoTributariaPIS As Integer
Public CodigoBeneficio As String


Private Sub cboCodigoNBM_Click()
  cboCodigoNBM.Text = cboCodigoNBM.Columns(0).Text
End Sub

Private Sub cboCodigoNBM_CloseUp()
  cboCodigoNBM.Text = cboCodigoNBM.Columns(0).Text
  cboCodigoNBM_LostFocus
End Sub

Private Sub cboCodigoNBM_LostFocus()
  Dim rstCodigoNBM As Recordset
  
  lblNomeCodigoNBM.Caption = ""
  lbl_CodigoCEST.Caption = ""
  If Len(cboCodigoNBM.Text) <= 0 Then Exit Sub
  
  Set rstCodigoNBM = Db.OpenRecordset("SELECT Codigo, Nome, CEST FROM AliquotasNCM WHERE Codigo = '" & CStr(cboCodigoNBM.Text) & "'", dbOpenSnapshot)
  
  With rstCodigoNBM
    If Not (.BOF And .EOF) Then
      lblNomeCodigoNBM.Caption = .Fields("Nome") & ""
      lbl_CodigoCEST.Caption = .Fields("CEST") & ""
    End If
    
    If Not rstCodigoNBM Is Nothing Then .Close
    Set rstCodigoNBM = Nothing
  End With
  
  If cboCodigoNBM.Text <> "" And lblNomeCodigoNBM.Caption = "" Then
      MsgBox "Antes de vincular este NCM " & cboCodigoNBM.Text & " neste produto você deve realizar o cadastro do NCM no sistema. O caminho é pelo menu principal, Aba 'Cadastro', opção 'Códigos NCM'.", vbInformation, "Realize o cadastro"
  End If

End Sub

Private Sub cmd_salvar_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim sDado As String
  Dim sDadoArray() As String
  Dim sSituacaoTributaria As String
  
  sSituacaoTributaria = cmb_SituacaoTributaria.Text
  
  If Trim(sSituacaoTributaria) = "" Then
      sSituacaoTributaria = "0"
  End If

  sSql = "Update Produtos set CodigoNBM='" & cboCodigoNBM.Text & "',"
  sSql = sSql & " [Situação Tributária]='" & sSituacaoTributaria & "',"
  
  If Trim(cmb_SituacaoTributariaEntrada.Text) <> "" Then
      sSql = sSql & " SituacaoTributariaEntrada='" & Trim(cmb_SituacaoTributariaEntrada.Text) & "',"
  End If
  
  If Len(Trim(cboCodigoNBM.Text)) > 0 Then
      Dim rsNCM As Recordset
      Set rsNCM = Db.OpenRecordset("Select AliqNacional,AliqImportacao from [AliquotasNCM] Where Codigo = '" & Trim(cboCodigoNBM.Text) & "'")

      If Not rsNCM.EOF Then
      
          If Not IsNull(rsNCM.Fields("AliqNacional").Value) Then
              sSql = sSql & " AliqNCM= " & Replace(rsNCM.Fields("AliqNacional").Value, ",", ".") & ","
          Else
              sSql = sSql & " AliqNCM= 0, "
          End If
              
          ' Dependendo da Situação de Tributação do produto será AliquotaNacional ou de Importação (ver com detalhes como aplicar)
          '.Fields("AliqNCM") = rsNCM("AliqNacional")
          '.Fields("AliqNCM") = rsNCM("AliqImportacao")
      End If
  Else
      sSql = sSql & " AliqNCM= 0, "
  End If
  
  If Len(Trim(pRedBC_ICM.Text)) > 0 Then
      sSql = sSql & " [Redução ICM]=" & Replace(pRedBC_ICM.Text, ",", ".") & ","
  Else
      sSql = sSql & " [Redução ICM]= 0, "
  End If

  
  If Len(Trim(Perc_ICMS_Saida.Text)) > 0 Then
      sSql = sSql & " [Percentual ICM Saida]=" & Replace(Perc_ICMS_Saida.Text, ",", ".") & ","
  Else
      sSql = sSql & " [Percentual ICM Saida]= 0, "
  End If
  
  If Len(Trim(Perc_ICMS_Entrada.Text)) > 0 Then
      sSql = sSql & " [Percentual ICM Entrada]=" & Replace(Perc_ICMS_Entrada.Text, ",", ".") & ","
  Else
      sSql = sSql & " [Percentual ICM Entrada]= 0, "
  End If
  
  If Len(Trim(Perc_ICMSST_Saida.Text)) > 0 Then
      sSql = sSql & " [Percentual_ICMSST_Saida]=" & Replace(Perc_ICMSST_Saida.Text, ",", ".") & ","
  Else
      sSql = sSql & " [Percentual_ICMSST_Saida]= 0, "
  End If
  
  If Len(Trim(Valor_baseCalculo_ICMSST_Saida.Text)) > 0 Then
      sSql = sSql & " [BaseCalculoICMSST_Saida]=" & Replace(Valor_baseCalculo_ICMSST_Saida.Text, ",", ".") & ","
  Else
      sSql = sSql & " [BaseCalculoICMSST_Saida]= 0, "
  End If
  
  If Len(Trim(Perc_ICMSST_Entrada.Text)) > 0 Then
      sSql = sSql & " [Percentual_ICMSST_Entrada]=" & Replace(Perc_ICMSST_Entrada.Text, ",", ".") & ","
  Else
      sSql = sSql & " [Percentual_ICMSST_Entrada]= 0, "
  End If
  
  If Len(Trim(Valor_baseCalculo_ICMSST_Entrada.Text)) > 0 Then
      sSql = sSql & " [BaseCalculoICMSST_Entrada]=" & Replace(Valor_baseCalculo_ICMSST_Entrada.Text, ",", ".") & ","
  Else
      sSql = sSql & " [BaseCalculoICMSST_Entrada]= 0, "
  End If
  
  If Len(Trim(Perc_IPI_Saida.Text)) > 0 Then
      sSql = sSql & " [Percentual IPI]=" & Replace(Perc_IPI_Saida.Text, ",", ".") & ","
  Else
      sSql = sSql & " [Percentual IPI]= 0, "
  End If
  
  If Len(Trim(Perc_IPI_Entrada.Text)) > 0 Then
      sSql = sSql & " [Percentual_IPI_Entrada]=" & Replace(Perc_IPI_Entrada.Text, ",", ".") & ","
  Else
      sSql = sSql & " [Percentual_IPI_Entrada]= 0, "
  End If
  
  'Tratamento da combo SITUAÇÃO TRIBUTÁRIO DO PIS
  If Len(cmb_situacaoTributariaDoPIS.Text) > 0 Then
      sSql = sSql & " [TipoSituacaoTributariaPIS]=" & Left(cmb_situacaoTributariaDoPIS.Text, 2) & ","
  Else
      sSql = sSql & " [TipoSituacaoTributariaPIS]=0,"
  End If
  'Fim tratamento combo

  If Len(cmb_codigoBeneficio.Text) > 0 Then
      sSql = sSql & " CodigoBeneficio='" & cmb_codigoBeneficio.Text & "' "
  Else
      sSql = sSql & " CodigoBeneficio='SEM CBENEF' "
  End If
  
  sSql = sSql & " WHERE Código='" & LTrim(RTrim(lbl_codigoProduto.Caption)) & "'"
  
  Call ws.BeginTrans
  Db.Execute sSql
  Call ws.CommitTrans
  
  MsgBox "Tributos salvos com sucesso", vbInformation, "Sucesso"
  
  Exit Sub
Erro:
  MsgBox "Erro ao tentar salvar os tributos...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_Load()
On Error GoTo Erro

  Dim rsCodigoBeneficio As Recordset
  Set rsCodigoBeneficio = Db.OpenRecordset("SELECT * FROM CodigoBeneficio ", dbOpenDynaset)
  If rsCodigoBeneficio.RecordCount > 0 Then
      rsCodigoBeneficio.MoveFirst
      While Not rsCodigoBeneficio.EOF
          If Not IsNull(rsCodigoBeneficio.Fields("CodigoBenef").Value) Then
              cmb_codigoBeneficio.AddItem rsCodigoBeneficio.Fields("CodigoBenef").Value
          End If
          rsCodigoBeneficio.MoveNext
      Wend
  Else
      cmb_codigoBeneficio.Enabled = False
      cmb_codigoBeneficio.BackColor = &H8000000C
      lblCodBeneficio.ForeColor = &H8000000C
  End If
  rsCodigoBeneficio.Close
  Set rsCodigoBeneficio = Nothing

  ' Verificar se é empresa SIMPLES NACIONAL ou LUCRO REAL
  If gblnSimplesNacional = True Then
    'Perc_ICMS_Saida.Enabled = False
    'Perc_ICMS_Saida.BackColor = &H8000000C
    'lbl_tit_ICMSSaida.ForeColor = &H8000000C
    'lbl_perc1.ForeColor = &H8000000C
    
    Perc_IPI_Saida.Enabled = False
    Perc_IPI_Saida.BackColor = &H8000000C
    lbl_tit_IPISaida.ForeColor = &H8000000C
    lbl_perc4.ForeColor = &H8000000C
  
    cmb_SituacaoTributaria.Text = SituacaoTributaria
    cmb_SituacaoTributariaEntrada.Text = SituacaoTributariaEntrada
  Else
    If SituacaoTributaria = "" Or SituacaoTributaria = "0" Then
      cmb_SituacaoTributaria.ListIndex = -1
    Else
      cmb_SituacaoTributaria.Text = SituacaoTributaria
    End If
    
    If SituacaoTributariaEntrada = "" Or SituacaoTributariaEntrada = "0" Then
      cmb_SituacaoTributariaEntrada.ListIndex = -1
    Else
      cmb_SituacaoTributariaEntrada.Text = SituacaoTributariaEntrada
    End If
    
    Perc_IPI_Saida.Text = Perc_IPISaida
  End If
    
  Perc_ICMS_Saida.Text = ICMS_Saida

  datCodigoNBM.DatabaseName = gsQuickDBFileName

  lbl_codigoProduto.Caption = CodigoProduto
  lbl_NomeProduto.Caption = nomeProduto
  
  cboCodigoNBM.Text = NCM
  Perc_ICMS_Entrada.Text = ICMS_EntradaDev
  Perc_ICMSST_Saida.Text = ICMSST_Saida
  Valor_baseCalculo_ICMSST_Saida.Text = ICMSST_SaidaBaseCalculo
  Perc_ICMSST_Entrada.Text = ICMSST_Entrada
  Valor_baseCalculo_ICMSST_Entrada.Text = ICMSST_EntradaBaseCalculo
  Perc_IPI_Entrada.Text = Perc_IPIEntrada
  pRedBC_ICM.Text = Perc_RedBC_ICM
  
  
  'Tratamento da combo SITUAÇÃO TRIBUTÁRIO DO PIS
  '  '01 – Operação Tributável - Base de Cálculo = Valor da Operação Alíquota Normal (Cumulativo/Não Cumulativo) = PISAliq
  '  '02 - Operação Tributável - Base de Calculo = Valor da Operação (Alíquota Diferenciada) = PISAliq
  '  '03 - Operação Tributável - Base de Calculo = Quantidade Vendida x Alíquota por Unidade de Produto = PISQtde
  '  '04 - Operação Tributável - Tributação Monofásica - (Alíquota Zero) = PISNT
  '  '06 - Operação Tributável - Alíquota Zero = PISNT
  '  '07 - Operação Isenta da contribuição = PISNT
  '  '08 - Operação Sem Incidência da contribuição = PISNT
  '  '09 - Operação com suspensão da contribuição = PISNT
  '  '99 - Outras Operações = PISOutr
  
  '01 – Operação Tributável com Alíquota Básica
  '02 - Operação Tributável com Alíquota Diferenciada
  '03 - Operação Tributável com Alíquota por Unidade de Medida de Produto
  '04 - Operação Tributável Monofásica - Revenda a Alíquota Zero
  '05 - Operação Tributável por Substituição Tributária
  '06 - Operação Tributável a Alíquota Zero
  '07 - Operação Isenta da Contribuição
  '08 - Operação sem Incidência da Contribuição
  '09 - Operação com Suspensão da Contribuição
  '49 - Outras Operações de Saída
  '50 - Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Tributada no Mercado Interno
  '51 - Operação com Direito a Crédito – Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno
  '52 - Operação com Direito a Crédito - Vinculada Exclusivamente a Receita de Exportação
  '53 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno
  '54 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas no Mercado Interno e de Exportação
  '55 - Operação com Direito a Crédito - Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação
  '56 - Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação
  '60 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Tributada no Mercado Interno
  '61 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Não-Tributada no Mercado Interno
  '62 - Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita de Exportação
  '63 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno
  '64 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas no Mercado Interno e de Exportação
  '65 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação
  '66 - Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação
  '67 - Crédito Presumido - Outras Operações
  '70 - Operação de Aquisição sem Direito a Crédito
  '71 - Operação de Aquisição com Isenção
  '72 - Operação de Aquisição com Suspensão
  '73 - Operação de Aquisição a Alíquota Zero
  '74 - Operação de Aquisição sem Incidência da Contribuição
  '75 - Operação de Aquisição por Substituição Tributária
  '98 - Outras Operações de Entrada
  '99 - Outras Operações
  If Not IsNull(SituacaoTributariaPIS) Then
      If SituacaoTributariaPIS = 0 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 0
      ElseIf SituacaoTributariaPIS = 1 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 1
      ElseIf SituacaoTributariaPIS = 2 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 2
      ElseIf SituacaoTributariaPIS = 3 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 3
      ElseIf SituacaoTributariaPIS = 4 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 4
      ElseIf SituacaoTributariaPIS = 5 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 5
      ElseIf SituacaoTributariaPIS = 6 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 6
      ElseIf SituacaoTributariaPIS = 7 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 7
      ElseIf SituacaoTributariaPIS = 8 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 8
      ElseIf SituacaoTributariaPIS = 9 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 9
      ElseIf SituacaoTributariaPIS = 49 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 10
      ElseIf SituacaoTributariaPIS = 50 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 11
      ElseIf SituacaoTributariaPIS = 51 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 12
      ElseIf SituacaoTributariaPIS = 52 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 13
      ElseIf SituacaoTributariaPIS = 53 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 14
      ElseIf SituacaoTributariaPIS = 54 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 15
      ElseIf SituacaoTributariaPIS = 55 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 16
      ElseIf SituacaoTributariaPIS = 56 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 17
      ElseIf SituacaoTributariaPIS = 60 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 18
      ElseIf SituacaoTributariaPIS = 61 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 19
      ElseIf SituacaoTributariaPIS = 62 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 20
      ElseIf SituacaoTributariaPIS = 63 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 21
      ElseIf SituacaoTributariaPIS = 64 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 22
      ElseIf SituacaoTributariaPIS = 65 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 23
      ElseIf SituacaoTributariaPIS = 66 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 24
      ElseIf SituacaoTributariaPIS = 67 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 25
      ElseIf SituacaoTributariaPIS = 70 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 26
      ElseIf SituacaoTributariaPIS = 71 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 27
      ElseIf SituacaoTributariaPIS = 72 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 28
      ElseIf SituacaoTributariaPIS = 73 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 29
      ElseIf SituacaoTributariaPIS = 74 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 30
      ElseIf SituacaoTributariaPIS = 75 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 31
      ElseIf SituacaoTributariaPIS = 98 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 32
      ElseIf SituacaoTributariaPIS = 99 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 33
      End If
  Else
      cmb_situacaoTributariaDoPIS.ListIndex = -1
  End If
  'Fim tratamento da combo
  
  If Not IsNull(CodigoBeneficio) Then
      cmb_codigoBeneficio.Text = CodigoBeneficio
  Else
      cmb_codigoBeneficio.Text = "SEM CBENEF"
  End If
  
  If Len(Trim(NCM)) > 0 Then
      cboCodigoNBM_LostFocus
  End If

  Exit Sub
Erro:
  MsgBox "Erro na carga da tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub
