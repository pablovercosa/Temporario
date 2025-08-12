VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmParametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"ParamFilial.frx":0000
   ClientHeight    =   8295
   ClientLeft      =   225
   ClientTop       =   510
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1210
   Icon            =   "ParamFilial.frx":00B7
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8295
   ScaleWidth      =   11940
   Begin VB.Data datFilial 
      Caption         =   "Filial"
      Connect         =   "Access 2000"
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
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin SSDataWidgets_B.SSDBCombo Código 
      Bindings        =   "ParamFilial.frx":4EA11
      Height          =   420
      Left            =   720
      TabIndex        =   0
      Top             =   135
      Width           =   945
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3281
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Filial"
      Columns(0).FieldLen=   256
      Columns(1).Width=   7197
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1).FieldLen=   256
      _ExtentX        =   1667
      _ExtentY        =   741
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Filial"
   End
   Begin VB.Frame Frame17 
      BorderStyle     =   0  'None
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
      Left            =   11820
      TabIndex        =   156
      Top             =   480
      Visible         =   0   'False
      Width           =   1005
      Begin VB.TextBox Ult_Nota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         DataField       =   "Última Nota"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   195
         Top             =   30
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Última nota"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   196
         Top             =   30
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.TextBox Ult_Mov 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      DataField       =   "Última Movimentação"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   330
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   172
      Width           =   1455
   End
   Begin VB.Data datOperSaida 
      Caption         =   "Oper. Saida"
      Connect         =   "Access 2000"
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
      Left            =   6720
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Op_Saída"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data5 
      Caption         =   "Conta"
      Connect         =   "Access 2000"
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
      Left            =   4800
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Cliente"
      Connect         =   "Access 2000"
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
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      DataField       =   "Nome"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1725
      MaxLength       =   35
      TabIndex        =   1
      Top             =   135
      Width           =   6690
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   7110
      Left            =   90
      TabIndex        =   100
      Top             =   765
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   12541
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
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
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "ParamFilial.frx":4EA29
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblTitle(9)"
      Tab(0).Control(1)=   "lblTitle(10)"
      Tab(0).Control(2)=   "lblTitle(11)"
      Tab(0).Control(3)=   "lblTitle(63)"
      Tab(0).Control(4)=   "lbl_TaxaMultaParcelaVencida"
      Tab(0).Control(5)=   "lbl_multaDiasAposParcelaVencida"
      Tab(0).Control(6)=   "Line1"
      Tab(0).Control(7)=   "Line2"
      Tab(0).Control(8)=   "mskTaxaMultaParcelaVencida"
      Tab(0).Control(9)=   "mskTaxaDesconto"
      Tab(0).Control(10)=   "Juros"
      Tab(0).Control(11)=   "Frame2(0)"
      Tab(0).Control(12)=   "Frame2(1)"
      Tab(0).Control(13)=   "Frame6"
      Tab(0).Control(14)=   "txtDiasBloqueioVenda"
      Tab(0).Control(15)=   "txt_multaDiasAposParcelaVencida"
      Tab(0).Control(16)=   "chk_cobrarMulta"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Venda Rápida"
      TabPicture(1)   =   "ParamFilial.frx":4EA45
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame19"
      Tab(1).Control(1)=   "Frame18"
      Tab(1).Control(2)=   "Frame15"
      Tab(1).Control(3)=   "chkProcuraCodigoENome"
      Tab(1).Control(4)=   "Frame7"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Saídas"
      TabPicture(2)   =   "ParamFilial.frx":4EA61
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblNomeOperSaida_S"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblTitle(22)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboOperSaida_S"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame13(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame13(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame16"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame10"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Nota Fiscal"
      TabPicture(3)   =   "ParamFilial.frx":4EA7D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtUltimaNFCe"
      Tab(3).Control(1)=   "txtUltimaNFe"
      Tab(3).Control(2)=   "fraCabecalhoListaPreco"
      Tab(3).Control(3)=   "fraAliquotaAproveitamentoCreditoIcms"
      Tab(3).Control(4)=   "fraNFe"
      Tab(3).Control(5)=   "fraCC"
      Tab(3).Control(6)=   "fraTicket"
      Tab(3).Control(7)=   "Frame4"
      Tab(3).Control(8)=   "lblUltimaNFCe"
      Tab(3).Control(9)=   "lblUltimaNFe"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Outros"
      TabPicture(4)   =   "ParamFilial.frx":4EA99
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Qtde_Balança"
      Tab(4).Control(1)=   "chk_viaRDP"
      Tab(4).Control(2)=   "chk_viaRDP_ticket"
      Tab(4).Control(3)=   "chkPermitir5Casas"
      Tab(4).Control(4)=   "cmb_casasDecimaisValorUnitario"
      Tab(4).Control(5)=   "Frame5"
      Tab(4).Control(6)=   "Frame1"
      Tab(4).Control(7)=   "Pesq3"
      Tab(4).Control(8)=   "Pesq2"
      Tab(4).Control(9)=   "Pesq1"
      Tab(4).Control(10)=   "Frame12"
      Tab(4).Control(11)=   "Frame3"
      Tab(4).Control(12)=   "fraImpostos"
      Tab(4).Control(13)=   "chk0aEsquerda"
      Tab(4).Control(14)=   "chkAlterVendedorCliFor"
      Tab(4).Control(15)=   "SenhaGerReimpTicket"
      Tab(4).Control(16)=   "SenhaGerVendaAtraso"
      Tab(4).Control(17)=   "NaoPermiteDuplicarCNPJ"
      Tab(4).Control(18)=   "CommonDialog1"
      Tab(4).Control(19)=   "lblTitle(12)"
      Tab(4).Control(20)=   "lblTitle(32)"
      Tab(4).Control(21)=   "Label1"
      Tab(4).Control(22)=   "lblTitle(58)"
      Tab(4).Control(23)=   "lblTitle(57)"
      Tab(4).Control(24)=   "lblTitle(56)"
      Tab(4).ControlCount=   25
      Begin VB.Frame Frame10 
         Caption         =   "Transferência"
         Height          =   2235
         Left            =   6060
         TabIndex        =   295
         Top             =   4590
         Width           =   5415
         Begin VB.ComboBox cboTabPrecosTransf 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   180
            TabIndex        =   302
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtOpEntradaTransf 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   299
            TabStop         =   0   'False
            Top             =   1125
            Width           =   3825
         End
         Begin VB.TextBox txtOpSaidaTransf 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   296
            TabStop         =   0   'False
            Top             =   570
            Width           =   3825
         End
         Begin SSDataWidgets_B.SSDBCombo cboOpSaidaTransf 
            Bindings        =   "ParamFilial.frx":4EAB5
            Height          =   285
            Left            =   210
            TabIndex        =   297
            Top             =   570
            Width           =   1215
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
            BackColorOdd    =   16777152
            Columns(0).Width=   3200
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
            DataFieldToDisplay=   "Código"
         End
         Begin SSDataWidgets_B.SSDBCombo cboOpEntradaTransf 
            Bindings        =   "ParamFilial.frx":4EAD4
            Height          =   285
            Left            =   210
            TabIndex        =   300
            Top             =   1125
            Width           =   1215
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
            BackColorOdd    =   16777152
            Columns(0).Width=   3200
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
            DataFieldToDisplay=   "Código"
         End
         Begin VB.Label lblTitle 
            Caption         =   "Tabela de Preços"
            Height          =   255
            Index           =   35
            Left            =   180
            TabIndex        =   303
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Operação de Entrada para transferência"
            Height          =   195
            Index           =   34
            Left            =   210
            TabIndex        =   301
            Top             =   900
            Width           =   2925
         End
         Begin VB.Label lblTitle 
            Caption         =   "Operação de Saída para tranferência"
            Height          =   255
            Index           =   33
            Left            =   210
            TabIndex        =   298
            Top             =   330
            Width           =   2925
         End
      End
      Begin VB.CheckBox chk_cobrarMulta 
         Appearance      =   0  'Flat
         Caption         =   "Cobrar multa após vencimento de parcela"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   294
         Top             =   6000
         Width           =   3315
      End
      Begin VB.TextBox txt_multaDiasAposParcelaVencida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   -73440
         TabIndex        =   293
         Text            =   "0"
         Top             =   6300
         Width           =   435
      End
      Begin VB.TextBox Qtde_Balança 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   -64110
         MaxLength       =   1
         TabIndex        =   289
         ToolTipText     =   "Nota: O Valor deste campo irá formatar a quantidade de caracteres após a "","" para o cálculo na coluna de Qtde em VR."
         Top             =   2070
         Width           =   645
      End
      Begin VB.CheckBox chk_viaRDP 
         Appearance      =   0  'Flat
         Caption         =   "Impressora acesso via RDP (CUPOM)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -67140
         TabIndex        =   286
         Top             =   3930
         Width           =   3255
      End
      Begin VB.CheckBox chk_viaRDP_ticket 
         Appearance      =   0  'Flat
         Caption         =   "Impressora acesso via RDP (TICKET)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -67140
         TabIndex        =   305
         Top             =   3630
         Width           =   3255
      End
      Begin VB.CheckBox chkPermitir5Casas 
         Appearance      =   0  'Flat
         Caption         =   "Na tela de Entradas: Permitir 5 casas após a vírgula em Preço Unitário"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -67140
         TabIndex        =   285
         Top             =   960
         Width           =   3705
      End
      Begin VB.ComboBox cmb_casasDecimaisValorUnitario 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "ParamFilial.frx":4EAF5
         Left            =   -64410
         List            =   "ParamFilial.frx":4EB02
         Style           =   2  'Dropdown List
         TabIndex        =   284
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Situação Tributária do PIS"
         Height          =   705
         Left            =   -74895
         TabIndex        =   281
         Top             =   4350
         Width           =   11475
         Begin VB.ComboBox cmb_situacaoTributariaDoPIS 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            ItemData        =   "ParamFilial.frx":4EB0F
            Left            =   120
            List            =   "ParamFilial.frx":4EB79
            Style           =   2  'Dropdown List
            TabIndex        =   282
            Top             =   240
            Width           =   11175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Configurações da impressora  "
         Height          =   2070
         Left            =   -63420
         TabIndex        =   264
         Top             =   2880
         Visible         =   0   'False
         Width           =   1365
         Begin VB.TextBox c_oito3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1455
            MaxLength       =   3
            TabIndex        =   273
            Top             =   2505
            Width           =   585
         End
         Begin VB.TextBox c_oito2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   825
            MaxLength       =   3
            TabIndex        =   272
            Top             =   2505
            Width           =   585
         End
         Begin VB.TextBox c_oito1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            MaxLength       =   3
            TabIndex        =   271
            Top             =   2505
            Width           =   585
         End
         Begin VB.TextBox c_comp3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1455
            MaxLength       =   3
            TabIndex        =   270
            Top             =   1725
            Width           =   585
         End
         Begin VB.TextBox c_comp2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   825
            MaxLength       =   3
            TabIndex        =   269
            Top             =   1725
            Width           =   585
         End
         Begin VB.TextBox c_comp1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            Height          =   315
            Left            =   210
            MaxLength       =   3
            TabIndex        =   268
            Top             =   1725
            Width           =   585
         End
         Begin VB.TextBox c_comp_pag1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   210
            MaxLength       =   3
            TabIndex        =   267
            Top             =   3540
            Width           =   585
         End
         Begin VB.TextBox c_comp_pag2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   825
            MaxLength       =   3
            TabIndex        =   266
            Top             =   3540
            Width           =   585
         End
         Begin VB.TextBox c_comp_pag3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1455
            MaxLength       =   3
            TabIndex        =   265
            Top             =   3540
            Width           =   585
         End
         Begin VB.Label lblTitle 
            Caption         =   "Códigos para impressão 1/8"""
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
            Index           =   54
            Left            =   210
            TabIndex        =   277
            Top             =   2205
            Width           =   2295
         End
         Begin VB.Label lblTitle 
            Caption         =   "Códigos para impressão comprimida"
            Height          =   255
            Index           =   53
            Left            =   210
            TabIndex        =   276
            Top             =   1440
            Width           =   2265
         End
         Begin VB.Label lblTitle 
            Caption         =   "Códigos para definição do comprimento da página em polegadas"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   55
            Left            =   210
            TabIndex        =   275
            Top             =   3000
            Width           =   3315
         End
         Begin VB.Label lblTitle 
            Caption         =   $"ParamFilial.frx":4F501
            Height          =   885
            Index           =   52
            Left            =   180
            TabIndex        =   274
            Top             =   420
            Width           =   3210
         End
      End
      Begin VB.TextBox Pesq3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   -65115
         MaxLength       =   10
         TabIndex        =   263
         Top             =   6435
         Width           =   1485
      End
      Begin VB.TextBox Pesq2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   -65115
         MaxLength       =   10
         TabIndex        =   262
         Top             =   6000
         Width           =   1485
      End
      Begin VB.TextBox Pesq1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   -65115
         MaxLength       =   10
         TabIndex        =   261
         Top             =   5550
         Width           =   1485
      End
      Begin VB.Frame Frame12 
         Caption         =   "Etiqueta de Roupa"
         Height          =   1125
         Left            =   -74895
         TabIndex        =   255
         Top             =   3225
         Width           =   7620
         Begin VB.TextBox Mens_Etiq2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   4965
            MaxLength       =   20
            TabIndex        =   258
            Top             =   705
            Width           =   2415
         End
         Begin VB.TextBox Mens_Etiq1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2205
            MaxLength       =   20
            TabIndex        =   257
            Top             =   705
            Width           =   2415
         End
         Begin VB.TextBox Mensagem_Troca 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2205
            MaxLength       =   140
            TabIndex        =   256
            Top             =   285
            Width           =   5175
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Mensagens"
            Height          =   195
            Index           =   60
            Left            =   420
            TabIndex        =   260
            Top             =   750
            Width           =   810
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem para troca"
            Height          =   195
            Index           =   59
            Left            =   420
            TabIndex        =   259
            Top             =   330
            Width           =   1560
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tabelas de Preço para a tela de Consulta"
         Height          =   2850
         Left            =   -74895
         TabIndex        =   248
         Top             =   375
         Width           =   3330
         Begin VB.ComboBox Com_Tab3 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   254
            Top             =   1125
            Width           =   3015
         End
         Begin VB.ComboBox Com_Tab1 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   253
            Top             =   285
            Width           =   3015
         End
         Begin VB.ComboBox Com_Tab4 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   252
            Top             =   1545
            Width           =   3015
         End
         Begin VB.ComboBox Com_Tab5 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   251
            Top             =   1965
            Width           =   3015
         End
         Begin VB.ComboBox Com_Tab6 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   250
            Top             =   2385
            Width           =   3015
         End
         Begin VB.ComboBox Com_Tab2 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   249
            Top             =   705
            Width           =   3015
         End
      End
      Begin VB.Frame fraImpostos 
         Caption         =   "Percentuais de Impostos"
         ForeColor       =   &H00000000&
         Height          =   2865
         Left            =   -71475
         TabIndex        =   237
         Top             =   375
         Width           =   4200
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "QuickStore acesso via RDP"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   304
            Top             =   0
            Width           =   2445
         End
         Begin VB.TextBox txtCOFINS 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   242
            Top             =   1170
            Width           =   975
         End
         Begin VB.TextBox txtPIS 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   3120
            MaxLength       =   5
            TabIndex        =   241
            Top             =   375
            Width           =   975
         End
         Begin VB.TextBox txtIRRF 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   240
            Top             =   1980
            Width           =   975
         End
         Begin VB.TextBox txtCSLL 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   239
            Top             =   375
            Width           =   975
         End
         Begin VB.TextBox txtValorIsencaoPisCofinsCsll 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   3120
            TabIndex        =   238
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "CSLL (%)"
            Height          =   195
            Index           =   48
            Left            =   120
            TabIndex        =   247
            Top             =   420
            Width           =   675
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "COFINS (%)"
            Height          =   195
            Index           =   49
            Left            =   120
            TabIndex        =   246
            Top             =   1215
            Width           =   900
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "PIS (%)"
            Height          =   195
            Index           =   50
            Left            =   2160
            TabIndex        =   245
            Top             =   420
            Width           =   570
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "IRRF (%)"
            Height          =   195
            Index           =   51
            Left            =   120
            TabIndex        =   244
            Top             =   2025
            Width           =   690
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Isenção ($)"
            Height          =   195
            Index           =   64
            Left            =   2160
            TabIndex        =   243
            Top             =   1215
            Width           =   825
         End
      End
      Begin VB.CheckBox chk0aEsquerda 
         Appearance      =   0  'Flat
         Caption         =   "Permitir 0 ""zero"" a esquerda para código de produtos (cadastro, vendas e compras)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74835
         TabIndex        =   236
         Top             =   5475
         Width           =   6900
      End
      Begin VB.CheckBox chkAlterVendedorCliFor 
         Appearance      =   0  'Flat
         Caption         =   "Apenas o Superusuário pode alterar o campo Vendedor no cadastro de Clientes / Fornecedores"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74835
         TabIndex        =   235
         Top             =   5160
         Width           =   7695
      End
      Begin VB.CheckBox SenhaGerReimpTicket 
         Appearance      =   0  'Flat
         Caption         =   "Exigir senha do gerente para reimprimir ticket "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74835
         TabIndex        =   234
         Top             =   5880
         Width           =   4095
      End
      Begin VB.CheckBox SenhaGerVendaAtraso 
         Appearance      =   0  'Flat
         Caption         =   "Exigir senha do gerente ao efetuar vendas para clientes em atraso"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74835
         TabIndex        =   233
         Top             =   6600
         Width           =   5490
      End
      Begin VB.CheckBox NaoPermiteDuplicarCNPJ 
         Appearance      =   0  'Flat
         Caption         =   "Não permitir CNPJ / CPF duplicados para Cli / For"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74835
         TabIndex        =   232
         Top             =   6240
         Width           =   4095
      End
      Begin VB.TextBox txtUltimaNFCe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   -68280
         TabIndex        =   228
         Top             =   6690
         Width           =   1485
      End
      Begin VB.TextBox txtUltimaNFe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         DataField       =   "Última Nota"
         DataSource      =   "Data1"
         Height          =   330
         Left            =   -64845
         MaxLength       =   9
         TabIndex        =   213
         Top             =   6690
         Width           =   1455
      End
      Begin VB.Frame fraCabecalhoListaPreco 
         Caption         =   "Cabeçalho para Listas de Preço"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   212
         Top             =   480
         Width           =   5175
         Begin VB.TextBox Lista 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   120
            MaxLength       =   80
            TabIndex        =   106
            Top             =   1590
            Width           =   4980
         End
         Begin VB.TextBox Lista 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   120
            MaxLength       =   80
            TabIndex        =   105
            Top             =   1260
            Width           =   4980
         End
         Begin VB.TextBox Lista 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   80
            TabIndex        =   104
            Top             =   930
            Width           =   4980
         End
         Begin VB.TextBox Lista 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   80
            TabIndex        =   103
            Top             =   600
            Width           =   4980
         End
         Begin VB.TextBox Lista 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   80
            TabIndex        =   102
            Top             =   270
            Width           =   4980
         End
      End
      Begin VB.Frame fraAliquotaAproveitamentoCreditoIcms 
         Caption         =   "Alíquota para aproveitamento do crédito de ICMS"
         Height          =   855
         Left            =   -74880
         TabIndex        =   211
         Top             =   4920
         Width           =   5175
         Begin VB.TextBox txtAliquotaAproveitamentoCreditoIcms 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   3840
            TabIndex        =   113
            Text            =   "0"
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraNFe 
         Caption         =   "Nota Fiscal Eletrônica"
         Height          =   6150
         Left            =   -69630
         TabIndex        =   204
         Top             =   480
         Width           =   6255
         Begin VB.TextBox txtBancoPDV 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            TabIndex        =   230
            Top             =   4110
            Width           =   5055
         End
         Begin VB.CommandButton btnProcurar 
            BackColor       =   &H00C0FFC0&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5340
            Style           =   1  'Graphical
            TabIndex        =   229
            Top             =   4110
            Width           =   585
         End
         Begin VB.ComboBox cboPadraoArquivoIntegracao 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            ItemData        =   "ParamFilial.frx":4F589
            Left            =   3150
            List            =   "ParamFilial.frx":4F596
            Style           =   2  'Dropdown List
            TabIndex        =   223
            Top             =   5385
            Width           =   2775
         End
         Begin VB.TextBox txtLayoutEnvio 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   3870
            TabIndex        =   218
            Top             =   840
            Width           =   2235
         End
         Begin VB.ComboBox cboCodigoRegimeTributario 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            ItemData        =   "ParamFilial.frx":4F5AC
            Left            =   120
            List            =   "ParamFilial.frx":4F5B9
            Style           =   2  'Dropdown List
            TabIndex        =   215
            Top             =   1515
            Width           =   3705
         End
         Begin VB.CheckBox chkHabilitarNotaFiscalEletronica 
            Appearance      =   0  'Flat
            Caption         =   "Habilitar Uso de Nota Fiscal Eletrônica"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   285
            Width           =   3255
         End
         Begin VB.CommandButton cmdSelecionarPastaNfe 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Index           =   1
            Left            =   5370
            Picture         =   "ParamFilial.frx":4F61C
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   4740
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton cmdSelecionarPastaNfe 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Index           =   0
            Left            =   5340
            Picture         =   "ParamFilial.frx":4F766
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   3465
            Width           =   585
         End
         Begin VB.TextBox txtPastaRetornoNfe 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            TabIndex        =   121
            Top             =   4740
            Width           =   5055
         End
         Begin VB.TextBox txtPastaEnvioNfe 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            TabIndex        =   119
            Top             =   3465
            Width           =   5055
         End
         Begin VB.ComboBox cboModDetBaseCalculoIcmsSt 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            ItemData        =   "ParamFilial.frx":4F8B0
            Left            =   150
            List            =   "ParamFilial.frx":4F8C9
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   2835
            Width           =   5985
         End
         Begin VB.ComboBox cboModDetBaseCalculoIcms 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            ItemData        =   "ParamFilial.frx":4F98F
            Left            =   120
            List            =   "ParamFilial.frx":4F99F
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   2190
            Width           =   6015
         End
         Begin VB.ComboBox cboFormatoImpressaoDanfeNfe 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            ItemData        =   "ParamFilial.frx":4FA0D
            Left            =   3870
            List            =   "ParamFilial.frx":4FA17
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   1515
            Width           =   2265
         End
         Begin VB.ComboBox cboAmbienteNfe 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            ItemData        =   "ParamFilial.frx":4FA36
            Left            =   120
            List            =   "ParamFilial.frx":4FA40
            Style           =   2  'Dropdown List
            TabIndex        =   115
            Top             =   840
            Width           =   3705
         End
         Begin MSMask.MaskEdBox txtPercentualSimplesNacional 
            Height          =   315
            Left            =   1710
            TabIndex        =   219
            Top             =   5355
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648384
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPercentualReducaoBC_SN 
            Height          =   315
            Left            =   3930
            TabIndex        =   222
            Top             =   5790
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###.00"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBancoPDV 
            Caption         =   "Caminho Banco PDV"
            Height          =   210
            Left            =   120
            TabIndex        =   231
            Top             =   3870
            Width           =   1605
         End
         Begin VB.Label lblPadraoArquivoIntegracao 
            AutoSize        =   -1  'True
            Caption         =   "Padrão do Arquivo de Integração"
            Height          =   195
            Index           =   0
            Left            =   3000
            TabIndex        =   224
            Top             =   5145
            Width           =   2400
         End
         Begin VB.Label lblPercentualReducaoBC_SN 
            AutoSize        =   -1  'True
            Caption         =   "% Redução Base de Cálculo - Simples Nacional"
            Height          =   195
            Index           =   65
            Left            =   120
            TabIndex        =   221
            Top             =   5820
            Width           =   3330
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "% Simples Nacional"
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   220
            Top             =   5385
            Width           =   1380
         End
         Begin VB.Label lblLayoutEnvio 
            AutoSize        =   -1  'True
            Caption         =   "Layout de Envio"
            Height          =   195
            Index           =   29
            Left            =   3870
            TabIndex        =   217
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label lblCodigoRegimeTributario 
            AutoSize        =   -1  'True
            Caption         =   "Código do Regime Tributario"
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   216
            Top             =   1275
            Width           =   2025
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Endereço Site Benefix"
            Height          =   195
            Index           =   78
            Left            =   120
            TabIndex        =   210
            Top             =   4500
            Width           =   1575
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Pasta XML"
            Height          =   195
            Index           =   77
            Left            =   120
            TabIndex        =   209
            Top             =   3225
            Width           =   735
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Modalidade de determinação da Base de Cálculo do ICMS ST"
            Height          =   195
            Index           =   76
            Left            =   120
            TabIndex        =   208
            Top             =   2595
            Width           =   4320
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Modalidade de determinação da Base de Cálculo do ICMS"
            Height          =   195
            Index           =   75
            Left            =   120
            TabIndex        =   207
            Top             =   1950
            Width           =   4095
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Formato Impressão do DANFE"
            Height          =   195
            Index           =   74
            Left            =   3900
            TabIndex        =   206
            Top             =   1275
            Width           =   2160
         End
         Begin VB.Label lblIdentificacaoAmbiente 
            AutoSize        =   -1  'True
            Caption         =   "Identificação do Ambiente"
            Height          =   195
            Index           =   73
            Left            =   120
            TabIndex        =   205
            Top             =   600
            Width           =   1875
         End
      End
      Begin VB.Frame fraCC 
         Caption         =   "Adição de Centros de Custo do Plano de Contas"
         Height          =   735
         Left            =   -74880
         TabIndex        =   191
         Top             =   4080
         Width           =   5175
         Begin VB.CommandButton cmdPlanodeContas 
            BackColor       =   &H00C0FFFF&
            Caption         =   "P&lano de Contas"
            Height          =   375
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   112
            ToolTipText     =   "Ao clicar neste botão o sistema carregará a tela de plano de contas"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraTicket 
         Caption         =   "Ticket Padrão p/ Fatur. Automático"
         Height          =   855
         Left            =   -74880
         TabIndex        =   190
         Top             =   5760
         Width           =   5175
         Begin VB.CommandButton cmdProcurarTicket 
            Height          =   375
            Left            =   4440
            Picture         =   "ParamFilial.frx":4FA63
            Style           =   1  'Graphical
            TabIndex        =   124
            Top             =   330
            Width           =   495
         End
         Begin MSComDlg.CommonDialog cdgFileOpen2 
            Left            =   120
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtTicket 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1680
            TabIndex        =   123
            Top             =   330
            Width           =   2655
         End
      End
      Begin VB.TextBox txtDiasBloqueioVenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   -64260
         TabIndex        =   183
         Text            =   "0"
         Top             =   6270
         Width           =   495
      End
      Begin VB.Frame Frame16 
         Caption         =   "Consignação"
         Height          =   3435
         Left            =   6060
         TabIndex        =   101
         Top             =   1020
         Width           =   5415
         Begin VB.Data datOperacaoEntrada 
            Caption         =   "datOperacaoEntrada"
            Connect         =   "Access 2000"
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
            Left            =   1920
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT Código, Nome FROM [Operações Entrada]"
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtNomeOperacaoFechamento 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   1200
            Width           =   3855
         End
         Begin SSDataWidgets_B.SSDBCombo cboOperacaoFechamento 
            Bindings        =   "ParamFilial.frx":4FBAD
            Height          =   285
            Left            =   210
            TabIndex        =   97
            Top             =   1200
            Width           =   1215
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
            BackColorOdd    =   16777152
            Columns(0).Width=   3200
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
            DataFieldToDisplay=   "Código"
         End
         Begin VB.Data datCaixa 
            Caption         =   "datCaixa"
            Connect         =   "Access 2000"
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
            Left            =   4200
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT DISTINCTROW Caixa, Descrição FROM [Caixas em Uso]"
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Data datOperacaoSaida 
            Caption         =   "datOperacaoSaida"
            Connect         =   "Access 2000"
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
            Left            =   3120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT Código, Nome FROM [Operações Saída]"
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtUltimaConsignacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3630
            Locked          =   -1  'True
            TabIndex        =   173
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2880
            Width           =   1695
         End
         Begin VB.ComboBox cboTabelaPrecoConsignacao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   210
            TabIndex        =   99
            Top             =   2865
            Width           =   3135
         End
         Begin VB.TextBox txtCaixa 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   172
            TabStop         =   0   'False
            Top             =   1740
            Width           =   3855
         End
         Begin VB.TextBox txtNomeOperacaoEntrada 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   2295
            Width           =   3825
         End
         Begin VB.TextBox txtNomeOperacaoSaida 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   170
            TabStop         =   0   'False
            Top             =   645
            Width           =   3855
         End
         Begin SSDataWidgets_B.SSDBCombo cboOperacaoSaida 
            Bindings        =   "ParamFilial.frx":4FBCC
            Height          =   285
            Left            =   210
            TabIndex        =   96
            Top             =   645
            Width           =   1215
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
            BackColorOdd    =   16777152
            Columns(0).Width=   3200
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
            DataFieldToDisplay=   "Código"
         End
         Begin SSDataWidgets_B.SSDBCombo cboCaixa 
            Bindings        =   "ParamFilial.frx":4FBEB
            Height          =   285
            Left            =   210
            TabIndex        =   98
            Top             =   1740
            Width           =   1215
            DataFieldList   =   "Caixa"
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
            Columns(0).Width=   3200
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
            DataFieldToDisplay=   "Caixa"
         End
         Begin SSDataWidgets_B.SSDBCombo cboOperacaoEntrada 
            Bindings        =   "ParamFilial.frx":4FC02
            Height          =   285
            Left            =   210
            TabIndex        =   95
            Top             =   2295
            Width           =   1215
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
            BackColorOdd    =   16777152
            Columns(0).Width=   3200
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
            DataFieldToDisplay=   "Código"
         End
         Begin VB.Label lblTitle 
            Caption         =   "Operação para Venda Estadual de consignado"
            Height          =   285
            Index           =   25
            Left            =   210
            TabIndex        =   179
            Top             =   960
            Width           =   4155
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            Caption         =   "Última Consignação"
            Height          =   255
            Index           =   28
            Left            =   3555
            TabIndex        =   178
            Top             =   2625
            Width           =   1695
         End
         Begin VB.Label lblTitle 
            Caption         =   "Tabela de Preços"
            Height          =   255
            Index           =   27
            Left            =   210
            TabIndex        =   177
            Top             =   2625
            Width           =   2175
         End
         Begin VB.Label lblTitle 
            Caption         =   "Caixa"
            Height          =   240
            Index           =   26
            Left            =   210
            TabIndex        =   176
            Top             =   1500
            Width           =   1455
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Op. de Entrada para consignados de mercadorias de terceiros"
            Height          =   195
            Index           =   23
            Left            =   210
            TabIndex        =   175
            Top             =   2070
            Width           =   4455
         End
         Begin VB.Label lblTitle 
            Caption         =   "Operação de Saída para consignado"
            Height          =   255
            Index           =   24
            Left            =   210
            TabIndex        =   174
            Top             =   405
            Width           =   2925
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Opções operacionais"
         Height          =   3900
         Index           =   0
         Left            =   120
         TabIndex        =   167
         Top             =   360
         Width           =   5745
         Begin VB.CheckBox chkProdutoNomeNFe 
            Appearance      =   0  'Flat
            Caption         =   "Editar nome do PRODUTO"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2940
            TabIndex        =   308
            Top             =   3480
            Width           =   2535
         End
         Begin VB.CheckBox chkComandas 
            Appearance      =   0  'Flat
            Caption         =   "Trabalhar com Comandas"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   226
            Top             =   3540
            Width           =   2295
         End
         Begin VB.CheckBox chkCFOP 
            Appearance      =   0  'Flat
            Caption         =   "Exibir coluna de CFOP para Produtos e Serviços na tela de Saídas"
            ForeColor       =   &H80000008&
            Height          =   675
            Left            =   2385
            TabIndex        =   194
            Top             =   1050
            Width           =   3285
         End
         Begin VB.CheckBox chkVendaSemEstoqueSaidas 
            Appearance      =   0  'Flat
            Caption         =   "Permite movimentação se não houver estoque suficiente"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            TabIndex        =   83
            Top             =   1725
            Value           =   1  'Checked
            Width           =   5150
         End
         Begin VB.CheckBox chkExibirFabricante 
            Appearance      =   0  'Flat
            Caption         =   "Exibir a coluna Fabricante na lista do campo código de produto nas telas de Saídas e Venda Rápida"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            TabIndex        =   86
            Top             =   3015
            Width           =   5055
         End
         Begin VB.CheckBox chkVerificaLimiteCli 
            Appearance      =   0  'Flat
            Caption         =   "Verificar o limite de compra do cliente antes da gravação da saída. (Obs: Válido para Venda Rápida também)"
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   120
            TabIndex        =   85
            ToolTipText     =   "Importante para empresas que trabalham com pronta entrega."
            Top             =   2445
            Width           =   4815
         End
         Begin VB.CheckBox chkSaida_Descr_Adicional 
            Appearance      =   0  'Flat
            Caption         =   "Incluir coluna para preenchimento da Descrição Adicional"
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   120
            TabIndex        =   84
            Top             =   2010
            Width           =   5070
         End
         Begin VB.CheckBox Crédito_Saídas 
            Appearance      =   0  'Flat
            Caption         =   "Verifica limite de crédito"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox Máximo_Serviço 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   4155
            MaxLength       =   3
            TabIndex        =   80
            Top             =   680
            Width           =   1215
         End
         Begin VB.TextBox Máximo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            DataField       =   "Linhas Digitação"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   4155
            MaxLength       =   3
            TabIndex        =   79
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkAlteraPreco 
            Appearance      =   0  'Flat
            Caption         =   "Permite Alterar Preço"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   1425
            Value           =   1  'Checked
            Width           =   2115
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(máx. 255)"
            Height          =   195
            Index           =   62
            Left            =   4185
            TabIndex        =   188
            Top             =   120
            Width           =   825
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Máximo de linhas para serviços na tela de saídas :"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   169
            Top             =   720
            Width           =   3585
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Máximo de linhas para produtos na tela de saídas :"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   168
            Top             =   405
            Width           =   3645
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Opções operacionais"
         Height          =   4725
         Left            =   -74880
         TabIndex        =   48
         Top             =   2280
         Width           =   5175
         Begin VB.CheckBox chkVRUtilizarTicketModoRelatorio 
            Appearance      =   0  'Flat
            Caption         =   "Utilizar Ticket em modo Relatório"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   4305
            Width           =   4575
         End
         Begin VB.CheckBox chkImprimeNotaMovEfetivada 
            Appearance      =   0  'Flat
            Caption         =   "Imprimir nota somente para movimentações efetivadas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   3975
            Width           =   4710
         End
         Begin VB.CheckBox chkImprimeTicketMovEfetivada 
            Appearance      =   0  'Flat
            Caption         =   "Imprimir ticket somente para movimentações efetivadas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   3645
            Width           =   4755
         End
         Begin VB.CheckBox chkVR_GravarExigeSenhaVend 
            Appearance      =   0  'Flat
            Caption         =   "Exigir senha do vendedor sempre que gravar"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   2025
            Width           =   3735
         End
         Begin VB.TextBox VR_Desconto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   3930
            MaxLength       =   5
            TabIndex        =   55
            Top             =   2445
            Width           =   1155
         End
         Begin VB.CheckBox VR_Cadastra_Cliente 
            Appearance      =   0  'Flat
            Caption         =   "Permite cadastrar novos clientes"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   705
            Width           =   3495
         End
         Begin VB.CheckBox VR_Altera_Cliente 
            Appearance      =   0  'Flat
            Caption         =   "Permite alterar cliente"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   375
            Width           =   2895
         End
         Begin VB.CheckBox Sem_Estoque 
            Appearance      =   0  'Flat
            Caption         =   "Permite venda rápida se não houver estoque suficiente"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1035
            Width           =   4635
         End
         Begin VB.CheckBox Crédito_Venda_Rápida 
            Appearance      =   0  'Flat
            Caption         =   "Verifica limite de crédito"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1365
            Width           =   2625
         End
         Begin VB.CheckBox VR_Recebimento_Normal 
            Appearance      =   0  'Flat
            Caption         =   "Permite usar tela de recebimento normal"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1695
            Width           =   3600
         End
         Begin VB.TextBox VR_Intervalo_Parc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   3930
            MaxLength       =   3
            TabIndex        =   56
            Top             =   2850
            Width           =   1155
         End
         Begin VB.CheckBox VR_Permite_Desconto 
            Appearance      =   0  'Flat
            Caption         =   "Permite fornecer descontos através da coluna Desconto %"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   3315
            Width           =   4845
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            Caption         =   "Desconto máximo sobre o total da nota, em % :"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   166
            Top             =   2505
            Width           =   3660
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            Caption         =   "Intervalo padrão entre parcelas (em dias) :"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   165
            Top             =   2910
            Width           =   3495
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Opções de tela"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   157
         Top             =   360
         Width           =   6615
         Begin VB.CheckBox chkVR_Tela_CheckOut 
            Appearance      =   0  'Flat
            Caption         =   "Estilo CheckOut"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3480
            TabIndex        =   38
            Top             =   240
            Width           =   1485
         End
         Begin VB.CheckBox VR_Altera_Preço 
            Appearance      =   0  'Flat
            Caption         =   "Permite alterar preços"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3480
            TabIndex        =   42
            Top             =   1150
            Width           =   2895
         End
         Begin VB.TextBox VR_Linhas 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1725
            MaxLength       =   3
            TabIndex        =   37
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox VR_Cód_Operação 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1725
            MaxLength       =   4
            TabIndex        =   39
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox VR_Altera_Tabela 
            Appearance      =   0  'Flat
            Caption         =   "Permite alterar tabela de preços"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3480
            TabIndex        =   41
            Top             =   900
            Width           =   2895
         End
         Begin SSDataWidgets_B.SSDBCombo VR_Combo_Preço 
            Height          =   285
            Left            =   1725
            TabIndex        =   40
            Top             =   1005
            Width           =   1695
            DataFieldList   =   "Column 0"
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
            BackColorOdd    =   16777152
            RowHeight       =   423
            Columns(0).Width=   3200
            Columns(0).Caption=   "Tabela"
            Columns(0).Name =   "Tabela"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   12648447
         End
         Begin SSDataWidgets_B.SSDBCombo VR_Combo_Cliente 
            Bindings        =   "ParamFilial.frx":4FC23
            DataSource      =   "Data2"
            Height          =   285
            Left            =   1725
            TabIndex        =   43
            Top             =   1440
            Width           =   1695
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
            Columns(0).Width=   3200
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   12648447
         End
         Begin VB.Label VR_Nome_Cliente 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3480
            TabIndex        =   164
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Cliente Padrão"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   163
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Máximo de linhas"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   162
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Código de Operação"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   161
            Top             =   600
            Width           =   1470
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Tabela de Preços"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   160
            Top             =   1020
            Width           =   1230
         End
         Begin VB.Label Nome_Operação 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3480
            TabIndex        =   159
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "(máx. 255)"
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   158
            Top             =   285
            Width           =   795
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Forma de pesquisa de produto"
         Height          =   1815
         Left            =   -68205
         TabIndex        =   44
         Top             =   360
         Width           =   4770
         Begin VB.ComboBox cboOrdenacao 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "ParamFilial.frx":4FC37
            Left            =   2220
            List            =   "ParamFilial.frx":4FC41
            TabIndex        =   47
            Text            =   "1 - Numérica"
            Top             =   1290
            Width           =   2145
         End
         Begin VB.OptionButton optLocalizarCodigoNome 
            Appearance      =   0  'Flat
            Caption         =   "Duas opções de pesquisa, por código e nome nas respectivas colunas"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   780
            Width           =   4170
         End
         Begin VB.OptionButton optLocalizarCodigo 
            Appearance      =   0  'Flat
            Caption         =   "Ordem alfabética pelo nome do produto na coluna código"
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   120
            TabIndex        =   45
            Top             =   300
            Width           =   4305
         End
         Begin VB.Label lblOrdenacao 
            Caption         =   "Ordenação"
            Height          =   255
            Left            =   1320
            TabIndex        =   154
            Top             =   1350
            Width           =   885
         End
      End
      Begin VB.CheckBox chkProcuraCodigoENome 
         Caption         =   "Procura produto por código e nome"
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
         Left            =   -67320
         TabIndex        =   151
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   147
         Top             =   4830
         Width           =   5895
         Begin VB.CheckBox O_Preços 
            Appearance      =   0  'Flat
            Caption         =   "Trabalhar com até 3 tabelas de preço + custo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   4215
         End
         Begin VB.TextBox Tabela1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   15
            TabIndex        =   19
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox Tabela3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   21
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox Tabela2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   20
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label_Tabela1 
            AutoSize        =   -1  'True
            Caption         =   "Nome tabela 1"
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label_Tabela2 
            AutoSize        =   -1  'True
            Caption         =   "Nome tabela 2"
            Height          =   195
            Left            =   2040
            TabIndex        =   149
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label_Tabela3 
            AutoSize        =   -1  'True
            Caption         =   "Nome tabela 3"
            Height          =   195
            Left            =   3960
            TabIndex        =   148
            Top             =   360
            Width           =   1035
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Recebimento"
         Height          =   2625
         Index           =   1
         Left            =   120
         TabIndex        =   145
         Top             =   4350
         Width           =   5745
         Begin VB.CheckBox chkWebCliCompraPrazo 
            Appearance      =   0  'Flat
            Caption         =   "Permitir que clientes oriundos da Loja Virtual sejam habilitados a fazer compras a prazo"
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   2205
            TabIndex        =   91
            Top             =   510
            Width           =   2895
         End
         Begin VB.Frame Frame9 
            Caption         =   "Parcelamento Padrão"
            Height          =   1320
            Left            =   120
            TabIndex        =   87
            Top             =   420
            Width           =   1935
            Begin VB.OptionButton Saída_Parcela_Carnet 
               Appearance      =   0  'Flat
               Caption         =   "Carnê"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   240
               TabIndex        =   90
               Top             =   855
               Value           =   -1  'True
               Width           =   885
            End
            Begin VB.OptionButton Saída_Parcela_Carteira 
               Appearance      =   0  'Flat
               Caption         =   "Carteira"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   240
               TabIndex        =   89
               Top             =   570
               Width           =   1005
            End
            Begin VB.OptionButton Saída_Parcela_Banco 
               Appearance      =   0  'Flat
               Caption         =   "Banco"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   240
               TabIndex        =   88
               Top             =   285
               Width           =   885
            End
         End
         Begin VB.CheckBox Saída_Altera_Parcela 
            Appearance      =   0  'Flat
            Caption         =   "Permite alterar parcelamento padrão"
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   2205
            TabIndex        =   92
            Top             =   1275
            Width           =   3015
         End
         Begin VB.TextBox Saída_Intervalo_Parc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   3600
            MaxLength       =   3
            TabIndex        =   93
            Top             =   1995
            Width           =   1470
         End
         Begin VB.Label lblTitle 
            Caption         =   "Intervalo padrão entre parcelas (em dias) :"
            Height          =   255
            Index           =   21
            Left            =   165
            TabIndex        =   146
            Top             =   2040
            Width           =   3225
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Arquivos de configuração da Nota Fiscal"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74880
         TabIndex        =   142
         Top             =   2520
         Width           =   5175
         Begin VB.CommandButton cmdProcurarArquivoNf 
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
            Index           =   1
            Left            =   4560
            MaskColor       =   &H00000000&
            Picture         =   "ParamFilial.frx":4FC66
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "Click neste botão para procurar o arquivo de nota fiscal"
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox chkMantemInformacaoUltimaNotaFiscal 
            Appearance      =   0  'Flat
            Caption         =   "Manter as observações digitadas na emissão da última nota fiscal."
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   120
            TabIndex        =   111
            Top             =   960
            Width           =   4920
         End
         Begin VB.CommandButton cmdProcurarArquivoNf 
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
            Index           =   0
            Left            =   2040
            MaskColor       =   &H00000000&
            Picture         =   "ParamFilial.frx":4FDB0
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "Click neste botão para procurar o arquivo de nota fiscal"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtConfigNFInp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   3360
            MaxLength       =   8
            TabIndex        =   109
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtConfigNFOut 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   840
            MaxLength       =   8
            TabIndex        =   107
            Top             =   360
            Width           =   1095
         End
         Begin MSComDlg.CommonDialog cdgFileOpen 
            Left            =   1200
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Entrada"
            Height          =   195
            Index           =   31
            Left            =   2640
            TabIndex        =   144
            Top             =   450
            Width           =   570
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Saída"
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   143
            Top             =   450
            Width           =   390
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opções operacionais"
         Height          =   5535
         Index           =   1
         Left            =   -68880
         TabIndex        =   140
         Top             =   390
         Width           =   5535
         Begin VB.CheckBox chkFiltrarProdutosInativos 
            Appearance      =   0  'Flat
            Caption         =   "Não filtrar produtos inativos na tela de cadastro de produtos"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            TabIndex        =   225
            Top             =   4650
            Width           =   5295
         End
         Begin VB.CheckBox chkVendedorSenhaGerente 
            Appearance      =   0  'Flat
            Caption         =   "Solicitar senha do gerente ao alterar vendedor em cadastro de clientes, venda rápida, check-out e saídas."
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   193
            Top             =   3510
            Width           =   5205
         End
         Begin VB.CheckBox chkSaldoAnterior 
            Appearance      =   0  'Flat
            Caption         =   "Considerar saldo anterior ao movimentar o caixa"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   192
            Top             =   3135
            Width           =   4515
         End
         Begin VB.CheckBox chkUsaVariosCaixas 
            Appearance      =   0  'Flat
            Caption         =   "Utilizar mais de um caixa"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   30
            Top             =   2790
            Width           =   2475
         End
         Begin VB.CheckBox Gerar_Conta_Paga 
            Appearance      =   0  'Flat
            Caption         =   "Gerar conta paga para recebimentos à vista"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   29
            Top             =   2475
            Width           =   3855
         End
         Begin VB.CheckBox O_Alfa 
            Appearance      =   0  'Flat
            Caption         =   "Aceitar códigos de produtos alfanuméricos"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   28
            Top             =   2100
            Width           =   3675
         End
         Begin VB.CheckBox chkWorkWeb 
            Appearance      =   0  'Flat
            Caption         =   "Trabalhar com Loja Virtual (requer o software Quick Web)"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   27
            Top             =   1770
            Width           =   4515
         End
         Begin VB.CheckBox Alterar_Serviços 
            Appearance      =   0  'Flat
            Caption         =   "Permite alterar descrição na venda"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   360
            TabIndex        =   26
            Top             =   1470
            Width           =   3075
         End
         Begin VB.CheckBox Usar_Serviços 
            Appearance      =   0  'Flat
            Caption         =   "Trabalhar com serviços"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   25
            Top             =   1170
            Width           =   2415
         End
         Begin VB.CheckBox O_Edições 
            Appearance      =   0  'Flat
            Caption         =   "Trabalhar com edições de produtos (revistas)"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   24
            Top             =   810
            Width           =   3915
         End
         Begin VB.CheckBox O_Grade 
            Appearance      =   0  'Flat
            Caption         =   "Trabalhar com grade de produtos"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   4005
         End
         Begin VB.CheckBox Verifica_Agenda 
            Appearance      =   0  'Flat
            Caption         =   "Verificar a agenda ao entrar no sistema"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   22
            Top             =   210
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.CheckBox chkUtilizarCodFornec 
            Appearance      =   0  'Flat
            Caption         =   $"ParamFilial.frx":4FEFA
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   $"ParamFilial.frx":4FF82
            Top             =   4020
            Width           =   5385
         End
         Begin VB.CheckBox chkCheckInstance 
            Appearance      =   0  'Flat
            Caption         =   "Não permitir que o Quick Store seja executado mais de uma vez na estação ao mesmo tempo"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   4980
            Width           =   5205
         End
         Begin VB.CheckBox chkDescSubTotalRateado 
            Appearance      =   0  'Flat
            Caption         =   "Desconto no Sub Total deve ser rateado entre os produtos"
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   4890
            TabIndex        =   31
            Top             =   1350
            Visible         =   0   'False
            Width           =   4755
         End
         Begin VB.CheckBox chkWorkTrafficLight 
            Appearance      =   0  'Flat
            Caption         =   "Utilizar Traffic Light no gerenciamento de vendas"
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   4920
            TabIndex        =   33
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblHelpTrafficLight 
            AutoSize        =   -1  'True
            Caption         =   "(clique aqui para saber mais)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   5070
            MouseIcon       =   "ParamFilial.frx":5001B
            MousePointer    =   99  'Custom
            TabIndex        =   181
            Top             =   3240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.Frame Frame7 
         Height          =   4725
         Left            =   -69600
         TabIndex        =   61
         Top             =   2280
         Width           =   6165
         Begin VB.CheckBox chkPrestServ 
            Appearance      =   0  'Flat
            Caption         =   "Permitir selecionar prestador de serviço"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   307
            Top             =   2610
            Width           =   3855
         End
         Begin VB.CheckBox chkOcultaOrc 
            Appearance      =   0  'Flat
            Caption         =   "Ocultar Orçamentos"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   306
            Top             =   2310
            Width           =   2865
         End
         Begin VB.CheckBox chkRecalculo 
            Appearance      =   0  'Flat
            Caption         =   "Realiza recálculo dos preços devido a possíveis modificações de desconto"
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   120
            TabIndex        =   69
            Top             =   1950
            Width           =   5835
         End
         Begin VB.CheckBox VR_Permite_Cheques 
            Appearance      =   0  'Flat
            Caption         =   "Permite Cheques"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   187
            Top             =   2000
            Width           =   1695
         End
         Begin VB.CheckBox VR_Permite_Dinheiro 
            Appearance      =   0  'Flat
            Caption         =   "Permite Dinheiro"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   186
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox VR_Permite_Parcela 
            Appearance      =   0  'Flat
            Caption         =   "Permite Parcelamento"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   185
            Top             =   1590
            Width           =   1980
         End
         Begin VB.Frame Frame14 
            Caption         =   "Parcelamento em Banco - Boleto"
            Height          =   1170
            Left            =   2400
            TabIndex        =   75
            Top             =   3435
            Width           =   3375
            Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
               Bindings        =   "ParamFilial.frx":508E5
               DataSource      =   "Data5"
               Height          =   270
               Left            =   1560
               TabIndex        =   78
               Top             =   660
               Width           =   735
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
               BackColorOdd    =   16777152
               Columns(0).Width=   3200
               _ExtentX        =   1296
               _ExtentY        =   476
               _StockProps     =   93
               ForeColor       =   -2147483640
               BackColor       =   12648447
            End
            Begin VB.OptionButton O_Conta_Fixa 
               Appearance      =   0  'Flat
               Caption         =   "Usar esta conta"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   105
               TabIndex        =   77
               Top             =   660
               Width           =   1500
            End
            Begin VB.OptionButton O_Conta_Cadastro 
               Appearance      =   0  'Flat
               Caption         =   "Usar a conta do cadastro do cliente"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   105
               TabIndex        =   76
               Top             =   345
               Value           =   -1  'True
               Width           =   3045
            End
            Begin VB.Label Nome_Conta 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2295
               TabIndex        =   141
               Top             =   660
               Width           =   975
            End
         End
         Begin VB.TextBox VR_Prazo_Parcela 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   4995
            MaxLength       =   6
            TabIndex        =   68
            Top             =   1545
            Width           =   735
         End
         Begin VB.TextBox VR_Prazo_Cheques 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   4995
            MaxLength       =   6
            TabIndex        =   66
            Top             =   585
            Width           =   735
         End
         Begin VB.CheckBox VR_Altera_Parcela 
            Appearance      =   0  'Flat
            Caption         =   "Permite alterar parcelamento padrão"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2490
            TabIndex        =   74
            Top             =   3090
            Width           =   3375
         End
         Begin VB.Frame Frame8 
            Caption         =   "Parcelamento Padrão"
            Height          =   1530
            Left            =   120
            TabIndex        =   70
            Top             =   3075
            Width           =   2175
            Begin VB.OptionButton VR_Parcela_Banco 
               Appearance      =   0  'Flat
               Caption         =   "Banco"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   255
               TabIndex        =   71
               Top             =   390
               Width           =   960
            End
            Begin VB.OptionButton VR_Parcela_Carteira 
               Appearance      =   0  'Flat
               Caption         =   "Carteira"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   255
               TabIndex        =   72
               Top             =   705
               Width           =   1005
            End
            Begin VB.OptionButton VR_Parcela_Carnet 
               Appearance      =   0  'Flat
               Caption         =   "Carnê"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   255
               TabIndex        =   73
               Top             =   1020
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.TextBox VR_Qtde_Parcela 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   2595
            MaxLength       =   3
            TabIndex        =   67
            Top             =   1545
            Width           =   495
         End
         Begin VB.TextBox VR_Qtde_Cheques 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   2595
            MaxLength       =   3
            TabIndex        =   65
            Top             =   585
            Width           =   495
         End
         Begin VB.CheckBox VR_Permite_Cartão 
            Appearance      =   0  'Flat
            Caption         =   "Permite Cartão"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1275
            Width           =   1455
         End
         Begin VB.CheckBox VR_Permite_Vales 
            Appearance      =   0  'Flat
            Caption         =   "Permite Vales"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   285
            Width           =   1455
         End
         Begin VB.CheckBox VR_Permite_Rec_Rápido 
            Appearance      =   0  'Flat
            Caption         =   "Permite recebimento simplificado na venda rápida"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5160
            TabIndex        =   62
            Top             =   180
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label57 
            Caption         =   "Prazo Máximo (em dias)"
            Height          =   255
            Left            =   3195
            TabIndex        =   139
            Top             =   1590
            Width           =   1755
         End
         Begin VB.Label Label56 
            Caption         =   "Qtde"
            Height          =   255
            Left            =   2115
            TabIndex        =   138
            Top             =   1590
            Width           =   375
         End
         Begin VB.Label Label55 
            Caption         =   "Prazo Máximo (em dias)"
            Height          =   255
            Left            =   3195
            TabIndex        =   137
            Top             =   630
            Width           =   1755
         End
         Begin VB.Label Label54 
            Caption         =   "Qtde"
            Height          =   255
            Left            =   2145
            TabIndex        =   136
            Top             =   630
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados fiscais da empresa"
         Height          =   4455
         Index           =   0
         Left            =   -74880
         TabIndex        =   126
         Top             =   390
         Width           =   5895
         Begin VB.TextBox txtCEP 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   8
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtCNAE 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   17
            Top             =   4080
            Width           =   2775
         End
         Begin VB.TextBox txtInscricaoSuframa 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   9
            TabIndex        =   16
            Top             =   4080
            Width           =   2775
         End
         Begin VB.TextBox txtInscricaoMunicipal 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   20
            TabIndex        =   14
            Top             =   3480
            Width           =   2775
         End
         Begin VB.TextBox txtPais 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3720
            MaxLength       =   64
            TabIndex        =   11
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox txtEnderecoComplemento 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtEnderecoNumero 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   5
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox Razão_Social 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   3
            Top             =   480
            Width           =   5655
         End
         Begin VB.TextBox Endereço 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1080
            Width           =   4335
         End
         Begin VB.TextBox Bairro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox Fone 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   30
            TabIndex        =   12
            Top             =   2880
            Width           =   2055
         End
         Begin VB.TextBox Cidade 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   30
            TabIndex        =   9
            Top             =   2280
            Width           =   2895
         End
         Begin VB.TextBox Estado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   10
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox CGC 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2280
            MaxLength       =   20
            TabIndex        =   13
            Top             =   2880
            Width           =   3495
         End
         Begin VB.TextBox Inscrição 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3000
            MaxLength       =   20
            TabIndex        =   15
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Index           =   72
            Left            =   4560
            TabIndex        =   203
            Top             =   1440
            Width           =   285
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "CNAE"
            Height          =   195
            Index           =   71
            Left            =   3000
            TabIndex        =   202
            Top             =   3840
            Width           =   405
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Suframa"
            Height          =   195
            Index           =   70
            Left            =   120
            TabIndex        =   201
            Top             =   3840
            Width           =   1290
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Municipal"
            Height          =   195
            Index           =   69
            Left            =   120
            TabIndex        =   200
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "País"
            Height          =   195
            Index           =   68
            Left            =   3720
            TabIndex        =   199
            Top             =   2040
            Width           =   285
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Complemento"
            Height          =   195
            Index           =   67
            Left            =   120
            TabIndex        =   198
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   66
            Left            =   4560
            TabIndex        =   197
            Top             =   840
            Width           =   555
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   134
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   133
            Top             =   840
            Width           =   675
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Index           =   3
            Left            =   1920
            TabIndex        =   132
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Fone"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   131
            Top             =   2640
            Width           =   360
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   130
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Index           =   8
            Left            =   3120
            TabIndex        =   129
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ"
            Height          =   195
            Index           =   6
            Left            =   2280
            TabIndex        =   128
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição  Estadual"
            Height          =   195
            Index           =   7
            Left            =   3000
            TabIndex        =   127
            Top             =   3240
            Width           =   1350
         End
      End
      Begin MSMask.MaskEdBox Juros 
         Height          =   315
         Left            =   -68730
         TabIndex        =   35
         Top             =   6285
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   15066597
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###.00"
         PromptChar      =   "_"
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperSaida_S 
         Bindings        =   "ParamFilial.frx":508F9
         DataSource      =   "datOperSaida"
         Height          =   285
         Left            =   6195
         TabIndex        =   94
         Top             =   700
         Width           =   1215
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
         BackColorOdd    =   12648384
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   6959
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1879
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin MSMask.MaskEdBox mskTaxaDesconto 
         Height          =   315
         Left            =   -68730
         TabIndex        =   36
         Top             =   6645
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   15066597
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###.00"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -63540
         Top             =   4050
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox mskTaxaMultaParcelaVencida 
         Height          =   315
         Left            =   -73440
         TabIndex        =   287
         Top             =   6645
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   15066597
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###.00"
         PromptChar      =   "_"
      End
      Begin VB.Line Line2 
         X1              =   -67140
         X2              =   -67140
         Y1              =   6060
         Y2              =   6990
      End
      Begin VB.Line Line1 
         X1              =   -71070
         X2              =   -71070
         Y1              =   6060
         Y2              =   6990
      End
      Begin VB.Label lbl_multaDiasAposParcelaVencida 
         Caption         =   "Somente após             dias da parcela vencida"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   -74520
         TabIndex        =   292
         Top             =   6308
         Width           =   3555
      End
      Begin VB.Label lblTitle 
         Caption         =   "Quantidade de dígitos do CÓDIGO DO PRODUTO emitido pela etiquetadora da balança (de 3 a 6):"
         ForeColor       =   &H00000000&
         Height          =   690
         Index           =   12
         Left            =   -66915
         TabIndex        =   291
         Top             =   1965
         Width           =   2715
      End
      Begin VB.Label lblTitle 
         Caption         =   "Etiqueta de Balança:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   32
         Left            =   -67140
         TabIndex        =   290
         Top             =   1710
         Width           =   1545
      End
      Begin VB.Label lbl_TaxaMultaParcelaVencida 
         Caption         =   "Taxa de multa                          %"
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   -74520
         TabIndex        =   288
         Top             =   6660
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Casas decimais para Preço Unitário"
         Height          =   255
         Left            =   -67170
         TabIndex        =   283
         Top             =   540
         Width           =   2745
      End
      Begin VB.Label lblTitle 
         Caption         =   "Nome Pesquisa 3"
         Height          =   240
         Index           =   58
         Left            =   -66720
         TabIndex        =   280
         Top             =   6480
         Width           =   1380
      End
      Begin VB.Label lblTitle 
         Caption         =   "Nome Pesquisa 2"
         Height          =   285
         Index           =   57
         Left            =   -66720
         TabIndex        =   279
         Top             =   6030
         Width           =   1380
      End
      Begin VB.Label lblTitle 
         Caption         =   "Nome Pesquisa 1"
         Height          =   285
         Index           =   56
         Left            =   -66720
         TabIndex        =   278
         Top             =   5580
         Width           =   1380
      End
      Begin VB.Label lblUltimaNFCe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Última NFCe"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -69480
         TabIndex        =   227
         Top             =   6690
         Width           =   1215
      End
      Begin VB.Label lblUltimaNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Última NFe"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -66645
         TabIndex        =   214
         Top             =   6690
         Width           =   1815
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Taxa mensal de Desconto"
         Height          =   195
         Index           =   63
         Left            =   -70680
         TabIndex        =   189
         Top             =   6705
         Width           =   1845
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "dias"
         Height          =   195
         Index           =   11
         Left            =   -63690
         TabIndex        =   184
         Top             =   6338
         Width           =   285
      End
      Begin VB.Label lblTitle 
         Caption         =   "Bloquear venda para cliente que não compra a mais de"
         Height          =   405
         Index           =   10
         Left            =   -66390
         TabIndex        =   182
         Top             =   6233
         Width           =   2085
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Op. de Venda a ser utilizada na transformação de Orçamento em Venda"
         Height          =   195
         Index           =   22
         Left            =   6195
         TabIndex        =   153
         Top             =   480
         Width           =   5160
      End
      Begin VB.Label lblNomeOperSaida_S 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7470
         TabIndex        =   152
         Top             =   700
         Width           =   3855
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Taxa mensal de Juros"
         Height          =   195
         Index           =   9
         Left            =   -70365
         TabIndex        =   135
         Top             =   6345
         Width           =   1560
      End
   End
   Begin VB.Label Última_Movimentação 
      Appearance      =   0  'Flat
      Caption         =   " Última sequência"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8820
      TabIndex        =   155
      Top             =   210
      Width           =   1425
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   135
      Top             =   7740
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
      Bands           =   "ParamFilial.frx":50914
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   61
      Left            =   135
      TabIndex        =   125
      Top             =   195
      Width           =   600
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Num_Registro As Variant
Private rsParametros As Recordset
Private rsPreços As Recordset
Private rsCliFor As Recordset
Private rsOp_Saída As Recordset
Private rsContas As Recordset
Private Mudou_Imp_Rel As Integer
Private Mudou_Imp_Nota As Integer
Private Mudou_Imp_Ticket As Integer
Private Mudou_Imp_Cheque As Integer
Private Mudou_Imp_Boleto As Integer
Private rsZZZLog As Recordset
Private rsParametros2 As Recordset

Public gsSenhaGerenteAtual As String

'Private Sub TrocaSenhaGerente()
'  Dim F As Form
'  Dim sSenhaGerente As String
'
'  If IsNull(Num_Registro) Then
'    DisplayMsg "Encontre ou grave um registro antes."
'    Exit Sub
'  End If
'
'  If Not frmGerente.gbSenhaGerente Then
''    gsSenhaGerente = sSenhaGerente
'    Exit Sub
'  End If
'
'  sSenhaGerente = gsSenhaGerente
'  gsSenhaGerente = gsSenhaGerenteAtual
'
'
'  gsSenhaGerente = sSenhaGerente
'  Set F = New frmTrocaSenhaGerente
'  F.Show vbModal
'  Set F = Nothing
'
'  If rsParametros("Filial").Value = gnCodFilial Then
'    gsSenhaGerente = gsSenhaGerenteAtual
'  End If
'
'End Sub

Public Function VerificaPAF()
  Set rsParametros2 = db.OpenRecordset("Select [BancoPDV] from [Parâmetros Filial] Where Filial = " & gnCodFilial & ";")
  
  Dim fso As New FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")

  With rsParametros2
  VerificaPAF = Not IsNull(rsParametros2("BancoPDV").Value)
  If VerificaPAF Then VerificaPAF = (Len(rsParametros2("BancoPDV").Value) > 0 And fso.FolderExists(rsParametros2("BancoPDV").Value))
  'If Len(rsParametros2("BancoPDV").Value) > 0 And fso.FolderExists(rsParametros2("BancoPDV").Value) Then
  'If rsParametros("Estado") & "" = "DF" Or rsParametros("Estado") & "" = "RJ" Or rsParametros("Estado") & "" = "CE" Then
  'VerificaPAF = True
  'Else
  'VerificaPAF = False
  'End If
  'rsParametros = Nothing
  End With
  'rsParametros2.Close
End Function

Private Sub ShowRecord()
  Dim sCod As String
  Dim intRet As Integer
  Dim bytRet As Byte
  
  On Error GoTo Processa_Erro
  
  Call StatusMsg("")
  
  
  If Not IsNull(rsParametros("CobrarMultaAposVencimentoParcela")) Then
      If rsParametros("CobrarMultaAposVencimentoParcela") = True Then
          chk_cobrarMulta.Value = vbChecked
      Else
          chk_cobrarMulta.Value = vbUnchecked
      End If
  End If
  If Not IsNull(rsParametros("TaxaMultaParcelaVencida")) Then
      mskTaxaMultaParcelaVencida.Text = rsParametros("TaxaMultaParcelaVencida")
  Else
      mskTaxaMultaParcelaVencida.Text = ""
  End If
  If Not IsNull(rsParametros("MultaDiasAposParcelaVencida")) Then
      txt_multaDiasAposParcelaVencida.Text = rsParametros("MultaDiasAposParcelaVencida")
  Else
      txt_multaDiasAposParcelaVencida.Text = ""
  End If
  
  'Tratamento da combo SITUAÇÃO TRIBUTÁRIO DO PIS
  '01 – Operação Tributável - Base de Cálculo = Valor da Operação Alíquota Normal (Cumulativo/Não Cumulativo) = PISAliq
  '02 - Operação Tributável - Base de Calculo = Valor da Operação (Alíquota Diferenciada) = PISAliq
  '03 - Operação Tributável - Base de Calculo = Quantidade Vendida x Alíquota por Unidade de Produto = PISQtde
  '04 - Operação Tributável - Tributação Monofásica - (Alíquota Zero) = PISNT
  '06 - Operação Tributável - Alíquota Zero = PISNT
  '07 - Operação Isenta da contribuição = PISNT
  '08 - Operação Sem Incidência da contribuição = PISNT
  '09 - Operação com suspensão da contribuição = PISNT
  '99 - Outras Operações = PISOutr
  If Not IsNull(rsParametros("TipoSituacaoTributariaPIS")) Then
      If rsParametros("TipoSituacaoTributariaPIS") = 0 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 0
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 1 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 1
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 2 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 2
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 3 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 3
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 4 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 4
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 6 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 5
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 7 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 6
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 8 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 7
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 9 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 8
      ElseIf rsParametros("TipoSituacaoTributariaPIS") = 99 Then
          cmb_situacaoTributariaDoPIS.ListIndex = 9
      End If
  Else
      cmb_situacaoTributariaDoPIS.ListIndex = -1
  End If
  'Fim tratamento da combo

  If Not IsNull(rsParametros("Quick_viaRDP_ticket").Value) And rsParametros("Quick_viaRDP_ticket").Value = 1 Then
      chk_viaRDP_ticket.Value = vbChecked
  Else
      chk_viaRDP_ticket.Value = vbUnchecked
  End If

  If Not IsNull(rsParametros("Quick_viaRDP").Value) And rsParametros("Quick_viaRDP").Value = 1 Then
      chk_viaRDP.Value = vbChecked
  Else
      chk_viaRDP.Value = vbUnchecked
  End If
  
  Código.Text = rsParametros("Filial")
  Nome.Text = rsParametros("Nome") & ""
  
  Razão_Social.Text = rsParametros("Razão Social") & ""
  Endereço.Text = rsParametros("Endereço") & ""
  Bairro.Text = rsParametros("Bairro") & ""
  Fone.Text = rsParametros("Fone") & ""
  Cidade.Text = rsParametros("Cidade") & ""
  Estado.Text = rsParametros("Estado") & ""
  CGC.Text = rsParametros("CGC") & ""
  Inscrição.Text = rsParametros("Inscrição") & ""
  
  '26/08/2009 - mpdea
  'Novos campos para NFe
  txtEnderecoNumero.Text = rsParametros.Fields("EnderecoNumero").Value & ""
  txtEnderecoComplemento.Text = rsParametros.Fields("EnderecoComplemento").Value & ""
  txtCEP.Text = rsParametros.Fields("CEP").Value & ""
  txtPais.Text = rsParametros.Fields("Pais").Value & ""
  txtInscricaoMunicipal.Text = rsParametros.Fields("InscricaoMunicipal").Value & ""
  txtInscricaoSuframa.Text = rsParametros.Fields("InscricaoSuframa").Value & ""
  txtCNAE.Text = rsParametros.Fields("CNAE").Value & ""
  
  
  Juros.Text = rsParametros("Juros")
  '17/08/2004 - Daniel
  mskTaxaDesconto.Text = rsParametros("TaxaDesconto").Value
  
  
  '07/08/2003 - mpdea
  'Modificado controle para o campo 'Usa Vários Caixas'
'  If rsParametros("Usa Vários Caixas") = False Then O_Um_Caixa.Value = True
'  If rsParametros("Usa Vários Caixas") = True Then O_Vários_Caixas.Value = True
  chkUsaVariosCaixas.Value = IIf(rsParametros.Fields("Usa Vários Caixas").Value, vbChecked, vbUnchecked)
  '20/11/2006 - Anderson
  'Considerar saldo anterior ao movimentar o caixa
  chkSaldoAnterior.Value = IIf(rsParametros.Fields("ConsiderarSaldoAnterior").Value, vbChecked, vbUnchecked)
  '17/01/2007 - Anderson
  'Solicitar senha do gerente ao alterar vendedor nas telas de cadastro de clientes, venda rápida, saídas e check-out
  chkVendedorSenhaGerente.Value = IIf(rsParametros.Fields("VendedorSenhaGerente").Value, vbChecked, vbUnchecked)
  txtDiasBloqueioVenda.Text = rsParametros.Fields("DiasBloqueioVenda").Value & ""
  
  O_Preços.Value = -rsParametros("Três Tabelas")
  O_Grade.Value = -rsParametros("Usar Grade")
  Usar_Serviços.Value = -rsParametros("Usar Serviços")
  Alterar_Serviços.Value = -rsParametros("Alterar Serviços")
  O_Alfa.Value = -rsParametros("Usar Códigos Alfa")
  O_Edições.Value = -rsParametros("Usar Edições")
  
  '13/03/2013-Alexandre Afornali
  'Campo para filtrar ou nao produtos inativos na tela de cadastro de produtos
  chkFiltrarProdutosInativos.Value = IIf(rsParametros.Fields("FiltrarProdutosInativos").Value, vbChecked, vbUnchecked)
  
  '17/05/2013-Alexandre Afornali
  'Campo para usar ou nao comandas
  chkComandas.Value = IIf(rsParametros.Fields("TrabalharComComanda").Value, vbChecked, vbUnchecked)
  
  '31/07/2002 - mpdea
  'Campo de utilização da Loja Virtual
  chkWorkWeb.Value = IIf(rsParametros.Fields("WorkWeb").Value, vbChecked, vbUnchecked)
  
  
  '29/05/2003 - mpdea
  'Campo de utilização do Traffic Light
  chkWorkTrafficLight.Value = IIf(rsParametros.Fields("WorkTrafficLight").Value, vbChecked, vbUnchecked)
  
  
  '07/08/2003 - mpdea
  'Campo de verificação de instâncias do Quick Store
  chkCheckInstance.Value = IIf(rsParametros.Fields("CheckInstance").Value, vbChecked, vbUnchecked)
  
  '13/04/2025 - pablo
  chkProdutoNomeNFe.Value = IIf(rsParametros.Fields("EditarNomeProduto").Value, vbChecked, vbUnchecked)
  
  
  Tabela1.Text = rsParametros("Tabela 1") & ""
  Tabela2.Text = rsParametros("Tabela 2") & ""
  Tabela3.Text = rsParametros("Tabela 3") & ""
  
  Lista(0).Text = rsParametros("Lista 1") & ""
  Lista(1).Text = rsParametros("Lista 2") & ""
  Lista(2).Text = rsParametros("Lista 3") & ""
  Lista(3).Text = rsParametros("Lista 4") & ""
  Lista(4).Text = rsParametros("Lista 5") & ""
  
  txtUltimaConsignacao.Text = rsParametros("UltimaConsignacao") & ""
  
  Ult_Nota.Text = rsParametros("Última Nota") & ""
  
 
  Máximo.Text = rsParametros("Linhas Digitação") & ""
  Máximo_Serviço.Text = rsParametros("Linhas Serviço") & ""
  ' Pede_Senha.Value = -rsParametros("Senha Sempre")
'  Super_Libera = -rsParametros("Superusuário Libera Telas")
  Sem_Estoque.Value = -rsParametros("Venda Sem Estoque")
  Ult_Mov.Text = rsParametros("Última Movimentação")
  txtConfigNFOut.Text = rsParametros("Nota Saída") & ""
  '18/08/2004 - Daniel
  If CheckSerialCaseMod("QS39823-684") Then txtTicket.Text = rsParametros("TicketPadrao").Value & ""
  
  txtConfigNFInp.Text = rsParametros("Nota Entrada") & ""
  
  '----------------------------------------------------------------------------
  '01/09/2009 - mpdea
  'NFe
  '
  'Identificação do Ambiente
  Call IsDataType(dtByte, rsParametros.Fields("AmbienteNfe").Value, bytRet)
  If bytRet < 1 Or bytRet > 2 Then bytRet = 0 Else bytRet = bytRet - 1
  cboAmbienteNfe.ListIndex = bytRet
  
  'Formato de Impressão do DANFE
  Call IsDataType(dtByte, rsParametros.Fields("FormatoImpressaoDanfeNfe").Value, bytRet)
  If bytRet < 1 Or bytRet > 2 Then bytRet = 0 Else bytRet = bytRet - 1
  cboFormatoImpressaoDanfeNfe.ListIndex = bytRet
  'Modalidade de determinação da Base de Cálculo do ICMS
  Call IsDataType(dtByte, rsParametros.Fields("ModDetBaseCalculoIcms").Value, bytRet)
  If bytRet < 0 Or bytRet > 3 Then bytRet = 0
  cboModDetBaseCalculoIcms.ListIndex = bytRet
  'Modalidade de determinação da Base de Cálculo do ICMS ST
  Call IsDataType(dtByte, rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value, bytRet)
  If bytRet < 0 Or bytRet > 5 Then bytRet = 0
  cboModDetBaseCalculoIcmsSt.ListIndex = bytRet
  'Pasta de envio
  txtPastaEnvioNfe.Text = rsParametros.Fields("PastaEnvioNfe").Value & ""
  'Pasta de retorno
  txtPastaRetornoNfe.Text = rsParametros.Fields("PastaRetornoNfe").Value & ""
  '17/09/2009 - mpdea
  chkHabilitarNotaFiscalEletronica.Value = IIf(rsParametros.Fields("HabilitarNotaFiscalEletronica").Value, vbChecked, vbUnchecked)
  Call chkHabilitarNotaFiscalEletronica_Click
  
  '30/09/2009 - Andrea
  txtUltimaNFe.Text = rsParametros.Fields("UltimaNFe").Value & ""
  
  txtUltimaNFCe.Text = rsParametros.Fields("UltimaNFCe").Value & ""
 
  '25/11/2010 - Andrea
  'Layout Envio
  txtLayoutEnvio.Text = rsParametros.Fields("VersaoLayoutEnvio").Value & ""
    
  'Codigo Regime Tributario
  Call IsDataType(dtByte, rsParametros.Fields("CodigoRegimeTributario").Value, bytRet)
  If bytRet < 1 Or bytRet > 3 Then bytRet = 0 Else bytRet = bytRet - 1
  cboCodigoRegimeTributario.ListIndex = bytRet
     
  ' Pilatti 2017-Setembro
  If IsNull(rsParametros.Fields("PadraoArquivoIntegracao").Value) Then
    cboPadraoArquivoIntegracao.Text = "TXT"
    ' Realizar UPDATE no DB para opção TXT
    Dim sSQLAt As String
    sSQLAt = "Update [Parâmetros Filial] Set PadraoArquivoIntegracao = 'TXT' "
    sSQLAt = sSQLAt & "Where Filial = " & rsParametros.Fields("Filial").Value
    
    db.Execute sSQLAt, dbFailOnError
    
  Else
    '16/11/2011 - Andrea
    'Padrão do Arquivo de Integração
    cboPadraoArquivoIntegracao.Text = rsParametros.Fields("PadraoArquivoIntegracao").Value
  End If
  ' Fim Pilatti
  
  '----------------------------------------------------------------------------
  
  '10/03/2009 - mpdea
  txtAliquotaAproveitamentoCreditoIcms.Text = rsParametros.Fields("AliquotaAprovCreditoIcms").Value & ""
  
  '14/03/2011 - Andrea
  txtPercentualSimplesNacional.Text = rsParametros.Fields("PercentualSimplesNacional").Value & ""
  
  '30/03/2011 - Andrea
  txtPercentualReducaoBC_SN.Text = rsParametros.Fields("PercentualReducaoBCSimplesNacional").Value & ""
  
  Verifica_Agenda.Value = -rsParametros("Verifica Agenda")
  optLocalizarCodigoNome = rsParametros("PesquisaCodigoENome_VR") 'chkProcuraCodigoENome
    
  Qtde_Balança.Text = rsParametros("Qtde Balança")
  
  
  '06/05/2003 - mpdea
  'Desconto no Sub Total rateado para Venda Rápida e Saídas
  chkDescSubTotalRateado.Value = IIf(rsParametros.Fields("DescSubTotalRateado").Value, vbChecked, vbUnchecked)
  
  '06/05/2005 - Daniel
  'Tratamento para o campo [Parâmetros Filial].UtilizarCodFornec
  chkUtilizarCodFornec.Value = IIf(rsParametros.Fields("UtilizarCodFornec").Value, vbChecked, vbUnchecked)
  
  c_comp1.Text = rsParametros("Cód Comp 1") & ""
  c_comp2.Text = rsParametros("Cód Comp 2") & ""
  c_comp3.Text = rsParametros("Cód Comp 3") & ""
  
  c_oito1.Text = rsParametros("Cód Oitavo 1") & ""
  c_oito2.Text = rsParametros("Cód Oitavo 2") & ""
  c_oito3.Text = rsParametros("Cód Oitavo 3") & ""
  
  sCod = rsParametros("Cód Comprim 1") & ""
  If Len(sCod) = 0 Then
   c_comp_pag1.Text = "27"
  Else
   c_comp_pag1.Text = sCod
  End If
  
  sCod = rsParametros("Cód Comprim 2") & ""
  If Len(sCod) = 0 Then
   c_comp_pag2.Text = "67"
  Else
   c_comp_pag2.Text = sCod
  End If
  
  sCod = rsParametros("Cód Comprim 3") & ""
  If Len(sCod) = 0 Then
   c_comp_pag3.Text = "0"
  Else
   c_comp_pag3.Text = sCod
  End If
  
  c_comp_pag2.Text = rsParametros("Cód Comprim 2") & ""
  c_comp_pag3.Text = rsParametros("Cód Comprim 3") & ""
  
  VR_Linhas.Text = rsParametros("VR Linhas Digitação")
  
  
  '16/01/2006 - mpdea
  'Utilização da tela de Venda Rápida em tela cheia
  chkVR_Tela_CheckOut.Value = IIf(rsParametros.Fields("VR_Tela_CheckOut").Value, vbChecked, vbUnchecked)
  
  
  VR_Cód_Operação.Text = rsParametros("VR Código Operação")
  VR_Cód_Operação_LostFocus
  VR_Combo_Preço.Text = rsParametros("VR Tab Preço") & ""
  VR_Altera_Tabela.Value = -rsParametros("VR Altera Tabela")
  VR_Altera_Preço.Value = -rsParametros("VR Altera Preço")
  
  
  '17/09/2003 - mpdea
  'Incluído chkAlteraPreco
  '
  '12/09/2003 - mpdea
  'Validação para o estado de SC
'  If UCase(gstrGetEstadoFilial(rsParametros.Fields("Filial").Value)) = "SC" Then
'    VR_Altera_Preço.Value = vbUnchecked
'    VR_Altera_Preço.Visible = False
'    chkAlteraPreco.Value = vbUnchecked
'    chkAlteraPreco.Visible = False
'  Else
'    VR_Altera_Preço.Visible = True
'    chkAlteraPreco.Visible = True
'  End If
  
  '29/12/2003 - mpdea
  'Senha obrigatória do vendedor ao gravar venda
  chkVR_GravarExigeSenhaVend.Value = IIf(rsParametros.Fields("VR_GravarExigeSenhaVend").Value, vbChecked, vbUnchecked)
  
  VR_Combo_Cliente.Text = rsParametros("VR Cliente")
  VR_Combo_Cliente_LostFocus
  VR_Altera_Cliente.Value = -rsParametros("VR Altera Cliente")
  VR_Cadastra_Cliente.Value = -rsParametros("VR Cadastra Cliente")
  VR_Permite_Desconto.Value = -rsParametros("VR Permite Desconto")
  VR_Desconto.Text = rsParametros("VR Desconto") & ""
  
  '03/07/2006 - mpdea
  'Permissão para imprimir ticket somente em movimentações efetivadas
  'Solicitante: Bem me quer
  chkImprimeTicketMovEfetivada.Value = IIf(rsParametros.Fields("ImprimeTicketMovEfetivada").Value, vbChecked, vbUnchecked)
  
  '04/12/2007 - Anderson
  'Permissão para imprimir nota somente em movimentações efetivadas
  chkImprimeNotaMovEfetivada.Value = IIf(rsParametros.Fields("ImprimeNotaMovEfetivada").Value, vbChecked, vbUnchecked)

  '13/09/2012 - mpdea
  chkVRUtilizarTicketModoRelatorio.Value = IIf(rsParametros.Fields("VRUtilizarTicketModoRelatorio").Value, vbChecked, vbUnchecked)

  'Venda_Mostrar_Estoque.Value = -rsParametros("VR Mostrar Estoque")
  
  
  VR_Permite_Rec_Rápido.Value = -rsParametros("VR Permite Rec Rápido")
  VR_Recebimento_Normal.Value = -rsParametros("VR Recebimento Normal")
  VR_Permite_Dinheiro.Value = -rsParametros("VR Permite Dinheiro")
  VR_Permite_Vales.Value = -rsParametros("VR Permite Vales")
  VR_Permite_Cartão.Value = -rsParametros("VR Permite Cartão")
  VR_Permite_Cheques.Value = -rsParametros("VR Permite Cheques")
  VR_Qtde_Cheques.Text = rsParametros("VR Qtde Cheques") & ""
  VR_Prazo_Cheques.Text = rsParametros("VR Prazo Cheques") & ""
  VR_Permite_Parcela.Value = -rsParametros("VR Permite Parcela")
  '25/05/2004 - Daniel
  'Inclusão do campo VR_RecalcularPreço
  If rsParametros("VR_RecalcularPreço").Value Then
    chkRecalculo.Value = vbChecked
  Else
    chkRecalculo.Value = vbUnchecked
  End If
  If rsParametros("VR_OcultaOrc").Value Then
    chkOcultaOrc.Value = vbChecked
  Else
    chkOcultaOrc.Value = vbUnchecked
  End If
  If rsParametros("comPrestServ").Value Then
    chkPrestServ.Value = vbChecked
  Else
    chkPrestServ.Value = vbUnchecked
  End If
  VR_Qtde_Parcela.Text = rsParametros("VR Qtde Parcela") & ""
  VR_Prazo_Parcela.Text = rsParametros("VR Prazo Parcela") & ""
  VR_Altera_Parcela.Value = -rsParametros("VR Altera Parcela")
  If rsParametros("VR Parcela Padrão") = "B" Then VR_Parcela_Banco.Value = True
  If rsParametros("VR Parcela Padrão") = "C" Then VR_Parcela_Carteira.Value = True
  If rsParametros("VR Parcela Padrão") = "E" Then VR_Parcela_Carnet.Value = True
  
  '-----------------------------------------------------------------------------------
  '24/07/2006 - Andrea
  'Incluído campo ExigeSenhaGerReimpTicket (Exigir senha do Gerente para
  'Reimpressão de ticket), solicitado por Rodrigo - TechnoMax.
  '-----------------------------------------------------------------------------------
  SenhaGerReimpTicket.Value = -rsParametros("ExigeSenhaGerReimpTicket")
  
  '-----------------------------------------------------------------------------------
  '04/12/2007 - Celso
  'Incluído campo ExigeSenhaGerVendaAtraso (Exigir senha do Gerente para
  'vendas a cliente com contas em atraso), solicitado por Valdeci - Vaplak
  '-----------------------------------------------------------------------------------
  SenhaGerVendaAtraso.Value = -rsParametros("ExigeSenhaGerVndContaAtraso")
  
  '-----------------------------------------------------------------------------------
  '11/12/2007 - Celso
  'Incluído campo NaoPermiteDuplicarCNPJ (Para não permitir duplicação no cadastro de
  ' Clientes / Fornecedores), solicitado por SMQ
  '-----------------------------------------------------------------------------------
  NaoPermiteDuplicarCNPJ.Value = -rsParametros("NaoPermiteDuplicarCNPJ")
  
  Crédito_Venda_Rápida.Value = -rsParametros("VR Verifica Limite")
  
  VR_Intervalo_Parc.Text = rsParametros("VR Intervalo Parc") & ""
  
  If rsParametros("VROrdenacaoCombo") Then
    cboOrdenacao.Text = "1 - Numérica"
  Else
    cboOrdenacao.Text = "2 - Alfanumérica"
  End If
    
  If rsParametros("VR Conta Padrão") = "C" Then
   O_Conta_Cadastro.Value = True
  Else
   O_Conta_Fixa.Value = True
  End If
  
  Combo_Conta.Text = rsParametros("VR Conta Usar") & ""
  Combo_Conta_LostFocus
  
  Com_Tab1.Text = rsParametros("Consulta TAB1") & ""
  Com_Tab2.Text = rsParametros("Consulta TAB2") & ""
  Com_Tab3.Text = rsParametros("Consulta TAB3") & ""
  Com_Tab4.Text = rsParametros("Consulta TAB4") & ""
  Com_Tab5.Text = rsParametros("Consulta TAB5") & ""
  Com_Tab6.Text = rsParametros("Consulta TAB6") & ""
  
  Mensagem_Troca.Text = rsParametros("Mensagem Troca") & ""
  Mens_Etiq1.Text = rsParametros("Mensagem Etiq 1") & ""
  Mens_Etiq2.Text = rsParametros("Mensagem Etiq 2") & ""
  
  '26/05/2004 - Daniel
  'Inclusão do campo [Zero a Esquerda]
  If rsParametros("Zero a Esquerda").Value = True Then
    chk0aEsquerda.Value = vbChecked
  Else
    chk0aEsquerda.Value = vbUnchecked
  End If
  
  '09/08/2005 - Daniel
  'Inclusão do campo AlterVendedorCliFor
  'Finalidade: Apenas o Superusuário poderá alterar o campo
  '            Vendedor no cadastro Cli / For
  If rsParametros("AlterVendedorCliFor").Value Then
    chkAlterVendedorCliFor.Value = vbChecked
  Else
    chkAlterVendedorCliFor.Value = vbUnchecked
  End If
  
  '17/08/2004 - Daniel
  'txtBoleto.Text = rsParametros("BoletoPadrao").Value & ""
  
  If Len(rsParametros("BancoPDV").Value) > 0 Then
    txtBancoPDV.Text = rsParametros("BancoPDV").Value
  Else
    txtBancoPDV.Text = ""
  End If
  
  gsSenhaGerenteAtual = rsParametros("Senha Gerente") & ""
  Crédito_Saídas.Value = -rsParametros("Saída Verifica Limite")
  
  '06/05/2007 - Anderson
  'Implementação da opção para exibir o campo CFOP na tela de Saídas
  chkCFOP.Value = -rsParametros("ExibeCFOP")
  
  If rsParametros("Saída Parcela Padrão") = "B" Then Saída_Parcela_Banco.Value = True
  If rsParametros("Saída Parcela Padrão") = "C" Then Saída_Parcela_Carteira.Value = True
  If rsParametros("Saída Parcela Padrão") = "E" Then Saída_Parcela_Carnet.Value = True
  
  '19/04/2005 - Daniel
  'Tratamento para o campo CliWebComprarPrazo
  If rsParametros("CliWebComprarPrazo").Value Then
    chkWebCliCompraPrazo.Value = vbChecked
  Else
    chkWebCliCompraPrazo.Value = vbUnchecked
  End If
  
  Saída_Altera_Parcela.Value = -rsParametros("Saída Altera Parcela")
  Saída_Intervalo_Parc.Text = rsParametros("Saída Intervalo Parc") & ""
  
  chkSaida_Descr_Adicional.Value = -rsParametros("Saida Descr Adicional") & ""
  '02/05/2005 - Daniel
  'Tratamento para o campo VerificaLimiteCli
  If rsParametros("VerificaLimiteCli").Value Then
    chkVerificaLimiteCli.Value = vbChecked
  Else
    chkVerificaLimiteCli.Value = vbUnchecked
  End If
  
  '12/05/2005 - Daniel
  'Tratamento para o campo ExibirFabricante
  'Finalidade...: Deixamos configurável à exibição nas telas de
  '               Saída e Venda Rápida da coluna Fabricante nos
  '               dropdowns de pesquisas
  'Solicitante..: Info Social
  If rsParametros("ExibirFabricante").Value Then
    chkExibirFabricante.Value = vbChecked
  Else
    chkExibirFabricante.Value = vbUnchecked
  End If
  
  chkAlteraPreco.Value = -rsParametros("Saida Altera Preco") & ""
  
  '19/08/2003 - mpdea
  'Modificado nome do campo e controle
  '
  '06/05/2003 - mpdea ????
  '10/11/2002 - maikel
  'Retirado o IIF, para inversão é possível utilizar o sinal '-'
  '08/11/2002 - mpdea
  'Modificado o texto do controle, inversão necessária de valores (vbChecked -> vbUnchecked)
  '07/10/2002 - mpdea
  'Verificação de estoque nas movimentações de Saída
  chkVendaSemEstoqueSaidas.Value = IIf(rsParametros("Venda Sem Estoque Saidas").Value, vbChecked, vbUnchecked)
  
  
  '13/11/2002 - mpdea
  'Código da operação de saída a ser utilizada na transformação de orçamento em venda
  Call IsDataType(dtInteger, rsParametros("OpSaidaOrcVenda").Value, intRet)
  cboOperSaida_S.Text = intRet
  Call cboOperSaida_S_LostFocus
  
  Pesq1.Text = rsParametros("Nome Pesquisa 1") & ""
  Pesq2.Text = rsParametros("Nome Pesquisa 2") & ""
  Pesq3.Text = rsParametros("Nome Pesquisa 3") & ""
  
  '30/01/2004 - Campos de Impostos sobre Serviços
  txtCSLL.Text = rsParametros("CSLL") & ""
  txtCOFINS.Text = rsParametros("COFINS") & ""
  txtPIS.Text = rsParametros("PIS") & ""
  
  '11/06/2008 - mpdea
  'Valor de isenção mensal no cálculo de impostos de serviços (PIS, COFINS e CSLL)
  txtValorIsencaoPisCofinsCsll.Text = rsParametros.Fields("ValorIsencaoPisCofinsCsll").Value & ""
  
  txtIRRF.Text = rsParametros("IRRF") & ""
  '----------------------------------------------
  
  '29/11/2004 - Daniel
  'Adicionado o campo Permitir5Casas
  'que terá impacto na tela de Entradas
  'em Preço Unitário
  If rsParametros("Permitir5Casas").Value Then
    chkPermitir5Casas.Value = vbChecked
  Else
    chkPermitir5Casas.Value = vbUnchecked
  End If
  '----------------------------------------------
  
  '---[ Campos de consignação ]---'
    cboOperacaoEntrada.Text = rsParametros.Fields("Consignacao_OpEntrada") & ""
    cboOperacaoEntrada_LostFocus
    cboOperacaoSaida.Text = rsParametros.Fields("Consignacao_OpSaida") & ""
    cboOperacaoSaida_LostFocus
    cboOperacaoFechamento.Text = rsParametros.Fields("Consignacao_OpFechamento") & ""
    cboOperacaoFechamento_LostFocus
    cboCaixa.Text = rsParametros.Fields("Consignacao_Caixa") & ""
    cboCaixa_LostFocus
    cboTabelaPrecoConsignacao.Text = rsParametros.Fields("Consignacao_TabelaPrecos") & ""
  '---[ Campos de consignação ]---'
  
  '---[ Campos de transferência ]---' Pablo
    cboOpEntradaTransf.Text = rsParametros.Fields("Transf_OpEntrada") & ""
    cboOpEntradaTransf_LostFocus
    cboOpSaidaTransf.Text = rsParametros.Fields("Transf_OpSaida") & ""
    cboOpSaidaTransf_LostFocus
    cboTabPrecosTransf.Text = rsParametros.Fields("Transf_TabelaPrecos") & ""
  '---[ Campos de transferência ]---'
  
  
  Tab1.TabEnabled(1) = True
  Tab1.TabEnabled(2) = True
  Tab1.TabEnabled(3) = True
  Tab1.TabEnabled(4) = True
  
  Gerar_Conta_Paga.Value = -rsParametros("Gerar Conta Paga")
  
  '15/05/2007 - Anderson
  'Indica se o Quick Store deve manter as observações impressas na última Nota Fiscal
  chkMantemInformacaoUltimaNotaFiscal.Value = -rsParametros("MantemInformacaoUltimaNotaFiscal")
  
  Mudou_Imp_Nota = False
  Mudou_Imp_Ticket = False
  Mudou_Imp_Rel = False
  Mudou_Imp_Cheque = False
  Mudou_Imp_Boleto = False
  
  If Not IsNull(rsParametros.Fields("NumCasasDecimais").Value) Then
      If rsParametros.Fields("NumCasasDecimais").Value = 2 Then
          cmb_casasDecimaisValorUnitario.ListIndex = 0
      ElseIf rsParametros.Fields("NumCasasDecimais").Value = 3 Then
          cmb_casasDecimaisValorUnitario.ListIndex = 1
      ElseIf rsParametros.Fields("NumCasasDecimais").Value = 5 Then
          cmb_casasDecimaisValorUnitario.ListIndex = 2
      Else
          cmb_casasDecimaisValorUnitario.ListIndex = -1
      End If
  Else
      cmb_casasDecimaisValorUnitario.ListIndex = -1
  End If
  
  Exit Sub
  
Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao apresentar registro."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
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
    Case "miOpPassword"
        frmTrocaSenhaGerente.Show vbModal
 '      Call GetPassword
    
    '30/01/2009 - mpdea
    'Configuração de envio de email
    Case "miOpConfigEnvioEmail"
      Call ConfigurarEnvioEmail
    
  End Select
End Sub

Private Sub MoveFirst()
  rsParametros.MoveFirst
  If rsParametros.NoMatch Then
     Beep
     If Not IsNull(Num_Registro) Then
       rsParametros.Bookmark = Num_Registro
     End If
     Exit Sub
  End If
  Num_Registro = rsParametros.Bookmark
  ShowRecord
End Sub

Private Sub MovePrevious()
  Dim Atual As Variant
  
  On Error GoTo Processa_Erro
  
  Atual = Código.Text
  If IsNull(Atual) Then Atual = 0
  If Not IsNumeric(Atual) Then Atual = 0
  If Atual < 0 Then Atual = 0
  If Atual > 999 Then Atual = 999
  
  rsParametros.Index = "Filial"
  
  rsParametros.Seek "<", Atual
  If rsParametros.NoMatch Then
     Beep
     If Not IsNull(Num_Registro) Then
       rsParametros.Bookmark = Num_Registro
     End If
     Exit Sub
  End If
  
  Num_Registro = rsParametros.Bookmark
  ShowRecord
  
  Exit Sub
  
Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apresentar o registro em Parâmetros."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Private Sub MoveNext()
  Dim Atual As Variant
  
  On Error Resume Next
  
  Atual = Código.Text
  If IsNull(Atual) Then Atual = 0
  If Not IsNumeric(Atual) Then Atual = 0
  If Atual < 0 Then Atual = 0
  If Atual > 999 Then Atual = 999
  
  rsParametros.Index = "Filial"
  
  rsParametros.Seek ">", Atual
  If rsParametros.NoMatch Then
     Beep
     If Not IsNull(Num_Registro) Then
       rsParametros.Bookmark = Num_Registro
     End If
     Exit Sub
  End If
  
  Num_Registro = rsParametros.Bookmark
  ShowRecord
  
End Sub

Private Sub MoveLast()
  rsParametros.MoveLast
  If rsParametros.NoMatch Then
     Beep
     If Not IsNull(Num_Registro) Then
       rsParametros.Bookmark = Num_Registro
     End If
     Exit Sub
  End If
  Num_Registro = rsParametros.Bookmark
  ShowRecord
End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  Dim Resp As Integer
  Dim i As Integer
  Dim sTexto As String
  Dim bytRet As Byte
  Dim intRet As Integer
  
  If txtLayoutEnvio.Text = "" Then
   txtLayoutEnvio.Text = 2#
  End If
  
  On Error GoTo Processa_Erro
  
  Call StatusMsg("")
  
  If IsNull(Qtde_Balança.Text) Then Qtde_Balança.Text = 3
  If Qtde_Balança.Text = "" Then Qtde_Balança.Text = 3
  
  '12/09/2003 - mpdea
  'Corrigido verificação do código que estava permitindo até 999
  'e o correto é 99
  '
  'Verifica código
  If IsNull(Código.Text) Then Erro = True
  If Not Erro Then If Not IsNumeric(Código.Text) Then Erro = True
  If Not Erro Then If Val(Código.Text) < 1 Then Erro = True
  If Not Erro Then If Val(Código.Text) > 99 Then Erro = True
  
  If Erro Then
    gsTitle = LoadResString(201)
    gsMsg = "Código deve ter valor entre 1 e 99."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Qtde_Balança.Text) Then Erro = True
  If Erro = False Then If Qtde_Balança.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Qtde_Balança.Text) Then Erro = True
  If Erro = False Then If Val(Qtde_Balança.Text) < 3 Or Val(Qtde_Balança.Text) > 6 Then Erro = True
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Número de dígitos para produto deve ser 3, 4, 5 ou 6."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Qtde_Balança.SetFocus
    Exit Sub
  End If
  
  If Len(Nome.Text & "") = 0 Then
    DisplayMsg "Nome da Filial inválido."
    Nome.SetFocus
    Exit Sub
  End If
  
  Nome.Text = Replace(Nome.Text, "'", "")
  Nome.Text = Replace(Nome.Text, "´", "")
  Nome.Text = Replace(Nome.Text, "`", "")
  
  '18/09/2003 - mpdea
  'Incluído a obrigatoriedade do preenchimento do campo estado
  If Estado.Text = "" Then
    DisplayMsg "Preencher o campo Estado."
    Estado.SetFocus
    Exit Sub
  End If
  
  
  If Tab1.TabEnabled(1) Then
    If VR_Combo_Preço.Text <> "" Then
      If Not gbCheckTabPreco(VR_Combo_Preço.Text) Then
        DisplayMsg "Tabela de Preços inválida."
        Tab1.Tab = 1
        VR_Combo_Preço.SetFocus
        Exit Sub
      End If
    ElseIf VR_Altera_Tabela.Value = vbUnchecked Then
      DisplayMsg "Escolha uma Tabela de Preços ou marque a opção 'Permite alterar tabela de preços'."
      Tab1.Tab = 1
      VR_Combo_Preço.SetFocus
      Exit Sub
    End If
    
    '20/07/2002 - mpdea
    'Verifica o limite de nr. de linhas da tela de Venda Rápida em 255
    If Not IsDataType(dtByte, VR_Linhas.Text, bytRet) Then
      DisplayMsg "Quantidade máxima de linhas para a tela de Venda Rápida incorreta."
      Tab1.Tab = 1
      VR_Linhas.SetFocus
      Exit Sub
    End If
    VR_Linhas.Text = bytRet
  Else
    VR_Altera_Tabela.Value = vbChecked
  End If
  
  
  If Tab1.TabEnabled(2) Then
    
    '13/08/2004 - mpdea
    'Verifica o limite de nr. de linhas da tela de Saídas em 255
    'Produtos e serviços
    If Not IsDataType(dtByte, Máximo.Text, bytRet) Then
      DisplayMsg "Quantidade máxima de linhas para produtos na tela de Saídas incorreta."
      Tab1.Tab = 2
      Máximo.SetFocus
      Exit Sub
    End If
    Máximo.Text = bytRet
    If Not IsDataType(dtByte, Máximo_Serviço.Text, bytRet) Then
      DisplayMsg "Quantidade máxima de linhas para serviços na tela de Saídas incorreta."
      Tab1.Tab = 2
      Máximo_Serviço.SetFocus
      Exit Sub
    End If
    Máximo_Serviço.Text = bytRet
    
    '13/11/2002 - mpdea
    'Validação para a tab Saídas
    Call cboOperSaida_S_LostFocus
    If lblNomeOperSaida_S.Caption = "" Then
      DisplayMsg "Selecione a operação de saída a ser utilizada na transformação de Orçamento em Venda"
      Tab1.Tab = 2
      cboOperSaida_S.SetFocus
      Exit Sub
    End If
  End If
  
  
  '17/09/2003 - mpdea
  'Incluído chkAlteraPreco
  '
  '12/09/2003 - mpdea
  'Validação para o estado de SC
'  If UCase(Estado.Text) = "SC" Then
'    VR_Altera_Preço.Value = vbUnchecked
'    VR_Altera_Preço.Visible = False
'    chkAlteraPreco.Value = vbUnchecked
'    chkAlteraPreco.Visible = False
'  Else
'    VR_Altera_Preço.Visible = True
'    chkAlteraPreco.Visible = True
'  End If
  
  '20/05/2005 - Daniel
  '
  'A operação de saída default em Venda Rápida não
  'poderá ser uma operação que aceite emissão de nota
  'manual
  If IsNumeric(VR_Cód_Operação.Text) Then 'Alteramos a condição em 28/06/2005 se não for numérico nem continua...
    If (VR_Cód_Operação.Text <> 0) Then
      If gbNotaManual(CInt(VR_Cód_Operação.Text), "SAIDA") Then
        MsgBox "A operação de saída padrão para VR não deverá ter a característca de impressão de nota manual, verifique.", vbExclamation, "Atenção"
        VR_Cód_Operação.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  Screen.MousePointer = vbHourglass
  
  If IsNull(Ult_Nota.Text) Then Ult_Nota.Text = 0
  If Ult_Nota.Text = "" Then Ult_Nota.Text = 0
  If Not IsNumeric(Ult_Nota.Text) Then Ult_Nota.Text = 0
  If Val(Ult_Nota.Text) < 0 Then Ult_Nota.Text = 0
  
  '30/09/2009 - Andrea
  If IsNull(txtUltimaNFe.Text) Then txtUltimaNFe.Text = 0
  If txtUltimaNFe.Text = "" Then txtUltimaNFe.Text = 0
  If Not IsNumeric(txtUltimaNFe.Text) Then txtUltimaNFe.Text = 0
  If Val(txtUltimaNFe.Text) < 0 Then txtUltimaNFe.Text = 0
  
  If IsNull(txtUltimaNFCe.Text) Then txtUltimaNFCe.Text = 0
  If txtUltimaNFCe.Text = "" Then txtUltimaNFCe.Text = 0
  If Not IsNumeric(txtUltimaNFCe.Text) Then txtUltimaNFCe.Text = 0
  If Val(txtUltimaNFCe.Text) < 0 Then txtUltimaNFCe.Text = 0
    
  If IsNull(Juros.Text) Then Juros.Text = 0
  '17/08/2004 - Daniel
  If Not IsNumeric(mskTaxaDesconto.Text) Then mskTaxaDesconto.Text = 0
  
  If Juros.Text = "" Then Juros.Text = 0
  
  '--------------------------------------------------------------------
  'Início da gravação dos Parâmetros
  '--------------------------------------------------------------------
  Call StatusMsg("Gravando ...")
  
  If IsNull(Num_Registro) Then
    rsParametros.AddNew
    sTexto = "Filial inserida com sucesso."
    gsSenhaGerenteAtual = "SENHA"
    
    '12/09/2003 - mpdea
    'Modificado atualização do código da filial
    'que estava permitindo alterar
    rsParametros.Fields("Filial").Value = Código.Text
    
  Else
    'Aplica a nova Senha do Gerente em caso de alteração
    If rsParametros("Filial").Value = gnCodFilial Then
      gsSenhaGerente = gsSenhaGerenteAtual
    End If
    rsParametros.Edit
    sTexto = "Filial alterada com sucesso."
  End If
  
  Call StatusMsg("Gravando ... Parâmetros gerais.")
  
  rsParametros("Nome") = Nome.Text
  
  rsParametros("Razão Social") = Razão_Social.Text
  rsParametros("Endereço") = Endereço.Text
  rsParametros("Bairro") = Bairro.Text
  rsParametros("Fone") = Fone.Text
  rsParametros("Cidade") = Cidade.Text
  rsParametros("Estado") = Estado.Text
  rsParametros("CGC") = CGC.Text
  rsParametros("Inscrição") = Inscrição.Text
  
  
  
  'Tratamento da combo SITUAÇÃO TRIBUTÁRIO DO PIS
  If Len(cmb_situacaoTributariaDoPIS.Text) > 0 Then
      rsParametros("TipoSituacaoTributariaPIS") = Left(cmb_situacaoTributariaDoPIS.Text, 2)
  Else
      rsParametros("TipoSituacaoTributariaPIS") = 0
  End If
  'Fim tratamento combo

  
  If chk_viaRDP_ticket.Value = vbChecked Then
      rsParametros("Quick_viaRDP_ticket") = 1
  Else
      rsParametros("Quick_viaRDP_ticket") = 0
  End If
  
  If chk_viaRDP.Value = vbChecked Then
      rsParametros("Quick_viaRDP") = 1
  Else
      rsParametros("Quick_viaRDP") = 0
  End If
    
  
  '26/08/2009 - mpdea
  'Novos campos para NFe
  rsParametros.Fields("EnderecoNumero").Value = txtEnderecoNumero.Text
  rsParametros.Fields("EnderecoComplemento").Value = txtEnderecoComplemento.Text
  rsParametros.Fields("CEP").Value = txtCEP.Text
  rsParametros.Fields("Pais").Value = txtPais.Text
  rsParametros.Fields("InscricaoMunicipal").Value = txtInscricaoMunicipal.Text
  rsParametros.Fields("InscricaoSuframa").Value = txtInscricaoSuframa.Text
  rsParametros.Fields("CNAE").Value = txtCNAE.Text
  '17/09/2009 - mpdea
  rsParametros.Fields("HabilitarNotaFiscalEletronica").Value = chkHabilitarNotaFiscalEletronica.Value = vbChecked
  
  
  rsParametros("Juros") = Format(CDbl(Juros.Text), "##0.00")
  '17/08/2004 - Daniel
  rsParametros("TaxaDesconto").Value = Format(CDbl(mskTaxaDesconto.Text), "##0.00")
  
  If Val(Ult_Nota.Text) <> rsParametros("Última Nota") Then
     rsZZZLog.AddNew
       rsZZZLog("Data") = Date
       rsZZZLog("Texto") = "Nota alterada de:" & rsParametros("Última Nota") & " para:" & Val(Ult_Nota.Text) & " por:" & gnUserCode & "-" & gsUserName
       rsZZZLog("Tipo") = "Alteração"
     rsZZZLog.Update
     rsParametros("Última Nota") = Val(Ult_Nota.Text)
  End If
  
  If IsNull(rsParametros("UltimaNFe")) Then
     rsZZZLog.AddNew
       rsZZZLog("Data") = Date
       rsZZZLog("Texto") = "Nota Fiscal Eletronica alterada de: 0 para: " & Val(txtUltimaNFe.Text) & " por:" & gnUserCode & "-" & gsUserName
       rsZZZLog("Tipo") = "Alteração"
     rsZZZLog.Update
     rsParametros("UltimaNFe") = Val(txtUltimaNFe.Text)
  ElseIf Val(txtUltimaNFe.Text) <> rsParametros("UltimaNFe") Then
     rsZZZLog.AddNew
       rsZZZLog("Data") = Date
       rsZZZLog("Texto") = "Nota Fiscal Eletronica alterada de:" & rsParametros("UltimaNFe") & " para:" & Val(txtUltimaNFe.Text) & " por:" & gnUserCode & "-" & gsUserName
       rsZZZLog("Tipo") = "Alteração"
     rsZZZLog.Update
     rsParametros("UltimaNFe") = Val(txtUltimaNFe.Text)
  End If
  
  '30/09/2009 - Andrea
  
  If IsNull(rsParametros("UltimaNFCe")) Then
     rsZZZLog.AddNew
       rsZZZLog("Data") = Date
       rsZZZLog("Texto") = "Nota Fiscal ao Consumidor Eletronica alterada de: 0 para:" & Val(txtUltimaNFCe.Text) & " por:" & gnUserCode & "-" & gsUserName
       rsZZZLog("Tipo") = "Alteração"
     rsZZZLog.Update
     rsParametros("UltimaNFCe") = Val(txtUltimaNFCe.Text)
  ElseIf Val(txtUltimaNFCe.Text) <> rsParametros("UltimaNFCe") Then
     rsZZZLog.AddNew
       rsZZZLog("Data") = Date
       rsZZZLog("Texto") = "Nota Fiscal ao Consumidor Eletronica alterada de:" & rsParametros("UltimaNFCe") & " para:" & Val(txtUltimaNFCe.Text) & " por:" & gnUserCode & "-" & gsUserName
       rsZZZLog("Tipo") = "Alteração"
     rsZZZLog.Update
     rsParametros("UltimaNFCe") = Val(txtUltimaNFCe.Text)
  End If
 
  If IsNull(Máximo.Text) Then Máximo.Text = 10
  If Máximo.Text = "" Then Máximo.Text = 10
  rsParametros("Linhas Digitação") = Máximo.Text
  
  If IsNull(Máximo_Serviço.Text) Then Máximo_Serviço.Text = 5
  If Máximo_Serviço.Text = "" Then Máximo_Serviço.Text = 5
  
  rsParametros("Linhas Serviço") = Máximo_Serviço.Text
  If IsNumeric(txtDiasBloqueioVenda.Text) Then rsParametros("DiasBloqueioVenda").Value = txtDiasBloqueioVenda.Text
  rsParametros("Senha Sempre") = True
  ' If Pede_Senha.Value = 0 Then rsParametros("Senha Sempre") = False
  
'  rsParametros("Superusuário Libera Telas") = True
'  If Super_Libera.Value = 0 Then rsParametros("Superusuário Libera Telas") = False
  
  rsParametros("Venda Sem Estoque") = Sem_Estoque.Value
  rsParametros("Qtde Balança") = Val(Qtde_Balança.Text)
  rsParametros("Nota Saída") = txtConfigNFOut.Text & ""
  rsParametros("Nota Entrada") = txtConfigNFInp.Text & ""
  
  '18/08/2004 - Daniel
  If CheckSerialCaseMod("QS39823-684") Then rsParametros("TicketPadrao").Value = CStr(txtTicket.Text)
  
  '----------------------------------------------------------------------------
  '01/09/2009 - mpdea
  'NFe
  '
  'Identificação do Ambiente
  rsParametros.Fields("AmbienteNfe").Value = cboAmbienteNfe.ItemData(cboAmbienteNfe.ListIndex)
  'Formato de Impressão do DANFE
  rsParametros.Fields("FormatoImpressaoDanfeNfe").Value = cboFormatoImpressaoDanfeNfe.ItemData(cboFormatoImpressaoDanfeNfe.ListIndex)
  'Modalidade de determinação da Base de Cálculo do ICMS
  rsParametros.Fields("ModDetBaseCalculoIcms").Value = cboModDetBaseCalculoIcms.ItemData(cboModDetBaseCalculoIcms.ListIndex)
  'Modalidade de determinação da Base de Cálculo do ICMS ST
  rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = cboModDetBaseCalculoIcmsSt.ItemData(cboModDetBaseCalculoIcmsSt.ListIndex)
  'Pasta de envio
  rsParametros.Fields("PastaEnvioNfe").Value = txtPastaEnvioNfe.Text
  'Pasta de retorno
  rsParametros.Fields("PastaRetornoNfe").Value = txtPastaRetornoNfe.Text
  
  '24/11/2010 - Andrea
  'Layout Envio
  rsParametros.Fields("VersaoLayoutEnvio").Value = txtLayoutEnvio.Text
  'Codigo do Regime Tributario
  rsParametros.Fields("CodigoRegimeTributario").Value = cboCodigoRegimeTributario.ItemData(cboCodigoRegimeTributario.ListIndex)
  
  '16/11/2011 - Andrea
  'Padrão do Arquivo de Integração
  rsParametros.Fields("PadraoArquivoIntegracao").Value = cboPadraoArquivoIntegracao.Text
  
  
  '----------------------------------------------------------------------------
  
  '10/03/2009 - mpdea
  rsParametros.Fields("AliquotaAprovCreditoIcms").Value = txtAliquotaAproveitamentoCreditoIcms.Text & ""
  
  '14/03/2011 - Andrea
  rsParametros.Fields("PercentualSimplesNacional").Value = txtPercentualSimplesNacional.Text & ""
  
  '30/03/2011 - Andrea
  rsParametros.Fields("PercentualReducaoBCSimplesNacional").Value = txtPercentualReducaoBC_SN.Text & ""
  
  rsParametros("Verifica Agenda") = Verifica_Agenda.Value
  
  rsParametros("Três Tabelas") = False
  If O_Preços.Value = 1 Then rsParametros("Três Tabelas") = True
  
  rsParametros("Usar Grade") = False
  If O_Grade.Value = 1 Then rsParametros("Usar Grade") = True
  
  rsParametros("Usar Edições") = False
  If O_Edições.Value = 1 Then rsParametros("Usar Edições") = True
  
  rsParametros("Usar Códigos Alfa") = False
  If O_Alfa.Value = 1 Then rsParametros("Usar Códigos Alfa") = True
  
  
  '31/07/2002 - mpdea
  'Campo de utilização da Loja Virtual
  rsParametros.Fields("WorkWeb").Value = IIf(chkWorkWeb.Value = vbChecked, True, False)
  
  
  '29/05/2003 - mpdea
  'Campo de utilização do Traffic Light
  rsParametros.Fields("WorkTrafficLight").Value = IIf(chkWorkTrafficLight.Value = vbChecked, True, False)
  
  
  '07/08/2003 - mpdea
  'Campo de verificação de instâncias do Quick Store
  rsParametros.Fields("CheckInstance").Value = IIf(chkCheckInstance.Value = vbChecked, True, False)
  
  
  rsParametros("Usar Serviços") = False
  If Usar_Serviços.Value = 1 Then rsParametros("Usar Serviços") = True
  
  rsParametros("Alterar Serviços") = False
  If Alterar_Serviços.Value = 1 Then rsParametros("Alterar Serviços") = True
  
  rsParametros("Tabela 1") = Tabela1.Text & ""
  rsParametros("Tabela 2") = Tabela2.Text & ""
  rsParametros("Tabela 3") = Tabela3.Text & ""
  
  rsParametros("Lista 1") = Lista(0).Text
  rsParametros("Lista 2") = Lista(1).Text
  rsParametros("Lista 3") = Lista(2).Text
  rsParametros("Lista 4") = Lista(3).Text
  rsParametros("Lista 5") = Lista(4).Text
  
  '07/08/2003 - mpdea
  'Modificado controle para o campo 'Usa Vários Caixas'
'  If O_Um_Caixa.Value = True Then rsParametros("Usa Vários Caixas") = False
'  If O_Vários_Caixas.Value = True Then rsParametros("Usa Vários Caixas") = True
  rsParametros.Fields("Usa Vários Caixas").Value = chkUsaVariosCaixas.Value = vbChecked
  
  '20/11/2006 - Anderson
  'Considerar saldo anterior ao movimentar o caixa
  rsParametros.Fields("ConsiderarSaldoAnterior").Value = chkSaldoAnterior.Value = vbChecked

  '17/01/2007 - Anderson
  'Solicitar senha do gerente ao alterar vendedor nas telas de cadastro de clientes, venda rápida, saídas e check-out
  rsParametros.Fields("VendedorSenhaGerente").Value = chkVendedorSenhaGerente.Value = vbChecked
  
  '06/05/2003 - mpdea
  'Desconto no Sub Total rateado para Venda Rápida e Saídas
  rsParametros.Fields("DescSubTotalRateado").Value = chkDescSubTotalRateado.Value = vbChecked
  
  '06/05/2005 - Daniel
  'Tratamento para o campo [Parâmetros Filial].UtilizarCodFornec
  rsParametros.Fields("UtilizarCodFornec").Value = chkUtilizarCodFornec.Value = vbChecked
  
  rsParametros("Cód Comp 1") = c_comp1.Text
  rsParametros("Cód Comp 2") = c_comp2.Text
  rsParametros("Cód Comp 3") = c_comp3.Text
  rsParametros("Cód Oitavo 1") = c_oito1.Text
  rsParametros("Cód Oitavo 2") = c_oito2.Text
  rsParametros("Cód Oitavo 3") = c_oito3.Text
  rsParametros("Cód Comprim 1") = c_comp_pag1.Text
  rsParametros("Cód Comprim 2") = c_comp_pag2.Text
  rsParametros("Cód Comprim 3") = c_comp_pag3.Text
  
  If IsNull(VR_Linhas.Text) Then VR_Linhas.Text = 5
  If VR_Linhas.Text = "" Then VR_Linhas.Text = 5
  If Not IsNumeric(VR_Linhas.Text) Then VR_Linhas.Text = 5
  If Val(VR_Linhas.Text) < 10 Then VR_Linhas.Text = 10
  
  '20/07/2002 - mpdea
  'Alterado o limite de nr. de linhas da tela de Venda Rápida para 255 (Antes - 99)
  If Val(VR_Linhas.Text) > 255 Then VR_Linhas.Text = 255
  
  rsParametros("VR Linhas Digitação") = VR_Linhas.Text
  
  
  '16/01/2006 - mpdea
  'Utilização da tela de Venda Rápida em tela cheia
  rsParametros.Fields("VR_Tela_CheckOut").Value = chkVR_Tela_CheckOut.Value = vbChecked
  
  
  If IsNull(VR_Cód_Operação.Text) Then VR_Cód_Operação.Text = 500
  If VR_Cód_Operação.Text = "" Then VR_Cód_Operação.Text = 500
  If Not IsNumeric(VR_Cód_Operação.Text) Then VR_Cód_Operação.Text = 500
  rsParametros("VR Código Operação") = VR_Cód_Operação.Text
  
  If Not IsNull(VR_Combo_Preço.Text) Then
   VR_Combo_Preço.Text = Left(VR_Combo_Preço.Text, 15)
  End If
  
  
  Call StatusMsg("Gravando ... Parâmetros Venda Rápida.")
  
  
  '29/12/2003 - mpdea
  'Senha obrigatória do vendedor ao gravar venda
  rsParametros.Fields("VR_GravarExigeSenhaVend").Value = (chkVR_GravarExigeSenhaVend.Value = vbChecked)
  
  '30/01/2004 - Daniel
  'Campos de Impostos sobre Serviços
  If Not IsNumeric(Trim(txtCSLL.Text)) Then
    rsParametros.Fields("CSLL").Value = 0
  Else
    rsParametros.Fields("CSLL").Value = (Trim(txtCSLL.Text))
  End If
  
  If Not IsNumeric(Trim(txtCOFINS.Text)) Then
    rsParametros.Fields("COFINS").Value = 0
  Else
    rsParametros.Fields("COFINS").Value = (Trim(txtCOFINS.Text))
  End If
  
  If Not IsNumeric(Trim(txtPIS.Text)) Then
    rsParametros.Fields("PIS").Value = 0
  Else
    rsParametros.Fields("PIS").Value = (Trim(txtPIS.Text))
  End If
  
  '11/06/2008 - mpdea
  'Valor de isenção mensal no cálculo de impostos de serviços (PIS, COFINS e CSLL)
  If Not IsNumeric(Trim(txtValorIsencaoPisCofinsCsll.Text)) Then
    rsParametros.Fields("ValorIsencaoPisCofinsCsll").Value = 0
  Else
    rsParametros.Fields("ValorIsencaoPisCofinsCsll").Value = (Trim(txtValorIsencaoPisCofinsCsll.Text))
  End If
  
  If Not IsNumeric(Trim(txtIRRF.Text)) Then
    rsParametros.Fields("IRRF").Value = 0
  Else
    rsParametros.Fields("IRRF").Value = (Trim(txtIRRF.Text))
  End If
  '----------------------------------------------------------
  
  '29/11/2004 - Daniel
  'Adicionado o campo Permitir5Casas
  'que terá impacto na tela de Entradas
  'em Preço Unitário
  If chkPermitir5Casas.Value = vbChecked Then
    rsParametros.Fields("Permitir5Casas").Value = True
  Else
    rsParametros.Fields("Permitir5Casas").Value = False
  End If
  '----------------------------------------------------------
  
  rsParametros("VR Tab Preço") = VR_Combo_Preço.Text
  rsParametros("VR Altera Tabela") = VR_Altera_Tabela.Value
  rsParametros("VR Altera Preço") = VR_Altera_Preço.Value
  rsParametros("VR Cliente") = 0
  If VR_Nome_Cliente.Caption <> "" Then rsParametros("VR Cliente") = Val(VR_Combo_Cliente.Text)
  rsParametros("VR Altera Cliente") = VR_Altera_Cliente.Value
  rsParametros("VR Cadastra Cliente") = VR_Cadastra_Cliente.Value
  If VR_Desconto.Text = "" Then VR_Desconto.Text = "0"
  rsParametros("VR Desconto") = CSng(VR_Desconto.Text)
  
  rsParametros("VR Permite Rec Rápido") = False
  If VR_Permite_Rec_Rápido.Value = 1 Then rsParametros("VR Permite Rec Rápido") = True
  
  rsParametros("VR Recebimento Normal") = False
  If VR_Recebimento_Normal.Value = 1 Then rsParametros("VR Recebimento Normal") = True
  
  rsParametros("VR Permite Dinheiro") = False
  If VR_Permite_Dinheiro.Value = 1 Then rsParametros("VR Permite Dinheiro") = True
  
  rsParametros("VR Permite Vales") = False
  If VR_Permite_Vales.Value = 1 Then rsParametros("VR Permite Vales") = True
  
  rsParametros("VR Permite Cartão") = False
  If VR_Permite_Cartão.Value = 1 Then rsParametros("VR Permite Cartão") = True
  
  rsParametros("VR Permite Cheques") = False
  If VR_Permite_Cheques.Value = 1 Then rsParametros("VR Permite Cheques") = True
  
  rsParametros("VR Mostrar Estoque") = True
  'If Venda_Mostrar_Estoque.Value = 1 Then rsParametros("VR Mostrar Estoque") = True
  
  '----------------------------------------------------------------------------------
  '24/07/2006 - Andrea
  'Inclusão campo ExigeSenhaGerReimpTicket
  rsParametros("ExigeSenhaGerReimpTicket") = False
  If SenhaGerReimpTicket.Value = 1 Then rsParametros("ExigeSenhaGerReimpTicket") = True
  '-----------------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------------
  '04/12/2007 - Celso
  'Inclusão campo ExigeSenhaGerVndContaAtraso
  rsParametros("ExigeSenhaGerVndContaAtraso") = False
  If SenhaGerVendaAtraso.Value = 1 Then rsParametros("ExigeSenhaGerVndContaAtraso") = True
  '-----------------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------------
  '11/12/2007 - Celso
  'Inclusão campo NaoPermiteDuplicarCNPJ
  rsParametros("NaoPermiteDuplicarCNPJ") = False
  If NaoPermiteDuplicarCNPJ.Value = 1 Then rsParametros("NaoPermiteDuplicarCNPJ") = True
  '-----------------------------------------------------------------------------------
  
  rsParametros("VR Verifica Limite") = False
  If Crédito_Venda_Rápida.Value = 1 Then rsParametros("VR Verifica Limite") = True
  
  rsParametros("VR Permite Desconto") = False
  If VR_Permite_Desconto.Value = 1 Then rsParametros("VR Permite Desconto") = True
   
  '03/07/2006 - mpdea
  'Permissão para imprimir ticket somente em movimentações efetivadas
  'Solicitante: Bem me quer
  rsParametros.Fields("ImprimeTicketMovEfetivada").Value = chkImprimeTicketMovEfetivada.Value = vbChecked
   
  '04/12/2007 - Anderson
  'Permissão para imprimir nota somente em movimentações efetivadas
  rsParametros.Fields("ImprimeNotaMovEfetivada").Value = chkImprimeNotaMovEfetivada.Value = vbChecked
  
  '13/09/2012 - mpdea
  rsParametros.Fields("VRUtilizarTicketModoRelatorio").Value = chkVRUtilizarTicketModoRelatorio.Value = vbChecked


  If IsNull(VR_Qtde_Cheques.Text) Then VR_Qtde_Cheques.Text = 0
  If VR_Qtde_Cheques.Text = "" Then VR_Qtde_Cheques.Text = 0
  If Not IsNumeric(VR_Qtde_Cheques.Text) Then VR_Qtde_Cheques.Text = 0
  If Val(VR_Qtde_Cheques.Text) < 0 Then VR_Qtde_Cheques.Text = 0
  rsParametros("VR Qtde Cheques") = VR_Qtde_Cheques.Text
  
  If IsNull(VR_Prazo_Cheques.Text) Then VR_Prazo_Cheques.Text = 0
  If VR_Prazo_Cheques.Text = "" Then VR_Prazo_Cheques.Text = 0
  If Not IsNumeric(VR_Prazo_Cheques.Text) Then VR_Prazo_Cheques.Text = 0
  If Val(VR_Prazo_Cheques.Text) < 0 Then VR_Prazo_Cheques.Text = 0
  rsParametros("VR Prazo Cheques") = VR_Prazo_Cheques.Text
  
  
  rsParametros("VR Permite Parcela") = False
  If VR_Permite_Parcela.Value = 1 Then rsParametros("VR Permite Parcela") = True
  
  '25/05/2004 - Daniel
  'Inclusão do campo VR_RecalcularPreço
  If chkRecalculo.Value Then
    rsParametros("VR_RecalcularPreço").Value = True
  Else
    rsParametros("VR_RecalcularPreço").Value = False
  End If
  If chkOcultaOrc.Value Then
    rsParametros("VR_OcultaOrc").Value = True
  Else
    rsParametros("VR_OcultaOrc").Value = False
  End If
  If chkPrestServ.Value Then
    rsParametros("comPrestServ").Value = True
  Else
    rsParametros("comPrestServ").Value = False
  End If
  
  If IsNull(VR_Qtde_Parcela.Text) Then VR_Qtde_Parcela.Text = 0
  If VR_Qtde_Parcela.Text = "" Then VR_Qtde_Parcela.Text = 0
  If Not IsNumeric(VR_Qtde_Parcela.Text) Then VR_Qtde_Parcela.Text = 0
  If Val(VR_Qtde_Parcela.Text) < 0 Then VR_Qtde_Parcela.Text = 0
  rsParametros("VR Qtde Parcela") = VR_Qtde_Parcela.Text
  
  If IsNull(VR_Prazo_Parcela.Text) Then VR_Prazo_Parcela.Text = 0
  If VR_Prazo_Parcela.Text = "" Then VR_Prazo_Parcela.Text = 0
  If Not IsNumeric(VR_Prazo_Parcela.Text) Then VR_Prazo_Parcela.Text = 0
  If Val(VR_Prazo_Parcela.Text) < 0 Then VR_Prazo_Parcela.Text = 0
  rsParametros("VR Prazo Parcela") = VR_Prazo_Parcela.Text
  
  
  rsParametros("VR Altera Parcela") = False
  If VR_Altera_Parcela.Value = 1 Then rsParametros("VR Altera Parcela") = True
  
  If VR_Parcela_Banco.Value = True Then rsParametros("VR Parcela Padrão") = "B"
  If VR_Parcela_Carteira.Value = True Then rsParametros("VR Parcela Padrão") = "C"
  If VR_Parcela_Carnet.Value = True Then rsParametros("VR Parcela Padrão") = "E"
  
  
  If IsNull(VR_Intervalo_Parc.Text) Then VR_Intervalo_Parc.Text = 30
  If VR_Intervalo_Parc.Text = "" Then VR_Intervalo_Parc.Text = 30
  If Not IsNumeric(VR_Intervalo_Parc.Text) Then VR_Intervalo_Parc.Text = 30
  If Val(VR_Intervalo_Parc.Text) < 1 Then VR_Intervalo_Parc.Text = 30
  If Val(VR_Intervalo_Parc.Text) > 150 Then VR_Intervalo_Parc.Text = 30
  
  rsParametros("VR Intervalo Parc") = Val(VR_Intervalo_Parc.Text)
  
  rsParametros("VR Conta Padrão") = "C"
  If O_Conta_Fixa.Value = True Then rsParametros("VR Conta Padrão") = "F"
  rsParametros("VR Conta Usar") = Val(Combo_Conta.Text)
  
  rsParametros("PesquisaCodigoENome_VR") = optLocalizarCodigoNome   ' -chkProcuraCodigoENome_Saida
    
  If GetCodigoCombos(cboOrdenacao.Text) = 1 Then    '1 é Numerica, 2 é alfa numerica
    rsParametros("VROrdenacaoCombo") = True
  Else
    rsParametros("VROrdenacaoCombo") = False
  End If
    
  Call StatusMsg("Gravando ... Parâmetros Outros.")
  
  rsParametros("Consulta TAB1") = Com_Tab1.Text
  rsParametros("Consulta TAB2") = Com_Tab2.Text
  rsParametros("Consulta TAB3") = Com_Tab3.Text
  rsParametros("Consulta TAB4") = Com_Tab4.Text
  rsParametros("Consulta TAB5") = Com_Tab5.Text
  rsParametros("Consulta TAB6") = Com_Tab6.Text
  
  rsParametros("Mensagem Troca") = Mensagem_Troca.Text
  rsParametros("Mensagem Etiq 1") = Mens_Etiq1.Text
  rsParametros("Mensagem Etiq 2") = Mens_Etiq2.Text
  
  '26/05/2004 - Daniel
  'Inclusão do campo [Zero a Esquerda]
  If chk0aEsquerda.Value = vbChecked Then
    rsParametros("Zero a Esquerda").Value = True
  Else
    rsParametros("Zero a Esquerda").Value = False
  End If
  
  '09/08/2005 - Daniel
  'Inclusão do campo AlterVendedorCliFor
  'Finalidade: Apenas o Superusuário poderá alterar o campo
  '            Vendedor no cadastro Cli / For
  If chkAlterVendedorCliFor.Value = vbChecked Then
    rsParametros("AlterVendedorCliFor").Value = True
  Else
    rsParametros("AlterVendedorCliFor").Value = False
  End If
  
  '17/08/2004 - Daniel
  'If Len(txtBoleto.Text) <= 0 Then
    'rsParametros("BoletoPadrao").Value = ""
  'Else
    'rsParametros("BoletoPadrao").Value = CStr(txtBoleto.Text)
  'End If
  
  If Len(txtBancoPDV.Text) <= 0 Then
    rsParametros("BancoPDV").Value = ""
  Else
    rsParametros("BancoPDV").Value = CStr(txtBancoPDV.Text)
  End If
  
  'Senha do Gerente
  rsParametros("Senha Gerente") = gsSenhaGerenteAtual
  
  rsParametros("Impressora Cheques") = ""
  
  rsParametros("Imprimir Centavos") = False
  
  Call StatusMsg("Gravando ... Parâmetros Cheques.")
  
  rsParametros("Saída Altera Parcela") = False
  
  '19/04/2005 - Daniel
  'Tratamento para o campo CliWebComprarPrazo
  If chkWebCliCompraPrazo.Value = vbChecked Then
    rsParametros("CliWebComprarPrazo").Value = True
  Else
    rsParametros("CliWebComprarPrazo").Value = False
  End If
  
  If Saída_Altera_Parcela.Value = 1 Then rsParametros("Saída Altera Parcela") = True
  
  rsParametros("Saída Verifica Limite") = False
  If Crédito_Saídas.Value = 1 Then rsParametros("Saída Verifica Limite") = True
  
  '06/05/2007 - Anderson
  'Implementação da opção para exibir o campo CFOP na tela de Saídas
  rsParametros("ExibeCFOP") = False
  If chkCFOP.Value = 1 Then rsParametros("ExibeCFOP") = True
  
  rsParametros("Saida Descr Adicional") = False
  If chkSaida_Descr_Adicional.Value = 1 Then rsParametros("Saida Descr Adicional") = True
  '02/05/2005 - Daniel
  'Tratamento para o campo VerificaLimiteCli
  If chkVerificaLimiteCli.Value = vbChecked Then
    rsParametros("VerificaLimiteCli").Value = True
  Else
    rsParametros("VerificaLimiteCli").Value = False
  End If
  
  '12/05/2005 - Daniel
  'Tratamento para o campo ExibirFabricante
  'Finalidade...: Deixamos configurável à exibição nas telas de
  '               Saída e Venda Rápida da coluna Fabricante nos
  '               dropdowns de pesquisas
  'Solicitante..: Info Social
  If chkExibirFabricante.Value Then
    rsParametros("ExibirFabricante").Value = True
  Else
    rsParametros("ExibirFabricante").Value = False
  End If
  
  rsParametros("Saida Altera Preco") = False
  If chkAlteraPreco.Value = 1 Then rsParametros("Saida Altera Preco") = True
  
  '19/08/2003 - mpdea
  'Modificado nome do campo e controle
  '
  '06/05/2003 - mpdea ????
  '19/02/2003 - maikel
  'Retirado o IIF, e colocado o sinal de '-', pois não estava funcionando com VBChecked e VBUnChecked
  '08/11/2002 - mpdea
  'Modificado o texto do controle, inversão necessária de valores (True -> False)
  '07/10/2002 - mpdea
  'Verificação de estoque nas movimentações de Saída
  rsParametros("Venda Sem Estoque Saidas").Value = chkVendaSemEstoqueSaidas.Value = vbChecked
  
  
  '13/11/2002 - mpdea
  'Código da operação de saída a ser utilizada na transformação de orçamento em venda
  If cboOperSaida_S.Text = "" Then cboOperSaida_S.Text = "500"
  Call IsDataType(dtInteger, cboOperSaida_S.Text, intRet)
  rsParametros("OpSaidaOrcVenda").Value = intRet
  
  'Tratar casas decimais
  If cmb_casasDecimaisValorUnitario.ListIndex = 0 Then
      rsParametros.Fields("NumCasasDecimais").Value = 2
      g_bln5CasasDecimais = False
      g_bln3CasasDecimais = False
  ElseIf cmb_casasDecimaisValorUnitario.ListIndex = 1 Then
      rsParametros.Fields("NumCasasDecimais").Value = 3
      g_bln5CasasDecimais = False
      g_bln3CasasDecimais = True
  ElseIf cmb_casasDecimaisValorUnitario.ListIndex = 2 Then
      rsParametros.Fields("NumCasasDecimais").Value = 5
      g_bln5CasasDecimais = True
      g_bln3CasasDecimais = False
  Else
      rsParametros.Fields("NumCasasDecimais").Value = 0
      g_bln5CasasDecimais = False
      g_bln3CasasDecimais = False
  End If
  
  
  If Saída_Parcela_Banco.Value = True Then
      rsParametros("Saída Parcela Padrão") = "B"
  End If
  
  If Saída_Parcela_Carteira.Value = True Then
      rsParametros("Saída Parcela Padrão") = "C"
  End If
  
  If Saída_Parcela_Carnet.Value = True Then
      rsParametros("Saída Parcela Padrão") = "E"
  End If
  
  rsParametros("Saída Altera Parcela") = False
  If Saída_Altera_Parcela.Value = 1 Then rsParametros("Saída Altera Parcela") = True
  
  If IsNull(Saída_Intervalo_Parc.Text) Then Saída_Intervalo_Parc.Text = 30
  If Saída_Intervalo_Parc.Text = "" Then Saída_Intervalo_Parc.Text = 30
  If Not IsNumeric(Saída_Intervalo_Parc.Text) Then Saída_Intervalo_Parc.Text = 30
  If Val(Saída_Intervalo_Parc.Text) < 1 Then Saída_Intervalo_Parc.Text = 30
  If Val(Saída_Intervalo_Parc.Text) > 150 Then Saída_Intervalo_Parc.Text = 30
  
  rsParametros("Saída Intervalo Parc") = Val(Saída_Intervalo_Parc.Text)
  
  
  If IsNull(Pesq1.Text) Then Pesq1.Text = ""
  If IsNull(Pesq2.Text) Then Pesq2.Text = ""
  If IsNull(Pesq3.Text) Then Pesq3.Text = ""
  Pesq1.Text = Trim(Pesq1.Text)
  Pesq2.Text = Trim(Pesq2.Text)
  Pesq3.Text = Trim(Pesq3.Text)
   
  '---[ Campos de consignação ]---'
    rsParametros.Fields("Consignacao_OpEntrada") = IIf(IsNumeric(cboOperacaoEntrada.Text), cboOperacaoEntrada.Text, 0)
    rsParametros.Fields("Consignacao_OpSaida") = IIf(IsNumeric(cboOperacaoSaida.Text), cboOperacaoSaida.Text, 0)
    rsParametros.Fields("Consignacao_OpFechamento") = IIf(IsNumeric(cboOperacaoFechamento.Text), cboOperacaoFechamento.Text, 0)
    rsParametros.Fields("Consignacao_Caixa") = IIf(IsNumeric(cboCaixa.Text), cboCaixa.Text, 0)
    rsParametros.Fields("Consignacao_TabelaPrecos") = cboTabelaPrecoConsignacao.Text
  '---[ Campos de consignação ]---'
   
  '---[ Campos de transferência ]---' Pablo
    rsParametros.Fields("Transf_OpEntrada") = IIf(IsNumeric(cboOpEntradaTransf.Text), cboOpEntradaTransf.Text, 0)
    rsParametros.Fields("Transf_OpSaida") = IIf(IsNumeric(cboOpSaidaTransf.Text), cboOpSaidaTransf.Text, 0)
    rsParametros.Fields("Transf_TabelaPrecos") = cboTabPrecosTransf.Text
  '---[ Campos de transferência ]---'
   
  rsParametros("Nome Pesquisa 1") = Pesq1.Text
  rsParametros("Nome Pesquisa 2") = Pesq2.Text
  rsParametros("Nome Pesquisa 3") = Pesq3.Text
  
  
  '13/04/2005 - pablo
  'Campo para alteração do nome do produto
  rsParametros.Fields("EditarNomeProduto").Value = IIf(chkProdutoNomeNFe.Value = vbChecked, True, False)

  
  ' ============================================================================
  ' MULTA DE JUROS APÓS VENCIMENTO DE PARCELA (CREDIARIO)
  If chk_cobrarMulta.Value = vbChecked Then
      rsParametros("CobrarMultaAposVencimentoParcela") = 1
      
      If Trim(mskTaxaMultaParcelaVencida.Text) = "0" Or Trim(mskTaxaMultaParcelaVencida.Text) = "" Then
          MsgBox "Você marcou que cobrará multa após vencimento de parcela em atraso. Então informe a taxa da multa.", vbInformation, "Atenção"
          Screen.MousePointer = vbDefault
          mskTaxaMultaParcelaVencida.SetFocus
          Exit Sub
      End If
  Else
      rsParametros("CobrarMultaAposVencimentoParcela") = 0
  End If
  
  If mskTaxaMultaParcelaVencida.Text <> "" Then
      rsParametros("TaxaMultaParcelaVencida") = mskTaxaMultaParcelaVencida.Text
  Else
      rsParametros("TaxaMultaParcelaVencida") = 0
  End If
  
  If txt_multaDiasAposParcelaVencida.Text <> "" Then
      rsParametros("MultaDiasAposParcelaVencida") = txt_multaDiasAposParcelaVencida.Text
  Else
      rsParametros("MultaDiasAposParcelaVencida") = 0
  End If
  ' ============================================================================
  
  rsParametros("Gerar Conta Paga") = False
  If Gerar_Conta_Paga.Value = 1 Then rsParametros("Gerar Conta Paga") = True
    
  '13/03/2013-Alexandre Afornali
  'Gravar filtrar produtos
  rsParametros("FiltrarProdutosInativos").Value = chkFiltrarProdutosInativos.Value
  
  '17/05/2013-Alexandre Afornali
  'Gravar trabalhar com comandas
  rsParametros("TrabalharComComanda").Value = chkComandas.Value

  '15/05/2007 - Anderson
  'Indica se o Quick Store deve manter as observações impressas na última Nota Fiscal
  rsParametros("MantemInformacaoUltimaNotaFiscal") = False
  If chkMantemInformacaoUltimaNotaFiscal.Value = 1 Then rsParametros("MantemInformacaoUltimaNotaFiscal") = True
  
  rsParametros.Update
  Num_Registro = rsParametros.LastModified
  rsParametros.Bookmark = Num_Registro
  
  Call StatusMsg("Gravação OK.")
  
  Glob_Cod_Alfa = rsParametros("Usar Códigos Alfa")
    
  gbGrade = rsParametros("Usar Grade") = True
  gbEdicao = rsParametros("Usar Edições") = True
  gbServico = rsParametros("Usar Serviços") = True
  gbCaixas = rsParametros("Usa Vários Caixas") = True
  '20/11/2006 - Anderson
  'Considerar saldo anterior ao movimentar o caixa
  gbSaldoAnterior = rsParametros("ConsiderarSaldoAnterior") = True
  
  '17/01/2007 - Anderson
  'Solicitar senha do gerente ao alterar vendedor nas telas de cadastro de clientes, venda rápida, saídas e check-out
  gbVendedorSenhaGerente = rsParametros("VendedorSenhaGerente") = True
  
  '31/07/2002 - mpdea
  'Campo de utilização da Loja Virtual
  gblnWorkWeb = rsParametros.Fields("WorkWeb").Value
  
  '17/09/2009 - mpdea
  'Habilitar uso de Nota Fiscal Eletrônica
  gblnNFe = rsParametros.Fields("HabilitarNotaFiscalEletronica").Value
  
  gsPesq1 = rsParametros("Nome Pesquisa 1") & ""
  gsPesq2 = rsParametros("Nome Pesquisa 2") & ""
  gsPesq3 = rsParametros("Nome Pesquisa 3") & ""
  
    
  If rsParametros("Filial") = gnCodFilial Then
    Call StatusMsg("Reabilitando Menus ...")
    
  '29/01/2009 - mpdea
  'Seta acessos ao menu
  SetMenuAcesso
    
'    Call SetEnabledMenus
    Call StatusMsg("")
  End If
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
'  DisplayMsg sTexto
  
  Tab1.TabEnabled(1) = True
  Tab1.TabEnabled(2) = True
  Tab1.TabEnabled(3) = True
  Tab1.TabEnabled(4) = True
  
  Tab1.Tab = 0
  
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
' Joga dados da empresa para o banco do GestoPDV por causa do PAF
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
   Dim GestoBD As Database
   Dim Parametro As Recordset
   If frmParametros.VerificaPAF = True Then
     Set rsParametros2 = db.OpenRecordset("Select [BancoPDV] from [Parâmetros Filial] Where Filial = " & gnCodFilial & ";")
     
     Dim fso As New FileSystemObject
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists(rsParametros2("BancoPDV").Value & "\Gesto.mde") Then
     
     Set GestoBD = OpenDatabase(rsParametros2("BancoPDV").Value & "\Gesto.mde", False, False)
     Set Parametro = GestoBD.OpenRecordset("Select * from Parametro Where CGC =  '" & CGC.Text & "';")
     If Parametro.EOF Then
     Parametro.AddNew
     Parametro!razao_Social = Razão_Social.Text
     If Nome.Text <> "" Then
     Parametro!NOME_FANTASIA = Nome.Text
     End If
     Parametro!ENDERECO = Endereço.Text
     If txtEnderecoComplemento.Text <> "" Then
       Parametro!Complemento_Endereco = txtEnderecoComplemento.Text
     End If
     If txtEnderecoNumero.Text <> "" Then
     Parametro!NumRua = txtEnderecoNumero.Text
     End If
     If Bairro.Text <> "" Then
     Parametro!Bairro = Bairro.Text
     End If
     If txtCEP.Text <> "" Then
     Parametro!CEP = txtCEP.Text
     End If
     Parametro!Cidade = Cidade.Text
     Parametro!UF = Estado.Text
     If Fone.Text <> "" Then
     Parametro!DDD = Left(Fone.Text, 3)
     Parametro!Telefone = Right(Fone.Text, 14)
     End If
     Parametro!CGC = CGC.Text
     If Inscrição.Text <> "" Then
     Parametro!INSC_ESTADUAL = Inscrição.Text
     End If
     If txtInscricaoMunicipal.Text <> "" Then
     Parametro!inscrMunicipal = txtInscricaoMunicipal.Text
     End If
     Parametro.Update
     Else
     Parametro.Edit
     Parametro!razao_Social = Razão_Social.Text
     If Nome.Text <> "" Then
     Parametro!NOME_FANTASIA = Nome.Text
     End If
     Parametro!ENDERECO = Endereço.Text
     If txtEnderecoComplemento.Text <> "" Then
       Parametro!Complemento_Endereco = txtEnderecoComplemento.Text
     End If
     If txtEnderecoNumero.Text <> "" Then
     Parametro!NumRua = txtEnderecoNumero.Text
     End If
     If Bairro.Text <> "" Then
     Parametro!Bairro = Bairro.Text
     End If
     If txtCEP.Text <> "" Then
     Parametro!CEP = txtCEP.Text
     End If
     Parametro!Cidade = Cidade.Text
     Parametro!UF = Estado.Text
     If Fone.Text <> "" Then
     Parametro!DDD = Left(Fone.Text, 3)
     Parametro!Telefone = Right(Fone.Text, 14)
     End If
     Parametro!CGC = CGC.Text
     If Inscrição.Text <> "" Then
     Parametro!INSC_ESTADUAL = Inscrição.Text
     End If
     If txtInscricaoMunicipal.Text <> "" Then
     Parametro!inscrMunicipal = txtInscricaoMunicipal.Text
     End If
     Parametro.Update
     End If
   End If
   End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

  Exit Sub
  
  
  
Processa_Erro:
  '12/09/2003 - mpdea
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar atualizar tabela de Parâmetros de Empresa/Filial."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  On Error GoTo 0

End Sub

Public Sub ClearScreen()
  
  On Error Resume Next
  
  Call StatusMsg("")
  
  cmb_casasDecimaisValorUnitario.ListIndex = -1
  
  Código.Text = ""
  Nome.Text = ""
  
  Razão_Social.Text = ""
  Endereço.Text = ""
  Bairro.Text = ""
  Fone.Text = ""
  Cidade.Text = ""
  Estado.Text = ""
  CGC.Text = ""
  Inscrição.Text = ""
  
  '26/08/2009 - mpdea
  'Novos campos para NFe
  txtEnderecoNumero.Text = ""
  txtEnderecoComplemento.Text = ""
  txtCEP.Text = ""
  txtPais.Text = ""
  txtInscricaoMunicipal.Text = ""
  txtInscricaoSuframa.Text = ""
  txtCNAE.Text = ""
  
  '07/08/2003 - mpdea
  'Modificado controle
'  O_Um_Caixa.Value = True
  chkUsaVariosCaixas.Value = vbUnchecked
  '20/11/2006 - Anderson
  'Considerar saldo anterior ao movimentar o caixa
  chkSaldoAnterior.Value = vbUnchecked
  '17/01/2007 - Anderson
  'Solicitar senha do gerente ao alterar vendedor nas telas de cadastro de clientes, venda rápida, saídas e check-out
  chkVendedorSenhaGerente.Value = vbUnchecked
  
  Juros.Text = 0
  '17/08/2004 - Daniel
  mskTaxaDesconto.Text = 0
  Verifica_Agenda.Value = 1
  
  O_Preços.Value = 0
  O_Grade.Value = 0
  O_Edições.Value = 0
  O_Alfa.Value = 0
  'chkProcuraCodigoENome = 0
  optLocalizarCodigo.Value = True
  
  
  '31/07/2002 - mpdea
  'Campo de utilização da Loja Virtual
  chkWorkWeb.Value = vbUnchecked
  
  
  '29/05/2003 - mpdea
  'Campo de utilização do Traffic Light
  chkWorkTrafficLight.Value = vbUnchecked
  
  
  '07/08/2003 - mpdea
  'Campo de verificação de instâncias do Quick Store
  chkCheckInstance.Value = vbUnchecked
  
  
  Usar_Serviços.Value = 0
  Alterar_Serviços.Value = 0
  Tabela1.Text = ""
  Tabela2.Text = ""
  Tabela3.Text = ""
  
  txtUltimaConsignacao.Text = ""
  Ult_Nota.Text = ""
  
  '30/09/2009 - Andrea
  txtUltimaNFe.Text = ""
  txtUltimaNFCe.Text = ""
  
  Máximo.Text = ""
  Máximo_Serviço.Text = ""
  ' Pede_Senha.Value = 0
'  Super_Libera.Value = 0
  Sem_Estoque.Value = False
  Ult_Mov.Text = ""
  txtConfigNFOut.Text = ""
  '18/08/2004 - Daniel
  If CheckSerialCaseMod("QS39823-684") Then txtTicket.Text = ""
  
  txtConfigNFInp.Text = ""
  
  '01/09/2009 - mpdea
  'NFe
  cboAmbienteNfe.ListIndex = 0
  cboFormatoImpressaoDanfeNfe.ListIndex = 0
  cboModDetBaseCalculoIcms.ListIndex = 0
  cboModDetBaseCalculoIcmsSt.ListIndex = 0
  txtPastaEnvioNfe.Text = ""
  txtPastaRetornoNfe.Text = ""
  cboPadraoArquivoIntegracao.ListIndex = 0
  
  '24/11/2010 - Andrea
  txtLayoutEnvio.Text = ""
  cboCodigoRegimeTributario.ListIndex = 0
   
  
  '17/09/2009 - mpdea
  chkHabilitarNotaFiscalEletronica.Value = vbUnchecked
  
  '10/03/2009 - mpdea
  txtAliquotaAproveitamentoCreditoIcms.Text = "0"
  
  '14/03/2011 - Andrea
  txtPercentualSimplesNacional.Text = "0"
  
  '30/03/2011 - Andrea
  txtPercentualReducaoBC_SN.Text = "0"

  Qtde_Balança.Text = ""
  
  Lista(0).Text = ""
  Lista(1).Text = ""
  Lista(2).Text = ""
  Lista(3).Text = ""
  Lista(4).Text = ""
  
  '06/05/2003 - mpdea
  'Desconto no Sub Total rateado para Venda Rápida e Saídas
  chkDescSubTotalRateado.Value = vbUnchecked
  
  '06/05/2005 - Daniel
  'Tratamento para o campo [Parâmetros Filial].UtilizarCodFornec
  chkUtilizarCodFornec.Value = vbUnchecked
  
  txtDiasBloqueioVenda.Text = 0
  
  c_comp1.Text = ""
  c_comp2.Text = ""
  c_comp3.Text = ""
  
  c_oito1.Text = ""
  c_oito2.Text = ""
  c_oito3.Text = ""
  
  c_comp_pag1.Text = ""
  c_comp_pag2.Text = ""
  c_comp_pag3.Text = ""
  
  
  VR_Linhas.Text = ""
  
  
  '16/01/2006 - mpdea
  'Utilização da tela de Venda Rápida em tela cheia
  chkVR_Tela_CheckOut.Value = vbUnchecked


  VR_Cód_Operação.Text = ""
  VR_Combo_Preço.Text = ""
  VR_Altera_Tabela.Value = False
  VR_Altera_Preço.Value = False
  VR_Altera_Preço.Visible = True
  chkAlteraPreco.Visible = True
  
  '29/12/2003 - mpdea
  'Senha obrigatória do vendedor ao gravar venda
  chkVR_GravarExigeSenhaVend.Value = vbUnchecked
  
  VR_Combo_Cliente.Text = ""
  VR_Altera_Cliente.Value = False
  VR_Cadastra_Cliente.Value = False
  VR_Desconto.Text = ""
  
  VR_Permite_Rec_Rápido.Value = 0
  VR_Recebimento_Normal.Value = 0
  VR_Permite_Dinheiro.Value = 0
  VR_Permite_Vales.Value = 0
  VR_Permite_Cartão.Value = 0
  VR_Permite_Cheques.Value = 0
  VR_Qtde_Cheques.Text = ""
  VR_Prazo_Cheques.Text = ""
  VR_Permite_Parcela.Value = 0
  '25/05/2004 - Daniel
  'Inclusão do campo VR_RecalcularPreço
  chkRecalculo.Value = ""
  chkOcultaOrc.Value = ""
  chkPrestServ.Value = ""
  VR_Qtde_Parcela.Text = ""
  VR_Prazo_Parcela.Text = ""
  VR_Altera_Parcela.Value = 0
  VR_Permite_Desconto.Value = 0
  VR_Parcela_Banco.Value = True
  
  '03/07/2006 - mpdea
  'Permissão para imprimir ticket somente em movimentações efetivadas
  'Solicitante: Bem me quer
  chkImprimeTicketMovEfetivada.Value = vbUnchecked
  
  '04/12/2007 - Anderson
  'Permissão para imprimir nota somente em movimentações efetivadas
  chkImprimeNotaMovEfetivada.Value = vbUnchecked

  '13/09/2012 - mpdea
  chkVRUtilizarTicketModoRelatorio.Value = vbUnchecked

'  Venda_Mostrar_Estoque.Value = 0

  
  
  

  VR_Intervalo_Parc.Text = ""
  
  cboOrdenacao.Text = "1 - Numérica"
  
  O_Conta_Cadastro.Value = True
  Combo_Conta.Text = ""
  Nome_Conta.Caption = ""
  
  
  Crédito_Venda_Rápida.Value = 0
  
  
  Com_Tab1.Text = ""
  Com_Tab2.Text = ""
  Com_Tab3.Text = ""
  Com_Tab4.Text = ""
  Com_Tab5.Text = ""
  Com_Tab6.Text = ""
  
  Mensagem_Troca.Text = ""
  Mens_Etiq1.Text = ""
  Mens_Etiq2.Text = ""
  
  '26/05/2004 - Daniel
  'Inclusão do campo [Zero a Esquerda]
  chk0aEsquerda.Value = vbUnchecked
  '09/08/2005 - Daniel
  'Inclusão do campo AlterVendedorCliFor
  'Finalidade: Apenas o Superusuário poderá alterar o campo
  '            Vendedor no cadastro Cli / For
  chkAlterVendedorCliFor.Value = vbUnchecked
  txtBancoPDV.Text = ""
  
  '19/04/2005 - Daniel
  'Tratamento para o campo CliWebComprarPrazo
  chkWebCliCompraPrazo.Value = vbUnchecked
  '------------------------------------------
  Saída_Altera_Parcela.Value = 0
  Saída_Parcela_Banco.Value = True
  Saída_Intervalo_Parc.Text = ""
  Crédito_Saídas.Value = 0
  
  '06/05/2007 - Anderson
  'Implementação da opção para exibir o campo CFOP na tela de Saídas
  chkCFOP.Value = 0
  
  chkSaida_Descr_Adicional.Value = 0
  '02/05/2005 - Daniel
  'Tratamento para o campo VerificaLimiteCli
  chkVerificaLimiteCli.Value = vbUnchecked
  '12/05/2005 - Daniel
  'Tratamento para o campo ExibirFabricante
  'Finalidade...: Deixamos configurável à exibição nas telas de
  '               Saída e Venda Rápida da coluna Fabricante nos
  '               dropdowns de pesquisas
  'Solicitante..: Info Social
  chkExibirFabricante.Value = vbUnchecked
  '------------------------------------------
  chkAlteraPreco.Value = 0
  
  
  '08/11/2002 - mpdea
  'Modificado o texto do controle, inversão necessária de valores (vbUnChecked -> vbchecked)
  '07/10/2002 - mpdea
  'Verificação de estoque nas movimentações de Saída
  chkVendaSemEstoqueSaidas.Value = vbChecked
  
  
  '13/11/2002 - mpdea
  'Novo controle
  cboOperSaida_S.Text = ""
  
  
  Pesq1.Text = ""
  Pesq2.Text = ""
  Pesq3.Text = ""
  
  '30/01/2004 - Daniel
  'Campos de Impostos sobre Serviços
  txtCSLL.Text = ""
  txtCOFINS.Text = ""
  txtPIS.Text = ""

  '11/06/2008 - mpdea
  'Valor de isenção mensal no cálculo de impostos de serviços (PIS, COFINS e CSLL)
  txtValorIsencaoPisCofinsCsll.Text = ""
  
  txtIRRF.Text = ""
  '----------------------------------
  
  '29/11/2004 - Daniel
  'Adicionado o campo Permitir5Casas
  'que terá impacto na tela de Entradas
  'em Preço Unitário
  chkPermitir5Casas.Value = vbUnchecked
  '----------------------------------
  
  Gerar_Conta_Paga.Value = 0
  
  '15/05/2007 - Anderson
  'Indica se o Quick Store deve manter as observações impressas na última Nota Fiscal
  chkMantemInformacaoUltimaNotaFiscal.Value = 0
  
  If Not rsParametros.EOF Then
    On Error Resume Next
    rsParametros.MoveFirst
    rsParametros.MovePrevious
    On Error GoTo 0
  End If
  
  '---[ Gera a combo de tabelas de preços para consignação ]---'
    Dim rstPrecos As Recordset
    
    Set rstPrecos = db.OpenRecordset("SELECT DISTINCT Tabela FROM [Tabela de Preços]", dbOpenSnapshot)
    
    With rstPrecos
      cboTabelaPrecoConsignacao.Clear
      cboTabPrecosTransf.Clear
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do While Not .EOF
          cboTabelaPrecoConsignacao.AddItem .Fields("Tabela") & ""
          cboTabPrecosTransf.AddItem .Fields("Tabela") & ""
          .MoveNext
        Loop
      End If
      
      If Not rstPrecos Is Nothing Then .Close
      Set rstPrecos = Nothing
    End With
  '---[ Gera a combo de tabelas de preços para consignação ]---'
  
  '---[ Campos de consignação ]---'
    cboOperacaoEntrada.Text = "0"
    cboOperacaoEntrada_LostFocus
    cboOperacaoSaida.Text = "0"
    cboOperacaoSaida_LostFocus
    cboOperacaoFechamento.Text = "0"
    cboOperacaoFechamento_LostFocus
    cboCaixa.Text = "1"
    cboTabelaPrecoConsignacao.Text = ""
    cboTabPrecosTransf.Text = ""
  '---[ Campos de consignação ]---'
  
  Num_Registro = Null
  
  Código.SetFocus
  
  Tab1.Tab = 0
  Tab1.TabEnabled(1) = False
  Tab1.TabEnabled(2) = False
  Tab1.TabEnabled(3) = False
  Tab1.TabEnabled(4) = False
  
  '13/03/2013-Alexandre Afornali
  'Campo para nao filtrar clientes inativos
  chkFiltrarProdutosInativos.Value = vbUnchecked
  
  '17/05/2013-Alexandre Afornali
  'Campo para trabalhar ou não com Comandas
  chkComandas.Value = vbUnchecked
  
End Sub

Private Sub GetPassword()
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Encontre ou grave um registro antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  
  frmTrocaSenhaGerente.Show vbModal
  
End Sub

Private Sub btnProcurar_Click()
'Dim sFileName As String
'
'  sFileName = gsOpenFile(Me, "Escolha o Arquivo de Banco de Dados", "Bancos Access | *.mde")
'
'  If Len(sFileName) = 0 Then
'    Exit Sub
'  End If
'
'  txtBancoPDV.Text = gsGetOnlyName(sFileName)
Dim sRet As String
  
  sRet = sFindDir("Selecione o diretório.", Me.hwnd)
  
  If Len(sRet) = 0 Then
    Exit Sub
  End If
  txtBancoPDV.Text = sRet
End Sub

Private Sub cboCaixa_CloseUp()
  cboCaixa.Text = cboCaixa.Columns(0).Text & ""
  cboCaixa_LostFocus
End Sub

Private Sub cboCaixa_LostFocus()
  Dim rstCaixa As Recordset
  
  txtCaixa.Text = ""
  
  If Not IsNumeric(cboCaixa.Text) Then Exit Sub
  Set rstCaixa = db.OpenRecordset("SELECT Caixa, Descrição FROM [Caixas em Uso] WHERE Caixa = " & CLng(cboCaixa.Text), dbOpenSnapshot)
  
  With rstCaixa
    If Not (.BOF And .EOF) Then
      txtCaixa.Text = .Fields("Descrição").Value & ""
    End If
    
    If Not rstCaixa Is Nothing Then .Close
    Set rstCaixa = Nothing
  End With
End Sub

Private Sub cboOpEntradaTransf_CloseUp()
  cboOpEntradaTransf.Text = cboOpEntradaTransf.Columns(0).Text & ""
  cboOpEntradaTransf_LostFocus
End Sub

Private Sub cboOpEntradaTransf_LostFocus()
  Dim rstOperacaoEntrada As Recordset
  
  txtOpEntradaTransf.Text = ""
  
  If Not IsNumeric(cboOpEntradaTransf.Text) Then Exit Sub
  Set rstOperacaoEntrada = db.OpenRecordset("SELECT Código, Nome FROM [Operações Entrada] WHERE Código = " & CLng(cboOpEntradaTransf.Text), dbOpenSnapshot)
  
  With rstOperacaoEntrada
    If Not (.BOF And .EOF) Then
      txtOpEntradaTransf.Text = .Fields("Nome").Value & ""
    End If
    
    If Not rstOperacaoEntrada Is Nothing Then .Close
    Set rstOperacaoEntrada = Nothing
  End With
End Sub

Private Sub cboOperacaoEntrada_CloseUp()
  cboOperacaoEntrada.Text = cboOperacaoEntrada.Columns(0).Text & ""
  cboOperacaoEntrada_LostFocus
End Sub

Private Sub cboOperacaoEntrada_LostFocus()
  Dim rstOperacaoEntrada As Recordset
  
  txtNomeOperacaoEntrada.Text = ""
  
  If Not IsNumeric(cboOperacaoEntrada.Text) Then Exit Sub
  Set rstOperacaoEntrada = db.OpenRecordset("SELECT Código, Nome FROM [Operações Entrada] WHERE Código = " & CLng(cboOperacaoEntrada.Text), dbOpenSnapshot)
  
  With rstOperacaoEntrada
    If Not (.BOF And .EOF) Then
      txtNomeOperacaoEntrada.Text = .Fields("Nome").Value & ""
    End If
    
    If Not rstOperacaoEntrada Is Nothing Then .Close
    Set rstOperacaoEntrada = Nothing
  End With
End Sub

Private Sub cboOperacaoFechamento_CloseUp()
  cboOperacaoFechamento.Text = cboOperacaoFechamento.Columns(0).Text & ""
  cboOperacaoFechamento_LostFocus
End Sub

Private Sub cboOperacaoFechamento_LostFocus()
  Dim rstOperacaoSaida As Recordset
  
  txtNomeOperacaoFechamento.Text = ""
  
  If Not IsNumeric(cboOperacaoFechamento.Text) Then Exit Sub
  Set rstOperacaoSaida = db.OpenRecordset("SELECT Código, Nome FROM [Operações Saída] WHERE Código = " & CLng(cboOperacaoFechamento.Text), dbOpenSnapshot)
  
  With rstOperacaoSaida
    If Not (.BOF And .EOF) Then
      txtNomeOperacaoFechamento.Text = .Fields("Nome").Value & ""
    End If
    
    If Not rstOperacaoSaida Is Nothing Then .Close
    Set rstOperacaoSaida = Nothing
  End With
End Sub

Private Sub cboOperacaoSaida_CloseUp()
  cboOperacaoSaida.Text = cboOperacaoSaida.Columns(0).Text & ""
  cboOperacaoSaida_LostFocus
End Sub

Private Sub cboOperacaoSaida_LostFocus()
  Dim rstOperacaoSaida As Recordset
  
  txtNomeOperacaoSaida.Text = ""
  
  If Not IsNumeric(cboOperacaoSaida.Text) Then Exit Sub
  Set rstOperacaoSaida = db.OpenRecordset("SELECT Código, Nome FROM [Operações Saída] WHERE Código = " & CLng(cboOperacaoSaida.Text), dbOpenSnapshot)
  
  With rstOperacaoSaida
    If Not (.BOF And .EOF) Then
      txtNomeOperacaoSaida.Text = .Fields("Nome").Value & ""
    End If
    
    If Not rstOperacaoSaida Is Nothing Then .Close
    Set rstOperacaoSaida = Nothing
  End With
End Sub

'--------------------------------[ Início ]-----------------------------------------
'13/11/2002 - mpdea
'Inclusão do controle de seleção da operação de saída (Orçamento -> Venda)
Private Sub cboOperSaida_S_Click()
  cboOperSaida_S.Text = cboOperSaida_S.Columns(1).Text
End Sub

Private Sub cboOperSaida_S_CloseUp()
  cboOperSaida_S.Text = cboOperSaida_S.Columns(1).Text
  cboOperSaida_S_LostFocus
End Sub

Private Sub cboOperSaida_S_DropDown()
  cboOperSaida_S.DataFieldList = "Código"
End Sub

Private Sub cboOperSaida_S_KeyPress(KeyAscii As Integer)
  If cboOperSaida_S.DroppedDown Then
    cboOperSaida_S.DataFieldList = "Nome"
  End If
End Sub

Private Sub cboOperSaida_S_LostFocus()
  Dim intCodOper As Integer
 
  lblNomeOperSaida_S.Caption = ""
 
  Call IsDataType(dtInteger, cboOperSaida_S.Text, intCodOper)
 
  If intCodOper > 0 Then
    With datOperSaida.Recordset
      .FindFirst "Código = " & intCodOper
      If Not .NoMatch Then
        lblNomeOperSaida_S.Caption = .Fields("Nome").Value & ""
      End If
    End With
  End If
  
  If intCodOper <> 0 And lblNomeOperSaida_S.Caption = "" Then
    Tab1.Tab = 2
    DisplayMsg "Operação de saída incorreta."
    SelectAllText cboOperSaida_S, True
  End If
 
End Sub
'----------------------------------[ Fim ]------------------------------------------


Private Sub cboOpSaidaTransf_CloseUp()
  cboOpSaidaTransf.Text = cboOpSaidaTransf.Columns(0).Text & ""
  cboOpSaidaTransf_LostFocus
End Sub

Private Sub cboOpSaidaTransf_LostFocus()
  Dim rstOperacaoSaida As Recordset
  
  txtOpSaidaTransf.Text = ""
  
  If Not IsNumeric(cboOpSaidaTransf.Text) Then Exit Sub
  Set rstOperacaoSaida = db.OpenRecordset("SELECT Código, Nome FROM [Operações Saída] WHERE Código = " & CLng(cboOpSaidaTransf.Text), dbOpenSnapshot)
  
  With rstOperacaoSaida
    If Not (.BOF And .EOF) Then
      txtOpSaidaTransf.Text = .Fields("Nome").Value & ""
    End If
    
    If Not rstOperacaoSaida Is Nothing Then .Close
    Set rstOperacaoSaida = Nothing
  End With
End Sub

Private Sub chk_cobrarMulta_Click()
  If chk_cobrarMulta.Value = vbChecked Then
      txt_multaDiasAposParcelaVencida.Enabled = True
      mskTaxaMultaParcelaVencida.Enabled = True
      lbl_multaDiasAposParcelaVencida.ForeColor = &H80000012
      lbl_TaxaMultaParcelaVencida.ForeColor = &H80000012
  Else
      mskTaxaMultaParcelaVencida.Text = ""
      txt_multaDiasAposParcelaVencida.Text = ""
      txt_multaDiasAposParcelaVencida.Enabled = False
      mskTaxaMultaParcelaVencida.Enabled = False
      lbl_multaDiasAposParcelaVencida.ForeColor = &H80000010
      lbl_TaxaMultaParcelaVencida.ForeColor = &H80000010
  End If
End Sub

'17/09/2009 - mpdea
Private Sub chkHabilitarNotaFiscalEletronica_Click()
  Dim blnEnabled As Boolean
  Dim intX As Integer
  
  blnEnabled = chkHabilitarNotaFiscalEletronica.Value = vbChecked
  
  'lblCodigoRegimeTributario.Enabled = blnEnabled
  'lblIdentificacaoAmbiente.Enabled = blnEnabled
  'lblLayoutEnvio.Enabled = blnEnabled
    
  For intX = 74 To 78
    lblTitle(intX).Enabled = blnEnabled
  Next
  
  cboAmbienteNfe.Enabled = blnEnabled
  cboFormatoImpressaoDanfeNfe.Enabled = blnEnabled
  cboModDetBaseCalculoIcms.Enabled = blnEnabled
  cboModDetBaseCalculoIcmsSt.Enabled = blnEnabled
  txtPastaEnvioNfe.Enabled = blnEnabled
  txtPastaRetornoNfe.Enabled = blnEnabled
  cmdSelecionarPastaNfe(0).Enabled = blnEnabled
  cmdSelecionarPastaNfe(1).Enabled = blnEnabled
  
  '24/11/2010 - Andrea
  txtLayoutEnvio.Enabled = blnEnabled
  cboCodigoRegimeTributario.Enabled = blnEnabled
  cboPadraoArquivoIntegracao.Enabled = blnEnabled
    
End Sub

Private Sub cmdPlanodeContas_Click()
  '12/05/2005 - Daniel
  'Carregar a tela de Plano de Contas
  frmPlanodeContas.Show
End Sub

'Private Sub cmdProcurar_Click()
'  Dim sFileName As String
'
'  sFileName = gsOpenFile(Me, "Escolha o Arquivo de Configuração", "Arquivo de Configuração |*.cbb")
'
'  If Len(sFileName) = 0 Then
'    Exit Sub
'  End If
'
'  txtBoleto.Text = gsGetOnlyName(sFileName)
'
'End Sub

Private Sub cmdProcurarArquivoNf_Click(Index As Integer)
  Dim sFileName As String
  
  sFileName = gsOpenFile(Me, "Escolha o Arquivo de Configuração", "Arquivo de Configuração |*.*")
  
  If Len(sFileName) = 0 Then
    Exit Sub
  End If
  
  If Index = 0 Then
    txtConfigNFOut.Text = gsGetOnlyName(sFileName)
  Else
    txtConfigNFInp.Text = gsGetOnlyName(sFileName)
  End If

End Sub

Private Sub cmdProcurarTicket_Click()
  Dim sFileName As String
  
  sFileName = gsOpenFile(Me, "Escolha o Arquivo de Configuração", "Arquivo de Configuração |*.CTI")
  
  If Len(sFileName) = 0 Then
    Exit Sub
  End If
  
  txtTicket.Text = gsGetOnlyName(sFileName)

End Sub

Private Function gsGetOnlyName(ByVal sFileName As String) As String
  Dim nI As Integer
  Dim sCh As String
  gsGetOnlyName = ""
  For nI = Len(sFileName) To 1 Step -1
    sCh = Mid(sFileName, nI, 1)
    If sCh = "." Then
      gsGetOnlyName = ""
    Else
      If sCh = "\" Then
        Exit For
      End If
      gsGetOnlyName = sCh & gsGetOnlyName
    End If
  Next nI
End Function

Private Sub cmdSelecionarPastaNfe_Click(Index As Integer)
  Dim sRet As String
  
  sRet = sFindDir("Selecione o diretório.", Me.hwnd)
  
  If Len(sRet) = 0 Then
    Exit Sub
  End If
  
  If Index = 0 Then
    txtPastaEnvioNfe.Text = sRet
  Else
    txtPastaRetornoNfe.Text = sRet
  End If

End Sub

Private Sub Código_CloseUp()
  Código.Text = Código.Columns(0).Text
  Código_LostFocus
End Sub

Private Sub Código_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub Código_LostFocus()
  If IsNull(Código.Text) Then Exit Sub
  If Not IsNumeric(Código.Text) Then Exit Sub
  If Val(Código.Text) < 1 Then Exit Sub
  If Val(Código.Text) > 999 Then Exit Sub
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Código.Text
  If Not rsParametros.NoMatch Then
    Num_Registro = rsParametros.Bookmark
    ShowRecord
  End If
  If rsParametros.NoMatch Then
    If Not IsNull(Num_Registro) Then rsParametros.Bookmark = Num_Registro
  End If
End Sub

Private Sub Combo_Conta_CloseUp()
 Combo_Conta.Text = Combo_Conta.Columns(2).Text
 Combo_Conta_LostFocus
End Sub

Private Sub Combo_Conta_LostFocus()


 Nome_Conta.Caption = ""
 If IsNull(Combo_Conta.Text) Then Exit Sub
 If Combo_Conta.Text = "" Then Exit Sub
 If Not IsNumeric(Combo_Conta.Text) Then Exit Sub
 If Val(Combo_Conta.Text) < 1 Then Exit Sub
 
 rsContas.Index = "Código"
 rsContas.Seek "=", Val(Combo_Conta.Text)
 If rsContas.NoMatch Then Exit Sub
 
 Nome_Conta.Caption = rsContas("Descrição") & ""

End Sub


Private Sub Form_Unload(Cancel As Integer)

  rsParametros.Close
  rsPreços.Close
  rsCliFor.Close
  rsOp_Saída.Close
  rsContas.Close
  rsZZZLog.Close
  
  Set rsParametros = Nothing
  Set rsPreços = Nothing
  Set rsCliFor = Nothing
  Set rsOp_Saída = Nothing
  Set rsContas = Nothing
  Set rsZZZLog = Nothing
End Sub

Private Sub lblHelpTrafficLight_Click()
  htmlhelp Me.hwnd, gsDefaultPath & "Ajuda\Traffic Light.chm", 0, 0
End Sub



Private Sub Máximo_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

'13/08/2004 - mpdea
'Modificado o limite para 255
Private Sub Máximo_Serviço_LostFocus()
  If IsNull(Máximo_Serviço.Text) Then Máximo_Serviço.Text = 5
  If Not IsNumeric(Máximo_Serviço.Text) Then Máximo_Serviço.Text = 5
  If Máximo_Serviço.Text > 255 Then Máximo_Serviço.Text = 255
End Sub

Private Sub mskTaxaMultaParcelaVencida_LostFocus()
    If Len(Trim(mskTaxaMultaParcelaVencida.Text)) > 0 Then
        If Not IsNumeric(mskTaxaMultaParcelaVencida.Text) Then
            MsgBox "Informe corretamente a taxa da multa que será cobrada para parcelas vencidas", vbInformation, "Atenção"
            mskTaxaMultaParcelaVencida.SetFocus
        End If
    End If
End Sub

Private Sub optLocalizarCodigo_Click()
  cboOrdenacao.Enabled = False
  lblOrdenacao.Enabled = False
End Sub

Private Sub optLocalizarCodigoNome_Click()
  cboOrdenacao.Enabled = True
  lblOrdenacao.Enabled = True
End Sub

Private Sub Saída_Intervalo_Parc_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub txt_multaDiasAposParcelaVencida_LostFocus()
    If Len(Trim(txt_multaDiasAposParcelaVencida.Text)) > 0 Then
        If Not IsNumeric(txt_multaDiasAposParcelaVencida.Text) Then
            MsgBox "Informe corretamente o número de dias", vbInformation, "Atenção"
            txt_multaDiasAposParcelaVencida.SetFocus
        End If
    End If
End Sub

Private Sub txtConfigNFInp_LostFocus()
  txtConfigNFInp.Text = UCase(txtConfigNFInp.Text & "")
End Sub

Private Sub txtConfigNFOut_LostFocus()
  txtConfigNFOut.Text = UCase(txtConfigNFOut.Text & "")
End Sub

Private Sub Estado_LostFocus()
  Estado.Text = UCase(Trim(Estado.Text & ""))
  
  '17/09/2003 - mpdea
  'Incluído chkAlteraPreco
  '
  '12/09/2003 - mpdea
  'Validação para o estado de SC
'  If UCase(Estado.Text) = "SC" Then
'    VR_Altera_Preço.Value = vbUnchecked
'    VR_Altera_Preço.Visible = False
'    chkAlteraPreco.Value = vbUnchecked
'    chkAlteraPreco.Visible = False
'  Else
'    VR_Altera_Preço.Visible = True
'    chkAlteraPreco.Visible = True
'  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
      'Call objHelp.Show(strfile, "QuickStore10Help")
      Call objHelp.Show(strfile, "QuickStore10Help", 10002)
      Set objHelp = Nothing
  Else
      Call HandleKeyDown(KeyCode, Shift)
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  Dim Aux As String
  Dim strSQL As String
  
  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  KeyPreview = True
  
  gsSenhaGerenteAtual = ""
  
  ActiveBar1.Tools("miOpSearch").Enabled = False
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial")
  Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsOp_Saída = db.OpenRecordset("Operações Saída", , dbReadOnly)
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsZZZLog = db.OpenRecordset("ZZZLog", , dbSeeChanges)

  Data5.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  '13/11/2002 - mpdea
  'Adicionado controle Data para seleção da op. saída (Tab. Saídas)
  datOperSaida.DatabaseName = gsQuickDBFileName
  strSQL = "SELECT DISTINCTROW Nome, Código, Tipo FROM [Operações Saída] " & _
           "WHERE Tipo = 'V' ORDER BY Nome"
  datOperSaida.RecordSource = strSQL
  
  datOperacaoEntrada.DatabaseName = gsQuickDBFileName
  datOperacaoSaida.DatabaseName = gsQuickDBFileName
  datCaixa.DatabaseName = gsQuickDBFileName
  

  rsPreços.Index = "Só Tabela"
  Aux = ""
  Do
   rsPreços.Seek ">", Aux
   If Not rsPreços.NoMatch Then
     Aux = rsPreços("Tabela")
     'Combo_Preço.AddItem rsPreços("Tabela")
     If UCase(Aux) <> "CUSTO" Then
       VR_Combo_Preço.AddItem Aux
     End If
     Com_Tab1.AddItem Aux
     Com_Tab2.AddItem Aux
     Com_Tab3.AddItem Aux
     Com_Tab4.AddItem Aux
     Com_Tab5.AddItem Aux
     Com_Tab6.AddItem Aux
   End If
  Loop Until rsPreços.NoMatch
  
  Tab1.TabEnabled(1) = False
  Tab1.TabEnabled(2) = False
  Tab1.TabEnabled(3) = False
  Tab1.TabEnabled(4) = False
  
  cboOrdenacao.Enabled = False
  lblOrdenacao.Enabled = False
  
  '17/08/2004 - Daniel
  'Tratamento para Impressão de Boletos automáticamente
  'na Manutenção do Contas a Receber
  'Case: De Mais Presentes (Loja do Nazareno) QS31735-849
  'Aberto também para o cliente F. Linhares QS37818-990
'  If CheckSerialCaseMod("QS31735-849", "QS37818-990") Then
'    lblBoleto.Visible = True
'    txtBoleto.Visible = True
'    cmdProcurar.Visible = True
'  Else
'    lblBoleto.Visible = False
'    txtBoleto.Visible = False
'    cmdProcurar.Visible = False
'  End If
  
  '18/08/2004 - Daniel
  'Frame com as configurações do ticket
  'Case: STC
  fraTicket.Visible = CheckSerialCaseMod("QS39823-684")
  
  '--------------------------------------------------
  ' SHOW FORM
  '--------------------------------------------------
  Me.Show
  DoEvents
  
  Call ActiveBarLoadToolTips(Me)
  
  Call ClearScreen
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Juros_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

'13/08/2004 - mpdea
'Modificado o limite para 255
Private Sub Máximo_LostFocus()
 If IsNull(Máximo.Text) Then Máximo.Text = 5
 If Not IsNumeric(Máximo.Text) Then Máximo.Text = 5
 If Máximo.Text > 255 Then Máximo.Text = 255
End Sub

Private Sub Máximo_Serviço_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub O_Preços_Click()
  If O_Preços.Value = 1 Then
   Label_Tabela1.Enabled = True
   Label_Tabela2.Enabled = True
   Label_Tabela3.Enabled = True
   
   Tabela1.Enabled = True
   Tabela2.Enabled = True
   Tabela3.Enabled = True
  End If
  
  If O_Preços.Value = 0 Then
   Label_Tabela1.Enabled = False
   Label_Tabela2.Enabled = False
   Label_Tabela3.Enabled = False
   
   Tabela1.Enabled = False
   Tabela2.Enabled = False
   Tabela3.Enabled = False
  End If
  
End Sub

Private Sub Qtde_Balança_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Tabela1_LostFocus()
  If IsNull(Tabela1.Text) Then Exit Sub
  If Tabela1.Text = "" Then Exit Sub
  Tabela1.Text = UCase(Tabela1.Text)
End Sub

Private Sub Tabela2_LostFocus()
  If IsNull(Tabela2.Text) Then Exit Sub
  If Tabela2.Text = "" Then Exit Sub
  Tabela2.Text = UCase(Tabela2.Text)
End Sub

Private Sub Tabela3_LostFocus()
  If IsNull(Tabela3.Text) Then Exit Sub
  If Tabela3.Text = "" Then Exit Sub
  Tabela3.Text = UCase(Tabela3.Text)
End Sub

Private Sub Usar_Serviços_Click()
  Alterar_Serviços.Enabled = Usar_Serviços.Value = vbChecked
End Sub

Private Sub VR_Cód_Operação_KeyPress(KeyAscii As Integer)
 KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub VR_Cód_Operação_LostFocus()
  Nome_Operação.Caption = ""

  If IsNull(VR_Cód_Operação.Text) Then Exit Sub
  If VR_Cód_Operação.Text = "" Then Exit Sub
  If Not IsNumeric(VR_Cód_Operação.Text) Then Exit Sub
  If Val(VR_Cód_Operação.Text) < 1 Then Exit Sub
  If Val(VR_Cód_Operação.Text) > 9999 Then Exit Sub
  
  rsOp_Saída.Index = "Código"
  rsOp_Saída.Seek "=", Val(VR_Cód_Operação.Text)
  If rsOp_Saída.NoMatch Then Exit Sub
  
  Nome_Operação.Caption = rsOp_Saída("Nome") & ""
End Sub

Private Sub VR_Combo_Cliente_CloseUp()
 VR_Combo_Cliente.Text = VR_Combo_Cliente.Columns(1).Text
 VR_Combo_Cliente_LostFocus
End Sub

Private Sub VR_Combo_Cliente_LostFocus()
  VR_Nome_Cliente.Caption = ""
  If IsNull(VR_Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(VR_Combo_Cliente.Text) Then Exit Sub
  If Val(VR_Combo_Cliente.Text) <= 0 Then Exit Sub
  If Val(VR_Combo_Cliente.Text) > 99999999 Then Exit Sub
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Val(VR_Combo_Cliente.Text)
  If Not rsCliFor.NoMatch Then
    VR_Nome_Cliente.Caption = rsCliFor("Nome")
  End If
End Sub

Private Sub VR_Desconto_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub VR_Intervalo_Parc_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub VR_Linhas_KeyPress(KeyAscii As Integer)
 KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub VR_Prazo_Cheques_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub VR_Prazo_Parcela_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub VR_Qtde_Cheques_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub VR_Qtde_Parcela_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

'30/01/2009 - mpdea
'Configuração de envio de email
Private Sub ConfigurarEnvioEmail()
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Encontre ou grave um registro antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Dim frm_x As New frmEmailConfigurar
  frm_x.CodigoFilial = rsParametros.Fields("Filial").Value
  frm_x.Show vbModal
  Set frm_x = Nothing
  
End Sub


