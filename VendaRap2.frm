VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmVendaRap2 
   Appearance      =   0  'Flat
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Venda Rápida"
   ClientHeight    =   9405
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   15420
   FillColor       =   &H00666666&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E5E5E5&
   HelpContextID   =   1230
   Icon            =   "VendaRap2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9405
   ScaleWidth      =   15420
   Begin VB.Data dataPrestadores 
      Caption         =   "dataPrestadores"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7140
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CommandButton btnComandaVendas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      Height          =   295
      Left            =   14850
      Picture         =   "VendaRap2.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   7260
      Width           =   375
   End
   Begin VB.TextBox Observacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   435
      Left            =   1200
      TabIndex        =   128
      Top             =   1380
      Width           =   14025
   End
   Begin VB.CommandButton cmd_opcoes 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   10740
      Picture         =   "VendaRap2.frx":4EE2F
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   7380
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmd_carneComRecibo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Carnê com recibo"
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
      Left            =   14850
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   8700
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmd_carne 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Carnê"
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
      Left            =   14790
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   8430
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmd_fecharTela 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14790
      Picture         =   "VendaRap2.frx":53F4D
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   8280
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmd_pesquisaAlfa 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consulta Produtos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   14910
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   8940
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame frm_produtoSemPrecoNaGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   4080
      TabIndex        =   112
      Top             =   4860
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CommandButton cmd_fecharFrameProdutoSemPrecoNaGrade 
         BackColor       =   &H008080FF&
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   930
         Width           =   1545
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Produto sem preço na grade. Se está correto, ignore este aviso."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   270
         TabIndex        =   114
         Top             =   90
         Width           =   4245
      End
   End
   Begin VB.CommandButton cmd_acharVenda 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Achar Vendas"
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
      Left            =   11190
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   923
      Width           =   1275
   End
   Begin VB.CommandButton cmd_tabelaDePrecos 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tabela de Preços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12900
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5130
      Width           =   1605
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   150
      Top             =   9060
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Referência 
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
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   9630
      MaxLength       =   20
      TabIndex        =   4
      Top             =   923
      Width           =   1515
   End
   Begin VB.Frame Frame_Recebimento 
      BackColor       =   &H00E5E5E5&
      Caption         =   "Recebimento Simplificado (F6)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   90
      TabIndex        =   69
      Top             =   4950
      Width           =   2955
      Begin MSMask.MaskEdBox Bom_Para 
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   23
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Frame fraButtonRecebeSimples 
         BackColor       =   &H00E5E5E5&
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
         Height          =   495
         Left            =   2160
         TabIndex        =   100
         Top             =   1560
         Width           =   1815
         Begin VB.CommandButton B_Recebe_Simples 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "&Confimar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   45
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.TextBox L_Receber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1710
         Width           =   2025
      End
      Begin VB.CheckBox Lançar_Débito 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "Lançar débito na conta do cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         TabIndex        =   15
         Top             =   180
         Width           =   3885
      End
      Begin MSMask.MaskEdBox Val_Cartão 
         Height          =   285
         Left            =   2790
         TabIndex        =   20
         Top             =   1215
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Dinheiro 
         Height          =   285
         Left            =   705
         TabIndex        =   16
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
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
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Parc 
         Height          =   250
         Index           =   4
         Left            =   9000
         TabIndex        =   50
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Frame Tipo_Parc 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   1575
         Left            =   9960
         TabIndex        =   82
         Top             =   390
         Width           =   1095
         Begin VB.OptionButton O_Carnet 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Carnet"
            Height          =   225
            Left            =   120
            TabIndex        =   53
            Top             =   1140
            Width           =   855
         End
         Begin VB.OptionButton O_Carteira 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Carteira"
            Height          =   225
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Width           =   950
         End
         Begin VB.OptionButton O_Banco 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Banco"
            Height          =   225
            Left            =   120
            TabIndex        =   51
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin MSMask.MaskEdBox Val_Parc 
         Height          =   255
         Index           =   3
         Left            =   9000
         TabIndex        =   48
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Parc 
         Height          =   250
         Index           =   2
         Left            =   9000
         TabIndex        =   46
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Parc 
         Height          =   250
         Index           =   1
         Left            =   9000
         TabIndex        =   44
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Parc 
         Height          =   255
         Index           =   0
         Left            =   9000
         TabIndex        =   42
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Cheque 
         Height          =   250
         Index           =   4
         Left            =   6840
         TabIndex        =   40
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Cheque 
         Height          =   250
         Index           =   3
         Left            =   6840
         TabIndex        =   36
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Cheque 
         Height          =   255
         Index           =   2
         Left            =   6840
         TabIndex        =   32
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Cheque 
         Height          =   255
         Index           =   1
         Left            =   6840
         TabIndex        =   28
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Val_Cheque 
         Height          =   255
         Index           =   0
         Left            =   6840
         TabIndex        =   24
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   14
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Cheque 
         Appearance      =   0  'Flat
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
         Height          =   250
         Index           =   4
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Cheque 
         Appearance      =   0  'Flat
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
         Height          =   250
         Index           =   3
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   34
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Cheque 
         Appearance      =   0  'Flat
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
         Height          =   250
         Index           =   2
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Cheque 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   26
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Cheque 
         Appearance      =   0  'Flat
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
         Height          =   250
         Index           =   0
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Cartão 
         Bindings        =   "VendaRap2.frx":5DA3F
         DataSource      =   "Data4"
         Height          =   285
         Left            =   705
         TabIndex        =   18
         Top             =   840
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
         BorderStyle     =   0
         BevelColorFrame =   15066597
         BevelColorHighlight=   16250871
         BevelColorShadow=   0
         CheckBox3D      =   0   'False
         BackColorEven   =   15066597
         BackColorOdd    =   12640511
         RowHeight       =   529
         Columns.Count   =   2
         Columns(0).Width=   9128
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
         _ExtentX        =   2461
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483630
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
      End
      Begin VB.TextBox Num_Cartão 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   705
         TabIndex        =   19
         Top             =   1170
         Width           =   1395
      End
      Begin MSMask.MaskEdBox Vale 
         Height          =   285
         Left            =   2610
         TabIndex        =   17
         Top             =   495
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
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
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Banco 
         Appearance      =   0  'Flat
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
         Height          =   250
         Index           =   4
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   37
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Banco 
         Appearance      =   0  'Flat
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
         Height          =   250
         Index           =   3
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   33
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Banco 
         Appearance      =   0  'Flat
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
         Height          =   250
         Index           =   2
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Banco 
         Appearance      =   0  'Flat
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
         Height          =   250
         Index           =   1
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   25
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Banco 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin MSMask.MaskEdBox Bom_Para 
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   27
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   955
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
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
      Begin MSMask.MaskEdBox Bom_Para 
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   31
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Bom_Para 
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   35
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Bom_Para 
         Height          =   255
         Index           =   4
         Left            =   5880
         TabIndex        =   39
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Parc 
         Height          =   250
         Index           =   0
         Left            =   7930
         TabIndex        =   41
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   720
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Parc 
         Height          =   250
         Index           =   1
         Left            =   7930
         TabIndex        =   43
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Parc 
         Height          =   255
         Index           =   2
         Left            =   7930
         TabIndex        =   45
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Parc 
         Height          =   250
         Index           =   3
         Left            =   7930
         TabIndex        =   47
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Parc 
         Height          =   250
         Index           =   4
         Left            =   7930
         TabIndex        =   49
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblRecebidoComCartao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "Recebido com mais de um cartão. Utilize a tela de recebimentos para visualização completa."
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
         Height          =   1215
         Left            =   2205
         TabIndex        =   102
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Line Line4 
         X1              =   4080
         X2              =   11120
         Y1              =   2080
         Y2              =   2080
      End
      Begin VB.Line Line3 
         X1              =   11120
         X2              =   11120
         Y1              =   120
         Y2              =   2100
      End
      Begin VB.Line Line2 
         X1              =   4080
         X2              =   4080
         Y1              =   240
         Y2              =   2100
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Valor"
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
         Left            =   2160
         TabIndex        =   92
         Top             =   1215
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Número"
         Height          =   255
         Left            =   75
         TabIndex        =   91
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label Label_Receber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "A Receber"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   90
         Top             =   1470
         Width           =   1935
      End
      Begin VB.Label L_Cheque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bom Para"
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
         Index           =   3
         Left            =   5880
         TabIndex        =   83
         Top             =   480
         Width           =   975
      End
      Begin VB.Label L_Parc3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor"
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
         Left            =   9000
         TabIndex        =   81
         Top             =   480
         Width           =   855
      End
      Begin VB.Label L_Parc2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data"
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
         Left            =   7930
         TabIndex        =   80
         Top             =   480
         Width           =   1100
      End
      Begin VB.Line Line1 
         X1              =   7800
         X2              =   7800
         Y1              =   240
         Y2              =   2100
      End
      Begin VB.Label L_Cheque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor"
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
         Index           =   4
         Left            =   6840
         TabIndex        =   79
         Top             =   480
         Width           =   855
      End
      Begin VB.Label L_Cheque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Num Cheque"
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
         Index           =   2
         Left            =   4800
         TabIndex        =   78
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label L_Cheque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
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
         Index           =   1
         Left            =   4200
         TabIndex        =   77
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Nome_Cartão 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   76
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label L_Parc1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parcelamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   255
         Left            =   7800
         TabIndex        =   74
         Top             =   120
         Width           =   3330
      End
      Begin VB.Label L_Cheque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheques"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   73
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label L_Vale 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Vale"
         Height          =   255
         Left            =   2205
         TabIndex        =   72
         Top             =   495
         Width           =   495
      End
      Begin VB.Label L_Cartão 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Cartão"
         Height          =   255
         Left            =   75
         TabIndex        =   71
         Top             =   840
         Width           =   735
      End
      Begin VB.Label L_Dinheiro 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   75
         TabIndex        =   70
         Top             =   495
         Width           =   735
      End
   End
   Begin VB.CommandButton B_Ret_NFCe 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Retorno NFCe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10260
      MaskColor       =   &H00F7F7F7&
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   8505
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
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
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   90
      TabIndex        =   99
      Top             =   7560
      Width           =   15165
      Begin VB.CommandButton cmd_esquerda 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13740
         Picture         =   "VendaRap2.frx":5DA53
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   480
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmd_direita 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14970
         Picture         =   "VendaRap2.frx":67545
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   480
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmd_Acima 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   14460
         Picture         =   "VendaRap2.frx":71037
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   510
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton cmd_abaixo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   14010
         Picture         =   "VendaRap2.frx":7AB29
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   510
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton B_NFC_e 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Emitir NFC-e"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10170
         MaskColor       =   &H00F7F7F7&
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   90
         Width           =   4965
      End
      Begin VB.CommandButton B_Recebe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Receber "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   510
         Width           =   4965
      End
      Begin VB.CommandButton B_Grava_Recebe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Gra&var e Receber"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   90
         Width           =   4965
      End
      Begin VB.CommandButton B_Ticket 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Ticket "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   4965
      End
      Begin VB.CommandButton B_Desconto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Desconto"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -90
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   4965
      End
      Begin VB.CommandButton B_programaFidelidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Programa de Fidelidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -90
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   945
         Width           =   4965
      End
      Begin VB.CommandButton B_Limpa 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         Caption         =   "&Limpar "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   12720
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   945
         Width           =   2415
      End
      Begin VB.CommandButton B_Grava 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Gravar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -90
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   510
         Width           =   4965
      End
   End
   Begin VB.TextBox L_Tot_IPI 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   14040
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   5805
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox L_Tot_Prod 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   12870
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5805
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtComanda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   13410
      MaxLength       =   13
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   7260
      Width           =   1395
   End
   Begin VB.TextBox txtDescSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   13410
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6510
      Width           =   1815
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   13410
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6150
      Width           =   1815
   End
   Begin VB.TextBox L_Tot_Pagar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   13410
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6870
      Width           =   1815
   End
   Begin VB.Data datCaixa 
      Caption         =   "datCaixa"
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
      Left            =   9510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Caixa, Descrição FROM [Caixas em Uso] ORDER BY Caixa"
      Top             =   9030
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Data datSequencias 
      Caption         =   "datSequencias"
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
      Left            =   7830
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9030
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Left            =   4110
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"VendaRap2.frx":8461B
      Top             =   9090
      Visible         =   0   'False
      Width           =   1800
   End
   Begin SSDataWidgets_B.SSDBDropDown DropDown2 
      Bindings        =   "VendaRap2.frx":846B4
      Height          =   1035
      Left            =   1680
      TabIndex        =   93
      Top             =   2745
      Width           =   10770
      DataFieldList   =   "Código"
      ListAutoValidate=   0   'False
      ListWidthAutoSize=   0   'False
      MaxDropDownItems=   15
      ListWidth       =   16140
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DividerStyle    =   2
      BevelColorFrame =   15066597
      BevelColorHighlight=   16250871
      BevelColorShadow=   15066597
      BevelColorFace  =   15066597
      CheckBox3D      =   0   'False
      ForeColorEven   =   0
      BackColorEven   =   16250871
      BackColorOdd    =   12648447
      RowHeight       =   450
      ExtraHeight     =   159
      Columns.Count   =   6
      Columns(0).Width=   5001
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Código"
      Columns(0).FieldLen=   256
      Columns(1).Width=   7488
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Nome"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1958
      Columns(2).Caption=   "Estoque"
      Columns(2).Name =   "Estoque"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2011
      Columns(3).Caption=   "Preço"
      Columns(3).Name =   "Preço"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "CódigoOrdenação"
      Columns(4).Name =   "CódigoOrdenação"
      Columns(4).DataField=   "Código Ordenação"
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "Fabricante"
      Columns(5).Name =   "Fabricante"
      Columns(5).Alignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   18997
      _ExtentY        =   1826
      _StockProps     =   77
      ForeColor       =   -2147483630
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
   Begin VB.CommandButton cmdInsertItens 
      Caption         =   "255!"
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
      Left            =   10980
      TabIndex        =   54
      Top             =   9000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Left            =   510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Preços"
      Top             =   9030
      Visible         =   0   'False
      Width           =   1860
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
      Left            =   5940
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cartão"
      Top             =   9045
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
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
      Left            =   11655
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo AND isPrestServ = FALSE ORDER BY Nome"
      Top             =   9030
      Visible         =   0   'False
      Width           =   1800
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
      Left            =   13545
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"VendaRap2.frx":846C8
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin SSDataWidgets_B.SSDBCombo cboCaixa 
      Bindings        =   "VendaRap2.frx":8475D
      Height          =   405
      Left            =   9630
      TabIndex        =   2
      Top             =   518
      Width           =   1515
      DataFieldList   =   "Caixa"
      BevelType       =   0
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Font3D          =   2
      BevelColorFrame =   6710886
      BevelColorHighlight=   16250871
      ForeColorEven   =   6710886
      ForeColorOdd    =   6710886
      BackColorEven   =   12648447
      BackColorOdd    =   15066597
      RowHeight       =   503
      Columns(0).Width=   3200
      _ExtentX        =   2672
      _ExtentY        =   714
      _StockProps     =   93
      Text            =   "Caixa"
      ForeColor       =   0
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
      DataFieldToDisplay=   "Caixa"
   End
   Begin SSDataWidgets_B.SSDBDropDown DropDown1 
      Bindings        =   "VendaRap2.frx":84774
      Height          =   1035
      Left            =   510
      TabIndex        =   106
      Top             =   3780
      Width           =   10530
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      MaxDropDownItems=   15
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      DividerStyle    =   2
      BevelColorFrame =   16250871
      BevelColorShadow=   15066597
      BevelColorFace  =   15066597
      CheckBox3D      =   0   'False
      ForeColorEven   =   0
      BackColorEven   =   16250871
      BackColorOdd    =   12648447
      RowHeight       =   450
      ExtraHeight     =   238
      Columns.Count   =   5
      Columns(0).Width=   11086
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4260
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2196
      Columns(2).Caption=   "Estoque"
      Columns(2).Name =   "Estoque"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2646
      Columns(3).Caption=   "Preço"
      Columns(3).Name =   "Preço"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Fabricante"
      Columns(4).Name =   "Fabricante"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   18574
      _ExtentY        =   1826
      _StockProps     =   77
      ForeColor       =   -2147483630
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
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Height          =   3195
      Left            =   60
      TabIndex        =   6
      Top             =   1860
      Width           =   15255
      _Version        =   196617
      DataMode        =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   6.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      GroupHeaders    =   0   'False
      GroupHeadLines  =   0
      DividerStyle    =   2
      BevelColorFrame =   15066597
      BevelColorHighlight=   15066597
      BevelColorShadow=   -2147483633
      BevelColorFace  =   15066597
      CheckBox3D      =   0   'False
      MultiLine       =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorEven   =   16250871
      BackColorOdd    =   12648447
      RowHeight       =   556
      ExtraHeight     =   79
      Columns.Count   =   14
      Columns(0).Width=   4207
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   40
      Columns(1).Width=   2090
      Columns(1).Caption=   "Qtde"
      Columns(1).Name =   "Qtde"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      Columns(2).Width=   9340
      Columns(2).Caption=   "Nome"
      Columns(2).Name =   "Nome"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2778
      Columns(3).Caption=   "Preço"
      Columns(3).Name =   "Preço"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2117
      Columns(4).Caption=   "Desc.%"
      Columns(4).Name =   "Desc.%"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3281
      Columns(5).Caption=   "Total"
      Columns(5).Name =   "Total"
      Columns(5).Alignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "ICM"
      Columns(6).Name =   "ICM"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   5
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "IPI"
      Columns(7).Name =   "IPI"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   5
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "Base_ICM"
      Columns(8).Name =   "Base_ICM"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "Valor_ICM"
      Columns(9).Name =   "Valor_ICM"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "Valor_Base_Unit"
      Columns(10).Name=   "Valor_Base_Unit"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "Redução_ICM"
      Columns(11).Name=   "Redução_ICM"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "Tipo_ICM"
      Columns(12).Name=   "Tipo_ICM"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   2011
      Columns(13).Caption=   "CFOP"
      Columns(13).Name=   "CFOP"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      UseDefaults     =   0   'False
      _ExtentX        =   26908
      _ExtentY        =   5636
      _StockProps     =   79
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel L_Preço 
      Height          =   255
      Left            =   12210
      TabIndex        =   75
      Top             =   5565
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   15
      ForeColor       =   16711680
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Threed.SSPanel L_Estoque 
      Height          =   255
      Left            =   13815
      TabIndex        =   68
      Top             =   5565
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   15
      ForeColor       =   16711680
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "VendaRap2.frx":84788
      DataSource      =   "Data3"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Vendedor"
      Top             =   533
      Width           =   1515
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      BevelType       =   0
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BevelColorFrame =   6710886
      BevelColorHighlight=   16250871
      BevelColorShadow=   0
      CheckBox3D      =   0   'False
      ForeColorEven   =   0
      BackColorEven   =   12648447
      BackColorOdd    =   15066597
      RowHeight       =   635
      Columns.Count   =   3
      Columns(0).Width=   6588
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Apelido"
      Columns(1).Name =   "Apelido"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Apelido"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1640
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   256
      _ExtentX        =   2672
      _ExtentY        =   661
      _StockProps     =   93
      Text            =   "Vendedor"
      ForeColor       =   0
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
   End
   Begin SSDataWidgets_B.SSDBCombo Número 
      Bindings        =   "VendaRap2.frx":8479C
      Height          =   405
      Left            =   13455
      TabIndex        =   5
      Top             =   938
      Width           =   1770
      DataFieldList   =   "Sequência"
      BevelType       =   0
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BevelColorFrame =   6710886
      BevelColorHighlight=   16250871
      ForeColorEven   =   6710886
      ForeColorOdd    =   6710886
      BackColorEven   =   12648447
      BackColorOdd    =   15066597
      RowHeight       =   476
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Caption=   "Sequência"
      Columns(0).Name =   "Sequência"
      Columns(0).DataField=   "Sequência"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Cliente"
      Columns(1).Name =   "Cliente"
      Columns(1).DataField=   "Cliente"
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Ref. Interna"
      Columns(2).Name =   "RefInterna"
      Columns(2).DataField=   "Referência"
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Valor Total"
      Columns(3).Name =   "valor"
      Columns(3).DataField=   "Total"
      Columns(3).FieldLen=   256
      _ExtentX        =   3122
      _ExtentY        =   714
      _StockProps     =   93
      ForeColor       =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Sequência"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Preço 
      Bindings        =   "VendaRap2.frx":847B8
      Height          =   405
      Left            =   1200
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   98
      Width           =   6840
      DataFieldList   =   "Tabela"
      BevelType       =   0
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BevelColorFrame =   6710886
      BevelColorHighlight=   16250871
      BevelColorShadow=   15066597
      ForeColorEven   =   6710886
      ForeColorOdd    =   6710886
      BackColorEven   =   12648447
      BackColorOdd    =   15066597
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   12065
      _ExtentY        =   714
      _StockProps     =   93
      ForeColor       =   0
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
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "VendaRap2.frx":847CC
      DataSource      =   "Data1"
      Height          =   405
      Left            =   1200
      TabIndex        =   3
      Top             =   938
      Width           =   1515
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      BevelType       =   0
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BevelColorFrame =   6710886
      BevelColorHighlight=   16250871
      ForeColorEven   =   6710886
      ForeColorOdd    =   6710886
      BackColorEven   =   12648447
      BackColorOdd    =   15066597
      RowHeight       =   476
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1561
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2672
      _ExtentY        =   714
      _StockProps     =   93
      ForeColor       =   0
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
   End
   Begin SSDataWidgets_B.SSDBCombo cboPrestador 
      Bindings        =   "VendaRap2.frx":847E0
      DataSource      =   "dataPrestadores"
      Height          =   405
      Left            =   2040
      TabIndex        =   131
      Top             =   7170
      Width           =   975
      DataFieldList   =   "cod"
      ListAutoValidate=   0   'False
      BevelType       =   0
      _Version        =   196617
      Cols            =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BevelColorFrame =   6710886
      BevelColorHighlight=   16250871
      ForeColorEven   =   6710886
      ForeColorOdd    =   6710886
      BackColorEven   =   12648447
      BackColorOdd    =   15066597
      RowHeight       =   582
      Columns(0).Width=   3200
      _ExtentX        =   1720
      _ExtentY        =   714
      _StockProps     =   93
      ForeColor       =   0
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
   End
   Begin VB.Label Apelido_Prestador 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   405
      Left            =   3060
      TabIndex        =   132
      Top             =   7170
      Width           =   1905
   End
   Begin VB.Label lblPrestador 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      Caption         =   "Prestador de serviços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   120
      TabIndex        =   130
      Top             =   7230
      Width           =   1845
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Observação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   45
      TabIndex        =   66
      Top             =   1477
      Width           =   1005
   End
   Begin VB.Label lblQtdeTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Height          =   255
      Left            =   2415
      TabIndex        =   117
      Top             =   5460
      Width           =   1245
   End
   Begin VB.Label lblTitleQtdeTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quantidade de produtos"
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
      Height          =   255
      Left            =   60
      TabIndex        =   116
      Top             =   5460
      Width           =   2325
   End
   Begin VB.Label lbl_retornoEnvioNFCe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Autorizado. Imprimindo NFC-e"
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
      Left            =   7515
      TabIndex        =   115
      Top             =   6120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Movimentação_Desfeita 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Desfeita"
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
      Left            =   10575
      TabIndex        =   110
      Top             =   6660
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   405
      Left            =   2760
      TabIndex        =   109
      Top             =   938
      Width           =   5280
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   45
      TabIndex        =   108
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Nome_Vendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   405
      Left            =   2760
      TabIndex        =   107
      ToolTipText     =   "Nome do vendedor"
      Top             =   518
      Width           =   5280
   End
   Begin VB.Label lblDescSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      Caption         =   "Desconto"
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
      Height          =   210
      Left            =   12555
      TabIndex        =   95
      Top             =   6555
      Width           =   780
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      Caption         =   "SubTotal"
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
      Height          =   210
      Left            =   12600
      TabIndex        =   94
      Top             =   6195
      Width           =   735
   End
   Begin VB.Label Efetivada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Efetivada"
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
      Left            =   10575
      TabIndex        =   97
      Top             =   5580
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total IPI"
      ForeColor       =   &H00666666&
      Height          =   195
      Left            =   14040
      TabIndex        =   62
      Top             =   5565
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Produtos"
      ForeColor       =   &H00666666&
      Height          =   195
      Left            =   14040
      TabIndex        =   60
      Top             =   5580
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   12690
      TabIndex        =   98
      Top             =   5490
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblComanda 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comanda"
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
      Height          =   240
      Left            =   12480
      TabIndex        =   104
      Top             =   7290
      Width           =   810
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   12720
      TabIndex        =   64
      Top             =   6870
      Width           =   615
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   -30
      Top             =   9060
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
      Bands           =   "VendaRap2.frx":847FE
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Ref. Interna"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   8580
      TabIndex        =   89
      Top             =   1020
      Width           =   1020
   End
   Begin VB.Label Cod_Caixa 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   11025
      TabIndex        =   88
      Top             =   9090
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Cod_Operador 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Operador"
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
      Height          =   405
      Left            =   9630
      TabIndex        =   87
      Top             =   98
      Width           =   1515
   End
   Begin VB.Label Nome_Caixa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   405
      Left            =   11175
      TabIndex        =   86
      Top             =   518
      Width           =   4050
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   8580
      TabIndex        =   85
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Operador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   8580
      TabIndex        =   84
      Top             =   180
      Width           =   810
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Seqüência"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   12510
      TabIndex        =   67
      Top             =   1020
      Width           =   885
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   45
      TabIndex        =   65
      Top             =   1020
      Width           =   585
   End
   Begin VB.Label Nome_Operador 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   405
      Left            =   11175
      TabIndex        =   59
      Top             =   98
      Width           =   4050
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Tab. Preços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   45
      TabIndex        =   58
      Top             =   180
      Width           =   1020
   End
End
Attribute VB_Name = "frmVendaRap2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim totalNCM_2 As Double    'Total em R$ de imposto pago na movimentação

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

'------------------------------------------------------------------------'
'                   !!!!!!! ATENÇÃO !!!!!!!
'
' Replicar código para a tela de Venda Rápida CheckOut
' (frmVendaRap2_CheckOut)
'
' Compatibilizar código para que funcione em ambas as telas
' da mesma forma, através da comparação da variável g_frmVendaRapida
'
' Data da última sincronização: 30/01/2009
'------------------------------------------------------------------------'

'19/10/2007 - Anderson
'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
Dim Tamanho As Integer
Dim Cor As Integer
Dim Edicao As Long
Dim Tipo As Integer
Dim Erro As Integer
Private m_bolLucroMinimoPermitido As Boolean
 
Private Type Tabela
  Código As String
  Nome As String
  '29/10/2004 - Daniel
  'Alterado tipo de dado para Preço_Final
  'Preço_Final As Single (old)
  'Preço_Final As Currency
  Preço_Final As String
  Qtde As Single
  Preço As Single
  Desconto As Single
  ICM As Single
  IPI As Single
  Base_ICM As Double
  Valor_ICM As Double
  Valor_Base_Unit As Double
  Redução_ICM As Double
  Tipo_ICM As String
  '05/09/2008 - mpdea
  'Descrição adicional (Mantém dados salvos em Saídas)
  DescricaoAdicional As String
  
  '05/04/2010 - Andrea
  'Registro de CFOP por Produto
  CFOP_Produto As String

End Type
Private Tabe(255) As Tabela

Private Num_Registro As Variant

Private gbBaseICMSomadoIPI As Boolean

Private Erro_Data As Integer
Private Erro_Data2 As Boolean

'07/10/2002 - mpdea
'Modificado verificação de estoque para acesso direto ao recordset
'Private Verifica_Estoque As Integer

Private rsParametros As Recordset
Private rsPreços As Recordset
Private rsProdutos As Recordset
Private rsProdutos2 As Recordset
Private rsOp_Saída As Recordset
Private rsFuncionarios As Recordset
Private rsUsuarios As Recordset
Private rsCliFor As Recordset
Private rsGrade As Recordset
Private rsSaidas As Recordset
Private rsSaída_Prod As Recordset
Private rsSaída_Cheques As Recordset
Private rsSaída_Parcelas As Recordset
Private rsTabelas As Recordset
Private rsCotacoes As Recordset
Private rsEstoque As Recordset
Private rsContas_Receber As Recordset
Private rsEstados As Recordset
Private rsCartoes As Recordset
Private rsLog As Recordset
Private Linhas_Grade As Integer
Private Total_Produtos As Double

'11/12/2009 - Andrea
Private rsSaída_Cartoes As Recordset

'05/04/2010 - Andrea
'Alteração realizada para o registro do CFOP por produto e serviço
Dim rsProdutoCFOP As Recordset

Private Total_Desconto As Double
Private gcDescInTotal As Currency

Private Total_IPI As Double
Private Total_Pagar As Double
Private Total_Recebido As Double
Private Total_Base_ICM As Double
Private Total_ICM As Double
Private Cód_Vendedor As Integer
Private Desconto_Cli As Double
Private Calcula_ICM As Integer
Private Calcula_IPI As Integer
Private Operador_Caixa As Integer
Private Estado As String
Private Primeira_Vez As Integer

'23/09/2002 - mpdea
'Desconto no SubTotal
Private mcurDescontoSubTotal As Currency
'Flag para forçar a atualização do registro
Private mblnForceUpdate As Boolean

'07/05/2003 - mpdea
'Desconto rateado
Private m_blnDescontoRateado As Boolean

'29/05/2003 - mpdea
'Utilização do Traffic Light
Private m_blnWorkTrafficLight As Boolean

'05/08/2002 - mpdea
'Objeto TrafficLight para tratamento de atualizações na base de dados
Private WithEvents TrafficLight As IQMod.TrafficLight
Attribute TrafficLight.VB_VarHelpID = -1

'07/01/2004 - Daniel
'Var que verificará se ocorreu troco
Private m_blnOcorreTroco As Boolean

'08/01/2004 - Daniel
'Armazenar a quantidade para posterior impressões
Public m_sngQtdeTotal As Single

'25/03/2004 - Daniel
'Implementação feita para evitar gravação de saída
'adulterada por usuário sem permissão
'Case: Casagrande
Public m_blnUserDanger As Boolean

'05/05/2004 - Daniel
'Flag de indicação que é o Cliente Embalavi
'realizará ações personalizadas para este Cliente
Private m_blnEmbalavi As Boolean

'05/05/2004 - Daniel
'Flag de indicação que é o Cliente Bic
'realizará ações personalizadas para este Cliente
Private m_blnBic      As Boolean

'21/05/2004 - Daniel
'Vars Public para implementação da Bic Amazônia
'Estas duas vars serão responsáveis pelo valor dos campos de saídas [Codigo Func Comprador] e [Status Venda Func]
Public g_intCodigoFuncComprador As Integer
Public g_blnStatusVendaFunc     As Boolean
'Var utilizada em caso do usuário precisar paralisar o
'processor e depois prosseguir com a ação de gravação
Public g_blnRetornar            As Boolean

'01/07/2004 - Daniel
'Var de controle da implementação da CONEG CAMPOS
Private m_blnClear As Boolean

'26/08/2004 - Daniel
'Criado validação para verificar se o usuário possui permissão
'para enchergar o estoque ou não a partir do click em Consultar Produto
'Case: Tendresse
Private m_blnPermitido As Boolean
'Private m_blnTendresse As Boolean '30/01/2007 - Anderson - Alterado para que a permissão de visualizar estoque funcione para diversos clientes

'28/10/2004 - Daniel
'Tratamento para o field [Saídas - Produtos].[Preço Final]
'Para o cliente A.S. Wijma (Belém - Pará) deverá ser Double
'para os demais clientes continua sendo Single
Dim m_dblPrecoFinalAuxi As Double
Dim m_blnASWijmaBelem   As Boolean

'09/11/2004 - Daniel
'Esta var é utilizada para identificar o cliente
'Teknika que possui tratamento especial para o ICMS
Dim m_blnTeknika As Boolean

'01/12/2004 - Daniel
'Esta var é utilizada para identificar o cliente
'De Mais (Nazareno) com a finalidade de mostrar
'após o recebimento os cheques e parcelas
Dim m_blnDeMais         As Boolean
'06/05/2005 - Daniel
'
'Implementação.: Trabalhar com o código para fornecedor cadastrado na tela de produtos.
'                Impacto: Ao entrar com o código para o fornecedor no campo código do produto
'                o sistema deverá trazer o código do produto que estiver amarrado nele
'Solicitação...: Cristiano Pavinato - PSI RS
Private m_blnUsaCodFornec As Boolean
'12/05/2005 - Daniel
'
'Solicitante..: Info Social
'
'Finalidade...: Deixamos configurável em Parâmetros à exibição
'               nas telas de Saída e Venda Rápida da coluna Fabricante
'               nos dropdowns de pesquisas
Private m_blnExibirColunaFabricante As Boolean

'23/05/2006 - mpdea
'Otimizado a verificação do cliente isento em IPI
Private m_blnIsentoIPI As Boolean

'14/12/2007 - Celso
'Utilizada para armazenar o cliente para o qual já tenha sido solicitada senha
'do gerente qdo o mesmo tiver contas em atraso
Private m_blnSenhaGerJaInformada As Boolean
Private m_strCodigoClienteContas As String

'11/11/2008 - mpdea
'Verifica se deve somar o ICMS Retido ao total da nota
Dim m_blnSomaIcmsRetidoTotalNota As Boolean

'05/04/2010 - Andrea
Private m_CodOper As Integer

Private bProdutoSemPrecoNaGrade As Boolean


Public Sub Calcula_Linha()
  'Calcula preço total e final da linha
  Dim nX As Integer
  Dim Preço_Total As Double
  Dim Preço_Final As Double
  Dim Preço_Final2 As Double
  Dim Qtde As Double
  Dim Preço As Double
  Dim Desconto As Double
  Dim Valor_Desconto As Double
  Dim IPI As Double
  Dim Valor_IPI As Double
  Dim Vpreco As String
  
  
  With Grade1
    For nX = 1 To 12
      Select Case nX
        Case Is <> 2
          If .Columns(nX).Text = "" Then
            .Columns(nX).Text = 0
          End If
      End Select
    Next nX
    
    Qtde = .Columns("Qtde").Text
    '05/05/2004 - Daniel
    'Personalização Embalavi
    'Tratamento de Máscara
    If g_bln5CasasDecimais Then
      Preço = Format((.Columns("Preço").Text), "##,###,##0.00000")
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      Preço = Format((.Columns("Preço").Text), "##,###,##0.000")
    Else
      Preço = .Columns("Preço").Text
    End If
    
    Desconto = .Columns("Desc.%").Text
          
          
    ' ==============================================================
    ' Tratar IPI
    If rsParametros("CodigoRegimeTributario") <> 1 Then
        IPI = .Columns("IPI").Text
    End If
    
'''    '------------------------------------------------------
'''    If Not IsNull(rsProdutos2("IPI_ValidoEntradaSaida").Value) And rsProdutos2("IPI_ValidoEntradaSaida").Value = 1 Then
'''        IPI = .Columns("IPI").Text
''''    ElseIf Not IsNull(rsProdutos2("IPI_ValidoEntradaSaida").Value) And rsProdutos2("IPI_ValidoEntradaSaida").Value <> 1 Then
''''        If cboFinalidade.ListIndex = 3 Then
''''            'Finalidade devolução
''''            IPI = .Columns("IPI").Text
''''        Else
''''            .Columns("IPI").Text = "0"
''''        End If
'''    Else
'''        .Columns("IPI").Text = "0"
'''    End If
    
    
    
''''    If m_blnEmbalavi Then
''''      If Len(Nome_Cliente.Caption) > 0 Then
''''        If IsencaoIPI(CLng(Combo_Cliente.Text)) Then 'Cliente É Isento de IPI
'''        If m_blnIsentoIPI Then
'''          IPI = 0
'''        Else
'''          IPI = .Columns("IPI").Text
'''        End If
''''      Else 'Len...
''''        IPI = .Columns("IPI").Text
''''      End If
''''
''''    Else 'Não Embalavi
''''      IPI = .Columns("IPI").Text
''''    End If
    '------------------------------------------------------
    
    'Preço_Total = Format(Qtde * Preço, "#0.00")
    Preço_Total = Format(Qtde * Preço, "########0.00")
    Vpreco = Format(Preço_Total, "##,###,##0.00")
    .Columns("Total").Text = Vpreco 'Preço_Total
    
'    'MsgBox "Quantidade: " + CStr(Qtde) + " - Preço: " + CStr(Preço) + " - Vpreco: " + CStr(Vpreco)
'
'    ' PILATTI INICIO 2017/07/03
'    Dim vAuxI As Integer
'    Dim vAuxI2 As Integer
'    Dim vAuxPreco As String
'
'    vAuxI = InStr(Preço_Total, ",")
'    vAuxI2 = Len(Preço_Total)
'    If vAuxI > 1 Then
'      vAuxPreco = Mid(Preço_Total, 1, vAuxI)
'      vAuxPreco = vAuxPreco + Mid(Preço_Total, vAuxI + 1, vAuxI2 - (vAuxI + 1))
'    End If
'    .Columns("Total").Text = vAuxPreco
'    ' PILATTI FIM
'
'    'MsgBox "vAuxPreco: " + vAuxPreco

    
    Valor_Desconto = Format(Preço_Total * Desconto / 100, "#0.00")
    Preço_Final = Format((Preço_Total - Valor_Desconto), "#0.00")
    Valor_IPI = Format(Preço_Final * IPI / 100, "#0.00")
    If Not Calcula_IPI Then
      Valor_IPI = 0
    End If
    
    Preço_Final2 = Format((Preço_Final + Valor_IPI), "#0.00")
    Vpreco = Format(Preço_Final2, "##,###,##0.00")
    .Columns("Total").Text = Vpreco
    
'    'MsgBox "Preço_Final: " + CStr(Preço_Final) + " - Valor_IPI: " + CStr(Valor_IPI) + " - Preço_Final2: " + CStr(Preço_Final2)
'
'    ' PILATTI INICIO 2017/07/03
'    vAuxI = InStr(Preço_Final2, ",")
'    vAuxI2 = Len(Preço_Final2)
'    If vAuxI > 1 Then
'      vAuxPreco = Mid(Preço_Final2, 1, vAuxI)
'      vAuxPreco = vAuxPreco + Mid(Preço_Final2, vAuxI + 1, vAuxI2 - (vAuxI + 1))
'    End If
'    .Columns("Total").Text = vAuxPreco
'    ' PILATTI FIM
   
    'Calculo do ICM
    If .Columns("Tipo_ICM").Text = "N" Then
      'ICM Normal
      If gbBaseICMSomadoIPI = True Then
        .Columns("Base_ICM").Text = Preço_Final2
        .Columns("Valor_ICM").Text = Preço_Final2 * CSng(gsHandleNull(.Columns("ICM").Text & "")) / 100
      Else
        .Columns("Base_ICM").Text = Preço_Final
        .Columns("Valor_ICM").Text = Preço_Final * CSng(gsHandleNull(.Columns("ICM").Text & "")) / 100
      End If
    ElseIf .Columns("Tipo_ICM").Text = "R" Then
      'ICM Retido
      If CDbl(.Columns("Valor_Base_Unit").Text) <> 0 Then
        'Base Fixa
        .Columns("Base_ICM").Text = CDbl(.Columns("Qtde").Text) * CDbl(.Columns("Valor_Base_Unit").Text)
        .Columns("Valor_ICM").Text = CDbl(.Columns("Base_ICM").Text) * CDbl(.Columns("ICM").Text) / 100
      End If
      If CDbl(.Columns("Redução_ICM").Text) <> 0 Then
        'Base Reduzida
        .Columns("Base_ICM").Text = Preço_Final * CDbl(.Columns("Redução_ICM").Text) / 100
        .Columns("Valor_ICM").Text = CDbl(.Columns("Base_ICM").Text) * CDbl(.Columns("ICM").Text) / 100
      End If
    ElseIf .Columns("Tipo_ICM").Text = "Z" Then
      'ICM Reduzido
      If CDbl(.Columns("Valor_Base_Unit").Text) <> 0 Then
        'Base Fixa
        .Columns("Base_ICM").Text = CDbl(.Columns("Qtde").Text) * CDbl(.Columns("Valor_Base_Unit").Text)
        .Columns("Valor_ICM").Text = CDbl(.Columns("Base_ICM").Text) * CDbl(.Columns("ICM").Text) / 100
      End If
      If CDbl(.Columns("Redução_ICM").Text) <> 0 Then
        'Base Reduzida
        .Columns("Base_ICM").Text = Preço_Final * CDbl(.Columns("Redução_ICM").Text) / 100
        .Columns("Valor_ICM").Text = CDbl(.Columns("Base_ICM").Text) * CDbl(.Columns("ICM").Text) / 100
      End If
    End If
  End With
End Sub

Private Sub Calcula_Linha_Tabe(ByVal nRow As Long)
  Dim Qtde As Double
  Dim Preço As Double
  Dim Desconto As Double
  Dim Valor_Desconto As Double
  Dim IPI As Double
  Dim Valor_IPI As Double
  Dim Preço_Total As Double
  Dim Preço_Final As Double
  Dim Vpreco As String
  
  Qtde = Tabe(nRow).Qtde
  '05/05/2004 - Daniel
  'Personalização Embalavi
  'Tratamento de Máscara
  If g_bln5CasasDecimais Then
    Preço = Format((Tabe(nRow).Preço), "##,###,##0.00000")
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    Preço = Format((Tabe(nRow).Preço), "##,###,##0.000")
  Else
    Preço = Tabe(nRow).Preço
  End If
  
  Desconto = Tabe(nRow).Desconto
  
  '------------------------------------------------------
  '23/05/2006 - mpdea
  'Comentado restrição de isenção de IPI para a Embalavi
  'É utilizado configuração do cadastro de clientes
  '
  '07/05/2004 - Daniel
  'Personalização Embalavi
  'Exatamente neste ponto que temos em mãos
  'o percentual do IPI do produto
  'Tratamento através da função IsencaoIPI para
  'verificação se suspende ou não a taxa de IPI conforme
  'o cliente e não o produto
'  If m_blnEmbalavi Then
'    If Len(Nome_Cliente.Caption) > 0 Then
'      If IsencaoIPI(CLng(Combo_Cliente.Text)) Then 'Cliente É Isento de IPI
      If m_blnIsentoIPI Then
        IPI = 0
      Else
        IPI = Tabe(nRow).IPI
      End If
'    Else 'Len...
'      IPI = Tabe(nRow).IPI
'    End If
'
'  Else 'Não Embalavi
'    IPI = Tabe(nRow).IPI
'  End If
  '------------------------------------------------------
  
  'Preço_Total = Format(Qtde * Preço, "#0.00")
  Preço_Total = Format(Qtde * Preço, "########0.00")
  Valor_Desconto = Preço_Total * Desconto / 100
  Preço_Final = (Preço_Total - Valor_Desconto)
  Valor_IPI = Preço_Final * IPI / 100
  If Not Calcula_IPI Then
    Valor_IPI = 0
  End If
  Preço_Final = Preço_Final + Valor_IPI
  Vpreco = Format(Preço_Final, "##,###,##0.00")
  Tabe(nRow).Preço_Final = Vpreco
End Sub

Private Function Grava_Venda() As Integer
  Dim i As Integer
  Dim Conta As Integer
  'Dim Linha As Integer
  Dim Ordem As Integer
  Dim Aux_Cod_Prod As String
  Dim Limite_Usado As Double
  Dim Máximo As Double
  Dim Aux_Texto As String
  
  Dim nSequencia As Long
  'Variáveis de Tratamento de Erro
  Dim bSequencia As Boolean
  Dim bSeqChanged As Boolean
  Dim nRepeatUpdate3022 As Integer
  Dim nRepeatUpdateLocked As Integer
      
  Dim nPercMaxDesc As Single
  Dim cDescMax As Currency
  
  Dim sUnidade As String
  Dim sTributaria As String
  
  '03/08/2002 - mpdea
  Dim blnInTransaction As Boolean
  
  '05/02/2004 - mpdea
  Dim intCodVendedor As Integer
  
  
  On Error GoTo ErrHandler

  L_Estoque.Caption = ""
  
  totalNCM_2 = 0#
  
  ' Cód_Vendedor = Val(Right(Nome_Vendedor.Caption, 4))
  
  '28/06/2004 - Daniel e mpdea
  'Trocado estas 03 linhas abaixo pelas 04 próximas para
  'evitar problemas na gravação caso o user aumente o tamanho
  'da grid com o mouse
  'Linha = Grade1.Row
  'Grade1.Row = 2
  'Grade1.Row = 1
  With Grade1
    .MoveLast
    .MoveFirst
  End With
  
  DoEvents
  
  Call Combo_Vendedor_LostFocus
  
  If Efetivada.Visible = True Then
    DisplayMsg "Esta operação já foi efetivada e não pode ser alterada."
    Grava_Venda = 1
    Exit Function
  End If

  '07/05/2003 - mpdea
  'Verifica se a movimentação foi efetivada
  If Not IsNull(Num_Registro) Then
    If rsSaidas.Fields("Efetivada").Value Then
      DisplayMsg "Esta operação já foi efetivada e não pode ser alterada."
      Grava_Venda = 1
      Exit Function
    End If
  End If

  If IsNull(Num_Registro) And gbDemoVersion Then
    rsSaidas.MoveLast
    rsSaidas.MoveFirst
    If rsSaidas.RecordCount >= NMAXREGDEMO Then
      gsTitle = LoadResString(201)
      gsMsg = LoadResString(13)
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Grava_Venda = 1
      Exit Function
    End If
  End If
 
  '22/10/2004 - Daniel
  'Verifica se o caixa está incorreto
  'Flexibilidade de troca de caixa
  'Case: Solicitado por Casagrande
  If Len(Nome_Caixa.Caption) <= 0 Then
     DisplayMsg "Caixa incorreto, verifique."
     Grava_Venda = 1
     Exit Function
  End If
 
  If IsNull(Combo_Preço.Text) Then
     DisplayMsg "Tabela de preços incorreta, verifique."
     Grava_Venda = 1
     Exit Function
  End If

  If Not IsNull(Combo_Preço.Text) Then
    If Len(Combo_Preço.Text) > 15 Then
     DisplayMsg "Tabela de preços incorreta, verifique."
     Grava_Venda = 1
     Exit Function
    End If
  End If
 
  rsTabelas.Index = "Tabela"
  rsTabelas.Seek "=", Combo_Preço.Text
  If rsTabelas.NoMatch Then
    DisplayMsg "Tabela de preços não existe, verifique."
    Combo_Preço.SetFocus
    Grava_Venda = 1
    Exit Function
  End If
  
  '29/08/2003 - mpdea
  'Força a atualização do recordset rsFuncionarios
  Call Combo_Vendedor_LostFocus
  
  If Nome_Vendedor.Caption = "" Then
    DisplayMsg "Vendedor não digitado."
    Combo_Vendedor.SetFocus
    Grava_Venda = 1
    Exit Function
  End If
  
  'Verifica os dados digitados
  Call Combo_Cliente_LostFocus
  
  '-------------------------------------------------------------------------
  '18/09/2002 - mpdea
  'Incluído/modificado verificação para cliente inativo e bloqueado
  'Alterado mensagem para cliente não localizado
  If Nome_Cliente.Caption = "" Then
    DisplayMsg "Cliente inativo, bloqueado ou inexistente."
    If Combo_Cliente.Enabled = True Then Combo_Cliente.SetFocus
    Grava_Venda = 1
    Exit Function
  End If
  
  If rsCliFor("Bloqueado") Then
    DisplayMsg "Este cliente está bloqueado, impossível gravar."
    If Combo_Cliente.Enabled = True Then Combo_Cliente.SetFocus
    Grava_Venda = 1
    Exit Function
  End If
  
  If rsCliFor("Inativo") Then
    DisplayMsg "Este cliente está inativo, impossível gravar."
    If Combo_Cliente.Enabled = True Then Combo_Cliente.SetFocus
    Grava_Venda = 1
    Exit Function
  End If
  '-------------------------------------------------------------------------
  
  Conta = 0
  For i = 0 To (Grade1.Rows - 1)
   If Tabe(i).Código <> "" Then Conta = Conta + 1
  Next i
  
  If Conta = 0 Then
   DisplayMsg "Nenhum produto digitado, impossível gravar."
   Grade1.SetFocus
   Grava_Venda = 1
   Exit Function
  End If
   
  If Not IsNull(Num_Registro) Then
   If rsSaidas("Nota Cancelada") = True Then
     Beep
     DisplayMsg "A nota fiscal desta movimentação já foi cancelada. A movimentação não pode ser alterada."
     Grava_Venda = 1
     Exit Function
   End If
  End If
    
  If rsCliFor("Bloqueado") = True Then
    DisplayMsg "Este cliente está bloqueado, impossível gravar."
    Grava_Venda = 1
    Exit Function
  End If
    
  If CDbl(gsHandleNull(Total_Produtos)) = 0 Then
    DisplayMsg "Total da venda com valor igual a zero. Verifique."
    Grava_Venda = 1
    Exit Function
  End If
  
'  If rsParametros("VR Verifica Limite") = True And rsCliFor("Limite Crédito") <> 0 Then
'    Limite_Usado = Pega_Limite_Usado(rsCliFor("Código"))
'    If (Limite_Usado + Retorna_Valor(L_Tot_Pagar.Text)) > rsCliFor("Limite Crédito") Then
'      Máximo = rsCliFor("Limite Crédito") - Limite_Usado
'      DisplayMsg "Limite de crédito excedido. Não é possível vender. Venda máxima = " + Format(Máximo, "###,###,##0.00")
'      Grava_Venda = 1
'      Exit Function
'    End If
'  End If


  '29/12/2003 - mpdea
  'Verifica exigência da senha do vendedor para gravação
  If rsParametros.Fields("VR_GravarExigeSenhaVend").Value Then
    
    '05/02/2003 - mpdea
    'Se a movimentação já estiver gravada, somente o vendedor
    'que a gravou poderá alterá-la, caso contrário será o vendedor
    'que estiver selecionado
    If IsNull(Num_Registro) Then
      Call IsDataType(dtInteger, Combo_Vendedor.Text, intCodVendedor)
    Else
      Call IsDataType(dtInteger, rsSaidas.Fields("Digitador").Value, intCodVendedor)
    End If
    
    If Not frmSenhaFuncionario.CheckSenha(intCodVendedor) Then
      Grava_Venda = 1
      Exit Function
    End If
  End If
  
  
  '=======================================================================================
  '07/11/2002 - mpdea
  'Variável mcurDescontoSubTotal não estava incluída na verificação do desconto máximo
  
  'Tratamento Jun/2019 para verifiar limite de desconto pelo operador (e não pelo VENDEDOR)
  
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Cod_Operador.Caption)
  If rsFuncionarios.NoMatch Then Exit Function
  
  'Verifica a aplicação do desconto, de acordo com o limite do funcionário
  nPercMaxDesc = IIf(rsFuncionarios("nPercDesconto") = 0, _
    rsParametros("VR Desconto"), rsFuncionarios("nPercDesconto"))
  cDescMax = (Total_Pagar + Total_Desconto + mcurDescontoSubTotal) * nPercMaxDesc / 100
  
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Combo_Vendedor.Text)
  If rsFuncionarios.NoMatch Then Exit Function
  '=======================================================================================
  
  
  
  '20/09/2002 - mpdea
  'Incluído o Desconto no SubTotal
  If gParticipaProgramaFidelidade = 1 And gClienteEntregouResgatePontos = True And gSaldoCdGuidResgate > 0 Then
  
      If Total_Desconto + mcurDescontoSubTotal - gSaldoCdGuidResgate > cDescMax Then
        DisplayMsg "Desconto superior ao permitido."
        Grava_Venda = 1
        Exit Function
      End If
  
  Else
      If Total_Desconto + mcurDescontoSubTotal > cDescMax Then
        DisplayMsg "Desconto superior ao permitido."
        Grava_Venda = 1
        Exit Function
      End If
  End If
  
  
  '09/10/2002 - mpdea
  'Verifica estoque conforme configurações
  If Not rsParametros.Fields("Venda Sem Estoque").Value And rsOp_Saída.Fields("Estoque").Value Then
    If Not mblnCheckStock() Then
      Grava_Venda = 1
      Exit Function
    End If
  End If
  
  
  '02/05/2005 - Daniel
  '
  'Solicitante..: Jorge Marcos - PSI MT
  '
  'Finalidade...: Verificar o limite de crédito do cliente antes da gravação
  '               Isto é essencial para todas as empresas que trabalham
  '               com pronta entrega
  If Len(Nome_Cliente.Caption) > 0 Then
    If rsParametros("VerificaLimiteCli").Value Then
      Dim dblLimiteCli     As Double
      Dim dblLimiteCredito As Double
      
      Call GetLimiteCliente(Combo_Cliente.Text, dblLimiteCli)
      
      dblLimiteCredito = Format(dblLimiteCli - Pega_Limite_Usado(Combo_Cliente.Text), FORMAT_VALUE)
      
      If ((L_Tot_Pagar.Text) > dblLimiteCredito) Or ((L_Tot_Pagar.Text) > dblLimiteCli) Then
        MsgBox "O cliente ao qual você está fazendo a venda tem R$ " & _
               Format(dblLimiteCredito, FORMAT_VALUE) & " de saldo para novas compras. O recebimento estará sendo de R$ " & _
               Format(L_Tot_Pagar.Text, FORMAT_VALUE) & ". Não é possivel continuar !! ", vbCritical, "Quick Store"
        
        Grava_Venda = 1
        Exit Function
      End If
    End If 'If rsParametros("VerificaLimiteCli").Value
  End If
  
  'acha unidade de venda e situação tributária
  
  '18/01/2007 - Anderson
  'Solicitação senha do gerente ao alterar o vendedor relacionado ao cliente
  If rsParametros("VendedorSenhaGerente").Value Then
    If rsCliFor("Vendedor") <> 0 And rsCliFor("Vendedor") <> Combo_Vendedor.Text Then
      If MsgBox("O código do vendedor não corresponde ao cliente selecionado. A senha do gerente será necessária para concluir a gravação da venda." & Chr(13) & "Deseja continuar assim mesmo?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
        If Not frmGerente.gbSenhaGerente Then
          Grava_Venda = 1
          Exit Function
        End If
      Else
        Grava_Venda = 1
        Exit Function
      End If
    End If
  End If
    
  '11/12/07 - Celso
  'Se o cliente tem contas em atraso, exige senha do gerente para continuar com a venda
   If rsParametros.Fields("ExigeSenhaGerVndContaAtraso").Value Then
      If Not m_blnSenhaGerJaInformada Then
         Dim Total_atrasado As Double
         Total_atrasado = Pega_Atrasado_Cliente(Combo_Cliente.Text)
         If Total_atrasado > 0 Then
            DisplayMsg "Cliente [" & rsCliFor.Fields("Nome").Value & "] tem contas em atraso."
            'Senha do gerente
            If Not frmGerente.gbSenhaGerente Then
              Call SelectAllText(Combo_Cliente, True)
              Grava_Venda = 1
              Exit Function
            End If
            m_blnSenhaGerJaInformada = True
            m_strCodigoClienteContas = Combo_Cliente.Text
         End If
      End If
   End If
   
  
  '29/04/2008 - mpdea
  'Verifica número de documento do cliente
  Dim str_numero_documento_cliente As String
  If Not IsNull(Num_Registro) Then
    str_numero_documento_cliente = rsSaidas.Fields("NumeroDocumentoCliente").Value & ""
  End If
  str_numero_documento_cliente = g_str_GetNumeroDocumento(CInt(rsParametros("VR Código Operação")), CLng(Combo_Cliente.Text), str_numero_documento_cliente)
   
  '------------------------------------------------------
  
  
  Call StatusMsg("Gravando venda....")
    
    
  '11/08/2003 - mpdea
  'Desabilita controles
  Call EnableControls(False)

  
  '----------------------------------------------------------------------------------
  '29/05/2003 - mpdea
  'Atualizado
  '
  '05/08/2002 - mpdea
  'Requisição de bloqueio para gravação de venda
  If m_blnWorkTrafficLight Then
    Call TrafficLight.StartRequest(CLng(-1))
  End If
  '----------------------------------------------------------------------------------
  
  Call ws.BeginTrans
  blnInTransaction = True
  
  'Pega número da nova movimentação
  If IsNull(Num_Registro) Then
    nSequencia = gnGetNextSequencia(gnCodFilial)
    rsParametros.Edit
    rsParametros("Última Movimentação") = nSequencia
    rsParametros.Update
  End If
    
  If Erro_Data2 = True Then  'grava log
    rsLog.AddNew
    rsLog("Tipo") = "MOVIMENTAÇÃO"
    rsLog("Data") = Date
    Aux_Texto = "Movimentação " + CStr(nSequencia) + " gravada com data incorreta. Filial " + str(gnCodFilial)
    rsLog("Texto") = Aux_Texto
    rsLog.Update
  End If
    
  If IsNull(Num_Registro) Then
    rsSaidas.AddNew
    rsSaidas("Filial") = gnCodFilial
    rsSaidas("Sequência") = nSequencia
    Número.Text = ""
  Else
    rsSaidas.Bookmark = Num_Registro
    rsSaidas.Edit
    nSequencia = Val(Número.Text)
  End If

  rsSaidas("Data") = Data_Atual
  rsSaidas("Operação") = rsParametros("VR Código Operação")
  rsSaidas("Tabela") = Combo_Preço.Text
  rsSaidas("Digitador") = Val(Combo_Vendedor.Text)
  rsSaidas("Operador") = Val(Cod_Operador.Caption)
  rsSaidas("Cliente") = Val(Combo_Cliente.Text)
  rsSaidas("Observações") = Trim(Observacao.Text)
  
  'cboPrestador_LostFocus
  If PrestadorServicoSelecionado <> "" Then rsSaidas("PrestadorServico") = Val(cboPrestador.Text)
  
  '29/04/2008 - mpdea
  'Número de documento do cliente
  rsSaidas.Fields("NumeroDocumentoCliente").Value = str_numero_documento_cliente
  
  'rsSaidas("Nota Impressa") = 0
  rsSaidas("Produtos") = Format(Total_Produtos, "#############0.00")
  rsSaidas("Desconto") = Format(Total_Desconto, "#########0.00")
  
  '20/05/2005 - Daniel
  '
  'Solicitante: Pedágio Calçados
  'Tratamento para o campo [Nota Fiscal] e Nr de Série
  'em Venda Rápida não otimizamos às rotinas de tratamento
  'de notas geradas manualmente
  rsSaidas.Fields("Nota Fiscal").Value = 0
  rsSaidas.Fields("SerieNF").Value = ""
  '
  rsSaidas.Fields("Consumidor_Final").Value = 1
  rsSaidas.Fields("Presenca_Comprador").Value = 1
  '17/09/2009 - mpdea
  'Modelo de documento fiscal
  rsSaidas.Fields("ModeloDocumentoFiscal").Value = gstrGetModeloDocumentoFiscalOperacao(tmSaidas, rsSaidas.Fields("Operação").Value)
  
  '23/09/2002 - mpdea
  'Desconto no SubTotal
  rsSaidas("DescontoSubTotal") = mcurDescontoSubTotal
  
  rsSaidas("IPI") = Format(Total_IPI, "#############0.00")
  rsSaidas("Frete") = 0
  rsSaidas("Base ICM") = Format(Total_Base_ICM, "###############0.00")
  rsSaidas("Valor ICM") = Format(Total_ICM, "################0.00")
  rsSaidas("Base ICM Subs") = 0
  rsSaidas("Valor ICM Subs") = 0
  rsSaidas("Total") = Format(Total_Pagar, "##############0.00")
  rsSaidas("Serviços") = 0
  rsSaidas("Recebimento") = False
  rsSaidas("Caixa") = CByte(cboCaixa.Text)
  rsSaidas("Referência") = Referência.Text
  
  '07/01/2004 - Daniel
  'Alimentando os fields Valor Recebido
  'e Troco da Tabela Saídas
  rsSaidas.Fields("Valor Recebido").Value = frmRecebimento.g_dblValorRecebidoFrmRec
  rsSaidas.Fields("Troco").Value = frmRecebimento.g_dblTrocoFrmRec
  
  '21/05/2004 - Daniel
  'Tratamento de campos especiais da Bic Amazônia
  If m_blnBic Then
    rsSaidas("Codigo Func Comprador").Value = g_intCodigoFuncComprador
  Else
    rsSaidas("Codigo Func Comprador").Value = 0
  End If
        
  rsSaidas("Status Venda Func").Value = False
  '--------------------------------------------------------------------
  
  '12/05/2004 - Daniel
  'Não haverá serviços em venda rápida => rsSaidas("Serviços") = 0
  'Então os percentuais e totais sobre serviços serão zerados...
  'Percentuais
  rsSaidas.Fields("Percentual CSLL").Value = CSng(0)
  rsSaidas.Fields("Percentual COFINS").Value = CSng(0)
  rsSaidas.Fields("Percentual PIS").Value = CSng(0)
  rsSaidas.Fields("Percentual IRRF").Value = CSng(0)
  'Totais
  rsSaidas.Fields("Total CSLL").Value = CDbl(0)
  rsSaidas.Fields("Total COFINS").Value = CDbl(0)
  rsSaidas.Fields("Total PIS").Value = CDbl(0)
  rsSaidas.Fields("Total IRRF").Value = CDbl(0)
  'TotalMenosServ
  rsSaidas.Fields("TotalMenosServ").Value = CDbl(0)
  '-----------------------------------------------------------------------
  
  '23/04/2004 - Daniel
  'Case: PSV
  'O campo FaturaSourceReserva sempre será False até o momento
  'que a partir dele seja clonado uma saída para venda na tela de
  'Manutenção de Reservas
  rsSaidas.Fields("FaturaSourceReserva").Value = False
  '-----------------------------------------------------------------------
  
  bSeqChanged = False
  bSequencia = True
  rsSaidas.Update
  bSequencia = False
  'Grava novamente a última movimentação
  'se a mesma foi alterada
  If bSeqChanged Then
    With rsParametros
      .Edit
      .Fields("Última Movimentação") = nSequencia
      .Update
    End With
  End If
  
  rsSaidas.Bookmark = rsSaidas.LastModified
  Num_Registro = rsSaidas.Bookmark
  
  'Apaga produtos
  Call EraseTypeMoviment(tmSaidasProdutos, gnCodFilial, nSequencia)
  
    
  bProdutoSemPrecoNaGrade = False
  
  'Grava Produtos
  Conta = 1
  rsProdutos.Index = "Código"
  For i = 0 To (Grade1.Rows - 1)
    If Tabe(i).Código <> "" Then
      If Tabe(i).Qtde <> 0 Then
        If Tabe(i).Nome <> "" Then
          rsProdutos.Seek "=", UCase(Tabe(i).Código)
          If rsProdutos.NoMatch Then
             sUnidade = ""
             sTributaria = ""
          Else
             If Not IsNull(rsProdutos("Unidade Venda")) Then
                sUnidade = rsProdutos("Unidade Venda")
              Else
                sUnidade = " "
              End If
             If Not IsNull(rsProdutos("Situação Tributária")) Then
                sTributaria = rsProdutos("Situação Tributária")
              Else
                sTributaria = " "
              End If
          End If
          
          rsSaída_Prod.AddNew
            rsSaída_Prod("Filial") = gnCodFilial
            rsSaída_Prod("Sequência") = nSequencia
            rsSaída_Prod("Linha") = Conta
            rsSaída_Prod("Código") = UCase(Tabe(i).Código)
            '08/01/2004 - Daniel
            'Armazenar a quantidade para posterior impressões
            m_sngQtdeTotal = m_sngQtdeTotal + (Tabe(i).Qtde)
            '-------------------------------------------------
            rsSaída_Prod("Qtde") = Tabe(i).Qtde
            '05/05/2004 - Daniel
            'Personalização Embalavi
            'Tratamento de Máscara
            If g_bln5CasasDecimais Then
              rsSaída_Prod("Preço") = Format(Tabe(i).Preço, "##,###,##0.00000")
            '30/04/2007 - Anderson - Implementação de 3 casas decimais
            ElseIf g_bln3CasasDecimais Then
              rsSaída_Prod("Preço") = Format(Tabe(i).Preço, "##,###,##0.000")
            Else
              rsSaída_Prod("Preço") = Format(Tabe(i).Preço, "##############0.00")
            End If
            
            If rsSaída_Prod("Preço") = 0 Then
                bProdutoSemPrecoNaGrade = True
            End If
            
            rsSaída_Prod("Desconto") = Format(Tabe(i).Desconto, "##0.000")
            rsSaída_Prod("ICM") = Tabe(i).ICM
            rsSaída_Prod("IPI") = Tabe(i).IPI
            rsSaída_Prod("Etiqueta") = False
            
            '28/10/2004 - Daniel
            'Tratamento para o field [Saídas - Produtos].[Preço Final]
            'Para o cliente A.S. Wijma (Belém - Pará) deverá ser Double
            'para os demais clientes continua sendo Single
            If m_blnASWijmaBelem Then
              Call IsDataType(dtDouble, Format(Tabe(i).Preço_Final, "#########0.00"), m_dblPrecoFinalAuxi)
              rsSaída_Prod("Preço Final") = m_dblPrecoFinalAuxi
            Else
              rsSaída_Prod("Preço Final") = Format(Tabe(i).Preço_Final, "#########0.00")
            End If
            
            Aux_Cod_Prod = Tabe(i).Código
            Aux_Cod_Prod = Acha_Grade(Aux_Cod_Prod)
            rsSaída_Prod("Código Sem Grade") = Aux_Cod_Prod
            
            If sUnidade = "" Or IsNull(sUnidade) Then
               sUnidade = "  "
               rsSaída_Prod("Unidade Venda") = sUnidade
            Else
               rsSaída_Prod("Unidade Venda") = sUnidade
            End If
            If sTributaria = "" Or IsNull(sTributaria) Then
               sTributaria = " "
               rsSaída_Prod("Situação Tributária") = sTributaria
            Else
               rsSaída_Prod("Situação Tributária") = sTributaria
            End If
            
            '09/08/2007 - Anderson
            'Alteração realizada para armazenar o custo do produto no momento da venda
            rsSaída_Prod("PrecoCusto") = gcGetPrecoProduto(rsSaída_Prod("Código"), "CUSTO")
            
            '05/09/2008 - mpdea
            'Descrição adicional
            rsSaída_Prod("Descricao Adicional") = Tabe(i).DescricaoAdicional
            
            '05/04/2010 - Andrea
            'Alteração para o registro de CFOP por produto
            rsSaída_Prod("CFOP") = Tabe(i).CFOP_Produto

            '************************
            'Trata tributos
            Call UpdateTotalNCM_2(rsSaída_Prod("Código"))
            'Fim trata tributos
                        
            rsSaída_Prod.Update
            Conta = Conta + 1
            
        End If
      End If
    End If
  Next i
  'MsgBox "m_sngQtdeTotal ==" & m_sngQtdeTotal
  
  
  If bProdutoSemPrecoNaGrade = True Then
      If Me.Height > 8000 Then
          frm_produtoSemPrecoNaGrade.Left = 4110
          frm_produtoSemPrecoNaGrade.Top = 5580
          frm_produtoSemPrecoNaGrade.Visible = True
      Else
          frm_produtoSemPrecoNaGrade.Left = 7300
          frm_produtoSemPrecoNaGrade.Top = 4280
          frm_produtoSemPrecoNaGrade.Visible = True
      End If
  Else
      frm_produtoSemPrecoNaGrade.Visible = False
  End If
  
  rsSaidas.Edit
      If totalNCM_2 > 0 Then
          rsSaidas("TotalNCM") = totalNCM_2
      End If
  rsSaidas.Update
  
            
  B_Recebe.Enabled = True
  
  Call StatusMsg("")
  
  'Fim da transação
  Call ws.CommitTrans
  blnInTransaction = False
  
  Grava_Venda = 0

  '23/09/2002 - mpdea
  'Registro atualizado, desativa flag para forçar atualização
  mblnForceUpdate = False

  '----------------------------------------------------------------------------------
  '29/05/2003 - mpdea
  'Atualizado
  '
  '05/08/2002 - mpdea
  'Remoção de bloqueio para gravação de venda
  If m_blnWorkTrafficLight Then
    Call TrafficLight.FinishRequest
  End If
  '----------------------------------------------------------------------------------
  
  
  '29/05/2003 - mpdea
  'Atualiza controle para exibição das sequências
  datSequencias.Refresh
  
  
  '11/08/2003 - mpdea
  'Habilita controles
  Call EnableControls(True)
  

  If Número.Text = "" Then
    Número.Text = nSequencia
  End If
  
'-----------------------------------------------------------------------------------------------------------------
' Joga dados da movimentação para o banco do GestoPDV por causa do PAF
'-----------------------------------------------------------------------------------------------------------------
  Dim GestoBD As Database
  Dim SaidaEstoque As Recordset
  Dim SaidaEstoqueItem As Recordset
  Dim ItemEstoqueAlmox As Recordset
  Dim QuickBD As Database
  Dim produtos As Recordset
  Dim cad_prod As Recordset
  If frmParametros.VerificaPAF = True Then
    Set rsParametros = db.OpenRecordset("Select * from [Parâmetros Filial] Where Filial = " & gnCodFilial & ";")
         
    Dim fso As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(rsParametros("BancoPDV").Value & "\Gesto.mde") Then

    Set GestoBD = OpenDatabase(rsParametros("BancoPDV").Value & "\Gesto.mde", False, False)
    Set SaidaEstoque = GestoBD.OpenRecordset("Select * from SaidaEstoque Where NUMERO = " & Número.Text & ";")
    If SaidaEstoque.EOF Then
      SaidaEstoque.AddNew
      SaidaEstoque!Numero = Número.Text
      SaidaEstoque!CODIGO_CLIENTE = Combo_Cliente.Text
      SaidaEstoque!Cliente = Left(Nome_Cliente.Caption, 40)
      SaidaEstoque!DATA_SAIDA = Data_Atual
      'If Obs.Text <> "" Then
        'SaidaEstoque!OBSERVACAO = Obs.Text
      'End If
      If txtDescSubTotal.Text <> "" Then
        SaidaEstoque!VL_DESCONTO = txtDescSubTotal.Text
      End If
      SaidaEstoque!COD_Vendedor = Combo_Vendedor.Text
      SaidaEstoque.Update
    Else
      SaidaEstoque.Edit
      SaidaEstoque!CODIGO_CLIENTE = Combo_Cliente.Text
      SaidaEstoque!Cliente = Left(Nome_Cliente.Caption, 40)
      SaidaEstoque!DATA_SAIDA = Data_Atual
      'If Obs.Text <> "" Then
        'SaidaEstoque!OBSERVACAO = Obs.Text
      'End If
      If txtDescSubTotal.Text <> "" Then
        SaidaEstoque!VL_DESCONTO = txtDescSubTotal.Text
      End If
      SaidaEstoque!COD_Vendedor = Combo_Vendedor.Text
      SaidaEstoque.Update
      Set SaidaEstoqueItem = GestoBD.OpenRecordset("Select * from SaidaEstoqueItem Where NUMERO = " & Número.Text & "")

        Do Until SaidaEstoqueItem.EOF = True
          If SaidaEstoqueItem.EOF = False Then
            SaidaEstoqueItem.Delete
            SaidaEstoqueItem.MoveNext
          End If
        Loop

    End If

    'Continuar PAF saidas produtos
    Dim Nome_Prod As Recordset
    Dim Estoque_Prod As Recordset
    Set produtos = db.OpenRecordset("Select * from [Saídas - Produtos] where Filial = " & gnCodFilial & " and Sequência = " & Número.Text & "")
    produtos.MoveFirst
    Do Until produtos.EOF
      Set cad_prod = db.OpenRecordset("Select * from Produtos where Código = '" & produtos("Código sem Grade") & "'")
      Set Nome_Prod = GestoBD.OpenRecordset("SELECT DESCRICAO From ItemEstoque WHERE CODIGO_FORNECEDOR = '" & produtos("Código sem Grade") & "'")
      Set ItemEstoqueAlmox = GestoBD.OpenRecordset("Select * from ItemEstoqueAlmox Where Codigo_Item = '" & produtos("Código sem Grade") & "'")
      If Nome_Prod.EOF Then
        MsgBox "O produto de código: " & produtos("Código sem Grade") & " não esta cadastrado no Gesto, para que o erro não volte a ocorrer entre no cadastro do produto e mande gravar."
        Exit Function
      End If
      If cad_prod("Tipo") = "N" Then
        Set Estoque_Prod = db.OpenRecordset("Select [Estoque Atual] From [Estoque Final] where Filial = " & gnCodFilial & " and Produto = '" & produtos("Código sem Grade") & "'")
        Set SaidaEstoqueItem = GestoBD.OpenRecordset("Select * from SaidaEstoqueItem Where NUMERO = " & Número.Text & " AND CODIGO_ITEM = '" & produtos("Código sem Grade") & "'")
        'If SaidaEstoqueItem.EOF Then
          SaidaEstoqueItem.AddNew
          SaidaEstoqueItem!Numero = Número.Text
          SaidaEstoqueItem!Item = produtos("Linha")
          SaidaEstoqueItem!Codigo_Item = produtos("Código")
          SaidaEstoqueItem!DESCRICAO_ITEM = Nome_Prod("DESCRICAO")
          SaidaEstoqueItem!Quantidade = produtos("Qtde")
          SaidaEstoqueItem!VALOR_UNIT_DESC = produtos("Preço") - (produtos("Preço") * produtos("Desconto") / 100)
          SaidaEstoqueItem!Valor_Total = produtos("Preço Final")
          SaidaEstoqueItem.Update
          If Estoque_Prod.EOF Then
            MsgBox "O produto " & Nome_Prod("DESCRICAO") & " esta com estoque não inicializado. Estoque não atualizado no Gesto. Favor inicializar estoque do produto na tela de cadastro de produto ou ficará com estoque errado."
          Else
            If Not ItemEstoqueAlmox.EOF Then
              ItemEstoqueAlmox.Edit
              ItemEstoqueAlmox!Qtde_Disponivel = Estoque_Prod("Estoque Atual")
              ItemEstoqueAlmox.Update
            End If
          End If
        'Produto Grade PAF
        ElseIf cad_prod("Tipo") = "G" Then
          Tamanho = 0
          Cor = 0
          Edicao = 0
          Tipo = 1
          Erro = 0
          Acha_Produto produtos("Código"), produtos("Código"), Tamanho, Cor, Edicao, Tipo, Erro
          Set Estoque_Prod = db.OpenRecordset("Select [Estoque Atual] From [Estoque Final] where Filial = " & gnCodFilial & " and Produto = '" & produtos("Código sem Grade") & "' AND Cor = " & Cor & " And Tamanho = " & Tamanho & "")
          Set SaidaEstoqueItem = GestoBD.OpenRecordset("Select * from SaidaEstoqueItem Where NUMERO = " & Número.Text & " AND CODIGO_ITEM = '" & produtos("Código sem Grade") & "'")
          If SaidaEstoqueItem.EOF Then
            SaidaEstoqueItem.AddNew
            SaidaEstoqueItem!Numero = Número.Text
            SaidaEstoqueItem!Item = produtos("Linha")
            SaidaEstoqueItem!Codigo_Item = produtos("Código")
            SaidaEstoqueItem!DESCRICAO_ITEM = Nome_Prod("DESCRICAO")
            SaidaEstoqueItem!Quantidade = produtos("Qtde")
            SaidaEstoqueItem!VALOR_UNIT_DESC = produtos("Preço") - (produtos("Preço") * produtos("Desconto") / 100)
            SaidaEstoqueItem!Valor_Total = produtos("Preço Final")
            SaidaEstoqueItem.Update
            If Estoque_Prod.EOF Then
              MsgBox "O produto " & Nome_Prod("DESCRICAO") & " esta com estoque não inicializado. Estoque não atualizado no Gesto. Favor inicializar estoque do produto na tela de cadastro de produto ou ficará com estoque errado."
            Else
              If Not ItemEstoqueAlmox.EOF Then
                ItemEstoqueAlmox.Edit
                ItemEstoqueAlmox!Qtde_Disponivel = Estoque_Prod("Estoque Atual")
                ItemEstoqueAlmox.Update
              End If
            End If
        End If
     End If
    produtos.MoveNext
    Loop
    Dim rsLibera As Recordset
    Set rsLibera = GestoBD.OpenRecordset("SELECT id From parametro")
  End If
  End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  Exit Function
    
ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3022 And bSequencia 'Duplicidade de movimentação
      If nRepeatUpdate3022 < 1000 Then
        Call StatusMsg("Verificando registro...")
        nRepeatUpdate3022 = nRepeatUpdate3022 + 1
        nSequencia = gnGetNextSequencia(gnCodFilial)
        bSeqChanged = True
        rsSaidas("Sequência") = nSequencia
        Resume
      End If
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If nRepeatUpdateLocked < 30 Then
        Call frmAvisoBloqueio.ShowTentativas(30 - nRepeatUpdateLocked)
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        nRepeatUpdateLocked = nRepeatUpdateLocked + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          nRepeatUpdateLocked = 0
          Resume
        Else
          Grava_Venda = -1 'Ação cancelada
          'Cancelamento da transação
          If blnInTransaction Then ws.Rollback
          GoTo EnableControls
          Exit Function
        End If
        
'        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
'          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Venda Rápida - Gravar") = vbOK Then
'          nRepeatUpdateLocked = 0
'          Resume
'        Else
'          Grava_Venda = -1 'Ação cancelada
'          'Cancelamento da transação
'          If blnInTransaction Then ws.Rollback
'          GoTo EnableControls
'          Exit Function
'        End If
      End If
    Case Else
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Venda Rápida - Gravar")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Grava_Venda = -1 'Ação cancelada
          GoTo EnableControls
          Exit Function
        Case 3 'Encerrar
          End
      End Select
  End Select

  Exit Function

EnableControls:
  '11/08/2003 - mpdea
  'Habilita controles
  Call EnableControls(True)
  

End Function

Sub Limpa_Tela(Ação As Integer)
  Dim Linha As Integer
  Dim i As Integer

  On Error GoTo Processa_Erro

  Call StatusMsg("")
  
  txtComanda.Text = ""
  btnComandaVendas.Visible = False
  txtComanda.Width = 1785
  
  Observacao.Text = ""

  '10/05/2024 - Pablo
  'If rsParametros("comPrestServ") Then
    cboPrestador.Value = 0
    Apelido_Prestador.Caption = ""
  'End If
 
  '29/05/2003 - mpdea
  'Atualiza controle para exibição das sequências
  datSequencias.Refresh


  Efetivada.Visible = False
  Movimentação_Desfeita.Visible = False
  
  Total_Recebido = 0
  
  Erase Tabe
  
  
  '26/07/2004 - mpdea
  'Setado como default 0 (zero) para a coluna código
  'e _ (underscore) para a coluna nome
  'devido a performance do controle de lista associado
  For Linha = 0 To (Grade1.Rows - 1)
    'Comentado em Junho/2019 ....tirando o ZERO na coluna do grade1
    ''''Tabe(Linha).Código = "0"
    Tabe(Linha).Código = ""
    Tabe(Linha).Nome = "_"
'   Tabe(Linha).Preço_Final = 0
'   Tabe(Linha).Qtde = 0
'   Tabe(Linha).Preço = 0
'   Tabe(Linha).Desconto = 0
'   Tabe(Linha).ICM = 0
'   Tabe(Linha).IPI = 0
  Next Linha
  
  Grade1.MoveLast
  Grade1.MoveFirst

  Num_Registro = Null
  Total_Desconto = 0
  gcDescInTotal = 0
  
  '23/09/2002 - mpdea
  'Desconto no SubTotal
  mcurDescontoSubTotal = 0
  txtSubTotal.Text = Format("0", FORMAT_VALUE)
  txtDescSubTotal.Text = Format("0", FORMAT_VALUE)
  
  Recalcula
  
  B_Recebe.Enabled = False
    
  frmRecebimento.Limpa_Tela (0)

  Rem Limpa recebimento
  Label_Receber.Caption = "A Receber"
  Lançar_Débito.Value = 0
  Dinheiro.Text = ""
  Vale.Text = ""
  Combo_Cartão.Text = ""
  Combo_Cartão_LostFocus
  Num_Cartão.Text = ""
  Val_Cartão.Text = ""
  
  L_Preço.Caption = ""
  
  For i = 0 To 4
   Banco(i).Text = ""
   Cheque(i).Text = ""
   '29/06/2004 - Daniel
   'Trocado componente Bom_Para de Text para Mask
   Bom_Para(i).Mask = ""
   Bom_Para(i).Text = ""
   Bom_Para(i).Mask = "##/##/####"
   '---------------------------------------------
   Val_Cheque(i).Text = ""
   '30/06/2004 - Daniel
   'Trocado componente Data_Parc de Text para Mask
   Data_Parc(i).Mask = ""
   Data_Parc(i).Text = ""
   Data_Parc(i).Mask = "##/##/####"
   '---------------------------------------------
   Val_Parc(i).Text = ""
  Next i
  
  Combo_Cliente.Text = rsParametros("VR Cliente")
  Combo_Cliente_LostFocus


  If ActiveBar1.Tools("miOpFreezeVendedor").Checked = False Then
    Combo_Vendedor.Text = ""
    Combo_Vendedor_LostFocus
  End If

  '14/12/2009 - Andrea
  lblRecebidoComCartao.Visible = False
  B_Recebe_Simples.Visible = False
  fraButtonRecebeSimples.Visible = True
  Val_Cartão.Visible = True


  Número.Text = ""
  Referência.Text = ""
  
  '21/05/2004 - Daniel
  'Limpamos as Vars Public
  g_intCodigoFuncComprador = 0
  g_blnStatusVendaFunc = False
  g_blnRetornar = False
  '---------------------------------------
  
  '23/09/2002 - mpdea
  'Novo registro, desativa flag para forçar atualização
  mblnForceUpdate = False
     
  Exit Sub
 
  If Ação = 0 Then
    Grade1.SetFocus
    'SendKeys "{TAB}"
    Exit Sub
  End If
  
  
  '------------------------------------------------
  '26/07/2004 - mpdea
  'Modificado seleção de controles com a função
  'SelectAllText que previne ocorrência de erro
  'ao setar o controle
  If Combo_Preço.Enabled Then
    Call SelectAllText(Combo_Preço, True)
    Exit Sub
  End If
  '
  If Combo_Cliente.Enabled Then
    Call SelectAllText(Combo_Cliente, True)
    Exit Sub
  End If
  '------------------------------------------------
  
Processa_Erro:

  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Venda Rápida - Limpar")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Sub
    Case 3 'Encerrar
      End
  End Select
  
End Sub

Private Sub cboPrestador_Click()
  cboPrestador_LostFocus
End Sub

Private Sub cboPrestador_LostFocus()
  Apelido_Prestador.Caption = IIf(PrestadorServicoSelecionado = "", "", cboPrestador.Columns(1).Text)

  Screen.MousePointer = vbDefault
    With Grade1
      .MoveLast
      .MoveFirst
    End With
  Screen.MousePointer = vbDefault
End Sub

Private Function PrestadorServicoSelecionado() As String
  Dim Retorno As String
  Retorno = IIf(Not rsParametros("comPrestServ"), "", cboPrestador.Columns(0).Text)
  If Retorno <> "" Then Retorno = IIf(IsNull(Retorno), "", Retorno)
  If Retorno <> "" Then Retorno = IIf(Not IsNumeric(Retorno), "", Retorno)
  If Retorno <> "" Then Retorno = IIf(Val(Retorno) < 0, "", Retorno)
  If Retorno <> "" Then Retorno = IIf(Val(Retorno) > 9999, "", Retorno)
  
  PrestadorServicoSelecionado = Retorno
End Function

Sub Mostra_Dados_Recebimento()

  Dim i As Integer
  Dim Erro As Integer
  Dim Ordem As Integer
 
  Lançar_Débito.Value = -rsSaidas("Recebe - Conta")
  Dinheiro.Text = rsSaidas("Recebe - Dinheiro")

  
  '--------------------------------------------------------------------------------------------
  '15/03/2012 - mpdea
  'Corrigido associação com recordset modular
  Dim rsSaída_Cartoes_local As Recordset
  '14/12/2009 - Andrea
  'Soma os recebimentos em cartões
  Dim strSQL As String
  Dim int_nro_cartoes As Integer
  Dim dbl_valor_recebido_cartao As Double
  Dim str_numero_cartao As String
  Dim str_administradora_cartao As String
  Dim bln_credito As Boolean
  
  Ordem = 0
  Erro = False
  int_nro_cartoes = 0
  dbl_valor_recebido_cartao = 0
  
  strSQL = "SELECT * "
  strSQL = strSQL & "FROM [Movimento - Cartoes] WHERE [Movimento - Cartoes].Filial = " & gnCodFilial & "  AND "
  strSQL = strSQL & "[Movimento - Cartoes].Sequência = " & rsSaidas("Sequência") & " ORDER BY [Movimento - Cartoes].Ordem "
  Set rsSaída_Cartoes_local = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsSaída_Cartoes_local
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        
        int_nro_cartoes = int_nro_cartoes + 1
        dbl_valor_recebido_cartao = dbl_valor_recebido_cartao + rsSaída_Cartoes_local.Fields("Valor").Value
        str_numero_cartao = rsSaída_Cartoes_local.Fields("NumeroCartao").Value
        str_administradora_cartao = rsSaída_Cartoes_local.Fields("Administradora").Value
        bln_credito = rsSaída_Cartoes_local.Fields("Credito").Value
        
       .MoveNext
      
      Loop
    End If
    .Close
  End With
  Set rsSaída_Cartoes_local = Nothing
  '--------------------------------------------------------------------------------------------
  
  '14/12/2009 - Andrea
  If int_nro_cartoes > 0 Then ' Foi recebido em cartão na tela de recebimento
    If int_nro_cartoes > 1 Then 'Foi recebido em mais de um cartão
      lblRecebidoComCartao.Visible = True
      B_Recebe_Simples.Visible = False
      fraButtonRecebeSimples.Visible = False
      Val_Cartão.Visible = False
    Else 'Foi recebido só em 1 cartão, move os dados para a tela
      
      rsCartoes.Index = "Nome"
      rsCartoes.Seek "=", str_administradora_cartao
      
      
      If rsCartoes.RecordCount <> 0 Then
        If Not rsCartoes.NoMatch Then
          Combo_Cartão.Text = rsCartoes("Código").Value
        End If
        

      End If
      Nome_Cartão.Caption = str_administradora_cartao & ""
      Num_Cartão.Text = str_numero_cartao & ""
      Val_Cartão.Text = dbl_valor_recebido_cartao
    End If
  Else
    Combo_Cartão.Text = rsSaidas("Recebe - Emp Cartão") & ""
    Combo_Cartão_LostFocus
    Num_Cartão.Text = rsSaidas("Recebe - Num Cartão") & ""
    Val_Cartão.Text = rsSaidas("Recebe - Cartão")
  End If
  
  Vale.Text = rsSaidas("Recebe - Vale")
  
  
 rsSaída_Cheques.Index = "Ordem"
 Ordem = 0
 i = 0
 Erro = False
 Do
   rsSaída_Cheques.Seek ">", gnCodFilial, rsSaidas("Sequência"), Ordem
   If rsSaída_Cheques.NoMatch Then Erro = True
   If Erro = False Then If rsSaída_Cheques("Filial") <> gnCodFilial Then Erro = True
   If Erro = False Then If rsSaída_Cheques("Sequência") <> rsSaidas("Sequência") Then Erro = True
   If Erro = False Then
     Ordem = rsSaída_Cheques("Ordem")
     If i < 5 Then
      Banco(i).Text = rsSaída_Cheques("Banco")
      Cheque(i).Text = rsSaída_Cheques("Cheque")
      Bom_Para(i).Text = rsSaída_Cheques("Bom")
      Val_Cheque(i).Text = rsSaída_Cheques("Valor")
      i = i + 1
     End If
   End If
 Loop Until Erro = True
 
 
 
 rsSaída_Parcelas.Index = "Ordem"
 Ordem = 0
 Erro = False
 i = 0
 Do
   rsSaída_Parcelas.Seek ">", gnCodFilial, rsSaidas("Sequência"), Ordem
   If rsSaída_Parcelas.NoMatch Then Erro = True
   If Erro = False Then If rsSaída_Parcelas("Filial") <> gnCodFilial Then Erro = True
   If Erro = False Then If rsSaída_Parcelas("Sequência") <> rsSaidas("Sequência") Then Erro = True
   If Erro = False Then
     Ordem = rsSaída_Parcelas("Ordem")
     If i < 5 Then
      Data_Parc(i).Text = rsSaída_Parcelas("Bom")
      Val_Parc(i).Text = rsSaída_Parcelas("Valor")
      i = i + 1
     End If
   End If
 Loop Until Erro = True
 
End Sub

Sub Mostra_Mov(ByVal Num As Long)

  Dim Linha As Integer
  Dim Erro As Integer
  Dim Nome_Prod As String
  Dim Aux_Prod As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim Código As String

  On Error GoTo Processa_Erro
  
  cboPrestador.Value = 0
  Apelido_Prestador.Caption = ""
  
  '07/05/2003 - mpdea
  'Zera o total de desconto concedido
  Total_Desconto = 0
  gcDescInTotal = 0


  Limpa_Tela (1)
  
  Call StatusMsg("")
  
  '25/03/2004 - Daniel
  If m_blnUserDanger Then
    B_Recebe.Enabled = True
  End If
  '---------------------
  
  Número.Text = Num
  
  rsSaidas.Index = "Sequência"
  rsSaidas.Seek "=", gnCodFilial, Num
  If rsSaidas.NoMatch Then
    DisplayMsg "Venda não encontrada."
    Exit Sub
  End If
  
  Num_Registro = rsSaidas.Bookmark
  
'''  DoEvents
'''  Sleep 300
'''  DoEvents
  
  Combo_Preço.Text = rsSaidas("Tabela") & ""
  Combo_Cliente.Text = rsSaidas("Cliente")
  Combo_Cliente_LostFocus
  
  Referência.Text = rsSaidas("Referência") & ""
  
  rsSaída_Prod.Index = "Sequência"
  rsProdutos.Index = "Código"
  Linha = 0
  Erro = False
  
'''  DoEvents
'''  Sleep 300
  DoEvents
  
  Do
    rsSaída_Prod.Seek ">", gnCodFilial, Num, Linha
    If rsSaída_Prod.NoMatch Then Erro = True
    If Erro = False Then If rsSaída_Prod("Filial") <> gnCodFilial Then Erro = True
    If Erro = False Then If rsSaída_Prod("Sequência") <> Num Then Erro = True
    If Erro = False Then
      Linha = rsSaída_Prod("Linha")
      Tabe(Linha - 1).Código = rsSaída_Prod("Código")
      Nome_Prod = "Produto inexistente ou apagado."
      
      Aux_Prod = rsSaída_Prod("Código")
      Acha_Produto Aux_Prod, Código, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Tipo, Aux_Erro
      If Aux_Erro = 0 Then
         rsProdutos.Seek "=", Código
         If Not rsProdutos.NoMatch Then Nome_Prod = rsProdutos("Nome")
      End If
      
      Tabe(Linha - 1).Nome = Nome_Prod
      Tabe(Linha - 1).Preço_Final = rsSaída_Prod("Preço Final")
      Tabe(Linha - 1).Qtde = rsSaída_Prod("Qtde")
      '05/05/2004 - Daniel
      'Personalização Embalavi
      ''Tratamento de Máscara
      If g_bln5CasasDecimais Then
        Tabe(Linha - 1).Preço = Format(rsSaída_Prod("Preço"), "##,###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Tabe(Linha - 1).Preço = Format(rsSaída_Prod("Preço"), "##,###,##0.000")
      Else
        Tabe(Linha - 1).Preço = rsSaída_Prod("Preço")
      End If
      
      Tabe(Linha - 1).Desconto = rsSaída_Prod("Desconto")
      Tabe(Linha - 1).ICM = rsSaída_Prod("ICM")
      Tabe(Linha - 1).IPI = rsSaída_Prod("IPI")
      Tabe(Linha - 1).Base_ICM = rsSaidas("Base ICM")
      Tabe(Linha - 1).Valor_ICM = rsSaidas("Valor ICM")
      Tabe(Linha - 1).Tipo_ICM = rsProdutos("Tipo ICM") & ""
            
      '13/11/2008 - mpdea
      'Não estava preenchendo a redução e base de cálculo
      'ocasionando erros em movimentações já gravadas
      Tabe(Linha - 1).Redução_ICM = rsProdutos("Redução ICM")
      Tabe(Linha - 1).Valor_Base_Unit = rsProdutos("Base Cálculo")
      
      '05/09/2008 - mpdea
      'Descrição adicional
      Tabe(Linha - 1).DescricaoAdicional = rsSaída_Prod("Descricao Adicional") & ""
      
      '05/04/2010 - Andrea
      'Alteração para registro de CFOP por produto
      Tabe(Linha - 1).CFOP_Produto = rsSaída_Prod("CFOP") & ""

    End If
  Loop Until Erro = True
  
  
  '29/05/2003 - mpdea
  'Ativado controle Redraw
  With Grade1
    .Redraw = False
    .MoveLast
    .MoveFirst
    .Redraw = True
    '03/09/2003 - mpdea
    'Força exibição dos registros
    .Refresh
  End With
  
'''  DoEvents
  
  '08/11/2002 - mpdea
  'Verificação de nulo
  '23/09/2002 - mpdea
  'Desconto no SubTotal
  Call IsDataType(dtCurrency, rsSaidas.Fields("DescontoSubTotal").Value, mcurDescontoSubTotal)

  
  L_Tot_Prod.Text = Format(rsSaidas("Produtos"), "###,###,##0.00")
  L_Tot_IPI.Text = Format(rsSaidas("IPI"), "###,###,##0.00")
  L_Tot_Pagar.Text = Format(rsSaidas("Total"), "###,###,##0.00")
  Total_Pagar = rsSaidas("Total")

  
  '20/09/2002 - mpdea
  'Exibição com o Desconto no SubTotal
  txtSubTotal.Text = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
  txtDescSubTotal.Text = Format(mcurDescontoSubTotal, FORMAT_VALUE)
  
  
  Call Recalcula
  
  DoEvents
  
  Efetivada.Visible = False
  If rsSaidas("Efetivada") = True Then Efetivada.Visible = True
  If rsSaidas("Movimentação desfeita") = True Then Movimentação_Desfeita.Visible = True
  
  rsUsuarios.Index = "Código"
  rsUsuarios.Seek "=", rsSaidas("Digitador")
  If rsUsuarios.NoMatch Then Exit Sub
  
'''  Me.Caption = " Venda Rápida - Caixa: " + rsUsuarios("Apelido")
  Me.Caption = " Venda Rápida - Operador: " + rsUsuarios("Apelido")
  
  Combo_Vendedor.Text = rsSaidas("Digitador")
  Combo_Vendedor_LostFocus
  
  If rsParametros("comPrestServ") And Not IsNull(rsSaidas("PrestadorServico")) Then
    cboPrestador.Value = rsSaidas("PrestadorServico")
    
    Dim sqlPrest As String
    sqlPrest = "SELECT f.Apelido AS apelido " & _
               "FROM Funcionários AS f " & _
               "WHERE (((f.Ativo)=True) AND ((f.Liberado)=True) AND ((f.isPrestServ)=True) AND f.Código = " & rsSaidas("PrestadorServico") & ");"
    Dim rsPrest_tmp As Recordset
    Set rsPrest_tmp = db.OpenRecordset(sqlPrest, dbOpenDynaset, dbReadOnly)
  
    With rsPrest_tmp
      'If Not (.BOF And .EOF) Then
        Do Until .EOF
          Apelido_Prestador.Caption = .Fields("apelido").Value
         .MoveNext
        Loop
      'End If
      .Close
    End With
    Set rsPrest_tmp = Nothing
  End If
  
  If rsSaidas("Recebimento") = True Then B_Recebe.Enabled = True

  If Frame_Recebimento.Visible = True Then
    Mostra_Dados_Recebimento
  End If
  
  If txtComanda.Visible = True Then
      CarregaComanda
  End If
  
  Observacao.Text = rsSaidas("Observações")
  
  Exit Sub
  
Processa_Erro:

  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Venda Rápida - Mostrar Registro")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Sub
    Case 3 'Encerrar
      End
  End Select
  
End Sub


Sub Recalcula_Recebido()
  Dim i As Integer
  

  Total_Recebido = 0
  
  If Not IsNumeric(Dinheiro.Text) Then Dinheiro.Text = 0
  
  If Lançar_Débito.Value = 1 Then
    Total_Recebido = Total_Pagar
    GoTo Fim_Loop
  End If
  
  
  If Vale.Text = "" Then Vale.Text = 0
  If Val_Cartão.Text = "" Then Val_Cartão.Text = 0
  If Dinheiro.Text = "" Then Dinheiro.Text = 0
  
  If Not IsNumeric(Vale.Text) Then Vale.Text = 0
  If Not IsNumeric(Dinheiro.Text) Then Dinheiro.Text = 0
  If Not IsNumeric(Val_Cartão.Text) Then Val_Cartão.Text = 0
  
  If CDbl(Vale.Text) < 0 Then Vale.Text = 0
  If CDbl(Dinheiro.Text) < 0 Then Dinheiro.Text = 0
  If CDbl(Val_Cartão.Text) < 0 Then Val_Cartão.Text = 0
  
  
  Total_Recebido = Total_Recebido + CDbl(Dinheiro.Text)
  Total_Recebido = Total_Recebido + CDbl(Vale.Text)
  Total_Recebido = Total_Recebido + CDbl(Val_Cartão.Text)
  
  For i = 0 To 4
    If Val_Cheque(i).Text = "" Then Val_Cheque(i).Text = 0
    If Val_Parc(i).Text = "" Then Val_Parc(i).Text = 0
    Total_Recebido = Total_Recebido + CDbl(Val_Cheque(i))
    Total_Recebido = Total_Recebido + CDbl(Val_Parc(i))
  Next i
  
Fim_Loop:
  L_Receber.Text = Format(Abs(Total_Pagar - Total_Recebido), FORMAT_VALUE)
  If (Total_Pagar - Total_Recebido) >= 0 Then
    Label_Receber.Caption = "A Receber"
  Else
    Label_Receber.Caption = "TROCO"
    L_Receber.Text = Format(CStr(Abs(Total_Pagar - Total_Recebido)), FORMAT_VALUE)
    '07/01/2004 - Daniel
    'Alimentando a var g_dblTrocoFrmRec
    frmRecebimento.g_dblTrocoFrmRec = (CDbl(L_Receber.Text))
    m_blnOcorreTroco = True
  End If
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  nReciboVALOR = IIf(IsNumeric(L_Tot_Pagar.Text), L_Tot_Pagar.Text, 0)
  nReciboACRESCIMO = 0
  nReciboDESCONTO = 0
  
  Select Case Tool.Name
    Case "miConsClientes"
      Dim objFrmPesquisaCliFor As frmPesquisaCliFor
      Set objFrmPesquisaCliFor = New frmPesquisaCliFor
      objFrmPesquisaCliFor.iOrigemVendaRapida = True
      objFrmPesquisaCliFor.Show
    Case "miOpLeitorOtico"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigVR", "Scanner", Tool.Checked)
    Case "miOpClearAfterVenda"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigVR", "Limpar Tela Automatico", Tool.Checked)
    Case "miOpEtiquetas"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigVR", "Etiqueta Balanca", Tool.Checked)
    Case "miOpFreezeVendedor"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigVR", "Mantem Vendedor", Tool.Checked)
    Case "miOpFindVenda"
      Call FindVenda
    Case "miOpCadastraCliente"
      Call CadastraCliente
    Case "mnuProdutosFavoritos"
        frmProdutosFavoritos.Show 1
    Case "miOpInfoCliente"
      
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          Call InfoCliente
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      '***************************************
    Case "miOpUndoMovimentacao"
      Call UndoMovimento
    Case "miEmisRecibo"
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          Call EmiteRecibo
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      '***************************************
    Case "miEmisFatura"
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          Call EmiteFatura
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      '***************************************
    Case "miEmisFaturaParcelados"
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          Call EmiteFaturaParcelados
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      '***************************************
    Case "miEmisCarnes"
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          Call EmiteCarnes
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      '***************************************
    Case "miCarneTp1"
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          Call EmiteCarnesNOVOS(0, Nome_Cliente.Caption)
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      '***************************************
    Case "miEmisBoletos"
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          Call EmiteBoletos
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      '***************************************
    Case "miEmisTicketRel"
      Call EmisTicketRel
    Case "miComplConsultaProdutos"
      nChamaConsulta = 1
      Call ConsultaProduto
    Case "miOpProdValorXQtde"
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          Call IncluirProdutoXValor
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      '***************************************
    
    '30/01/2009 - mpdea
    Case "miEnviarEmail"
      '***************************************
      'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
      If gbUsuarioAcessoApenasTelaVendaRapida = False Then
          ImprimirTicket True
      Else
          MsgBox "Sem acesso", vbInformation, "Atenção"
      End If
      
    Case "carneRapido"
      Call EmiteCarnesNOVOS(1, Nome_Cliente.Caption)
    
    Case "carneRapidoRecibo"
      Call EmiteCarnesNOVOS(2, Nome_Cliente.Caption)
    
      '***************************************
    '17/01/2006 - mpdea
    'Menu de saída para a tela estilo CheckOut
    Case "miSair"
      DoEvents
      Unload Me
  End Select
End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  Cancel = True
End Sub

'30/04/2003 - mpdea
'Dividido rotina em RealizaDescontoSubTotal e RealizaDescontoRateado
Private Sub B_Desconto_Click()
  
  '02/06/2005 - Daniel
  'Criado rotina para verificar se existe um ou mais produtos que não permitem desconto para usuários
  'que não estão habilitados para conceder desconto para produtos configurados para não conceder descontos
  If Len(Nome_Vendedor.Caption) > 0 Then
    If Not UserSemPermissao(Combo_Vendedor.Text) Then
      If ValidarDesconto Then Exit Sub
    End If
  End If
  
  'Case: BIC Amazônia
  'Caso seja BIC chamaremos a tela de venda para funcionários
  'Nela identificaremos se a venda é para funcionário e quem é
  'o funcionário que está comprando
  '
  '17/05/2004 - Daniel
  'Adicionamos rotina no botão de desconto
  If m_blnBic Then
    frmVendaParaFuncionario.Show vbModal
    
    If g_blnRetornar Then
      'Alteramos para False e saímos da rotina
      g_blnRetornar = False
      Exit Sub
    End If
  
  End If
  '-----------------------------------------------------------
  
  If m_blnDescontoRateado Then
    Call RealizaDescontoRateado
  Else
    Call RealizaDescontoSubTotal
  End If
  
End Sub


'23/09/2002 - mpdea
'Totalmente reformulado para suportar o desconto no sub total
'
'O Desconto no SubTotal deverá ser a última operação antes do recebimento,
'portanto não poderá haver venda de mais itens após o desconto
'
Private Sub RealizaDescontoSubTotal()
  Dim sngMaxDescPerc As Single
  Dim curDesconto As Currency
  Dim curNewTotal As Currency
  Dim blnHasItem As Boolean
  Dim intX As Integer
  Dim strSQL As String
  Dim intRet As Integer
  
  
  '29/11/2002 - mpdea
  'Ajustes da Base de ICM
  Dim dblBaseICM As Double
  Dim dblValorICM As Double
  Dim sngDescPerc As Single
  
  '03/09/2003 - mpdea
  'Ajustes de IPI
  Dim dblValorIPI As Double
  
  
  Call StatusMsg("")
  
  'Operação
  If rsOp_Saída.NoMatch Then
    DisplayMsg "Operação não selecionada ou incorreta."
    Exit Sub
  End If
  
  If Not rsOp_Saída.Fields("Dinheiro").Value Then
    DisplayMsg "Tipo de operação não movimenta dinheiro para utilizar esta função."
    Exit Sub
  End If
  
  For intX = 0 To (Grade1.Rows - 1)
    If Tabe(intX).Código <> "0" And Tabe(intX).Código <> "" Then
      blnHasItem = True
    End If
  Next intX

  If Not blnHasItem Then
    DisplayMsg "Não existe nenhum produto digitado, impossível fornecer desconto."
    Exit Sub
  End If
  
  If Total_Pagar = 0 Then
    DisplayMsg "Total igual a zero, impossível fornecer desconto."
    Exit Sub
  End If
  
  If Efetivada.Visible Then
    DisplayMsg "Movimentação já efetivada."
    Exit Sub
  End If
  
  
  '20/11/2002 - mpdea
  'Verificação de desconto já concedido
  If Not IsNull(Num_Registro) Then
    If mcurDescontoSubTotal > 0 Then
      DisplayMsg "Desconto no SubTotal já concedido."
      Exit Sub
    End If
  End If
  
  
  'Percentual de desconto para o funcionário / Filial
  'rsFuncionarios.Index = "Código"
  'rsFuncionarios.Seek "=", Val(rsSaidas("PrestadorServico"))
  sngMaxDescPerc = IIf(rsFuncionarios("nPercDesconto") = 0, _
    rsParametros("VR Desconto"), rsFuncionarios("nPercDesconto"))
    
  'Exibe o form de desconto
  '06/11/2007 - Anderson
  'Verifica se os produtos devem ser somados a nota
  'If frmDesconto.Start(CCur(Total_Pagar), sngMaxDescPerc, curDesconto, curNewTotal, False) Then
  If frmDesconto.Start(IIf(rsOp_Saída("SomarProdutosTotalNota"), CCur(Total_Pagar), 0), sngMaxDescPerc, curDesconto, curNewTotal, False) Then
    
    '03/09/2003 - mpdea
    'Incluído IPI
    '
    '29/11/2002 - mpdea
    'Armazena temporariamente valores de ICM (normal)
    dblBaseICM = Total_Base_ICM
    dblValorICM = Total_ICM
    dblValorIPI = CSng("0" & L_Tot_IPI.Text)
    
    
    '03/09/2003 - mpdea
    'Removido formatação do percentual de desconto
    'ocasionava erro de arredondamento
    '
    'Desconto concedido em percentual
    sngDescPerc = CSng(curDesconto / Total_Pagar)
    
    
    'Atualiza valores
    Total_Pagar = CDbl(curNewTotal)
    mcurDescontoSubTotal = curDesconto
    
    'Atualiza exibição
    txtSubTotal.Text = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
    txtDescSubTotal.Text = Format(mcurDescontoSubTotal, FORMAT_VALUE)
    L_Tot_Pagar.Text = Format(Total_Pagar, FORMAT_VALUE)
    
    
    '03/09/2003 - mpdea
    'Incluído IPI
    '
    '29/11/2002 - mpdea
    'Atualiza valores de ICM
    Total_Base_ICM = Format(dblBaseICM * (1 - sngDescPerc), FORMAT_VALUE)
    Total_ICM = Format(dblValorICM * (1 - sngDescPerc), FORMAT_VALUE)
    L_Tot_IPI.Text = Format(dblValorIPI * (1 - sngDescPerc), FORMAT_VALUE)
    
    
    'Atualiza registro
    intRet = Grava_Venda
    
    'Verifica erro
    If intRet <> 0 Or mblnForceUpdate Then
      'Ativa flag para forçar nova atualização de registro
      mblnForceUpdate = True
      
      'Cancela desconto
      Total_Pagar = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
      mcurDescontoSubTotal = 0
      
      'Atualiza exibição
      txtSubTotal.Text = Format(Total_Pagar, FORMAT_VALUE)
      txtDescSubTotal.Text = Format(0, FORMAT_VALUE)
      L_Tot_Pagar.Text = Format(Total_Pagar, FORMAT_VALUE)
      
      
      '03/09/2003 - mpdea
      'Incluído IPI
      '
      '28/11/2002 - mpdea
      'Restaura valores de ICM
      Total_Base_ICM = dblBaseICM
      Total_ICM = dblValorICM
      Total_IPI = dblValorIPI
      
      Exit Sub
    End If

    '******************************************************
    ' DEZEMBRO/2019
    ' Coloquei esta condição aqui também...pois caso o operador logado deseje realizar o desconto, mas
    ' ele não possui permissão para realizar o recebimento (parametrizado no cadastro de usuário - tem um checkbox)
    ' então ele realiza o desconto e volta para a tela de Venda Rápida
    rsUsuarios.Index = "Código"
    rsUsuarios.Seek "=", Cod_Operador.Caption
    If rsUsuarios.NoMatch Then
        MsgBox ("Operador não encontrado.")
        Exit Sub
    End If
    If rsUsuarios("Recebimento") = False Then
        Beep
        DisplayMsg "Desconto concedido com sucesso!" & vbCrLf & vbCrLf & "Retornaremos agora para a tela de Venda Rápida."
        
        Exit Sub
    End If
    '******************************************************
    '
    
    'Realiza recebimento
    Call B_Recebe_Click
    
    
    '08/11/2002 - mpdea
    'Incluído verificação do número da movimentação
    'Necessário caso esteja ativado a opção de limpar a tela após concluir a venda
    
    'Verifica confirmação do recebimento
    'Caso contrário restaura valores anteriores ao desconto
    If Not Efetivada.Visible And Número.Text <> "" Then
      'Ativa flag para forçar nova atualização de registro
      mblnForceUpdate = True
      
      'Atualiza valores
      Total_Pagar = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
      mcurDescontoSubTotal = 0
      
      
      '--------------------------------------------------------------------------
      '03/09/2003 - mpdea
      'Restaura valores do registro para os campos: Base ICM, Valor ICM e IPI
      '
      '07/11/2002 - mpdea
      'Corrigido argumento de valor para a string SQL (RT-3144)
      '
      'Restaura valores do registro gravado
      strSQL = "UPDATE Saídas SET DescontoSubTotal = 0, Total = " & _
        Replace(Total_Pagar, ",", ".") & _
        ", [Base ICM] = " & Replace(dblBaseICM, ",", ".") & _
        ", [Valor ICM] = " & Replace(dblValorICM, ",", ".") & _
        ", IPI = " & Replace(dblValorIPI, ",", ".") & _
        " WHERE Filial = " & gnCodFilial & " AND Sequência = " & CLng(Número.Text)
      db.Execute strSQL, dbFailOnError
      '--------------------------------------------------------------------------
      
      
      'Atualiza exibição
      txtSubTotal.Text = Format(Total_Pagar, FORMAT_VALUE)
      txtDescSubTotal.Text = Format(0, FORMAT_VALUE)
      L_Tot_Pagar.Text = Format(Total_Pagar, FORMAT_VALUE)
      
      
      '03/09/2003 - mpdea
      'Incluído IPI
      '
      '28/11/2002 - mpdea
      'Restaura valores de ICM
      Total_Base_ICM = dblBaseICM
      Total_ICM = dblValorICM
      Total_IPI = dblValorIPI
      
      
      Exit Sub
    End If
    
    'Desativa flag, operação completada com sucesso
    mblnForceUpdate = False
    
  End If
  
End Sub

Private Sub RealizaDescontoRateado()

  Dim Conta As Integer, i As Integer
  Dim Desc_Max As Double
  Dim Desc As Double
  Dim Desc_Perc As Double
  Dim Novo_Total As Double
  Dim Tot_IPI As Double
  Dim F As Form
  Dim nPercMaxDesc As Single
  '23/04/2009 - mpdea
  Dim str_format_casas_decimais As String


  Call StatusMsg("")

  Conta = 0
  For i = 0 To (Grade1.Rows - 1)
    If Tabe(i).Código <> "0" And Tabe(i).Código <> "" Then
      Conta = Conta + 1
      Exit For
    End If
  Next i

  If Conta = 0 Then
    DisplayMsg "Não existe nenhum produto digitado, impossível fornecer desconto."
    Exit Sub
  End If

  If Efetivada.Visible = True Then
    DisplayMsg "Movimentação já efetivada."
    Exit Sub
  End If


  '07/05/2003 - mpdea
  'Adaptações para o desconto rateado
  '---------------------------------------------------------------------------------
  'Percentual de desconto para o funcionário / Filial
'  Desc_Max = Total_Pagar * rsParametros("VR Desconto") / 100
  nPercMaxDesc = IIf(rsFuncionarios("nPercDesconto") = 0, _
    rsParametros("VR Desconto"), rsFuncionarios("nPercDesconto"))
  
  
  'Exibe o form de desconto
  '06/11/2007 - Anderson
  'Verifica se os produtos devem ser somados a nota
  'If Not frmDesconto.Start(CCur(Total_Pagar), nPercMaxDesc, _
  '                         0, 0, True, Total_Desconto) Then Exit Sub
  If Not frmDesconto.Start(IIf(rsOp_Saída("SomarProdutosTotalNota"), CCur(Total_Pagar), 0), nPercMaxDesc, 0, 0, True, IIf(rsOp_Saída("SomarProdutosTotalNota"), Total_Desconto, 0)) Then Exit Sub

  
'  Set F = New frmDesconto
'  F.Desc_Fornecido.Caption = Format(Total_Desconto, "###,###,##0.00")
'  F.Total.Caption = Total_Pagar
'  F.Desconto.Text = ""
'  F.Show vbModal
'  Set F = Nothing
'
'  If gnDesconto = 0# Then Exit Sub
  '---------------------------------------------------------------------------------

  B_Recebe.Enabled = False
  
  '23/04/2009 - mpdea
  'Formatação do preço unitário
  If g_bln3CasasDecimais Then
    str_format_casas_decimais = "000"
  ElseIf g_bln5CasasDecimais Then
    str_format_casas_decimais = "00000"
  Else
    str_format_casas_decimais = "00"
  End If

  Desc_Max = (Total_Pagar + Total_Desconto) * nPercMaxDesc / 100
  If (Desc_Max + 0.1) < (Total_Desconto + gnDesconto) Then
    '29/08/2003 - mpdea
    'Incluído apelido do vendedor na mensagem
    DisplayMsg "Desconto superior ao permitido para o vendedor " & _
      rsFuncionarios.Fields("Apelido").Value & ""
    Exit Sub
  End If

  Total_Desconto = Total_Desconto + gnDesconto
  'Adicionado para manter o total em desconto no Total Geral
  gcDescInTotal = gcDescInTotal + gnDesconto

  '23/04/2009 - mpdea
  'Modificado para que o índice de desconto não seja formatado
  'Desc_Perc = Format(gnDesconto / Total_Pagar, FORMAT_VALUE)
  Desc_Perc = gnDesconto / Total_Pagar
  Desc_Perc = 1 - Desc_Perc
  Novo_Total = Total_Pagar - gnDesconto

  For i = 0 To (Grade1.Rows - 1)
    '02/06/2005 - Daniel
    'Adicionado: And Tabe(i).Código <> "0"
    If Tabe(i).Código <> "" And Tabe(i).Código <> "0" Then
      '23/04/2009 - mpdea
      'Modificado para que o cálculo do preço seja formatado de acordo com as casas decimais de preço
      'Tabe(i).Preço = Format((Tabe(i).Preço * Desc_Perc), "###########.00")
      'Tabe(i).Preço = Format((Tabe(i).Preço * Desc_Perc), "#0." & str_format_casas_decimais)
      Tabe(i).Preço_Final = Format((Tabe(i).Qtde * (Tabe(i).Preço - (Tabe(i).Preço * Tabe(i).Desconto / 100))), "#0." & str_format_casas_decimais)
      Tot_IPI = Tabe(i).Preço_Final * Tabe(i).IPI / 100
      Tot_IPI = Format(Tot_IPI, "#0.00")
      Tabe(i).Preço_Final = Format(Tabe(i).Preço_Final + Tot_IPI, "#0." & str_format_casas_decimais)
    End If
  Next i

  Call Recalcula

  If Total_Pagar <> Novo_Total Then
    Desc = Total_Pagar - Novo_Total
    For i = 0 To (Grade1.Rows - 1)
      '23/04/2009 - mpdea
      'Adicionado: And Tabe(i).Código <> "0"
      If Tabe(i).Código <> "" And Tabe(i).Código <> "0" Then
        If Tabe(i).Qtde = 1 Then
          '23/04/2009 - mpdea
          'Modificado para que o cálculo do preço seja formatado de acordo com as casas decimais de preço
          'Tabe(i).Preço = Format((Tabe(i).Preço - Desc), "###########.00")
          Tabe(i).Preço = Format((Tabe(i).Preço - Desc), "#0." & str_format_casas_decimais)
          Tabe(i).Preço_Final = Tabe(i).Qtde * Tabe(i).Preço
          Desc = 0
        End If
      End If
    Next i
    Call Recalcula
  End If
  
  '23/04/2009 - mpdea
  'Ajusta desconto caso haja resíduo
  gcDescInTotal = Format(gcDescInTotal - Desc, FORMAT_VALUE)
  Total_Desconto = Format(Total_Desconto - Desc, FORMAT_VALUE)

  Grade1.MoveLast
  Grade1.MoveFirst
  
End Sub

Function Acha_Grade(Aux As String) As String

  rsProdutos.Index = "Código"
  rsProdutos.Seek "=", Aux
  If Not rsProdutos.NoMatch Then
    Acha_Grade = Aux
    Exit Function
  End If
  If rsProdutos.NoMatch Then 'procura o código na grade
    rsGrade.Index = "Código"
    rsGrade.Seek "=", Aux
    If rsGrade.NoMatch Then
        Acha_Grade = Aux
        Exit Function
    End If
    Acha_Grade = rsGrade("Código Original")
    Exit Function
  End If

End Function


Private Sub B_Grava_Click()


  If cboCaixa.Text = "" Then
    MsgBox "Informe o caixa que será utilizado nesta venda.", vbInformation, "Atenção"
    cboCaixa.SetFocus
    Exit Sub
  End If


  '01/07/2004 - Daniel
  'Case: Coneg Campos
  'Caso o usuário Gravou libera a limpeza da tela
  m_blnClear = True
  '----------------------------------------------
  
  '08/01/2004 - Daniel
  m_sngQtdeTotal = 0
  '-------------------
  
  '21/05/2004 - Daniel
  'Case: BIC Amazônia
  'Caso seja BIC chamaremos a tela de venda para funcionários
  'Nela identificaremos se a venda é para funcionário e quem é
  'o funcionário que está comprando
  If m_blnBic Then
    frmVendaParaFuncionario.Show vbModal
    
    If g_blnRetornar Then
      'Alteramos para False e saímos da rotina
      g_blnRetornar = False
      Exit Sub
    End If
  
  End If
  '-----------------------------------------------------------
  
  If Grava_Venda = 0 Then
      '25/06/2013-Alexandre Afornali
    If (Número.Text <> "") Then
      Call UpdateTotalNCM
      Call GravarComanda
    End If
    If Frame_Recebimento.Visible Then
      Frame_Recebimento.Enabled = True
      B_Recebe_Simples.Visible = True
      Lançar_Débito.Enabled = True
      If Not rsCliFor("Tem Conta") Then
        Lançar_Débito.Enabled = False
      End If
    End If
    L_Preço.Caption = ""
    L_Estoque.Caption = ""
  End If
  
End Sub

Private Sub B_Grava_Recebe_Click()
On Error GoTo Erro

  If cboCaixa.Text = "" Then
    MsgBox "Informe o caixa que será utilizado nesta venda.", vbInformation, "Atenção"
    cboCaixa.SetFocus
    Exit Sub
  End If

  '01/07/2004 - Daniel
  'Case: Coneg Campos
  'Caso o usuário Gravou libera a limpeza da tela
  m_blnClear = True
  '----------------------------------------------
  
  '08/01/2004 - Daniel
  m_sngQtdeTotal = 0
  '-------------------
  
  ' ************************************************
  ' PROGRAMA FIDELIDADE
  If gParticipaProgramaFidelidade = 1 Then
    '1-SIM PARTICIPA;
    '0-NÃO PARTICIPA Empresa/filial;
    If gClienteEntregouResgatePontos = True And gSaldoCdGuidResgate_clicou_ok_telaDesconto = False Then
      B_Desconto_Click
      Exit Sub
    End If
  End If
  ' ************************************************
        
  
  '21/05/2004 - Daniel
  'Case: BIC Amazônia
  'Caso seja BIC chamaremos a tela de venda para funcionários
  'Nela identificaremos se a venda é para funcionário e quem é
  'o funcionário que está comprando
  If m_blnBic Then
    frmVendaParaFuncionario.Show vbModal
    
    If g_blnRetornar Then
      'Alteramos para False e saímos da rotina
      g_blnRetornar = False
    
      Exit Sub
    End If
  
  End If
  '-----------------------------------------------------------
  
  Call StatusMsg("")
  
  If Grava_Venda = 0 Then
    Call B_Recebe_Click
  End If

  Exit Sub
Erro:
  MsgBox "Erro " + Err.Number + " " + Err.Description, vbInformation, "Atenção"

End Sub

Private Sub B_Limpa_Click()
  txtComanda.Text = ""
  btnComandaVendas.Visible = False
  txtComanda.Width = 1785
  
  Observacao.Text = ""
  Refresh
  
  Call Limpa_Tela(1)
  '24/06/2004 - Daniel
  'Criado rotina de validação para checar se o user tem permissão ou
  'não de limpar os campos. Solicitado pelo cliente Coneg Campos e
  'aproveitado para os demais
  Dim rstFuncionarios As Recordset
  Dim strQuery        As String
  Dim blnPermissao    As Boolean
  
  
  frm_produtoSemPrecoNaGrade.Visible = False
  bProdutoSemPrecoNaGrade = False
  
  
  blnPermissao = True
  
  strQuery = "SELECT Código, Nome, SenhaClear "
  strQuery = strQuery & " FROM Funcionários "
  strQuery = strQuery & " WHERE Funcionários.Código = " & CInt(Trim(Cod_Operador.Caption))
    
  Set rstFuncionarios = db.OpenRecordset(strQuery, dbOpenDynaset)

  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If Not .Fields("SenhaClear").Value Then
        blnPermissao = False
      End If
      
    End If
    .Close
  End With
  
  Set rstFuncionarios = Nothing
  
  If Not blnPermissao And m_blnClear = False Then 'm_blnClear = False o user não clicou nos botões de gravar ou gravar e receber
  
    If Not frmGerente.gbSenhaGerente Then
      Exit Sub
    Else
      Limpa_Tela (0)
      If Frame_Recebimento.Visible = True Then
        Frame_Recebimento.Enabled = False
        '25/03/2004 - Daniel
        If m_blnUserDanger = True Then
           B_Recebe_Simples.Visible = True
        Else
           B_Recebe_Simples.Visible = False
        End If
      End If
    End If
  Else
    Limpa_Tela (0)
    If Frame_Recebimento.Visible = True Then
      Frame_Recebimento.Enabled = False
      
      '25/03/2004 - Daniel
      If m_blnUserDanger = True Then
         B_Recebe_Simples.Visible = True
      Else
         B_Recebe_Simples.Visible = False
      End If
      
    End If
    
    m_blnClear = False '01/07/2004 - Daniel - Seto para False novamente
    txtComanda.Text = ""
  End If
  
  'Variaveis globais do programa fidelidade
  gSaldoCdGuidResgate_clicou_ok_telaDesconto = False
  gCdGuidResgate = ""
  gSaldoCdGuidResgate = 0
  gCdClienteCdGuidResgate = 0
  gNmClienteCdGuidResgate = ""
  gClienteEntregouResgatePontos = False
  lbl_retornoEnvioNFCe.Visible = False
  
End Sub

Private Sub B_NFC_e_Click()
  If Número.Text = "" Then
    DisplayMsg "NFC-e só pode ser emitido a partir de uma venda efetivada. Encontre uma venda efetivada."
    Exit Sub
  End If
  
  DoEvents
  
  ' PILATTI MAIO/2018 busca desconto no DB para a venda
  Dim sDescontoDaVenda As String
  sDescontoDaVenda = ""
  Dim strSQL_buscaDesconto As String
  Dim rsSaidasDesconto As Recordset
  strSQL_buscaDesconto = "SELECT Desconto FROM [Saídas] WHERE Sequência = " & Número.Text
  Set rsSaidasDesconto = db.OpenRecordset(strSQL_buscaDesconto, dbOpenSnapshot)
  With rsSaidasDesconto
    If .RecordCount > 0 Then
      sDescontoDaVenda = .Fields("desconto").Value
    End If
    .Close
  End With
  Set rsSaidasDesconto = Nothing
  
  If sDescontoDaVenda <> "" And sDescontoDaVenda <> "0" And gcDescInTotal = 0 Then
    gcDescInTotal = CCur(sDescontoDaVenda)
  End If
  '
    
  Dim EnviaNFCe As New clsNFCe
  Dim bRetEnviaNFCE As Boolean

  EnviaNFCe.EnviaNFCe (Número.Text), gcDescInTotal
  
  If sRetornoEnvioNFCe <> "" Then
      lbl_retornoEnvioNFCe.Visible = True
  Else
      lbl_retornoEnvioNFCe.Visible = False
  End If
  sRetornoEnvioNFCe = ""
End Sub

'04/05/2004 - mpdea
'Corrigido e otimizado o código em geral
'
'08/04/2003 - mpdea
'Implementado tratamento de erro
''Private Sub B_Nota_Click()
''  Dim frmX As Form
''
''  Dim rsTempOpSaidas As Recordset
''  Dim strSQL As String
''  Dim blnExit As Boolean
''  Dim blnShowObs As Boolean
''  Dim intX As Integer
''
''  Dim strFileNF As String
''  Dim intRet As Integer
''  Dim lngNotaFiscal As Long
''  Dim blnInTransaction As Boolean
''  Dim intRepeatUpdateLocked As Integer
''
''
''  On Error GoTo ErrHandler
''
''
''  Call StatusMsg("")
''
''  If Número.Text = "" Then
''    DisplayMsg "Ache ou grave uma venda antes."
''    Exit Sub
''  End If
''
''  If rsSaidas.Fields("Nota Cancelada").Value Then
''    DisplayMsg "Esta nota está cancelada e não pode ser reimpressa."
''    Exit Sub
''  End If
''
''  '04/12/2007 - Anderson
''  'Verifica permissão para imprimir nota somente em movimentações efetivadas
''  If rsParametros.Fields("ImprimeNotaMovEfetivada").Value Then
''    If Not rsSaidas.Fields("Efetivada").Value Then
''      DisplayMsg "Movimentação não efetivada. Não é possível imprimir a nota fiscal."
''      Exit Sub
''    End If
''  End If
''
''  'Verificações referente a operação de Saída
''  strSQL = "SELECT * FROM [Operações Saída] WHERE Código = " & rsSaidas.Fields("Operação").Value
''  Set rsTempOpSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
''  With rsTempOpSaidas
''    If .RecordCount > 0 Then
''      If Not .Fields("Nota").Value Then
''        DisplayMsg "Operação não permite Nota Fiscal."
''        blnExit = True
''      End If
''      blnShowObs = .Fields("InTelaObsTransp").Value
''    Else
''      DisplayMsg "Operação de Saída não encontrada."
''      blnExit = True
''    End If
''    .Close
''  End With
''  Set rsTempOpSaidas = Nothing
''  If blnExit Then Exit Sub
''
'''  Call RecalculaPesos
''
''  If blnShowObs Then
''    Set frmX = New frmObsNota
''    frmX.gsCliente = rsCliFor.Fields("Transportadora").Value
''    frmX.lngSequencia = rsSaidas.Fields("Sequência").Value
''    frmX.bytTipoTabela = 1
''    frmX.Show vbModal
''    Set frmX = Nothing
''    If gsRetornoDoc <> "OK" Then
''      StatusMsg "Nota não impressa."
''      Exit Sub
''    End If
''  Else
''    For intX = 0 To 7
''      gsObsDoc(intX) = ""
''    Next intX
''    gsPlaca = ""
''    gsUfrmPlaca = ""
''    gsQtdeTrans = ""
''    gsMarcaTrans = ""
''    gsEspecieTrans = ""
''    gsPesoBruto = ""
''    gsPesoLiquido = ""
''    gsTransportadora = ""
''  End If
''
''  Call IsDataType(dtLong, rsSaidas.Fields("Nota Impressa").Value, lngNotaFiscal)
''  If lngNotaFiscal <> 0 Then
''    If MsgBox("A Nota fiscal já foi impressa, deseja imprimir novamente?", _
''      vbQuestion + vbYesNo + vbDefaultButton2, "Atenção") = vbNo Then
''      Exit Sub
''    End If
''  End If
''
''
''  '--------------------------------------------------------------------------
''  'Grava nova NF
''  '--------------------------------------------------------------------------
''  If lngNotaFiscal = 0 Then
''    'Modificado leitura e gravação do número da última nota fiscal
''    'Incluído transação durante gravação
''    ws.BeginTrans
''    blnInTransaction = True
''    '
''    lngNotaFiscal = g_lngNextNotaFiscal(rsSaidas.Fields("Filial").Value)  ' rsParametros.Fields("Última Nota").Value + 1
''    '
'''    With rsParametros
'''      .Edit
'''      .Fields("Última Nota").Value = lngNotaFiscal
'''      .Update
'''    End With
''    '
''    With rsSaidas
''      .Edit
''      .Fields("Nota Impressa").Value = lngNotaFiscal
''      'Gravação dos campos de observações na tela de saídas
''      'For intX = 0 To 7
''      '  .Fields("obs_Obs" & intX + 1).Value = gsObsDoc(intX)
''      'Next intX
''      For intX = 0 To 1
''        .Fields("obs_infCpl" & intX + 1).Value = gsObsDoc(intX)
''      Next intX
''      .Fields("obs_Transportadora") = gsTransportadora
''      .Fields("obs_Placa") = gsPlaca
''      .Fields("obs_Uf") = gsUfrmPlaca
''      .Fields("obs_Especie") = gsEspecieTrans
''      .Fields("obs_Qtde") = gsQtdeTrans
''      .Fields("obs_Marca") = gsMarcaTrans
''      .Fields("obs_PesoBruto") = IIf(IsNumeric(gsPesoBruto), gsPesoBruto, 0)
''      .Fields("obs_PesoLiquido") = IIf(IsNumeric(gsPesoLiquido), gsPesoLiquido, 0)
''      .Fields("obs_FretePago") = IIf(IsNumeric(gsFretePago), gsFretePago, 0)
''      .Update
''    End With
''
''    '---------------------
''
''    '05/05/2005 - mpdea
''    'Atualiza a Nota Fiscal e Fatura do Contas a Receber
''    Call StatusMsg("Verificando e atualizando contas a receber...")
''    '
''    strSQL = "UPDATE [Contas a Receber] SET Nota = " & lngNotaFiscal
''    strSQL = strSQL & ", Fatura = '" & lngNotaFiscal & "/ ' & Parcela"
''    strSQL = strSQL & " WHERE Tipo = 'R'"
''    strSQL = strSQL & " AND Filial = " & rsSaidas.Fields("Filial").Value
''    strSQL = strSQL & " AND Sequência = " & rsSaidas.Fields("Sequência").Value
''    '
''    db.Execute strSQL, dbFailOnError
''
''    '10/09/2007 - Anderson
''    'Gera arquivo log do sistema
''    If g_bolSystemLog Then
''      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
''      strSQL, _
''      "frmVendaRap2_B_Nota_Click", _
''      "Contas a Receber", g_strArquivoSystemLog
''    End If
''
''
'''    Rem Acha as contas a pagar e atualiza os campos nota e fatura
'''    Call StatusMsg("Verificando e atualizando contas a receber...")
'''
''''    Aux_Data = CDate("01/01/1980")
'''    Aux_Int = 1
'''    Aux_Conta = 0
'''    rsContas_Receber.Index = "Cliente"
''''    Erro = False
'''Lp1_Receber:
'''    rsContas_Receber.Seek ">", "R", rsSaidas("Cliente"), Aux_Conta
'''    If rsContas_Receber.NoMatch Then GoTo Fim_Receber
'''    If rsContas_Receber("Tipo") <> "R" Then GoTo Fim_Receber
'''    If rsContas_Receber("Cliente") <> rsSaidas("Cliente") Then GoTo Fim_Receber
'''    Aux_Conta = rsContas_Receber("Contador")
'''    If rsContas_Receber("Sequência") <> rsSaidas("Sequência") Then GoTo Lp1_Receber
'''    rsContas_Receber.Edit
'''      rsContas_Receber("Nota") = rsSaidas("Nota Impressa")
'''      rsContas_Receber("Fatura") = str(rsSaidas("Nota Impressa")) + "/" + str(Aux_Int)
'''      Aux_Int = Aux_Int + 1
'''    rsContas_Receber.Update
'''    GoTo Lp1_Receber
'''
'''Fim_Receber:
''
''    Call StatusMsg("")
''
''    'Finaliza transação
''    ws.CommitTrans
''    blnInTransaction = False
''  End If
''  '--------------------------------------------------------------------------
''
''
''  '--------------------------------------------------------------------------
''  'Imprime NF
''  '--------------------------------------------------------------------------
''  strFileNF = gsConfigPath + rsParametros.Fields("Nota Saída").Value + ".CNF"
''  intRet = Imprime_Nota(strFileNF, rsSaidas.Fields("Filial").Value, rsSaidas.Fields("Sequência").Value)
''  If intRet = 0 Then
''    '14/04/2003 - mpdea
''    'Atualiza a data da impressão da nota fiscal
''    strSQL = "UPDATE Saídas SET DataEmissaoNota = #"
''    strSQL = strSQL & Format(Date, "mm/dd/yyyy") & "# "
''    strSQL = strSQL & "WHERE Filial = " & rsSaidas.Fields("Filial").Value
''    strSQL = strSQL & " AND Sequência = " & rsSaidas.Fields("Sequência").Value
''    db.Execute strSQL, dbFailOnError
''    '
''    DisplayMsg "Nota [" & lngNotaFiscal & "] impressa com sucesso."
''  Else
''    DisplayMsg "Houve o erro " & intRet & " durante a impressão da Nota."
''  End If
''  '--------------------------------------------------------------------------
''
''  Exit Sub
''
''ErrHandler:
''  Screen.MousePointer = vbDefault
''  Call StatusMsg("")
''  Select Case Err.Number
''    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
''      If intRepeatUpdateLocked < 30 Then
''        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
''        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
''        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
''        Call WaitSeconds(1, False) 'Aguarda um segundo
''        Resume
''      Else
''
''        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
''          intRepeatUpdateLocked = 0
''          Resume
''        Else
''          'Cancelamento da transação
''          If blnInTransaction Then ws.Rollback
''          Exit Sub
''        End If
''
'''        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
'''          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
'''          "uma nova tentativa.", vbExclamation + vbOKCancel, "Saídas - Imprimir Nota Fiscal") = vbOK Then
'''          intRepeatUpdateLocked = 0
'''          Resume
'''        Else
'''          'Cancelamento da transação
'''          If blnInTransaction Then ws.Rollback
'''          Exit Sub
'''        End If
''      End If
''    Case Else
''      'Cancelamento da transação
''      If blnInTransaction Then ws.Rollback
''      'Outros Erros
''      MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
''  End Select
''
''
''
''
'''  Dim Aux As Variant
'''  Dim Nome_Arq As String
'''  Dim Texto As String
'''  Dim Final As Integer
'''  Dim Str_Impre As String
'''  Dim Cód1, Cód2, Cód3 As Integer
'''  Dim Num_cod As Integer
'''  Dim Final_Linha As Integer
'''  Dim Linhas As Integer
'''  Dim Especial2 As Integer
'''  Dim Aux_Data As Date
'''  Dim Aux_Int As Integer
'''  Dim Aux_Conta As Long
'''  Dim Erro As Integer
'''  Dim Resp As Integer
'''  Dim F As Form
'''  Dim strSQL As String
'''  Dim intX As Integer
'''
'''  Dim lngUltimaNotaFiscal As Long
'''  Dim blnInTransaction As Boolean
'''
'''
'''  On Error GoTo ErrHandler
'''
'''
'''  Call StatusMsg("")
'''
'''  Aux = Número.Text
'''  If IsNull(Aux) Or Aux = "" Then
'''    DisplayMsg "Ache ou grave uma venda antes."
'''    Exit Sub
'''  End If
'''
'''  If rsSaidas.Fields("Nota Cancelada").Value Then
'''    DisplayMsg "Esta nota está cancelada e não pode ser reimpressa."
'''    Exit Sub
'''  End If
'''
'''  'Verificações referente a operação de Saída
'''  strSQL = "SELECT * FROM [Operações Saída] WHERE Código = " & rsSaidas.Fields("Operação").Value
'''  Set rsTempOpSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
'''  With rsTempOpSaidas
'''    If .RecordCount > 0 Then
'''      If Not .Fields("Nota").Value Then
'''        DisplayMsg "Operação não permite Nota Fiscal."
'''        blnExit = True
'''      End If
'''      blnShowObs = .Fields("InTelaObsTransp").Value
'''    Else
'''      DisplayMsg "Operação de Saída não encontrada."
'''      blnExit = True
'''    End If
'''    .Close
'''  End With
'''  Set rsTempOpSaidas = Nothing
'''  If blnExit Then Exit Sub
'''
'''  Call RecalculaPesos
'''
'''  If blnShowObs Then
'''    Set F = New frmObsNota
'''    F.gsCliente = rsCliFor.Fields("Transportadora").Value
'''    F.lngSequencia = rsSaidas.Fields("Sequência").Value
'''    F.bytTipoTabela = 1
'''    F.Show vbModal
'''    Set F = Nothing
'''    If gsRetornoDoc <> "OK" Then
'''      StatusMsg "Nota não impressa."
'''      Exit Sub
'''    End If
'''  Else
'''    For intX = 0 To 7
'''      gsObsDoc(intX) = ""
'''    Next intX
'''    gsPlaca = ""
'''    gsUfrmPlaca = ""
'''    gsQtdeTrans = ""
'''    gsMarcaTrans = ""
'''    gsEspecieTrans = ""
'''    gsPesoBruto = ""
'''    gsPesoLiquido = ""
'''    gsTransportadora = ""
'''  End If
'''
'''
'''  '11/08/2003 - maikel
'''  '             Gravação dos campos de observações na tela de saídas
'''  '----------------------------------------------------------------'
'''    rsSaidas.Edit
'''
'''    For intX = 0 To 7
'''      rsSaidas.Fields("obs_Obs" & intX + 1).Value = gsObsDoc(intX)
'''    Next intX
'''
'''    rsSaidas.Fields("obs_Transportadora") = gsTransportadora
'''    rsSaidas.Fields("obs_Placa") = gsPlaca
'''    rsSaidas.Fields("obs_Uf") = gsUfrmPlaca
'''    rsSaidas.Fields("obs_Especie") = gsEspecieTrans
'''    rsSaidas.Fields("obs_Qtde") = gsQtdeTrans
'''    rsSaidas.Fields("obs_Marca") = gsMarcaTrans
'''    rsSaidas.Fields("obs_PesoBruto") = IIf(IsNumeric(gsPesoBruto), gsPesoBruto, 0)
'''    rsSaidas.Fields("obs_PesoLiquido") = IIf(IsNumeric(gsPesoLiquido), gsPesoLiquido, 0)
'''
'''    rsSaidas.Fields("obs_FretePago") = IIf(IsNumeric(gsFretePago), gsFretePago, 0)
'''    rsSaidas.Update
'''  '----------------------------------------------------------------'
'''
'''  Rem pega próxima nota e grava no arquivo
'''  Aux = rsSaidas("Nota Impressa")
'''  If IsNull(Aux) Then Aux = 0
'''
'''  If Aux <> 0 Then
'''    If MsgBox("A Nota fiscal já foi impressa, deseja imprimir novamente?", _
'''      vbQuestion + vbYesNo + vbDefaultButton2, "Atenção") = vbNo Then
'''      Exit Sub
'''    End If
'''  End If
'''
'''
'''
'''  If Not IsNumeric(Aux) Then Aux = 0
'''  If Val(Aux) = 0 Then
'''
'''    '-------------------------------------------------------------------
'''    '28/11/2003 - mpdea
'''    'Modificado leitura e gravação do número da última nota fiscal
'''    'Incluído transação durante gravação
'''    lngUltimaNotaFiscal = rsParametros.Fields("Última Nota").Value + 1
'''    '
'''    ws.BeginTrans
'''    blnInTransaction = True
'''    '
'''    With rsParametros
'''      .Edit
'''      .Fields("Última Nota").Value = lngUltimaNotaFiscal
'''      .Update
'''    End With
'''    '
'''    With rsSaidas
'''      .Edit
'''      .Fields("Nota Impressa").Value = lngUltimaNotaFiscal
'''      .Update
'''    End With
'''    '
'''    ws.CommitTrans
'''    blnInTransaction = False
'''    '
''''    rsParametros.Edit
''''      rsParametros("Última Nota") = rsParametros("Última Nota") + 1
''''    rsParametros.Update
''''    rsSaidas.Edit
''''      rsSaidas("Nota Impressa") = rsParametros("Última Nota")
''''    rsSaidas.Update
'''    '-------------------------------------------------------------------
'''
'''
'''    Rem Acha as contas a pagar e atualiza os campos nota e fatura
'''    Call StatusMsg("Verificando e atualizando contas a receber...")
'''    Aux_Data = CDate("01/01/1980")
'''    Aux_Int = 1
'''    Aux_Conta = 0
'''    rsContas_Receber.Index = "Cliente"
'''    Erro = False
'''Lp1_Receber:
'''    rsContas_Receber.Seek ">", "R", rsSaidas("Cliente"), Aux_Conta
'''    If rsContas_Receber.NoMatch Then GoTo Fim_Receber
'''    If rsContas_Receber("Tipo") <> "R" Then GoTo Fim_Receber
'''    If rsContas_Receber("Cliente") <> rsSaidas("Cliente") Then GoTo Fim_Receber
'''    Aux_Conta = rsContas_Receber("Contador")
'''    If rsContas_Receber("Sequência") <> rsSaidas("Sequência") Then GoTo Lp1_Receber
'''    rsContas_Receber.Edit
'''      rsContas_Receber("Nota") = rsSaidas("Nota Impressa")
'''      rsContas_Receber("Fatura") = str(rsSaidas("Nota Impressa")) + "/" + str(Aux_Int)
'''      Aux_Int = Aux_Int + 1
'''    rsContas_Receber.Update
'''    GoTo Lp1_Receber
'''
'''
'''Fim_Receber:
'''    Call StatusMsg("")
'''  End If
'''
'''
'''  Rem Pegar o nome do arquivo de configuração
'''  Nome_Arq = gsConfigPath & rsParametros("Nota Saída") + ".CNF"
'''
'''
'''  Resp = Imprime_Nota(Nome_Arq, rsSaidas("Filial"), rsSaidas("Sequência"))
'''
'''  If Resp = 0 Then
'''    DisplayMsg "Nota impressa com sucesso."
'''  Else
'''    DisplayMsg "Houve o erro " + CStr(Resp) + " durante a impressão da nota."
'''    Exit Sub
'''  End If
'''
'''
'''  '14/04/2003 - mpdea
'''  'Atualiza a data da impressão da nota fiscal
'''  strSQL = "UPDATE Saídas SET DataEmissaoNota = #"
'''  strSQL = strSQL & Format(Date, "mm/dd/yyyy") & "# "
'''  strSQL = strSQL & "WHERE Filial = " & rsSaidas.Fields("Filial").Value
'''  strSQL = strSQL & " AND Sequência = " & rsSaidas.Fields("Sequência").Value
'''  db.Execute strSQL, dbFailOnError
'''
'''  Exit Sub
'''
'''ErrHandler:
'''  Call StatusMsg("")
'''  If blnInTransaction Then ws.Rollback
'''  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
''End Sub

Private Sub B_programaFidelidade_Click()
  Dim frmX As New frmProgramaFidelidadeResgatePontos_Rapido

  If Combo_Cliente.Text <> "" Then
    frmX.lCodigoCliente = Combo_Cliente.Text
    frmX.vTotalDaVendaEmAndamento = L_Tot_Pagar.Text
    frmX.Show
  Else
    MsgBox "Selecione um cliente", vbInformation, "Atenção"
    Exit Sub
  End If

End Sub

Private Sub B_Recebe_Click()
        Dim nRet As Integer
        Dim nRepeatUpdateLocked As Integer
        
        Dim Ordem As Integer
        Dim Fim As Integer
        Dim Resposta As Integer
        Dim R_Banco As Integer
        Dim R_Cheque As String
        Dim R_Bom As Date
        Dim R_Valor As Double
        Dim Resp As Integer
        Dim Total_Prazo As Double
        Dim Parcelas As Integer
        
        '03/08/2002 - mpdea
        Dim blnInTransaction As Boolean
        
10      On Error GoTo ProcessErr

        If frm_produtoSemPrecoNaGrade.Visible = True Then
            Dim retMsg As Variant
            retMsg = MsgBox("Nesta venda tem produto SEM PREÇO, deseja realizar o recebimento mesmo assim?", vbYesNo, "Atenção")
      
            If retMsg = vbNo Then
                Exit Sub
            End If
        End If
        
        '08/01/2004 - Daniel
20      m_sngQtdeTotal = 0
        '-------------------
        
30      Total_Prazo = 0

       
        '07/06/2005 - Daniel
        'Solicitação...: PSI Ayrton
        'Finalidade....: Ocorre validação do caixa antes do recebimento também;
        'Foi notificado de gravações com caixa inválido na empresa Casagrande Armarinhos
40      If Len(Nome_Caixa.Caption) <= 0 Then
50        MsgBox "Caixa inválido, verifique.", vbExclamation, "Quick Store"
60        cboCaixa.SetFocus
70        Exit Sub
80      End If

        ' ************************************************
        ' PROGRAMA FIDELIDADE
        If gParticipaProgramaFidelidade = 1 Then
          '1-SIM PARTICIPA;
          '0-NÃO PARTICIPA Empresa/filial;
          If gClienteEntregouResgatePontos = True And gSaldoCdGuidResgate_clicou_ok_telaDesconto = False Then
            B_Desconto_Click
            Exit Sub
          End If
        End If
        ' ************************************************
        
        '25/03/2004 - Daniel
        'Implementação criada para evitar gravação adulterada por usuário
90      If m_blnUserDanger Then
100       B_Grava_Click
110     End If
        '-----------------------------------------------------------------
        
120     Call StatusMsg("")
        
130     If IsNull(Num_Registro) Then
140       DisplayMsg "Encontre ou grave uma venda antes."
150       Exit Sub
160     End If
        
        
        '23/09/2002 - mpdea
        'Força a atualização do registro
170     If mblnForceUpdate Then
180       DisplayMsg "Valores alterados, grave a venda antes."
190       Exit Sub
200     End If
        
        
210     rsUsuarios.Index = "Código"
220     rsUsuarios.Seek "=", Cod_Operador.Caption
230     If rsUsuarios.NoMatch Then
240       MsgBox ("Operador não encontrado.")
250       Exit Sub
260     End If
270     If rsUsuarios("Recebimento") = False Then
280        Beep
290        DisplayMsg "Este usuário não tem permissão para usar a tela de recebimento."
300        Exit Sub
310     End If
        
320     If IsNumeric(rsParametros("DiasBloqueioVenda").Value) Then
330       If rsParametros.Fields("DiasBloqueioVenda") > 0 Then
340         If IsDate(rsCliFor.Fields("Última Compra")) Then
350           If (CDate(Data_Atual) - CDate(rsCliFor.Fields("Última Compra"))) > CInt(rsParametros.Fields("DiasBloqueioVenda")) Then
360             If MsgBox("O cliente que você escolheu não compra há " & (CDate(Data_Atual) - CDate(rsCliFor.Fields("Última Compra"))) & " dias, deseja continuar ? ", vbQuestion + vbYesNo, "Quick Store") = vbNo Then
370               Exit Sub
380             End If
390           End If
400         End If
410       End If
420     End If
        
430     If rsParametros("VR Recebimento Normal") = False Then
440       Beep
450       DisplayMsg "Não é permitido o recebimento normal, use o recebimento simplificado."
460       Exit Sub
470     End If
         
480     If rsSaidas("Recebimento") = True Then
490       Resp = MsgBox("Esta operação já foi efetivada. Os dados de recebimento estão disponíveis apenas para visualização. Caso queira alterar os dados do recebimento, use a opção DESFAZ movimentação no menu Opções antes.", vbInformation, "Atenção")
          
500       frmRecebimento.Limpa_Tela (0)
510       frmRecebimento.Receber.Caption = Total_Pagar
520       frmRecebimento.L_Sequência = rsSaidas("Sequência")
530       frmRecebimento.Só_Leitura.Value = 1
          
540       frmRecebimento.Show vbModal
550       Exit Sub
          
560     End If
        
        
        '09/10/2002 - mpdea
        'Verifica estoque conforme configurações
570     If Not rsParametros.Fields("Venda Sem Estoque").Value And rsOp_Saída.Fields("Estoque").Value Then
580       If Not mblnCheckStock() Then Exit Sub
590     End If
        
        
600     Call StatusMsg("")
        
610     frmRecebimento.Limpa_Tela (0)
620     frmRecebimento.Só_Leitura.Value = 0
630     frmRecebimento.L_Sequência = rsSaidas("Sequência")
640     frmRecebimento.Receber.Caption = Total_Pagar
650     frmRecebimento.Max_Cheques.Caption = 0
660     frmRecebimento.Max_Parcelas.Caption = 0
670     frmRecebimento.Intervalo_Parc.Caption = rsParametros("VR Intervalo Parc")
680     frmRecebimento.Combo_Banco.Text = rsCliFor("Conta Cobrança")
        
        
        '30/07/2003 - Maikel
        '             Adicionada propriedade do código de cliente para verificação de limite de crédito e se o cliente pode comprar a prazo.
690     frmRecebimento.lngCodigoCliente = CLng(Combo_Cliente.Text)
700     frmRecebimento.bytTelaChamada = 1  'Venda rápida
        '----------------------------------------------------------------------'
        
710     If rsCliFor("Tem Conta") = False Then frmRecebimento.Conta.Enabled = False
        
720     rsTabelas.Index = "Tabela"
730     rsTabelas.Seek "=", Combo_Preço.Text
740     If Not rsTabelas.NoMatch Then
750       If rsTabelas("Aceita Cartão") = False Then
760         frmRecebimento.Combo_Empresa.Enabled = False
770         frmRecebimento.Num_Cartão.Enabled = False
780         frmRecebimento.Cartão.Enabled = False
790       End If
800       If rsTabelas("Aceita Vale") = False Then
810         frmRecebimento.Vale.Enabled = False
820       End If
830       If rsTabelas("Aceita Pré") = False Then
840         frmRecebimento.Qtde_Cheques.Enabled = False
850         frmRecebimento.Grade_Cheque.Enabled = False
860       End If
870       If rsTabelas("Aceita Pré") = True Then frmRecebimento.Max_Cheques.Caption = rsTabelas("Prazo Pré")
880       If rsTabelas("Aceita Parcelamento") = False Then
890         frmRecebimento.Qtde_Parcelas.Enabled = False
900         frmRecebimento.Grade_Parcela.Enabled = False
910       End If
920       If rsTabelas("Aceita Parcelamento") = True Then frmRecebimento.Max_Parcelas.Caption = rsTabelas("Prazo Parcelamento")
930     End If
        
940     If rsCliFor("Faturado") = False Then
950       frmRecebimento.Max_Cheques.Caption = 1
960       frmRecebimento.Max_Parcelas.Caption = 1
970     End If

      '  '17/07/2003 - Maikel
      '  '   Se na tela de clientes, compra a prazo estiver desmarcada o frame de parcelamento ficará desabilitado
      '    frmRecebimento.frmParcela.Enabled = rsCliFor("Faturado")
      '  '-------------------------------------------------------------------------------------
        
       ' frmRecebimento.Dinheiro.SetFocus
980     frmRecebimento.Show vbModal
        
        
990     If frmRecebimento.Retorno.Caption <> "OK" Then
      '    DisplayMsg "Recebimento não efetivado."
1000      Exit Sub
1010    End If
        
        
        '27/03/2006 - mpdea
        'Solicitante: PSI Technomax - Rodrigo
        'Verifica o uso da gaveta em Venda Rápida
1020    If g_blnUsaGavetaVendaRapida() Then Call AbrirGaveta

        
        '11/08/2003 - mpdea
        'Desabilita controles
1030    Call EnableControls(False)
        
        
1040    Call WaitSeconds(1, True) 'Aguarda um segundo para o refresh
1050    Me.Refresh
        
1060    Screen.MousePointer = vbHourglass
        
1070    Call StatusMsg("Gravando recebimento...")
        
        '----------------------------------------------------------------------------------
        '29/05/2003 - mpdea
        'Atualizado
        '
        '05/08/2002 - mpdea
        'Requisição de bloqueio para gravação de venda
1080    If m_blnWorkTrafficLight Then
1090      Call TrafficLight.StartRequest(CLng(Número.Text))
1100    End If
        '----------------------------------------------------------------------------------
        
        'Início da transação
1110    ws.BeginTrans
1120    blnInTransaction = True
        
1130    Total_Prazo = frmRecebimento.Pega_Total_Parcelas
        
1140    rsSaidas.Edit
        '22/11/2004 - Daniel
        'Forçar mais uma vez a gravação do
        'caixa que estiver no objeto cboCaixa
        'Case: Casagrande
1150    rsSaidas.Fields("Caixa").Value = CByte(cboCaixa.Text)
        '----------------------------------------------------
1160    rsSaidas("Recebe - Conta") = False
1170    If frmRecebimento.Conta.Value = 1 Then
1180      rsSaidas("Recebe - Conta") = True
1190      rsSaidas("Total Prazo") = rsSaidas("Total")
1200    Else
1210      rsSaidas("Total Prazo") = frmRecebimento.Pega_Total_Parcelas
1220    End If
1230    rsSaidas("Recebe - Dinheiro") = CDbl(frmRecebimento.Dinheiro.Text)
1240    rsSaidas("Recebe - Emp Cartão") = Val(frmRecebimento.Combo_Empresa.Text)
1250    rsSaidas("Recebe - Num Cartão") = frmRecebimento.Num_Cartão.Text
1260    rsSaidas("Recebe - Cartão") = CDbl(frmRecebimento.Cartão.Text)
1270    rsSaidas("Recebe - Vale") = CDbl(frmRecebimento.Vale.Text)
1280    rsSaidas("Recebimento") = True
        rsSaidas("TotalCartaoDebito") = frmRecebimento.TxtDebito.Text
        rsSaidas("TotalCartaoCredito") = frmRecebimento.txtCredito.Text
1290    If rsSaidas("Recebe - Conta") = False Then rsSaidas("Total Prazo") = Total_Prazo
        
1300    If frmRecebimento.O_Banco.Value = True Then
1310      rsSaidas("Tipo Parcela") = "B"
1320      If rsSaidas("Total Prazo") <> 0 Then rsSaidas("Conta") = frmRecebimento.Combo_Banco.Text
1330    End If
        
1340    If frmRecebimento.O_Carteira.Value = True Then rsSaidas("Tipo Parcela") = "C"
1350    If frmRecebimento.O_Carnet.Value = True Then rsSaidas("Tipo Parcela") = "T"
         
        '11/12/2009 - Andrea
        'O recebimento em cartões agora será feito no grid de cartões (Grade_cartao)
        'e será salvo na tabela Movimento - Cartoes
        'If Len(Trim(frmRecebimento.Label_Cartão2.Caption)) > 0 Then
        '  rsSaidas("Parcela Cartão") = "S"
        '  rsSaidas("Qtde Parcelas") = frmRecebimento.Label_Cartão2.Caption
        '  rsSaidas("Valor Parcela") = CDbl(frmRecebimento.Label_Cartão4.Caption)
        'End If

        '07/01/2004 - Daniel
        'Alimentando os campos Valor Recebido e Troco
        'da tabela Saídas
1360    rsSaidas.Fields("Valor Recebido").Value = frmRecebimento.g_dblValorRecebidoFrmRec
1370    rsSaidas.Fields("Troco").Value = frmRecebimento.g_dblTrocoFrmRec

1380    rsSaidas.Update
        
        '-------------------------------------------------------------------------------------------------------------------------
        '11/12/2009 - Andrea
        'Apaga Cartoes
1390    Call EraseTypeMoviment(tmMovimentoCartoes, gnCodFilial, Val(Número.Text))
        'Grava Cartoes
        Dim lng_row As Long
        Dim var_book As Variant
        Dim str_administradora As String
        Dim dbl_valor As Double
        Dim int_qtde_parcelas As Double
        Dim dbl_valor_parcela As Double
        Dim str_numero As Double
        Dim bln_credito As Boolean


        'Valor em cartão
1400    With frmRecebimento.Grade_Cartoes
          'Verifica ocorrência
1410      If .Rows > 0 Then
            
1420        For lng_row = 0 To .Rows - 1
                
1430          var_book = .AddItemBookmark(lng_row)
                    
              'Verifica registro informado
1440          Call IsDataType(dtString, .Columns("Administradora").CellText(var_book), str_administradora)
1450          If str_administradora <> "" Then
                'Valores
1460            Call IsDataType(dtDouble, .Columns("Valor").CellText(var_book), dbl_valor)
1470            Call IsDataType(dtInteger, .Columns("Qtde Parcelas").CellText(var_book), int_qtde_parcelas)
1480            If int_qtde_parcelas = 0 Then int_qtde_parcelas = 1
1490            Call IsDataType(dtDouble, .Columns("Valor Parcelas").CellText(var_book), dbl_valor_parcela)
1500            Call IsDataType(dtString, .Columns("Numero").CellText(var_book), str_numero)
                Call IsDataType(dbBoolean, .Columns("Credito").CellValue(var_book), bln_credito)
                
1510            rsSaída_Cartoes.AddNew
1520              rsSaída_Cartoes("Filial") = gnCodFilial
1530              rsSaída_Cartoes("Sequência") = Val(Número.Text)
1540              rsSaída_Cartoes("Ordem") = (lng_row + 1)
1550              rsSaída_Cartoes("Administradora") = str_administradora
1560              rsSaída_Cartoes("Valor") = dbl_valor
1570              rsSaída_Cartoes("Parcelas") = int_qtde_parcelas
1580              rsSaída_Cartoes("ValorParcelas") = dbl_valor_parcela
                  '15/12/2009 - Andrea
                  'Maikel e Marcelo pediram para não gravar o número do cartão
                  rsSaída_Cartoes("NumeroCartao") = str_numero
                  rsSaída_Cartoes("Credito") = bln_credito
1590            rsSaída_Cartoes.Update
                
1600          End If
1610        Next lng_row
1620      End If
1630    End With
        '-------------------------------------------------------------------------------------------------------------------------
       
        'Apaga Cheques
1640    Call EraseTypeMoviment(tmMovimentoCheques, gnCodFilial, Val(Número.Text))
        'Grava Cheques
1650    Ordem = 1
1660    Do
1670      Resposta = frmRecebimento.Pega_Banco(Ordem, R_Banco, R_Cheque, R_Bom, R_Valor)
1680      If Resposta = 1 Then
1690        rsSaída_Cheques.AddNew
1700          rsSaída_Cheques("Filial") = gnCodFilial
1710          rsSaída_Cheques("Sequência") = Val(Número.Text)
1720          rsSaída_Cheques("Ordem") = Ordem
1730          rsSaída_Cheques("Banco") = R_Banco
1740          rsSaída_Cheques("Cheque") = R_Cheque
1750          rsSaída_Cheques("Bom") = R_Bom
1760          rsSaída_Cheques("Valor") = R_Valor
1770        rsSaída_Cheques.Update
1780      End If
1790      Ordem = Ordem + 1
' removido em 20/06/2022 (Pablo) - habilita parâmetro
'       Loop Until Ordem > 50
1800    Loop Until Ordem > pab_VR_Qtde_Cheques
          
        'Apaga Parcelas
1810    Call EraseTypeMoviment(tmMovimentoParcelas, gnCodFilial, Val(Número.Text))
        'Grava Parcelas
1820    Ordem = 1
1830    Do
1840      Resposta = frmRecebimento.Pega_Parcela(Ordem, R_Bom, R_Valor, Parcelas)
1850      If Resposta = 1 Then
1860        rsSaída_Parcelas.AddNew
1870        rsSaída_Parcelas("Filial") = gnCodFilial
1880        rsSaída_Parcelas("Sequência") = Val(Número.Text)
1890        rsSaída_Parcelas("Ordem") = Ordem
1900        rsSaída_Parcelas("Bom") = R_Bom
1910        rsSaída_Parcelas("Valor") = R_Valor
1920        rsSaída_Parcelas("Parcelas") = Parcelas
1930        rsSaída_Parcelas.Update
1940      End If
1950      Ordem = Ordem + 1
' removido em 20/06/2022 (Pablo) - habilita parâmetro
'       Loop Until Ordem > 50
1960    Loop Until Ordem > pab_VR_Qtde_Parcela
              
1970    Call StatusMsg("Aguarde, efetivando venda...")
        
1980    nRet = Efetiva_Saída(gnCodFilial, Val(Número.Text))
        
1990    If nRet <> 0 Then
2000      Select Case nRet
            Case -1
              'Ação cancelada
2010          Call StatusMsg("Ação cancelada.")
2020        Case 5
2030          Call DisplayMsg("Tabela de preços inexistente.")
2040        Case Else
2050          Call DisplayMsg("Operação NÃO efetivada. Erro" & str(nRet))
2060      End Select
2070      Efetivada.Visible = False
          Movimentação_Desfeita.Visible = False
          'Cancelamento da transação
2080      ws.Rollback
2090    Else
          'Fim da transação
2100      ws.CommitTrans
2110      blnInTransaction = False
2120      Efetivada.Visible = True
2130      m_blnSenhaGerJaInformada = False
2140      Call StatusMsg("")
2150    End If
        
        '----------------------------------------------------------------------------------
        '29/05/2003 - mpdea
        'Atualizado
        '
        '05/08/2002 - mpdea
        'Remoção de bloqueio para gravação de venda
2160    If m_blnWorkTrafficLight Then
2170      Call TrafficLight.FinishRequest
2180    End If
        '----------------------------------------------------------------------------------
        
        '11/08/2003 - mpdea
        'Habilita controles
2190    Call EnableControls(True)
        
        '01/12/2004 - Daniel
        'Case: De Mais Presentes (Nazareno)
        'Mostrar as informações do Recebimento
2200    If m_blnDeMais Then
2210      Mostra_Dados_Recebimento
2220    End If
        
        '12/08/2003 - mpdea
        'Somente executa se obteve retorno com sucesso
        '
        '25/10/2002 - mpdea
        'Verifica se deseja limpar a tela automaticamente
        'Código movido para após a ativação do form
2230    If nRet = 0 Then
2240      If ActiveBar1.Tools("miOpClearAfterVenda").Checked Then
2250        Call B_Limpa_Click
2260      End If
2270    End If
        
2280    Screen.MousePointer = vbDefault
          
2290    Exit Sub
        
ProcessErr:
2300    Screen.MousePointer = vbDefault
2310    Call StatusMsg("")
2320    Select Case Err.Number
          Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
2330        If nRepeatUpdateLocked < 30 Then
2340          Call frmAvisoBloqueio.ShowTentativas(30 - nRepeatUpdateLocked)
2350          Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
2360          nRepeatUpdateLocked = nRepeatUpdateLocked + 1
2370          Call WaitSeconds(1, False) 'Aguarda um segundo
2380          Resume
2390        Else
              
2400          If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
2410            nRepeatUpdateLocked = 0
2420            Resume
2430          Else
2440            Call StatusMsg("")
                'Cancelamento da transação
2450            If blnInTransaction Then ws.Rollback
2460            GoTo EnableControls
2470            Exit Sub
2480          End If
              
      '        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
      '          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
      '          "uma nova tentativa.", vbExclamation + vbOKCancel, "Venda Rápida - Recebimento") = vbOK Then
      '          nRepeatUpdateLocked = 0
      '          Resume
      '        Else
      '          Call StatusMsg("")
      '          'Cancelamento da transação
      '          If blnInTransaction Then ws.Rollback
      '          GoTo EnableControls
      '          Exit Sub
      '        End If
2490        End If
2500      Case Else
            'Cancelamento da transação
2510        If blnInTransaction Then ws.Rollback
            'Outros Erros
2520        Call StatusMsg("")
2530        MsgBox "Erro em Venda Rápida - Recebimento: " & Err.Number & " - " & Err.Description & "(Linha: " & Erl & ")", vbCritical, "Erro"
2540        GoTo EnableControls
2550        Exit Sub
            
      '      'Outros Erros
      '      Select Case frmErro.gnShowErr(Err.Number, "Venda Rápida - Recebimento")
      '        Case 0 'Repetir
      '          Resume
      '        Case 1 'Prosseguir
      '          Resume Next
      '        Case 2 'Sair
      '          Call StatusMsg("")
      '          GoTo EnableControls
      '          Exit Sub
      '        Case 3 'Encerrar
      '          End
      '      End Select
2560    End Select

2570    Exit Sub

EnableControls:
        '11/08/2003 - mpdea
        'Habilita controles
2580    Call EnableControls(True)

End Sub

Private Sub B_Recebe_Simples_Click()
  Dim nRet As Integer
  Dim nRepeatUpdateLocked As Integer
    
  Dim i As Integer
  Dim Fim As Integer
  Dim Ordem As Integer
  Dim Aux_Str As String
  Dim troco As Double
  Dim Parcelas As Integer
  Dim Total_Prazo As Double
  Dim rsCR As Recordset
  
  '08/01/2004 - Daniel
  m_sngQtdeTotal = 0
  '-------------------
  
  '10/08/2005 - Daniel
  'Adicionado a invocação da Private Recalcula_Recebido devido o problema que estava
  'ocorrendo com o Lost_Focus do campo Vale
  Call Recalcula_Recebido
  
  '07/01/2004 - Daniel
  'Alimentar as Variáveis Públicas g_dblTrocoFrmRec e g_dblValorRecebidoFrmRec
  'que popularão os fields [Valor Recebido] e [Troco] da tabela de Saídas
  If Not m_blnOcorreTroco Then '<Não Ocorre troco>
    'MsgBox "<Não Ocorre troco>"
    frmRecebimento.g_dblTrocoFrmRec = 0
    frmRecebimento.g_dblValorRecebidoFrmRec = (CDbl(L_Tot_Pagar.Text))
  Else  '<Ocorre troco>
    'MsgBox "<Ocorre troco>"
    '10/08/2005 - Daniel
    'Correção do Run-time error 13 Type mismatch
    'Incluímos Tratamento para os objetos nulos
    Dim bytAuxi As Byte
    '
    For bytAuxi = 0 To 4
      If Not IsNumeric(Val_Cheque(bytAuxi).Text) Then Val_Cheque(bytAuxi).Text = "0,00"
      If Not IsNumeric(Val_Parc(bytAuxi).Text) Then Val_Parc(bytAuxi).Text = "0,00"
    Next bytAuxi
    '
    If Not IsNumeric(Val_Cartão.Text) Then Val_Cartão.Text = "0,00"
    '----------------------------------------------------------------------------------------------------------
    frmRecebimento.g_dblValorRecebidoFrmRec = (CDbl(Format((Val_Cheque(0).Text), "###,###,##0.00")))
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Cheque(1).Text), "###,###,##0.00")))
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Cheque(2).Text), "###,###,##0.00")))
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Cheque(3).Text), "###,###,##0.00")))
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Cheque(4).Text), "###,###,##0.00")))
    
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Parc(0).Text), "###,###,##0.00")))
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Parc(1).Text), "###,###,##0.00")))
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Parc(2).Text), "###,###,##0.00")))
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Parc(3).Text), "###,###,##0.00")))
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Format((Val_Parc(4).Text), "###,###,##0.00")))
    
    frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Val_Cartão.Text))
    
    '10/08/2005 - Daniel
    'Incluímos os valores de Dinheiro e Vale na soma
    If IsNumeric(Dinheiro.Text) Then frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Dinheiro.Text))
    If IsNumeric(Vale.Text) Then frmRecebimento.g_dblValorRecebidoFrmRec = frmRecebimento.g_dblValorRecebidoFrmRec + (CDbl(Vale.Text))
    
  End If
    
  'Var m_blnOcorreTroco volta ao estado inicial para novas operações
  m_blnOcorreTroco = False
  '-----------------------------------------------------------------
  
  '28/06/2004 - Daniel
  '             Tratamento para funções não pegarem valores vazios para cálculos
  Dim intX As Integer

  For intX = 0 To 4
    If Val_Cheque(intX).Text = "" Then Val_Cheque(intX).Text = 0
    If Val_Parc(intX).Text = "" Then Val_Parc(intX).Text = 0
  Next intX
  
  '06/08/2003 - Maikel
  '             Chamada da função que verifica o crédito do cliente para o recebimento simplificado
  If Not AnalisaCreditoCliente Then Exit Sub
  '--------------------------------------------------------------------------------------------------
  
  '03/08/2002 - mpdea
  Dim blnInTransaction As Boolean
  
  On Error GoTo ProcessErr
  
  DoEvents
  
  
  '23/09/2002 - mpdea
  'Força a atualização do registro
  If mblnForceUpdate Then
    DisplayMsg "Valores alterados, grave a venda antes."
    Exit Sub
  End If
  
  
  troco = 0
  
  rsUsuarios.Index = "Código"
  rsUsuarios.Seek "=", Cod_Operador.Caption
  If rsUsuarios.NoMatch Then
    MsgBox ("Operador não encontrado.")
    Exit Sub
  End If
  If rsUsuarios("Recebimento") = False Then
     Beep
     DisplayMsg "Este usuário não tem permissão para realizar recebimento."
     Exit Sub
  End If
  
  If CDbl(L_Receber.Text) > 0 Then
    If Label_Receber.Caption = "A Receber" Then
      Beep
      DisplayMsg "Valor recebido insuficiente."
      Exit Sub
    Else
      troco = CDbl(L_Receber.Text)
      Aux_Str = "Troco = " + Format(troco, FORMAT_VALUE)
      MsgBox Aux_Str, vbInformation, "Aviso"
    End If
  End If
  
  If IsNull(Val_Cartão.Text) Then Val_Cartão.Text = 0
  If Val_Cartão.Text = "" Then Val_Cartão.Text = 0
  If CDbl(Val_Cartão.Text) <> 0 And Nome_Cartão.Caption = "" Then
    DisplayMsg "Escolha a administradora de cartão."
    Combo_Cartão.SetFocus
    Exit Sub
  End If
  
  i = Grava_Venda
  If i <> 0 Then
    Exit Sub
  End If
  
  Parcelas = 0
  For i = 0 To 4
   If IsNull(Val_Parc(i).Text) Then Val_Parc(i).Text = 0
   If Val_Parc(i).Text = "" Then Val_Parc(i).Text = 0
   If IsDate(Data_Parc(i).Text) And IsNumeric(Val_Parc(i).Text) Then Parcelas = Parcelas + 1
  Next i
  
  Screen.MousePointer = vbHourglass
  
  Call StatusMsg("Gravando Recebimento Simples...")
  
  
  '11/08/2003 - mpdea
  'Desabilita controles
  Call EnableControls(False)
  
  
  '----------------------------------------------------------------------------------
  '29/05/2003 - mpdea
  'Atualizado
  '
  '05/08/2002 - mpdea
  'Requisição de bloqueio para gravação de venda
  If m_blnWorkTrafficLight Then
    Call TrafficLight.StartRequest(CLng(Número.Text))
  End If
  '----------------------------------------------------------------------------------
  
  
  '24/03/2006 - mpdea
  'Solicitante: PSI Technomax - Rodrigo
  'Verifica o uso da gaveta em Venda Rápida
  If g_blnUsaGavetaVendaRapida() Then Call AbrirGaveta

  
  'Início de transação
  Call ws.BeginTrans
  blnInTransaction = True
  
  rsSaidas.Edit
  
  rsSaidas("Recebe - Conta") = False
  If Lançar_Débito.Value = 1 Then
    rsSaidas("Recebe - Conta") = True
    rsSaidas("Total Prazo") = rsSaidas("Total")
  Else
    rsSaidas("Total Prazo") = CDbl(Val_Parc(0).Text) + _
      CDbl(Val_Parc(1).Text) + CDbl(Val_Parc(2).Text) + _
      CDbl(Val_Parc(3).Text) + CDbl(Val_Parc(4).Text)
  End If
  
  If IsNull(Dinheiro.Text) Or Dinheiro.Text = "" Then Dinheiro.Text = 0
  If IsNull(Combo_Cartão.Text) Or Combo_Cartão.Text = "" Then Combo_Cartão.Text = 0
  If IsNull(Val_Cartão.Text) Or Val_Cartão.Text = "" Then Val_Cartão.Text = 0
  If IsNull(Vale.Text) Or Vale.Text = "" Then Vale.Text = 0
  
  rsSaidas("Recebe - Dinheiro") = CDbl(Dinheiro.Text) - troco
  rsSaidas("Recebe - Emp Cartão") = Val(Combo_Cartão.Text)
  rsSaidas("Recebe - Num Cartão") = Num_Cartão.Text
  rsSaidas("Recebe - Cartão") = CDbl(Val_Cartão.Text)
  rsSaidas("Recebe - Vale") = CDbl(Vale.Text)
  rsSaidas("Recebimento") = True
  rsSaidas("TotalCartaoDebito") = frmRecebimento.TxtDebito.Text
  rsSaidas("TotalCartaoCredito") = frmRecebimento.txtCredito.Text
  
  If O_Banco.Value = True Then
    rsSaidas("Tipo Parcela") = "B"
    If rsParametros("VR Conta Padrão") = "F" Then
      rsSaidas("Conta") = rsParametros("VR Conta Usar")
    Else
      rsSaidas("Conta") = rsCliFor("Conta Cobrança")
    End If
  End If
   
  If CDbl(Val_Cartão.Text) > 0 Then
    
    Call StatusMsg("Verificando cartão...")
    
    rsSaidas("Parcela Cartão") = "S"
    rsSaidas("Qtde Parcelas") = 1
    rsSaidas("Valor Parcela") = CDbl(Val_Cartão.Text)
  
    Set rsCR = db.OpenRecordset("Contas a Receber")
    
    rsCR.Index = "Contas"
    rsCR.Seek ">", "O", gnCodFilial, rsSaidas("Sequência"), 0
    If Not rsCR.NoMatch Then
      If rsCR("Tipo") = "O" Then
        If rsCR("Filial") = gnCodFilial Then
          If rsCR("Sequência") = rsSaidas("Sequência") Then
            '10/09/2007 - Anderson
            'Gera arquivo log do sistema
            If g_bolSystemLog Then
              SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
              "Cli:" & rsCR("Cliente") & "- Seq:" & rsCR("Sequência") & "- NF:" & rsCR("Nota") & "- Venc:" & rsCR("Vencimento") & "- Valor:" & rsCR("Valor"), _
              "frmVendaRap2_B_Recebe_Simples_CLick", _
              "Contas a Receber", g_strArquivoSystemLog
            End If
            rsCR.Delete
          End If
        End If
     End If
    End If
           
    rsCartoes.Index = "Código"
    rsCartoes.Seek "=", rsSaidas("Recebe - Emp Cartão")
    If Not rsCartoes.NoMatch Then
      rsCR.AddNew
      rsCR("Tipo") = "O"
      rsCR("Filial") = gnCodFilial
      rsCR("Sequência") = rsSaidas("Sequência")
      rsCR("Cliente") = rsSaidas("Cliente")
      rsCR("Administradora") = rsSaidas("Recebe - Emp Cartão")
      rsCR("Cartão") = rsSaidas("Recebe - Num Cartão")
      rsCR("Vencimento") = (rsSaidas("Data") + rsCartoes("Dias Pagar"))
      rsCR("Data Emissão") = rsSaidas("Data")
      rsCR("Valor Cartão") = rsSaidas("Recebe - Cartão")
      rsCR("Valor") = Round(CDbl(rsSaidas("Recebe - Cartão") * ((1 - rsCartoes("Taxa") / 100))), 2)
      rsCR("Data Alteração") = Format(Date, "dd/mm/yyyy")
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
        "Cli:" & rsCR("Cliente") & "- Seq:" & rsCR("Sequência") & "- NF:" & rsCR("Nota") & "- Venc:" & rsCR("Vencimento") & "- Valor:" & rsCR("Valor"), _
        "frmVendaRap2_B_Recebe_Simples_Click", _
        "Contas a Receber", g_strArquivoSystemLog
      End If
      rsCR.Update
    End If
    
    rsCR.Close
    Set rsCR = Nothing
  
  End If
   
  If O_Carteira.Value = True Then rsSaidas("Tipo Parcela") = "C"
  If O_Carnet.Value = True Then rsSaidas("Tipo Parcela") = "T"
  
  rsSaidas.Update
    
  Call StatusMsg("Verificando cheques...")
  
  'Apaga Cheques
  Call EraseTypeMoviment(tmMovimentoCheques, gnCodFilial, Val(Número.Text))
  'Grava Cheques
  Ordem = 1
  For i = 0 To 4
    If IsNull(Val_Cheque(i).Text) Then Val_Cheque(i).Text = 0
    If Val_Cheque(i).Text = "" Then Val_Cheque(i).Text = 0
    If CDbl(Val_Cheque(i).Text) <> 0 Then
      rsSaída_Cheques.AddNew
      rsSaída_Cheques("Filial") = gnCodFilial
      rsSaída_Cheques("Sequência") = Val(Número.Text)
      rsSaída_Cheques("Ordem") = Ordem
      If Banco(i).Text = "" Then Banco(i).Text = 0
      rsSaída_Cheques("Banco") = Banco(i).Text
      rsSaída_Cheques("Cheque") = Cheque(i).Text
      rsSaída_Cheques("Bom") = Bom_Para(i).Text
      rsSaída_Cheques("Valor") = Val_Cheque(i).Text
      rsSaída_Cheques.Update
      Ordem = Ordem + 1
     End If
  Next i
  
  Call StatusMsg("Verificando parcelas...")
  'Apaga Parcelas
  Call EraseTypeMoviment(tmMovimentoParcelas, gnCodFilial, Val(Número.Text))
  'Grava Parcelas
  Ordem = 1
  
  Erase pfParcelasFatura
  ReDim pfParcelasFatura(4) As ParcelasFatura
  
  For i = 0 To 4
   If CDbl(Val_Parc(i).Text) <> 0 Then
    rsSaída_Parcelas.AddNew
    rsSaída_Parcelas("Filial") = gnCodFilial
    rsSaída_Parcelas("Sequência") = Val(Número.Text)
    rsSaída_Parcelas("Ordem") = Ordem
    rsSaída_Parcelas("Bom") = Data_Parc(i).Text
    rsSaída_Parcelas("Valor") = Val_Parc(i).Text
    rsSaída_Parcelas("Parcelas") = Parcelas
    rsSaída_Parcelas.Update
    Ordem = Ordem + 1
    
    pfParcelasFatura(i).pfDataVencimento = Data_Parc(i).Text
    pfParcelasFatura(i).pfValor = Val_Parc(i).Text
   End If
  Next i
        
  Call StatusMsg("Aguarde, efetivando venda...")
  
  nRet = Efetiva_Saída(gnCodFilial, Val(Número.Text))
  
  If nRet <> 0 Then
    Efetivada.Visible = False
    Movimentação_Desfeita.Visible = False
    'Cancelamento da transação
    ws.Rollback
    Select Case nRet
      Case -1
        'Ação cancelada
        Call StatusMsg("Ação cancelada.")
      Case 5
        Call DisplayMsg("Tabela de preços inexistente.")
      Case Else
        Call DisplayMsg("Operação NÃO efetivada. Erro" & str(nRet))
    End Select
  Else
    'Fim da transação
    ws.CommitTrans
    blnInTransaction = False
    Efetivada.Visible = True
    m_blnSenhaGerJaInformada = False
    Call StatusMsg("")
  End If
  
  '----------------------------------------------------------------------------------
  '29/05/2003 - mpdea
  'Atualizado
  '
  '05/08/2002 - mpdea
  'Remoção de bloqueio para gravação de venda
  If m_blnWorkTrafficLight Then
    Call TrafficLight.FinishRequest
  End If
  '----------------------------------------------------------------------------------
  
  
  '11/08/2003 - mpdea
  'Habilita controles
  Call EnableControls(True)
  
  
  '25/10/2002 - mpdea
  'Verifica se deseja limpar a tela automaticamente
  'Código movido para após a ativação do form
  If ActiveBar1.Tools("miOpClearAfterVenda").Checked Then
    Call B_Limpa_Click
  End If
  
  Screen.MousePointer = vbDefault
    
  Exit Sub
    
ProcessErr:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If nRepeatUpdateLocked < 30 Then
        Call frmAvisoBloqueio.ShowTentativas(30 - nRepeatUpdateLocked)
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        nRepeatUpdateLocked = nRepeatUpdateLocked + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          nRepeatUpdateLocked = 0
          Resume
        Else
          Call StatusMsg("")
          'Cancelamento da transação
          If blnInTransaction Then ws.Rollback
          GoTo EnableControls
          Exit Sub
        End If
        
'        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
'          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Venda Rápida - Recebimento Simples") = vbOK Then
'          nRepeatUpdateLocked = 0
'          Resume
'        Else
'          Call StatusMsg("")
'          'Cancelamento da transação
'          If blnInTransaction Then ws.Rollback
'          GoTo EnableControls
'          Exit Sub
'        End If
      End If
    Case Else
      'Cancelamento da transação
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      Call StatusMsg("")
      MsgBox "Erro em Venda Rápida - Recebimento Simples: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
      GoTo EnableControls
      Exit Sub
      
'      'Outros Erros
'      Select Case frmErro.gnShowErr(Err.Number, "Venda Rápida - Recebimento Simples")
'        Case 0 'Repetir
'          Resume
'        Case 1 'Prosseguir
'          Resume Next
'        Case 2 'Sair
'          Call StatusMsg("")
'          GoTo EnableControls
'          Exit Sub
'        Case 3 'Encerrar
'          End
'      End Select
  End Select

  Exit Sub

EnableControls:
  '11/08/2003 - mpdea
  'Habilita controles
  Call EnableControls(True)

End Sub

Private Sub B_Ret_NFCe_Click()
  If Número.Text = "" Then
    Exit Sub
  End If
  Dim VerificaRetorno As New clsNFCe
  'VerificaRetorno.VerificaRetorno (Número.Text)
  VerificaRetorno.VerificaRetorno ("123")
  
End Sub

'Formata o valor de acordo com o número de casas decimais e substitui separador decimal por ponto
Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Function FormataValorTextoComVirgula(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTextoComVirgula = Format(dblValor, "#0." & String(lngCasasDecimais, "0"))
End Function

Private Sub GeraXML_ticket()
  On Error GoTo Erro
  
  Dim rsVenda           As Recordset
  Dim rsVendaProdutos   As Recordset
  Dim sSql              As String
  Dim sXML_Ticket       As String
  Dim sDataVenda        As String
  Dim sHoraVenda        As String
  Dim sTotalProdutos    As String
  Dim dTotalProdutos    As Double
  Dim sTotalPago        As String
  Dim dTotalPago        As Double
  Dim sAuxValor1        As String
  Dim sAuxValor2        As String
  Dim sTotalDesconto    As String
  Dim sCodCliente       As String
  Dim sNomCliente       As String
  Dim sNomVendedor      As String
  Dim sNumeroLinhasProd As String
  Dim sCaixa            As String
  
  Dim rsParametros      As Recordset
  Dim sBancoPDV         As String
  
  sSql = "SELECT * FROM [Parâmetros Filial] Where Filial = " & gnCodFilial
  Set rsParametros = db.OpenRecordset(sSql, dbOpenDynaset)
  sBancoPDV = rsParametros("BancoPDV").Value
  rsParametros.Close
  Set rsParametros = Nothing
  
  sSql = "SELECT S.Data, S.NSU_Hora, S.Produtos, S.Total, S.Digitador, S.Cliente, C.Nome, F.Nome as NomeVendedor, S.Caixa "
  sSql = sSql & " FROM Saídas S, Cli_For C, Funcionários F "
  sSql = sSql & " Where S.Filial = " & gnCodFilial & " and "
  sSql = sSql & " S.Sequência = " & Número.Text & " and "
  sSql = sSql & " S.Cliente = C.Código and "
  sSql = sSql & " S.Digitador = F.Código "
  Set rsVenda = db.OpenRecordset(sSql, dbOpenDynaset)
  
  If Not (rsVenda.EOF And rsVenda.BOF) Then
      sDataVenda = rsVenda.Fields("Data").Value
      
      If Not IsNull(rsVenda.Fields("NSU_Hora").Value) Then
          sHoraVenda = rsVenda.Fields("NSU_Hora").Value
      Else
          sHoraVenda = "00:00"
      End If
      
      dTotalProdutos = rsVenda.Fields("Produtos").Value
      sAuxValor1 = FormataValorTexto(dTotalProdutos, 2)
      sAuxValor1 = Replace(sAuxValor1, ".", ",")
      sTotalProdutos = sAuxValor1
  
      dTotalPago = rsVenda.Fields("Total").Value
      sAuxValor2 = FormataValorTexto(dTotalPago, 2)
      sAuxValor2 = Replace(sAuxValor2, ".", ",")
      sTotalPago = sAuxValor2
      
      sTotalDesconto = CDbl(sAuxValor1) - CDbl(sAuxValor2)
      sTotalDesconto = FormataValorTexto(sTotalDesconto, 2)
      sTotalDesconto = Replace(sTotalDesconto, ".", ",")
      
      sCodCliente = rsVenda.Fields("Cliente").Value
      sNomCliente = rsVenda.Fields("Nome").Value
      sNomVendedor = rsVenda.Fields("NomeVendedor").Value
      sCaixa = rsVenda.Fields("Caixa").Value
  End If
  rsVenda.Close
  Set rsVenda = Nothing
  
  sXML_Ticket = ""
  sXML_Ticket = "<TicketQS>"
  
  ' =========================================================
  ' Dados empresa/filial
  sXML_Ticket = sXML_Ticket & "<DadosCabecalho>"
  sXML_Ticket = sXML_Ticket & "<Nome>" & gsNomeFilial & "</Nome>"
  sXML_Ticket = sXML_Ticket & "<Endereco>" & gsFilialEndereco & "</Endereco>"
  sXML_Ticket = sXML_Ticket & "<Bairro>" & gsFilialBairro & "</Bairro>"
  sXML_Ticket = sXML_Ticket & "<CidadeEstado>" & gsFilialCidadeEstado & "</CidadeEstado>"
  sXML_Ticket = sXML_Ticket & "<Cep>" & gsFilialCep & "</Cep>"
  sXML_Ticket = sXML_Ticket & "<Fone>" & gsFilialFone & "</Fone>"
  
  sXML_Ticket = sXML_Ticket & "<CodCliente>" & sCodCliente & "</CodCliente>"
  sXML_Ticket = sXML_Ticket & "<NomCliente>" & sNomCliente & "</NomCliente>"
  sXML_Ticket = sXML_Ticket & "<NomVendedor>" & sNomVendedor & "</NomVendedor>"
  sXML_Ticket = sXML_Ticket & "<Sequencia>" & Número.Text & "</Sequencia>"
  sXML_Ticket = sXML_Ticket & "<DataVenda>" & sDataVenda & "</DataVenda>"
  sXML_Ticket = sXML_Ticket & "<HoraVenda>" & sHoraVenda & "</HoraVenda>"
  sXML_Ticket = sXML_Ticket & "<Caixa>" & sCaixa & "</Caixa>"
  
  sXML_Ticket = sXML_Ticket & "</DadosCabecalho>"
  ' =========================================================
  
  ' =========================================================
  ' Dados Produtos
  
  sSql = "SELECT S.Código, P.Nome, S.Qtde, S.Preço, S.Desconto, S.[Preço Final] as PrecoFinal, S.Linha "
  sSql = sSql & " FROM [Saídas - Produtos] S, Produtos P "
  sSql = sSql & " Where S.Filial = " & gnCodFilial & " and "
  sSql = sSql & " S.Sequência = " & Número.Text & " and "
  sSql = sSql & " S.[Código sem Grade] = P.Código "
  sSql = sSql & " Order by S.Linha "
  Set rsVendaProdutos = db.OpenRecordset(sSql, dbOpenDynaset)
  
  If Not (rsVendaProdutos.EOF And rsVendaProdutos.BOF) Then
    rsVendaProdutos.MoveLast
    sNumeroLinhasProd = rsVendaProdutos.RecordCount
    rsVendaProdutos.MoveFirst
    
    sXML_Ticket = sXML_Ticket & "<TotalLinhasProduto>" & sNumeroLinhasProd & "</TotalLinhasProduto>"
    sXML_Ticket = sXML_Ticket & "<Produtos>"
  
    While Not rsVendaProdutos.EOF
        sXML_Ticket = sXML_Ticket & "<LinhaProduto" & rsVendaProdutos.Fields("Linha").Value & ">"
    
        sXML_Ticket = sXML_Ticket & "<CodProduto>" & rsVendaProdutos.Fields("Código").Value & "</CodProduto>"
        sXML_Ticket = sXML_Ticket & "<NomProduto>" & rsVendaProdutos.Fields("Nome").Value & "</NomProduto>"
        sXML_Ticket = sXML_Ticket & "<QtdeProduto>" & rsVendaProdutos.Fields("Qtde").Value & "</QtdeProduto>"
        sXML_Ticket = sXML_Ticket & "<PrecoProduto>" & FormataValorTextoComVirgula(rsVendaProdutos.Fields("Preço").Value, 2) & "</PrecoProduto>"
        sXML_Ticket = sXML_Ticket & "<DescProduto>" & FormataValorTextoComVirgula(rsVendaProdutos.Fields("Desconto").Value, 2) & "</DescProduto>"
        sXML_Ticket = sXML_Ticket & "<PrecoFinalProduto>" & FormataValorTextoComVirgula(rsVendaProdutos.Fields("PrecoFinal").Value, 2) & "</PrecoFinalProduto>"
    
        sXML_Ticket = sXML_Ticket & "</LinhaProduto" & rsVendaProdutos.Fields("Linha").Value & ">"
    
        rsVendaProdutos.MoveNext
    Wend
  End If
  rsVendaProdutos.Close
  Set rsVendaProdutos = Nothing
  
  sXML_Ticket = sXML_Ticket & "</Produtos>"
  ' =========================================================
  
  ' =========================================================
  ' Dados Totais
  sXML_Ticket = sXML_Ticket & "<Totais>"
  sXML_Ticket = sXML_Ticket & "<SubTotal>" & sTotalProdutos & "</SubTotal>"
  sXML_Ticket = sXML_Ticket & "<TotalDesconto>" & sTotalDesconto & "</TotalDesconto>"
  sXML_Ticket = sXML_Ticket & "<Total>" & sTotalPago & "</Total>"
  sXML_Ticket = sXML_Ticket & "</Totais>"
  ' =========================================================
  
  sXML_Ticket = sXML_Ticket & "</TicketQS>"
  
  ' =========================================================
  ' Gravar no banco de dados local do IMPRESSOR
  Dim rsTesteConexao As Recordset
  On Error GoTo SegueFluxo
    
  Set rsTesteConexao = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = ''")
SegueFluxo:

  If Err.Number <> 0 Then
      Set BancoPDV = OpenDatabase(sBancoPDV & "\QuickStore.mdb", False, False, ";PWD=" & gsGetPValue())
  End If
  
  Dim rsNFCe_Job As Recordset
  Set rsNFCe_Job = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = 'XXXXXXX' And cnpj = '999999'")

  Dim iTipo As Integer
' iTipo = 1     ' Não esta em contingencia
' iTipo = 2     ' Esta em contingencia
  iTipo = 9     ' XML Impressão de Ticket do QuickStore

  If rsNFCe_Job.EOF Then
      rsNFCe_Job.AddNew
      rsNFCe_Job!CNPJ = "TICKET_QS"
      rsNFCe_Job!xml = sXML_Ticket
      rsNFCe_Job!Tipo = iTipo
      rsNFCe_Job!Serie = 1
      rsNFCe_Job!N_NF = Número.Text
      rsNFCe_Job!Chave = Número.Text
      rsNFCe_Job!CPF = "TICKET_QS"
      rsNFCe_Job!Nome_Consumidor = "TICKET_QS"
      rsNFCe_Job!Data_Emissao = ""
      rsNFCe_Job!Total_Tributos = ""
      rsNFCe_Job!Nome_Emitente = ""
      rsNFCe_Job!Endereco_Emitente = ""
      rsNFCe_Job!IE_Emitente = ""
      rsNFCe_Job!retFazenda = "SEM RET"
      rsNFCe_Job.Update
  End If

  rsNFCe_Job.Close
  Set rsNFCe_Job = Nothing
  ' =========================================================

  Exit Sub
Erro:
  MsgBox "Inconsistência em rotina GeraXML_ticket " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub B_Ticket_Click()
  
  If giQuick_viaRDP = 1 Then
    'Será impresso pelo IMPRESSOR EXE c#
    GeraXML_ticket
  Else
    'Impresso padrão antigo
    ImprimirTicket False
  End If

End Sub

Private Sub EmisTicketRel()
  Dim strFileNameTicket As String
  Dim frmX As Form
  Dim intX As Integer
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre ou grave uma venda antes."
    Exit Sub
  End If
  
  If rsOp_Saída("Ticket Imprimir") <> "" Then
    'Ticket pré-selecionado
    strFileNameTicket = gsConfigPath & rsOp_Saída.Fields("Ticket Imprimir").Value & ""
  Else
    'Escolha do ticket
    Set frmX = New frmEscolheTicket
    frmX.Show vbModal
    Set frmX = Nothing
    If gsRetornoDoc = "CANCELADO" Then
'      DisplayMsg "Ticket não impresso."
      Exit Sub
    End If
    strFileNameTicket = gsConfigPath & gsRetornoDoc
  End If
    
  'Verifica a existência do ticket
  If Dir(strFileNameTicket) = "" Then
    DisplayMsg "Arquivo """ & strFileNameTicket & """ não encontrado."
    Exit Sub
  End If
  
  With frmRelatorioTicket
    .Filial = rsSaidas.Fields("Filial").Value
    .Sequencia = rsSaidas.Fields("Sequência").Value
    .Show vbModal
  End With
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

'30/01/2009 - mpdea
'Implementado opção para email
'28/01/2004 - mpdea
'Incluído a impressão do ticket para o Quick Store em modo limitado
'Organizado código e comentado
'Movido código para gravação das observações para somente quando
'a tela de observações for solicitada
Private Sub ImprimirTicket(ByVal blnEmail As Boolean)
  Dim strFileNameTicket As String
  Dim frmX As Form
  Dim intX As Integer
  
  
  On Error GoTo ErrHandler
  
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre ou grave uma venda antes."
    Exit Sub
  End If

  '03/07/2006 - mpdea
  'Verifica permissão para imprimir ticket somente em movimentações efetivadas
  'Solicitante: Bem me quer
  If Not blnEmail Then
    If rsParametros.Fields("ImprimeTicketMovEfetivada").Value Then
      If Not rsSaidas.Fields("Efetivada").Value Then
        DisplayMsg "Movimentação não efetivada. Não é possível imprimir o Ticket."
        Exit Sub
      End If
    End If
  End If
  
  '13/09/2012 - mpdea
  'Desvia se o ticket é em modo relatório
  If Not rsParametros.Fields("VRUtilizarTicketModoRelatorio").Value Or blnEmail Then
    'Quick em modo limitado
    If Not gblnQuickFull Then
      'Ticket padrão
      strFileNameTicket = gsConfigPath & "TicketFC.cti"
    Else
      If rsOp_Saída("Ticket Imprimir") <> "" Then
        'Ticket pré-selecionado
        strFileNameTicket = gsConfigPath & rsOp_Saída.Fields("Ticket Imprimir").Value & ""
      Else
        'Escolha do ticket
        Set frmX = New frmEscolheTicket
        frmX.Show vbModal
        Set frmX = Nothing
        If gsRetornoDoc = "CANCELADO" Then
    '      DisplayMsg "Ticket não impresso."
          Exit Sub
        End If
        strFileNameTicket = gsConfigPath & gsRetornoDoc
      End If
    End If
    
    'Verifica a existência do ticket
    If Dir(strFileNameTicket) = "" Then
      DisplayMsg "Arquivo """ & strFileNameTicket & """ não encontrado."
      Exit Sub
    End If
  End If
  
''''  'Quick em modo limitado
''''  'Exibe a tela somente para a versão completa
''''  If gblnQuickFull Then
''''    If IsToShowTelaObsTransp() Then
''''      Set frmX = New frmObsNota
''''
''''      With frmX
''''        .gsCliente = rsCliFor.Fields("Transportadora").Value
''''        .lngSequencia = rsSaidas.Fields("Sequência").Value
''''        .bytTipoTabela = 1
''''        .Show vbModal
''''      End With
''''
''''      Set frmX = Nothing
''''      If gsRetornoDoc <> "OK" Then
''''        DisplayMsg "Operação cancelada."
''''        Exit Sub
''''      End If
''''
''''      '11/08/2003 - maikel
''''      '             Gravação dos campos de observações na tela de saídas
''''      '----------------------------------------------------------------'
''''      With rsSaidas
''''        .Edit
''''
''''        'For intX = 0 To 7
''''        '  .Fields("obs_Obs" & intX + 1).Value = gsObsDoc(intX)
''''        'Next intX
''''        For intX = 0 To 1
''''          .Fields("obs_infCpl" & intX + 1).Value = gsObsDoc(intX)
''''        Next intX
''''
''''        .Fields("obs_Transportadora") = gsTransportadora
''''        .Fields("obs_Placa") = gsPlaca
''''        .Fields("obs_Uf") = gsUfrmPlaca
''''        .Fields("obs_Especie") = gsEspecieTrans
''''        .Fields("obs_Qtde") = gsQtdeTrans
''''        .Fields("obs_Marca") = gsMarcaTrans
''''        .Fields("obs_PesoBruto") = IIf(IsNumeric(gsPesoBruto), gsPesoBruto, 0)
''''        .Fields("obs_PesoLiquido") = IIf(IsNumeric(gsPesoLiquido), gsPesoLiquido, 0)
''''
''''        .Fields("obs_FretePago") = IIf(IsNumeric(gsFretePago), gsFretePago, 0)
''''        .Update
''''      End With
''''      '----------------------------------------------------------------'
''''    End If
''''  End If
  
  If blnEmail Then
    'Prepara para enviar por email
    Call EnviarEmailModeloTicket(strFileNameTicket, gnCodFilial, rsSaidas.Fields("Sequência").Value, rsSaidas.Fields("Cliente").Value)
  Else
    '13/09/2012 - mpdea
    'Ticket em modo relatório
    If rsParametros.Fields("VRUtilizarTicketModoRelatorio").Value Then
      With frmRelatorioTicket
        .Filial = rsSaidas.Fields("Filial").Value
        .Sequencia = rsSaidas.Fields("Sequência").Value
        .Show vbModal
      End With
    Else
      'Imprime o ticket
      Call Imprime_Ticket(strFileNameTicket, gnCodFilial, rsSaidas.Fields("Sequência").Value)
    End If
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Function IsToShowTelaObsTransp() As Boolean
  Dim rsCadOper As Recordset
  Dim rsParametros As Recordset
  Dim nCodOper As Integer
  
  Set rsCadOper = db.OpenRecordset("SELECT * FROM [Operações Saída] ORDER BY Código", dbOpenDynaset)
  Set rsParametros = db.OpenRecordset("SELECT * FROM [Parâmetros Filial] ORDER BY Filial", dbOpenDynaset)
  
  'Determine qual é a operação de saída que reflete a VR
  nCodOper = 0
  With rsParametros
    If Not .EOF Then
      .FindFirst "Filial = " & gnCodFilial
      If Not .NoMatch Then
        nCodOper = .Fields("VR Código Operação")
      End If
    End If
  End With
  
 
  'Veja se a operação encontrada necessita ou não exibição da tela de Obs
  With rsCadOper
    .FindFirst "Código = " & nCodOper
    If Not .NoMatch Then
      IsToShowTelaObsTransp = .Fields("InTelaObsTransp")
      rsCadOper.Close
      Set rsCadOper = Nothing
      rsParametros.Close
      Set rsParametros = Nothing
      Exit Function
    End If
  End With
  
    
  IsToShowTelaObsTransp = False
  rsCadOper.Close
  Set rsCadOper = Nothing
  rsParametros.Close
  Set rsParametros = Nothing
  
End Function

Private Sub Banco_GotFocus(Index As Integer)
  Banco(Index).SelStart = 0
  Banco(Index).SelLength = Len(Banco(Index).Text)
End Sub

Private Sub Bom_Para_GotFocus(Index As Integer)
  Bom_Para(Index).SelStart = 0
  Bom_Para(Index).SelLength = Len(Bom_Para(Index).Text)
End Sub

'24/10/2002 - mpdea
'Adicionado acesso ao calendário
Private Sub Bom_Para_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Bom_Para(Index).Text = frmCalendario.gsDateCalender(Bom_Para(Index).Text)
  End Select
End Sub

Private Sub Bom_Para_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Data(KeyAscii)
End Sub

'24/10/2002 - mpdea
'Modificado foco do objeto após mensagem de alerta/erro, evitando RT-5
Private Sub Bom_Para_LostFocus(Index As Integer)

  Dim Nova_Data As Date
  
  Call StatusMsg("")
   
  If IsNull(Bom_Para(Index).Text) Then Exit Sub
  If Bom_Para(Index).Text = "" Then Exit Sub

  '04/04/2005 - Daniel
  'Alteração: Caso esteja vazio sai fora e não ficará
  'forçando sem parar a digitação de uma data
  If (Bom_Para(Index).Text) = "  /  /    " Then Exit Sub
  '-----------------------------------------------------
  
  If IsNumeric(Bom_Para(Index).Text) Then
    If Val(Bom_Para(Index).Text) > 0 And Val(Bom_Para(Index).Text) < 500 Then
      Nova_Data = Data_Atual + Val(Bom_Para(Index).Text)
      Bom_Para(Index).Text = Nova_Data
    End If
    
    If Val(Bom_Para(Index).Text) = 0 Then
       Bom_Para(Index).Text = Date
    End If
   End If
  

  If Not IsDate(Bom_Para(Index).Text) Then
    DisplayMsg "Data inválida, verifique."
'    Bom_Para(Index).SetFocus
    Call SelectAllText(Bom_Para(Index), True)
    Exit Sub
  End If
  
  If CDate(Bom_Para(Index).Text) < CDate(Data_Atual) Then
    DisplayMsg "Data inválida, anterior à atual."
'    Bom_Para(Index).SetFocus
    Call SelectAllText(Bom_Para(Index), True)
    Exit Sub
  End If
  
    
  Bom_Para(Index).Text = Ajusta_Data(Bom_Para(Index).Text)


  If rsParametros("VR Prazo Cheques") > 0 Then
    If CDate(Bom_Para(Index).Text) - CDate(Data_Atual) > rsParametros("VR Prazo Cheques") Then
      Bom_Para(Index).Mask = ""
      Bom_Para(Index).Text = ""
      Bom_Para(Index).Mask = "##/##/####"
'      Bom_Para(Index).SetFocus
    Call SelectAllText(Bom_Para(Index), True)
      DisplayMsg "Prazo superior ao permitido."
      Beep
      Exit Sub
    End If
  End If
  
  If rsCliFor("Faturado") = False Then
    If CDate(Bom_Para(Index).Text) > CDate(Data_Atual) Then
      DisplayMsg "Cliente não pode comprar à prazo."
'      Bom_Para(Index).SetFocus
      Call SelectAllText(Bom_Para(Index), True)
      Beep
      Exit Sub
    End If
  End If

End Sub

Private Sub btnComandaVendas_Click()
  If frmComanda.Total > 1 Then frmComanda.Show vbModal
End Sub

Private Sub cboCaixa_CloseUp()
  cboCaixa.Text = cboCaixa.Columns(0).Text
  cboCaixa_LostFocus
End Sub

'10/02/2006 - mpdea
'Incluído tratamento de erro
'Corrigido RT-6 (overflow) ao informar valores inválidos
Private Sub cboCaixa_LostFocus()
  Dim bytCaixa As Byte
  
  
  On Error GoTo ErrHandler
  
  
  Nome_Caixa.Caption = ""
  
  If cboCaixa.Text <> "" Then
    Call IsDataType(dtByte, cboCaixa.Text, bytCaixa)
    
    If bytCaixa > 0 And bytCaixa < 100 Then
      With datCaixa.Recordset
        .FindFirst "Caixa = " & bytCaixa
        If .NoMatch Then
          DisplayMsg "Caixa não encontrado."
        Else
          Nome_Caixa.Caption = .Fields("Descrição").Value & ""
        End If
      End With
    Else
      DisplayMsg "Caixa incorreto."
    End If
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Cheque_GotFocus(Index As Integer)
 Cheque(Index).SelStart = 0
 Cheque(Index).SelLength = Len(Cheque(Index).Text)
End Sub

Private Sub Cheque_LostFocus(Index As Integer)
 Call StatusMsg("")
End Sub

Private Sub cmd_abaixo_Click()
    Me.Top = Me.Top + 100
End Sub

Private Sub cmd_acharVenda_Click()
    lbl_retornoEnvioNFCe.Visible = False
    
    frmVendasHoje.bo_acaoSeleciona_e_fecha = True
    frmVendasHoje.Show vbModal
End Sub

Private Sub cmd_Acima_Click()
    Me.Top = Me.Top - 100
End Sub

Private Sub cmd_carne_Click()
    Call EmiteCarnesNOVOS(1, Nome_Cliente.Caption)
End Sub

Private Sub cmd_carneComRecibo_Click()
    Call EmiteCarnesNOVOS(2, Nome_Cliente.Caption)
End Sub

Private Sub cmd_direita_Click()
    Me.Left = Me.Left + 100
End Sub

Private Sub cmd_esquerda_Click()
    Me.Left = Me.Left - 100
End Sub

Private Sub cmd_fecharFrameProdutoSemPrecoNaGrade_Click()
    frm_produtoSemPrecoNaGrade.Visible = False
End Sub

Private Sub cmd_fecharTela_Click()
    Unload Me
End Sub

Private Sub cmd_opcoes_Click()
On Error GoTo Erro
    
    If cmd_opcoes.BackColor = &HC0FFFF Then
        ActiveBar1.Tools("miConsClientes").Visible = False
        ActiveBar1.Tools("miConsProduto").Visible = False
        ActiveBar1.Tools("carneRapido").Visible = False
        ActiveBar1.Tools("carneRapidoRecibo").Visible = False
        ActiveBar1.Tools("miEnviarEmail").Visible = False
        ActiveBar1.Attach
        cmd_opcoes.BackColor = &HFFC0C0
        Me.Height = Me.Height + 200
    Else
        ActiveBar1.Detach
        cmd_opcoes.BackColor = &HC0FFFF
        Me.Height = Me.Height - 200
    End If

    Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub
Private Sub cmd_pesquisaAlfa_Click()
    frmPesquisaProduto2.Show
End Sub

Private Sub cmd_tabelaDePrecos_Click()
  
  Dim bm As Variant
  Dim obj_formPreco As Form
  
  If Grade1.Columns("Código").CellValue(bm) = "" Or Grade1.Columns("Código").CellValue(bm) = "0" Then
      MsgBox "Selecione um produto na grade", vbInformation, "Atenção"
  Else
      Set obj_formPreco = New frmTabelasDePrecosProdutos
      
      obj_formPreco.valorProdutoAcatado = ""
      obj_formPreco.CodigoProduto = Grade1.Columns("Código").CellValue(bm)
      obj_formPreco.nomeProduto = Grade1.Columns("Nome").CellValue(bm)

      obj_formPreco.Show vbModal

      'MsgBox "Produto " & Grade1.Columns("Código").CellValue(bm)
      
      If obj_formPreco.valorProdutoAcatado <> "" Then
          Grade1.Columns("Preço").Value = obj_formPreco.valorProdutoAcatado
          Call Calcula_Linha
      End If
      
      Set obj_formPreco = Nothing
  End If
End Sub

Private Sub cmdInsertItens_Click()
  Dim nX As Integer
  Call B_Limpa_Click
  Grade1.MoveFirst
  Grade1.SetFocus
  SendKeys "^{HOME}", True
  For nX = 1 To 255
    SendKeys "1{DOWN}", True
  Next nX
  SendKeys "1{UP}", True
  B_Grava_Recebe.SetFocus
End Sub

Private Sub Combo_Cartão_CloseUp()
  Combo_Cartão.Text = Combo_Cartão.Columns(1).Text
  Combo_Cartão_LostFocus
End Sub

Private Sub Combo_Cartão_LostFocus()

  Nome_Cartão.Caption = ""
  
  If IsNull(Combo_Cartão.Text) Then Exit Sub
  If Combo_Cartão.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Cartão.Text) Then Exit Sub
  If Val(Combo_Cartão.Text) < 1 Then Exit Sub
    
  rsCartoes.Index = "Código"
  rsCartoes.Seek "=", Val(Combo_Cartão.Text)
  If rsCartoes.NoMatch Then Exit Sub
  
  Nome_Cartão.Caption = rsCartoes("Nome") & ""
  

End Sub

Private Sub Combo_Cliente_Click()
  Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
End Sub

Private Sub Combo_Cliente_CloseUp()
 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
 Nome_Cliente.Caption = ""
 
 'Indica que ainda não foi informada Senha Gerente para este cliente
 If m_strCodigoClienteContas <> Combo_Cliente.Text Then
    m_blnSenhaGerJaInformada = False
 End If
  
 Desconto_Cli = 0
 If IsNull(Combo_Cliente.Text) Then Exit Sub
 If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
 If Val(Combo_Cliente.Text) <= 0 Then Exit Sub
 If Val(Combo_Cliente.Text) > 99999999 Then Exit Sub
 
 rsCliFor.Index = "Código"
 rsCliFor.Seek "=", Val(Combo_Cliente.Text)
 If rsCliFor.NoMatch Then Exit Sub
 
  '18/09/2002 - mpdea
  'Verifica se o cliente está bloqueado ou inativo
  If rsCliFor("Bloqueado") Then
    DisplayMsg "Cliente [" & rsCliFor.Fields("Nome").Value & "] está bloqueado."
    Call SelectAllText(Combo_Cliente, True)
    Exit Sub
  End If
  
  If rsCliFor("Inativo") Then
    DisplayMsg "Cliente [" & rsCliFor.Fields("Nome").Value & "] está inativo."
    Call SelectAllText(Combo_Cliente, True)
    Exit Sub
  End If
 
  '03/09/2003 - Maikel
  '             Adicionada a verificação abaixo para analisar se o usuário pode ver o limite de crédito do cliente
  With rsFuncionarios
    If Not .NoMatch Then
      Nome_Cliente.Caption = IIf(.Fields("VR_PermiteVisualizarLimiteCredito"), " " & rsCliFor("Nome") & " - " & rsCliFor("Limite Crédito"), " " & rsCliFor("Nome"))
    '------------------------------------------------------------------------------------
    '04/08/2006 - Andrea
    'Quando o vendedor fosse inválido, não estava mostrando nem o nome do cliente
    ' inclui este else abaixo, para sempre trazer o nome do cliente, independente do
    ' vendedor ser válido.
    Else
      Nome_Cliente.Caption = rsCliFor("Nome")
    '-------------------------------------------------------------------------------------
    End If
    
  End With
  
  '21/02/2008 - mpdea
  'Corrigido valor nulo (RT-94)
  'Desconto_Cli = rsCliFor.Fields("Desconto").Value
  Call IsDataType(dtDouble, rsCliFor.Fields("Desconto").Value, Desconto_Cli)
 
 Rem Acha estado da empresa
 Estado = ""
 rsEstados.Index = "Estado"
 If IsNull(rsCliFor("Estado")) Then Exit Sub
 If rsCliFor("Estado") <> "" Then
     rsEstados.Seek "=", rsCliFor("Estado")
     If Not rsEstados.NoMatch Then
       Estado = rsEstados("Estado")
     End If
 End If

  If Nome_Vendedor.Caption = "" Then
    If rsCliFor("Vendedor") <> 0 Then
      Combo_Vendedor.Text = " " & rsCliFor("Vendedor") & ""
      Combo_Vendedor_LostFocus
    End If
  End If
    
  '28/10/2005 - mpdea
  'Movido código, pois dependia de possuir vendedor para o cliente
  'para exibir a tabela de preços
  If Len(Trim(rsCliFor.Fields("TabelaPrecoPadrao"))) > 0 Then
    Combo_Preço.Text = rsCliFor.Fields("TabelaPrecoPadrao") & ""
    '06/04/2004 - mpdea
    'Movido execução para o final do código onde sempre é realizado
'      Combo_Preço_LostFocus
  End If
  
  
  '23/05/2006 - mpdea
  'Cliente isento em IPI
  m_blnIsentoIPI = rsCliFor.Fields("IsentoIPI").Value
  
  
  '06/04/2004 - mpdea
  'Realiza sempre o recalculo dos preços devido a possíveis
  'modificações de desconto
  '25/05/2004 - Daniel
  'Foi criado um campo em Parâmetros [VR_RecalcularPreço] para
  'tornar o recalculo opcional aos clientes do quick
  Dim rstParametros As Recordset
  
  Set rstParametros = db.OpenRecordset("SELECT VR_RecalcularPreço FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      .MoveFirst
      'If .Fields("VR_RecalcularPreço").Value = True Then Call Combo_Preço_CloseUp 'Chama a rotina
      If .Fields("VR_RecalcularPreço").Value = True Then Call Recalcula
    End If
    .Close
  End With
  
  Set rstParametros = Nothing
  '---------[Fim da implementação feita em 25/05/2004]---------
  
End Sub

Private Sub Combo_Preço_Click()
  '07/10/2004 - Daniel
  Combo_Preço.Text = Combo_Preço.Columns(0).Text
End Sub

Private Sub Combo_Preço_CloseUp()
'  If Len(Trim(Combo_Preço.Text)) = 0 Then
'    Exit Sub
'  End If
'  Call RecalculaPrecos

  Dim nRow As Integer
  Dim bm As Variant
  Dim strFullCode As String
  Dim strCodProd As String
  Dim intErro As Integer
  Dim Aux_Preço As Double
  
  
  If Trim(Combo_Preço.Text) <> "" Then
    Screen.MousePointer = vbHourglass
    Call StatusMsg("Refazendo tabela...")
    'Refaz o preço ao alterar a tabela de preços
    rsPreços.Index = "Tabela"
    For nRow = 0 To Linhas_Grade - 1
      strFullCode = gsHandleNull(Tabe(nRow).Código)
      '-------------------------------------------------------------------------------
      '07/05/2002 - mpdea
      '
      'Alterado para que localize o preço para os produtos do tipo Grade e Edição
      '(Procura código principal)
      '-------------------------------------------------------------------------------
      Call Acha_Produto(strFullCode, strCodProd, 0, 0, 0, 0, intErro)
      If intErro = 0 Then
        If strFullCode <> "0" Then
'          rsPreços.Seek "=", Combo_Preço.Text, strCodProd
'          If rsPreços.NoMatch Then
'            Tabe(nRow).Preço = 0
'          Else
'            Tabe(nRow).Preço = rsPreços("Preço")
'          End If
        
          '---------------------------------------------------------------------------
          '06/04/2004 - mpdea
          '
          'Alterado para que inclua o desconto do produto, cliente e cotação no
          'cálculo do preço
          '---------------------------------------------------------------------------
          'Posiciona recordset
          rsProdutos.Index = "Código"
          rsProdutos.Seek "=", strCodProd
          If rsProdutos.NoMatch Then Exit Sub
          'Acha preço
          rsPreços.Index = "Tabela"
          rsPreços.Seek "=", Combo_Preço.Text, strCodProd
          If rsPreços.NoMatch Then
            Tabe(nRow).Preço = 0
          End If
          If Not rsPreços.NoMatch Then
            'Verifica permissão de desconto no produto
            If rsProdutos.Fields("DontAllowDesc").Value Then
              '05/05/2004 - Daniel
              'Personalização Embalavi
              'Tratamento de Máscara
              If g_bln5CasasDecimais Then
                Aux_Preço = (Format((rsPreços("Preço")), "##,###,##0.00000")) * ((100 - (rsProdutos("Desconto"))) / 100)
              '30/04/2007 - Anderson - Implementação de 3 casas decimais
              ElseIf g_bln3CasasDecimais Then
                Aux_Preço = (Format((rsPreços("Preço")), "##,###,##0.000")) * ((100 - (rsProdutos("Desconto"))) / 100)
              Else
                Aux_Preço = rsPreços("Preço") * ((100 - (rsProdutos("Desconto"))) / 100)
              End If
            Else
              '05/05/2004 - Daniel
              'Personalização Embalavi
              'Tratamento de Máscara
              If g_bln5CasasDecimais Then
                Aux_Preço = (Format((rsPreços("Preço")), "##,###,##0.00000")) * ((100 - (rsProdutos("Desconto") + Desconto_Cli)) / 100)
              '30/04/2007 - Anderson - Implementação de 3 casas decimais
              ElseIf g_bln3CasasDecimais Then
                Aux_Preço = (Format((rsPreços("Preço")), "##,###,##0.000")) * ((100 - (rsProdutos("Desconto") + Desconto_Cli)) / 100)
              Else
                Aux_Preço = rsPreços("Preço") * ((100 - (rsProdutos("Desconto") + Desconto_Cli)) / 100)
              End If
            End If
             
            If rsProdutos("Moeda") <> 1 Then
              rsCotacoes.Index = "Moeda"
              rsCotacoes.Seek "<=", rsProdutos("Moeda"), Data_Atual
              If Not rsCotacoes.NoMatch Then
                If rsCotacoes("Moeda") = rsProdutos("Moeda") Then
                  '05/05/2004 - Daniel
                  'Personalização Embalavi
                  'Tratamento de Máscara
                  If g_bln5CasasDecimais Then
                    Aux_Preço = (Format(Aux_Preço, "##,###,##0.00000")) * rsCotacoes("Cotação")
                  '30/04/2007 - Anderson - Implementação de 3 casas decimais
                  ElseIf g_bln3CasasDecimais Then
                    Aux_Preço = (Format(Aux_Preço, "##,###,##0.000")) * rsCotacoes("Cotação")
                  Else
                    Aux_Preço = Aux_Preço * rsCotacoes("Cotação")
                  End If
                End If
              End If
            End If
            '05/05/2004 - Daniel
            'Personalização Embalavi
            'Tratamento de Máscara
            If g_bln5CasasDecimais Then
              Tabe(nRow).Preço = Format(Aux_Preço, "##,###,##0.00000")
            '30/04/2007 - Anderson - Implementação de 3 casas decimais
            ElseIf g_bln3CasasDecimais Then
              Tabe(nRow).Preço = Format(Aux_Preço, "##,###,##0.000")
            Else
              Tabe(nRow).Preço = Format(Aux_Preço, FORMAT_VALUE)
            End If
          End If
          '---------------------------------------------------------------------------
        
        Else
          Tabe(nRow).Preço = 0
        End If
      Else
        Tabe(nRow).Preço = 0
      End If
      Call Calcula_Linha_Tabe(nRow)
    Next nRow
    'Recalcula valores
    Call Recalcula
    With Grade1
      .MoveLast
      .MoveFirst
    End With
    Screen.MousePointer = vbDefault
    Call StatusMsg("")
  End If

End Sub

Private Sub Combo_Preço_LostFocus()
  If IsNull(Combo_Preço.Text) Then
    Exit Sub
  ElseIf Combo_Preço.Text = "" Then
    Exit Sub
  Else
    Combo_Preço.Text = UCase(Combo_Preço.Text)
  End If
End Sub

Private Sub Combo_Vendedor_Click()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(2).Text
End Sub

Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(2).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_KeyPress(KeyAscii As Integer)
'Combo_Vendedor.Text = Combo_Vendedor.Columns(2).Text
'Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_LostFocus()
  B_Desconto.Enabled = False
 
  Nome_Vendedor.Caption = ""
  
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) <= 0 Then Exit Sub
  If Val(Combo_Vendedor.Text) > 9999 Then Exit Sub
  
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Combo_Vendedor.Text)
  If rsFuncionarios.NoMatch Then Exit Sub
  
  Nome_Vendedor.Caption = " " & rsFuncionarios("Apelido")
  
  B_Desconto.Enabled = rsFuncionarios("bPermiteDesconto")
  
  '27/12/2004 - Daniel
  'BUG..............: Tratamento para desprender o cursor
  '                   da grid de ítens
  'Encontrado por...: De Mais Presentes (Nazareno)
  Screen.MousePointer = vbDefault
    With Grade1
      .MoveLast
      .MoveFirst
    End With
  Screen.MousePointer = vbDefault
  '-------------------------------------------------------
  
End Sub

Private Sub Data_Parc_GotFocus(Index As Integer)
  Data_Parc(Index).SelStart = 0
  Data_Parc(Index).SelLength = Len(Data_Parc(Index).Text)
End Sub

'24/10/2002 - mpdea
'Adicionado acesso ao calendário
Private Sub Data_Parc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Parc(Index).Text = frmCalendario.gsDateCalender(Data_Parc(Index).Text)
  End Select
End Sub

'24/10/2002 - mpdea
'Modificado foco do objeto após mensagem de alerta/erro, evitando RT-5
Private Sub Data_Parc_LostFocus(Index As Integer)
  Dim Erro As Integer
  
  If IsNull(Data_Parc(Index).Text) Or Data_Parc(Index).Text = "" Then
     Val_Parc(Index).Text = 0
     Exit Sub
  End If
  
  '27/07/2004 - Daniel
  'Alteração: Caso esteja vazio sai fora e não ficará
  'forçando sem parar a digitação de uma data
  If (Data_Parc(Index).Text) = "  /  /    " Then Exit Sub
  
  If Not IsDate(Data_Parc(Index).Text) Then
    If Not IsNumeric(Data_Parc(Index).Text) Then
      DisplayMsg "Data inválida."
'      Data_Parc(Index).SetFocus
      Call SelectAllText(Data_Parc(Index), True)
      Exit Sub
    End If
    If Val(Data_Parc(Index).Text) < 0 Or Val(Data_Parc(Index).Text) > 360 Then
      DisplayMsg "Data inválida."
'      Data_Parc(Index).SetFocus
      Call SelectAllText(Data_Parc(Index), True)
      Exit Sub
    End If
    
    Data_Parc(Index).Text = Data_Atual + Val(Data_Parc(Index).Text)
    
  End If
  
    Data_Parc(Index).Text = Format(Data_Parc(Index).Text, "dd/mm/yyyy")
  
  If rsParametros("VR Prazo Parcela") > 0 Then
    If CDate(Data_Parc(Index).Text) - CDate(Data_Atual) > rsParametros("VR Prazo Cheques") Then
      
      '26/07/2004 - Daniel
      'Limpa o texto
      With Data_Parc(Index)
        .Mask = ""
        .Text = ""
        .Mask = "##/##/####"
      End With
      
      DisplayMsg "Prazo superior ao permitido."
      Beep
      Exit Sub
    End If
  End If
  
  If rsCliFor("Faturado") = False Then
    If CDate(Data_Parc(Index).Text) > CDate(Data_Atual) Then
      DisplayMsg "Cliente não pode comprar à prazo."
'      Data_Parc(Index).SetFocus
      Call SelectAllText(Data_Parc(Index), True)
      Beep
      Exit Sub
    End If
  End If

  
  
End Sub


Private Sub Dinheiro_GotFocus()
  Dinheiro.SelStart = 0
  Dinheiro.SelLength = Len(Dinheiro.Text)
End Sub


Private Sub Dinheiro_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Dinheiro_LostFocus()
 Recalcula_Recebido
End Sub

Private Sub DropDown1_Click()
' rsPreços.Index = "Tabela"
' rsPreços.Seek "=", Combo_Preço.Text, DropDown1.Columns(1).Text
' If rsPreços.NoMatch Then
'    L_Preço.Caption = ""
'    Grade1.Columns(3).Text = ""
' Else
'    L_Preço.Caption = "$ " + Format$(rsPreços("Preço"), "###,###,##0.00")
'    Grade1.Columns(3).Text = rsPreços("Preço")
' End If
' Grade1.Columns(0).Text = DropDown1.Columns(1).Text
' Grade1.Columns(2).Text = DropDown1.Columns(0).Text
' Call RecalculaPrecos
End Sub

Private Sub DropDown1_CloseUp()
'  DropDown1.DataFieldToDisplay = "Código"
'  'Grade1.Columns(Grade1.Col).Text = DropDown1.Columns(1).Text
'  Call StatusMsg("")

  With DropDown1
    rsPreços.Index = "Tabela"
    rsPreços.Seek "=", Combo_Preço.Text, .Columns("Código").Text
    If rsPreços.NoMatch Then
      Grade1.Columns("Preço").Text = "0.00"
    Else
      '05/05/2004 - Daniel
      'Personalização Embalavi
      'Tratamento de Máscara
      If g_bln5CasasDecimais Then
        Grade1.Columns("Preço").Text = Format$(rsPreços("Preço"), "###,###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Grade1.Columns("Preço").Text = Format$(rsPreços("Preço"), "###,###,##0.000")
      Else
        Grade1.Columns("Preço").Text = Format$(rsPreços("Preço"), "###,###,##0.00")
      End If
    End If
'    Grade1.Columns("Preço").Text = Format$(gsHandleNull(.Columns("Preço").Text), "###,###,##0.00")
    Grade1.Columns("Código").Text = .Columns("Código").Text
    Grade1.Columns("Nome").Text = .Columns("Nome").Text
    Call Calcula_Linha
 End With
End Sub

Private Sub Dropdown1_RowLoaded(ByVal Bookmark As Variant)
  Dim nEstoque As Double
  Dim sMsgEstoque As String
  Dim nErro As Integer
  
  On Error Resume Next
  
  With DropDown1
    'Estoque
    nEstoque = Acha_Estoque(gnCodFilial, .Columns("Código").Text, 0, 0, 0, nErro)
    Select Case nErro
      Case 0
        sMsgEstoque = nEstoque
      Case 1
        sMsgEstoque = "Estoque não iniciado"
      Case 2
        sMsgEstoque = "Depende da grade"
      Case 3
        sMsgEstoque = "Depende da edição"
      Case 4
        sMsgEstoque = "Produto não existe"
    End Select
    
    rsFuncionarios.Index = "Código"
    rsFuncionarios.Seek "=", Funcionario
    
    If rsFuncionarios.NoMatch Then
      .Columns("Estoque").Text = sMsgEstoque
    Else
      .Columns("Estoque").Text = IIf(rsFuncionarios.Fields("VRVisualizarEstoque"), sMsgEstoque, "Usuário não autorizado")
    End If
    
    '.Columns("Estoque").Text = sMsgEstoque
    'Preço
    If Combo_Preço.Text = "" Then
      .Columns("Preço").Text = "Preço não encontrado"
    Else
      rsPreços.Index = "Tabela"
      rsPreços.Seek "=", Combo_Preço.Text, .Columns("Código").Text
      
      If rsPreços.NoMatch Then
        .Columns("Preço").Text = "Preço não encontrado"
      Else
        '05/05/2004 - Daniel
        'Personalização Embalavi
        'Tratamento de Máscara
        If g_bln5CasasDecimais Then
          .Columns("Preço").Text = Format(rsPreços("Preço"), "##,###,##0.00000")
        '30/04/2007 - Anderson - Implementação de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
          .Columns("Preço").Text = Format(rsPreços("Preço"), "##,###,##0.000")
        Else
          .Columns("Preço").Text = Format(rsPreços("Preço"), Formato_Preço)
        End If
      End If
      
      If Not rsFuncionarios.Fields("VRVisualizarPreco") Then .Columns("Preço").Text = "Usuário não autorizado"
    End If
    
    '---[Coluna Fabricante]---
      '29/03/2005 - Daniel
      'Case: Elétrica Leal - Mostrar o fabricante
      '
      '12/05/2005 - Daniel
      'Rotina parametrizada conforme solicitação da Info Social
      If m_blnExibirColunaFabricante Then
        If Len(.Columns("Código").Text) > 0 Then .Columns("Fabricante").Text = GetFabricante(.Columns("Código").Text) & ""
      End If
    '---[Fim Fabricante]---
    
  End With
  
'
'
'  Dim Estoque As Double
'  Dim Erro As Integer
'
' If rsParametros("VR Mostrar Estoque") = False Then
'   DropDown1.Columns(2).Width = 1
'   DropDown1.Columns(3).Width = 1
'   Exit Sub
' End If
'
'
'  Estoque = Acha_Estoque(gnCodFilial, DropDown1.Columns(1).Text, 0, 0, 0, Erro)
'
'
'  If Erro = 0 Then DropDown1.Columns(2).Text = Estoque
'  If Erro = 1 Then DropDown1.Columns(2).Text = "estoque não iniciado"
'  If Erro = 2 Then DropDown1.Columns(2).Text = "depende da grade"
'  If Erro = 3 Then DropDown1.Columns(2).Text = "depende da edição"
'  If Erro = 4 Then DropDown1.Columns(2).Text = "produto não existe"
'
'
'  If Combo_Preço.Text = "" Then
'    DropDown1.Columns(3).Text = "preço não encontrado"
'    Exit Sub
'  End If
'
'  rsPreços.Index = "Tabela"
'  rsPreços.Seek "=", Combo_Preço.Text, DropDown1.Columns(1).Text
'  If rsPreços.NoMatch Then
'    DropDown1.Columns(3).Text = "preço não encontrado"
'    Exit Sub
'  End If
'
'  DropDown1.Columns(3).Text = Format(rsPreços("Preço"), Formato_Preço)
'
  
  
End Sub

Public Sub CheckMovimentacao()
  If Erro_Data Then
    Erro_Data = False
    If frmErroMov.gbContinue Then
      If Not frmGerente.gbSenhaGerente Then
        Unload Me
      End If
    Else
      Unload Me
    End If
  End If
End Sub

Private Sub DropDown2_CloseUp()
  With DropDown2
    rsPreços.Index = "Tabela"
    rsPreços.Seek "=", Combo_Preço.Text, .Columns("Código").Text
    If rsPreços.NoMatch Then
      Grade1.Columns("Preço").Text = "0.00"
    Else
      '05/05/2004 - Daniel
      'Personalização Embalavi
      'Tratamento de Máscara
      If g_bln5CasasDecimais Then
        Grade1.Columns("Preço").Text = Format$(rsPreços("Preço"), "###,###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Grade1.Columns("Preço").Text = Format$(rsPreços("Preço"), "###,###,##0.000")
      Else
        Grade1.Columns("Preço").Text = Format$(rsPreços("Preço"), "###,###,##0.00")
      End If
    End If
'    Grade1.Columns("Preço").Text = Format$(gsHandleNull(.Columns("Preço").Text), "###,###,##0.00")
    Grade1.Columns("Código").Text = .Columns("Código").Text
    Grade1.Columns("Nome").Text = .Columns("Nome").Text
    '21/12/2006 - Anderson
    'força a execução do evento para evitar erro de uso do Quick Store na tela de vendas
    'Descrição do Erro: AO digitar um código inválido, o quick exibe uma mensagem de erro e coloca o foco na
    'coluna do código do produto. O usuário usava as teclas de movimentação para a direita e depois para a
    'esquerda e abria a combo para selecionar um produto. Assim que escolhia o produto correto, o Quick não
    'estava atualizando os valores de impostos como por exemplo ICMS.
    Call Grade1_BeforeColUpdate(0, 0, 0)
    Call Calcula_Linha
 End With
End Sub

Private Sub DropDown2_RowLoaded(ByVal Bookmark As Variant)
  Dim nEstoque As Double
  Dim sMsgEstoque As String
  Dim nErro As Integer
  
  On Error Resume Next
  
  With DropDown2
    'Estoque
    nEstoque = Acha_Estoque(gnCodFilial, .Columns("Código").Text, 0, 0, 0, nErro)
    Select Case nErro
      Case 0
        sMsgEstoque = nEstoque
      Case 1
        sMsgEstoque = "Estoque não iniciado"
      Case 2
        sMsgEstoque = "Depende da grade"
      Case 3
        sMsgEstoque = "Depende da edição"
      Case 4
        sMsgEstoque = "Produto não existe"
    End Select
    .Columns("Estoque").Text = IIf(rsFuncionarios.Fields("VRVisualizarEstoque"), sMsgEstoque, "Usuário não autorizado")
    '.Columns("Estoque").Text = sMsgEstoque
    'Preço
    If Combo_Preço.Text = "" Then
      .Columns("Preço").Text = "Preço não encontrado"
    Else
      rsPreços.Index = "Tabela"
      rsPreços.Seek "=", Combo_Preço.Text, .Columns("Código").Text
      If rsPreços.NoMatch Then
        .Columns("Preço").Text = "Preço não encontrado"
      Else
        '05/05/2004 - Daniel
        'Personalização Embalavi
        'Tratamento de Máscara
        If g_bln5CasasDecimais Then
          .Columns("Preço").Text = Format(rsPreços("Preço"), "##,###,##0.00000")
        '30/04/2007 - Anderson - Implementação de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
          .Columns("Preço").Text = Format(rsPreços("Preço"), "##,###,##0.000")
        Else
          .Columns("Preço").Text = Format(rsPreços("Preço"), Formato_Preço)
        End If
      End If
      If Not rsFuncionarios.Fields("VRVisualizarPreco") Then .Columns("Preço").Text = "Usuário não autorizado"
    End If
    
    '---[Coluna Fabricante]---
      '29/03/2005 - Daniel
      'Case: Elétrica Leal - Mostrar o fabricante
      '
      '12/05/2005 - Daniel
      'Rotina parametrizada conforme solicitação da Info Social
      If m_blnExibirColunaFabricante Then
        If Len(.Columns("Código").Text) > 0 Then .Columns("Fabricante").Text = GetFabricante(.Columns("Código").Text) & ""
      End If
    '---[Fim Fabricante]---
    
  End With

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim bHandled As Boolean

  bHandled = ActiveBar1.OnKeyDown(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If
  
  Call HandleKeyDown(KeyCode, Shift)

' If KeyCode = vbKeyF2 Then
'   B_Grava_Click
'   Exit Sub
' End If
'
' If KeyCode = vbKeyF3 Then
'   B_Grava_Recebe_Click
'   Exit Sub
' End If
'
' If KeyCode = vbKeyF4 Then
'   B_Limpa_Click
'   Exit Sub
' End If
'
' If KeyCode = vbKeyF5 Then
'   B_Nota_Click
'   Exit Sub
' End If
'
' If KeyCode = vbKeyF7 Then
'   B_Recebe_Click
'   Exit Sub
' End If
'
' If KeyCode = vbKeyF9 Then
'   B_Ticket_Click
'   Exit Sub
' End If
'
' If KeyCode = vbKeyF10 Then
'   If B_Recebe_Simples.Enabled = True Then
'     B_Recebe_Simples.SetFocus
'     DoEvents
'     B_Recebe_Simples_Click
'   End If
'     Exit Sub
' End If
'
' If KeyCode = vbKeyF11 Then
'   Exit Sub
' End If
'
' If KeyCode = vbKeyF12 Then
'   Exit Sub
' End If
'
'
 If KeyCode = vbKeyF6 Then
   If Lançar_Débito.Enabled Then
     Lançar_Débito.SetFocus
   ElseIf Dinheiro.Enabled Then
     Dinheiro.SetFocus
   ElseIf Vale.Enabled Then
     Vale.SetFocus
   End If
 End If
'
'

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  Dim Aux As String
  Dim Aux_Int As Integer
  Dim i As Integer
  Dim Resp As Integer
  Dim Ret_Arquivo As String
  
  On Error GoTo ErrHandler
  
  btnComandaVendas.Visible = False
  txtComanda.Width = 1785
  
  '***************************************
  'TRATAMENTO PARA TELA TOTALMENTE RESTRITA
  If gbUsuarioAcessoApenasTelaVendaRapida = True Then
  
      If Funcionario = "" Then
          Combo_Vendedor.Text = gnUserCode
      Else
          Combo_Vendedor.Text = Funcionario
      End If
      
      'Combo_Vendedor_LostFocus
      'Combo_Vendedor.Enabled = False
      cboCaixa.Enabled = False
      B_Grava_Recebe.Visible = False
      B_Recebe.Visible = False
      B_NFC_e.Visible = False
      B_Ret_NFCe.Visible = False
      Cod_Operador.Enabled = False
      Nome_Operador.Enabled = False
      Nome_Caixa.Enabled = False
      Nome_Vendedor.Enabled = False
      Nome_Vendedor.Caption = gsVendedorVR
      B_Limpa.Top = 510
      B_Limpa.Width = 4965
      B_Limpa.Left = 5040
      B_programaFidelidade.Visible = False
      B_Ticket.Top = 90
  End If
  ' VENDA RÁPIDA (SOMENTE ESTA TELA) 195
  '***************************************
  
  
  
  If gParticipaProgramaFidelidade = 1 Then
    '1-SIM PARTICIPA;
    '0-NÃO PARTICIPA Empresa/filial;
    B_programaFidelidade.Enabled = True
  Else
    B_programaFidelidade.Enabled = False
  End If
    
  '17/01/2006 - mpdea
  'Centraliza tela de Venda Rápida normal
  If g_frmVendaRapida Is frmVendaRap2 Then
  
    If gbUsuarioAcessoApenasTelaVendaRapida = True Then
       Me.Top = 400
       Me.Left = 0
    Else
      Call CenterForm(Me)
    End If
  End If
  
  
  DropDown1.Columns(2).Visible = gbSuperUser

  KeyPreview = True
  Me.ActiveBar1.Refresh
  Screen.MousePointer = vbHourglass
    
    
  '03/04/2006 - mpdea
  'Solicitante: PSI Technomax - Rodrigo
  'Verifica o uso da gaveta em Venda Rápida e inicializa
  If g_blnUsaGavetaVendaRapida() Then Call InicializaGaveta
    
    
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  Data6.DatabaseName = gsQuickDBFileName
  dataPrestadores.DatabaseName = gsQuickDBFileName '10/05/2024 - Pablo
  
  '22/10/2004 - Daniel
  'Flexibilidade de troca de caixa
  'Case: Solicitado por Casagrande
  datCaixa.DatabaseName = gsQuickDBFileName
  
  '10/02/2006 - mpdea
  'Verifica utilização de vários caixas
  cboCaixa.Enabled = gbCaixas
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial")
  
  '01/05/2024 - Pablo
  '===================================================================================================
  cboPrestador.Visible = rsParametros("comPrestServ")
  cboPrestador.Enabled = rsParametros("comPrestServ")
  Apelido_Prestador.Visible = rsParametros("comPrestServ")
  Apelido_Prestador.Enabled = rsParametros("comPrestServ")
  lblPrestador.Visible = rsParametros("comPrestServ")
  lblPrestador.Enabled = rsParametros("comPrestServ")
  If rsParametros("comPrestServ") Then
    Dim sqlPrest As String
    sqlPrest = "SELECT 0 AS cod, '' AS apelido, '' AS nome FROM ZZZ UNION ALL " & _
               "SELECT f.Código AS cod, f.Apelido AS apelido, f.Nome AS nome " & _
               "FROM Funcionários AS f " & _
               "WHERE (((f.Ativo)=True) AND ((f.Liberado)=True) AND ((f.isPrestServ)=True));"
    'sqlPrest = "SELECT f.Código AS cod, f.Apelido AS apelido, f.Nome AS nome " & _
    '           "FROM Funcionários AS f " & _
    '           "WHERE (((f.Ativo)=True) AND ((f.Liberado)=True) AND ((f.isPrestServ)=True));"
    dataPrestadores.RecordSource = sqlPrest
  End If
  '===================================================================================================
  
  
  '24/09/2003 - mpdea
  'Criado índice para agilizar pesquisa
  datSequencias.DatabaseName = gsQuickDBFileName
  Dim sqlSequencias As String
  sqlSequencias = " SELECT s.Sequência, s.Cliente, s.Referência, s.Total " & _
                  " FROM Saídas AS s INNER JOIN [Operações Saída] AS os ON os.Código = s.Operação " & _
                  " WHERE s.Filial = " & gnCodFilial & " AND s.Efetivada = False " & _
                  " AND s.Data = #" & Format(Data_Atual, "mm/dd/yyyy") & "#"
  If rsParametros("VR_OcultaOrc").Value Then
    sqlSequencias = sqlSequencias & " AND os.Tipo <> 'O' "
  End If
  sqlSequencias = sqlSequencias & " ORDER BY s.Sequência DESC "
  datSequencias.RecordSource = sqlSequencias
  
  Data1.RecordSource = "Con_Cliente"
  
  Desconto_Cli = 0
   
  
  Set rsProdutos2 = modQSGeral.rsProdutos.Clone
  
  Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsOp_Saída = db.OpenRecordset("Operações Saída", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsUsuarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsSaidas = db.OpenRecordset("Saídas")
  Set rsSaída_Prod = db.OpenRecordset("Saídas - Produtos")
  Set rsSaída_Cheques = db.OpenRecordset("Movimento - Cheques")
  Set rsSaída_Parcelas = db.OpenRecordset("Movimento - Parcelas")
  Set rsTabelas = db.OpenRecordset("Tabela de Preços", , dbReadOnly)
  Set rsCotacoes = db.OpenRecordset("Cotações", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
  Set rsEstados = db.OpenRecordset("Estados", , dbReadOnly)
  Set rsCartoes = db.OpenRecordset("Cartões", , dbReadOnly)
  Set rsLog = db.OpenRecordset("ZZZLog")
  
  '11/12/2009 - Andrea
  Set rsSaída_Cartoes = db.OpenRecordset("Movimento - Cartoes")
  
  '20/12/2006 - Anderson - Registro de CFOP por produto e serviço
  Set rsProdutoCFOP = db.OpenRecordset("ProdutoCFOP", , dbReadOnly)
  
  '19/10/2007 - Anderson
  'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", gnUserCode
  
  If Not rsFuncionarios.NoMatch Then
    m_bolLucroMinimoPermitido = rsFuncionarios("LucroMinimoPermitido")
  Else
    m_bolLucroMinimoPermitido = False
  End If

 
  Ret_Arquivo = GetSetting("QuickStore", "ConfigVR", "Scanner", False)
  ActiveBar1.Tools("miOpLeitorOtico").Checked = CBool(Ret_Arquivo)
      
  Ret_Arquivo = GetSetting("QuickStore", "ConfigVR", "Limpar Tela Automatico", False)
  ActiveBar1.Tools("miOpClearAfterVenda").Checked = CBool(Ret_Arquivo)
      
  Ret_Arquivo = GetSetting("QuickStore", "ConfigVR", "Etiqueta Balanca", False)
  ActiveBar1.Tools("miOpEtiquetas").Checked = CBool(Ret_Arquivo)
      
  Ret_Arquivo = GetSetting("QuickStore", "ConfigVR", "Mantem Vendedor", False)
  ActiveBar1.Tools("miOpFreezeVendedor").Checked = CBool(Ret_Arquivo)
      
  Screen.MousePointer = vbDefault
  
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Funcionario
  
  If Not rsFuncionarios.NoMatch Then
    L_Preço.Visible = rsFuncionarios.Fields("VRVisualizarPreco")
    L_Estoque.Visible = rsFuncionarios.Fields("VRVisualizarEstoque")
  End If
  
  
  '------------------------------------------------------------------------------
  '29/08/2003 - mpdea
  'Verifica permissão para Tela Cadastro de Clientes
  ActiveBar1.Bands("mnuOpcoes").Tools("miOpCadastraCliente").Enabled = rsFuncionarios.Fields("Clientes").Value Or rsFuncionarios.Fields("Superusuário").Value
  'Verifica permissão para achar venda
  ActiveBar1.Bands("mnuOpcoes").Tools("miOpFindVenda").Enabled = rsFuncionarios.Fields("PermiteAcharVenda").Value Or rsFuncionarios.Fields("Superusuário").Value
  Número.Enabled = rsFuncionarios.Fields("PermiteAcharVenda").Value Or rsFuncionarios.Fields("Superusuário").Value
  
  cmd_acharVenda.Enabled = rsFuncionarios.Fields("PermiteAcharVenda").Value Or rsFuncionarios.Fields("Superusuário").Value
  '------------------------------------------------------------------------------
  
  '------------------------------------------------------------------------------
  '25/03/2004 - Daniel
  'Implementação feita para evitar gravação adulterada de usuário sem permissão
  'Case: Casagrande
  If CheckSerialCaseMod("QS40485-308", "QS39938-203", "QS39939-287", "QS40322-497") Then
  
    Dim rstAcessos As Recordset
 
    Set rstAcessos = db.OpenRecordset("SELECT * FROM Acessos WHERE Usuário =" & gnUserCode & " AND Numero =" & 25, dbOpenDynaset)
    
      With rstAcessos
        If Not (.BOF And .EOF) Then
          m_blnUserDanger = (.Fields("Gravar").Value = False Or .Fields("Apagar").Value = False)
        End If
        .Close
      End With
    
    Set rstAcessos = Nothing
  
    If m_blnUserDanger Then
      B_Recebe_Simples.Visible = True
      B_Grava.Enabled = False
      B_Grava_Recebe.Enabled = False
    End If
  
  End If
  '------------------------------------------------------------------------------
  
  '------------------------------------------------------------------------------
  '05/05/2004 - Daniel
  'Personalização Embalavi
  m_blnEmbalavi = CheckSerialCaseMod("QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162")
  
  With Grade1
    If g_bln5CasasDecimais Then
      .Columns("Preço").NumberFormat = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Columns("Preço").NumberFormat = "##,###,##0.000"
    Else
      .Columns("Preço").NumberFormat = "##,###,##0.00"
    End If
    '.Columns("Total").NumberFormat = "##,###,##0.00"
  End With
  
  With DropDown1
    If g_bln5CasasDecimais Then
      .Columns("Preço").NumberFormat = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Columns("Preço").NumberFormat = "##,###,##0.000"
    Else
      .Columns("Preço").NumberFormat = "##,###,##0.00"
    End If
  End With
  
  With DropDown2
    If g_bln5CasasDecimais Then
      .Columns("Preço").NumberFormat = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Columns("Preço").NumberFormat = "##,###,##0.000"
    Else
      .Columns("Preço").NumberFormat = "##,###,##0.00"
    End If
  End With
  '------------------------------------------------------------------------------
  
  '20/10/2004 - Daniel
  'Case.......: A.S. Wijman
  'Finalidade.: Tratamento para o campo [Saídas - Produtos].[Preço Final]
  m_blnASWijmaBelem = CheckSerialCaseMod("QS39881-068", "QS40377-377")
  '------------------------------------------------------------------------------
  
  '------------------------------------------------------------------------------
  '09/11/2004 - Daniel
  'Case: Cliente Teknika
  m_blnTeknika = CheckSerialCaseMod("QS40966-243")
  '------------------------------------------------------------------------------
  
  '------------------------------------------------------------------------------
  '01/12/2004 - Daniel
  'Case: De Mais (Nazareno)
  m_blnDeMais = CheckSerialCaseMod("QS31735-849")
  '------------------------------------------------------------------------------
  
  '------------------------------------------------------------------------------
  '21/05/2004 - Daniel
  'Flag de indicações para Personalizações da Bic Amazônia
  m_blnBic = CheckSerialCaseMod("QS35509-939", "QS37715-731")
  '------------------------------------------------------------------------------

  '30/01/2007 - Anderson - Alterado para que a permissão de visualizar estoque funcione para diversos clientes
  '------------------------------------------------------------------------------
  '26/08/2004 - Daniel
  'Criado validação para verificar se o usuário possui permissão para enchergar
  'o estoque ou não
  'Case: Tendresse
  'If CheckSerialCaseMod("QS37234-796", "QS37416-794") Then
  '  m_blnTendresse = True
    Call EnchergarEstoque
  'End If
  '------------------------------------------------------------------------------
  
  '06/05/2005 - Daniel
  '
  'Implementação.: Trabalhar com o código para fornecedor cadastrado na tela de produtos.
  '                Impacto: Ao entrar com o código para o fornecedor no campo código do produto
  '                o sistema deverá trazer o código do produto que estiver amarrado nele
  'Solicitação...: Cristiano Pavinato - PSI RS
  m_blnUsaCodFornec = g_blnVerificarUsoCodFornece
  '-------------------------------------------------------------
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then
      MsgBox "Filial não encontrada", vbCritical, "Erro"
      Exit Sub
  End If
  
  'Verifica se a coluna CFOP deve ser exibida na grade
  If rsParametros("ExibeCFOP") = True Then
     Grade1.Columns("CFOP").Visible = True
  Else
     Grade1.Columns("CFOP").Visible = False
  End If
  
  Num_Registro = Null
  
  Estado = ""
  
  Primeira_Vez = True
  
  ' Pilatti Novembro/2017 - comentei o primeiro select e criei o segundo (abaixo)
'''  Data5.RecordSource = "SELECT DISTINCT Tabela FROM Preços WHERE UCASE(Tabela) Not Like '*CUSTO*' ORDER BY Tabela"
'''  Data5.RecordSource = "SELECT DISTINCT P.Tabela FROM Preços P, AcessoTabelasDePrecosProdutos A WHERE UCASE(P.Tabela) Not Like '*CUSTO*' AND P.Tabela = A.Tabela AND A.Usuario = " & gnUserCode & " ORDER BY P.Tabela"
  Data5.RecordSource = "SELECT DISTINCT P.Tabela FROM Preços P, AcessoTabelasDePrecosProdutos A WHERE UCASE(P.Tabela) Not Like '*CUSTO*' AND P.Tabela = A.Tabela AND A.Usuario = " & Funcionario & " ORDER BY P.Tabela"
  ' Pilatti fim
  
  Data5.Refresh
    
  
 rsPreços.Index = "Só Tabela"

' Aux = ""
' Do
'  rsPreços.Seek ">", Aux
'  If Not rsPreços.NoMatch Then
'   Aux = rsPreços("Tabela")
'   If rsPreços("Tabela") <> "CUSTO" Then
'    Combo_Preço.AddItem rsPreços("Tabela")
'   End If
'  End If
' Loop Until rsPreços.NoMatch
'
 
 
 
 'Finalidade...: Deixamos configurável em Parâmetros à exibição
  '               nas telas de Saída e Venda Rápida da coluna Fabricante
  '               nos dropdowns de pesquisas
  '
  If rsParametros("ExibirFabricante").Value Then
    m_blnExibirColunaFabricante = True
    'DropDown1
    DropDown1.Columns("Fabricante").Visible = True
    'DropDown1.Columns("Nome").Width = 4004.788
    'DropDown2
    DropDown2.Columns("Fabricante").Visible = True
    'DropDown2.Columns("Nome").Width = 4004.788
  Else
    m_blnExibirColunaFabricante = False
    'DropDown1
    DropDown1.Columns("Fabricante").Visible = False
    'DropDown1.Columns("Nome").Width = 5500
    DropDown1.Refresh
    'DropDown2
    DropDown2.Columns("Fabricante").Visible = False
    'DropDown2.Columns("Nome").Width = 5500
    DropDown2.Refresh
  End If
  '----------------------------------------------------------------------
 
  If rsParametros.Fields("VR_Tela_CheckOut").Value Then
      Me.BorderStyle = 0
      ActiveBar1.Detach
  End If

 
  '07/05/2003 - mpdea
  'Desconto rateado
  m_blnDescontoRateado = rsParametros.Fields("DescSubTotalRateado").Value
    
  '07/05/2003 - mpdea
  'Objetos para Desconto rateado
  lblSubTotal.Visible = Not m_blnDescontoRateado
  txtSubTotal.Visible = Not m_blnDescontoRateado
  lblDescSubTotal.Visible = Not m_blnDescontoRateado
  txtDescSubTotal.Visible = Not m_blnDescontoRateado
 
  
  '29/05/2003 - mpdea
  'Utilização do Traffic Light
  m_blnWorkTrafficLight = rsParametros.Fields("WorkTrafficLight").Value

  '----------------------------------------------------------------------------------
  '29/05/2003 - mpdea
  'Atualizado
  '
  '05/08/2002 - mpdea
  'Objeto TrafficLight
  If m_blnWorkTrafficLight Then
    On Error Resume Next
    
    Set TrafficLight = New IQMod.TrafficLight
    TrafficLight.PathTarget = gsDefaultPath
    
    If Err.Number <> 0 Then
      m_blnWorkTrafficLight = False
      MsgBox "Componente 'Traffic Light' não instalado.", vbCritical, "Erro"
    End If
    
    On Error GoTo ErrHandler
  End If
  '----------------------------------------------------------------------------------



  '07/10/2002 - mpdea
  'Modificado verificação de estoque para acesso direto ao recordset
'   Verifica_Estoque = Not rsParametros("Venda Sem Estoque")
 
 Me.Caption = "Venda Rápida - Caixa: " + UCase(gsVendedorVR)
 
 
 'Vê se tem a tabela
 '
 '28/12/2004 - Daniel
 'Estava dando o erro 13 - Type mismatch quando este
 'campo rsParametros("VR Tab Preço") estava nulo
 'Solução para o problema...: Adicionamos o & ""
 'Erro encontrado por.......: Speed Auto Peças
 
 
  '
  If rsParametros("VR Tab Preço") <> "" Then
    Dim rsAcessosTabPrecoUsu As Recordset
    Dim iTemTabelasPreco As Integer
    Dim sSql As String
  
    iTemTabelasPreco = 0
  
    sSql = "Select Tabela From AcessoTabelasDePrecosProdutos Where Usuario=" & Funcionario & " And Tabela='" & rsParametros("VR Tab Preço") & "' "
  
    Set rsAcessosTabPrecoUsu = db.OpenRecordset(sSql, dbOpenDynaset)
  
    If Not (rsAcessosTabPrecoUsu.EOF And rsAcessosTabPrecoUsu.BOF) Then
        iTemTabelasPreco = 1
        Combo_Preço.Text = rsParametros("VR Tab Preço") & ""
        rsPreços.Seek "=", rsParametros("VR Tab Preço") & ""
        If rsPreços.NoMatch Then Combo_Preço.Text = ""
    Else
        iTemTabelasPreco = 0
    End If
    rsAcessosTabPrecoUsu.Close
    Set rsAcessosTabPrecoUsu = Nothing
  End If
  '
''' Combo_Preço.Text = rsParametros("VR Tab Preço") & ""
'' rsPreços.Seek "=", rsParametros("VR Tab Preço") & ""
'' If rsPreços.NoMatch Then Combo_Preço.Text = ""
 
 
 
 If rsParametros("VR Altera Tabela") = False Then Combo_Preço.Enabled = False
 If Combo_Preço.Text = "" Then Combo_Preço.Enabled = True
 
 
  '12/09/2003 - mpdea
  'Validação para o estado de SC
'  If rsParametros("VR Altera Preço") = False Then Grade1.Columns(3).Locked = True
  If UCase(gstrGetEstadoFilial(gnCodFilial)) = "SC" Then
    Grade1.Columns(3).Locked = True
  Else
    Grade1.Columns(3).Locked = Not rsParametros.Fields("VR Altera Preço").Value
  End If
  
  '09-07-2015 Jean Ricardo Zanella - Função para verificar se usuario tem permissão para alterar preços
  If blnPermissaoAlterarPrecos(Funcionario) = False Then
    Grade1.Columns(3).Locked = True
    'If Combo_Preço.Text = "" Then
      'Combo_Preço.Enabled = True
    'Else
      'Combo_Preço.Enabled = False
    'End If
  End If

 
  Combo_Cliente.Text = rsParametros("VR Cliente")
  
  '05/02/2004 - mpdea
  'Executa evento LostFocus do controle para validação das informações
  Combo_Cliente_LostFocus
  
  If rsParametros("VR Altera Cliente") = False Then
    Combo_Cliente.Enabled = False
  End If
  
  If rsParametros("VR Cadastra Cliente") = False Then
    ActiveBar1.Tools("miOpCadastraCliente").Enabled = False
  End If
  
  If rsParametros("VR Permite Desconto") = False Then
   Grade1.Columns(4).Locked = True
  End If
 
 
 
 Grade1.Rows = rsParametros("VR Linhas Digitação")
 Linhas_Grade = rsParametros("VR Linhas Digitação")
  
  
  '26/07/2004 - mpdea
  'Limpa a tela ao carregar o form
  Call Limpa_Tela(1)
 
 
 rsOp_Saída.Index = "Código"
 rsOp_Saída.Seek "=", rsParametros("VR Código Operação")
 If rsOp_Saída.NoMatch Then
   Beep
   Beep
   MsgBox "Operação de venda não encontrada", vbExclamation, "Aviso"
   
   '25/10/2002 - mpdea
   'Desabilitado controles que podem provocar erro
   Me.KeyPreview = False
   Grade1.Enabled = False
   B_Limpa.Enabled = False
   'Modificado propriedade de Visible para Enabled (padrão)
   B_Grava.Enabled = False
   B_Grava_Recebe.Enabled = False
  ' Unload Me
   Exit Sub
 End If
 Calcula_ICM = rsOp_Saída("ICM")
 Calcula_IPI = rsOp_Saída("IPI")
 gbBaseICMSomadoIPI = rsOp_Saída("Base ICM com IPI")
 
  '11/11/2008 - mpdea
  m_blnSomaIcmsRetidoTotalNota = rsOp_Saída.Fields("SomaIcmsRetidoTotalNota").Value
 
 Rem Configurações do Recebimento
 If rsParametros("VR Permite Rec Rápido") = False And rsParametros.Fields("VR_Tela_CheckOut").Value = False Then
   Frame_Recebimento.Visible = False
   'Grade1.Height = 3720
 Else
   If rsParametros("VR Permite Dinheiro") = False Then
     L_Dinheiro.Visible = False
     Dinheiro.Visible = False
   End If
   
   If rsParametros("VR Permite Vales") = False Then
     L_Vale.Visible = False
     Vale.Visible = False
   End If
   
   If rsParametros("VR Permite Cartão") = False Then
     L_Cartão.Visible = False
     Combo_Cartão.Visible = False
     Nome_Cartão.Visible = False
     Num_Cartão.Visible = False
     Val_Cartão.Visible = False
     Label13.Visible = False
     Label12.Visible = False
   End If
   
   
   Rem Cheques
   If rsParametros("VR Permite Cheques") = False Then
    For i = 0 To 4
       Banco(i).Visible = False
       Cheque(i).Visible = False
       Bom_Para(i).Visible = False
       Val_Cheque(i).Visible = False
       L_Cheque(i).Visible = False
     Next i
   End If
   If rsParametros("VR Permite Cheques") = True Then
     Aux_Int = rsParametros("VR Qtde Cheques")
     If Aux_Int <> 5 Then
       For i = Aux_Int To 4
         Banco(i).Visible = False
         Cheque(i).Visible = False
         Bom_Para(i).Visible = False
         Val_Cheque(i).Visible = False
       Next i
     End If
   End If
  
  
   Rem Parcelas
   If rsParametros("VR Permite Parcela") = False Then
    For i = 0 To 4
       Data_Parc(i).Visible = False
       Val_Parc(i).Visible = False
    Next i
    L_Parc1.Visible = False
    L_Parc2.Visible = False
    L_Parc3.Visible = False
    Tipo_Parc.Visible = False
   End If
   If rsParametros("VR Permite Parcela") = True Then
     Aux_Int = rsParametros("VR Qtde Parcela")
     If Aux_Int <> 5 Then
       For i = Aux_Int To 4
         Data_Parc(i).Visible = False
         Val_Parc(i).Visible = False
       Next i
     End If
     If rsParametros("VR Parcela Padrão") = "B" Then O_Banco.Value = True
     If rsParametros("VR Parcela Padrão") = "C" Then O_Carteira.Value = True
     If rsParametros("VR Parcela Padrão") = "E" Then O_Carnet.Value = True
     If rsParametros("VR Altera Parcela") = False Then Tipo_Parc.Enabled = False
   End If
  
  
 End If
 
 Operador_Caixa = 0
 
  Call ActiveBarLoadToolTips(Me)
 
  'Teste
  cmdInsertItens.Visible = gbTeste
  
  
  '22/01/2003 - mpdea
  'Quick em modo limitado
  If Not gblnQuickFull Then
'    B_Nota.Visible = False
    
    '27/01/2004 - mpdea
    'Botão de ticket agora está disponível
    '09/06/2004 - Daniel
    'Setando igual a True
    B_Ticket.Visible = True
    
    With ActiveBar1
      .Bands("mnuBand1").Tools("miEmissoes").Visible = False
      .RecalcLayout
      .Refresh
    End With
  End If
  
  '05/07/2013-Alexandre Afornali
  If (rsParametros("TrabalharComComanda") = -1) Then
    txtComanda.Visible = True
    lblComanda.Visible = True
  Else
    txtComanda.Visible = False
    lblComanda.Visible = False
  End If
 'Teste de erro de data de Saídas
 Erro_Data = False
 rsSaidas.Index = "Data"
 rsSaidas.Seek ">", gnCodFilial, CDate(Data_Atual + 1), 0
  
 
  If rsSaidas.NoMatch Then
      If gbUsuarioAcessoApenasTelaVendaRapida = True Then
          If Funcionario = "" Then
              Combo_Vendedor.Text = gnUserCode
          Else
              Combo_Vendedor.Text = Funcionario
          End If
          
'''          Combo_Vendedor_LostFocus
'''          Combo_Vendedor.Enabled = False
      End If
    
      Exit Sub
  End If
 
 If rsSaidas("Filial") <> gnCodFilial Then Exit Sub
 
 Erro_Data = True
 Erro_Data2 = True
 

  If gbUsuarioAcessoApenasTelaVendaRapida = True Then
'''      Combo_Vendedor_LostFocus
'''      Combo_Vendedor.Enabled = False
  Else
      Combo_Vendedor.Text = gnUserCode
      Combo_Vendedor_LostFocus
  End If
  Exit Sub
 
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

'17/01/2006 - mpdea
'Redimensionamento das colunas do grid de acordo com a proporção total
Private Sub Form_Resize()
  Dim sngGridWidth As Single
  
  On Error GoTo ErrHandler
  
  If rsParametros.Fields("VR_Tela_CheckOut").Value Then
  
      Frame_Recebimento.Visible = False
  
      If cmd_opcoes.BackColor = &HFFC0C0 Then
          Me.Width = 17610
          Me.Height = 5400 + 280
          Me.BackColor = &HFFA324
      Else
          Me.Width = 17610
          Me.Height = 5400
          Me.BackColor = &HFFA324
      End If
      
      ' Dados operador
      Label8.Visible = False
      Cod_Operador.Visible = False
      Nome_Operador.Visible = False
      
      ' Dados Combo sequencia
      Label6.Visible = False
      Número.Visible = False
      
      ' Dados ref. interna
      Label10.Visible = False
      Referência.Visible = False
      
      ' Dados tab. precos
      Label1.Left = 45
      Label1.Top = 90
      Label1.ForeColor = &HFFFFFF
      Combo_Preço.Top = 60
      Combo_Preço.Width = 4260
      
      ' Dados combo caixa
      Label9.Left = 5600
      Label9.Top = 90
      Label9.ForeColor = &HFFFFFF
      cboCaixa.Left = 6200
      cboCaixa.Top = 60
      Nome_Caixa.Left = 7750
      Nome_Caixa.Top = 60
      Nome_Caixa.Width = 4300
      
 
      ' Dados vendedor
      Label7.Left = 45
      Label7.Top = 550
      Label7.ForeColor = &HFFFFFF
      Combo_Vendedor.Left = 1200
      Combo_Vendedor.Top = 500
      Combo_Vendedor.Width = 900
      Nome_Vendedor.Left = 2150
      Nome_Vendedor.Top = 500
      Nome_Vendedor.Width = 3320
      
      ' Dados cliente
      Label5.Left = 5600
      Label5.Top = 550
      Label5.ForeColor = &HFFFFFF
      Combo_Cliente.Left = 6200
      Combo_Cliente.Top = 510
      Nome_Cliente.Left = 7750
      Nome_Cliente.Top = 510
      Nome_Cliente.Width = 4300
      
      ' Botão fecha tela
      cmd_fecharTela.Visible = True
      cmd_fecharTela.Left = 13050
      cmd_fecharTela.Top = 45
      cmd_fecharTela.Width = 1260
      
      ' Grade
      Grade1.Top = 1860
      Grade1.Height = 3195
      Grade1.Width = 15255
      Grade1.RowHeight = 500
      
      ' label contador Quantidade produtos
      lblTitleQtdeTotal.Left = 100
      lblTitleQtdeTotal.Top = 4250
      lblTitleQtdeTotal.BorderStyle = 0
      lblQtdeTotal.Left = 2520
      lblQtdeTotal.Top = 4250
      lblQtdeTotal.Width = 1100
      lblQtdeTotal.BorderStyle = 0
      
      L_Preço.Visible = False
      L_Estoque.Visible = True
'''      L_Preço.Left = 4000
'''      L_Preço.Top = 4250
      L_Estoque.Width = 2750
      L_Estoque.Left = 4300
      L_Estoque.Top = 4250
      
      ' Botoes lateral
      fraButtons.Top = 45
      fraButtons.Left = 14500
      fraButtons.Width = 3400
      fraButtons.Height = 5500
      fraButtons.BackColor = &HFFA324
      
      ' Botões move tela
      cmd_esquerda.Visible = True
      cmd_esquerda.Width = 900
      cmd_esquerda.Height = 880
      
      cmd_direita.Visible = True
      cmd_direita.Width = 900
      cmd_direita.Height = 880
      
      cmd_Acima.Visible = True
      cmd_Acima.Width = 1100
      cmd_Acima.Height = 435
      
      cmd_abaixo.Visible = True
      cmd_abaixo.Width = 1100
      cmd_abaixo.Height = 435
      
      cmd_esquerda.Left = 1
      cmd_esquerda.Top = 1
      
      cmd_Acima.Left = 970
      cmd_Acima.Top = 1
      
      cmd_abaixo.Left = 970
      cmd_abaixo.Top = 450
      
      cmd_direita.Left = 2150
      cmd_direita.Top = 1
      
      
      ' linha 1 de botões
      B_Grava.Left = 1
      B_Grava.Top = 920
      B_Grava.Width = 1500
      B_Grava.Height = 720
      
      B_Desconto.Left = 1540
      B_Desconto.Top = 920
      B_Desconto.Width = 1500
      B_Desconto.Height = 720
      B_Desconto.Font = "Tahoma, 8, Normal"
      
      ' linha 2 de botões
      B_Recebe.Left = 1
      B_Recebe.Top = 1720
      B_Recebe.Width = 1500
      B_Recebe.Height = 720
      
      B_programaFidelidade.Left = 1540
      B_programaFidelidade.Top = 1720
      B_programaFidelidade.Width = 1500
      B_programaFidelidade.Caption = "Pontos"
      B_programaFidelidade.Height = 720
      
      ' linha 3 de botões
      B_Grava_Recebe.Left = 1
      B_Grava_Recebe.Top = 2520
      B_Grava_Recebe.Width = 3030
      B_Grava_Recebe.Caption = "Gravar / Receber"
      B_Grava_Recebe.Height = 720
      
      ' linha 4 de botões
      B_NFC_e.Left = 1
      B_NFC_e.Top = 3320
      B_NFC_e.Width = 1500
      B_NFC_e.Height = 720
      B_NFC_e.Caption = "NFC-e"
      
      B_Ticket.Left = 1540
      B_Ticket.Top = 3320
      B_Ticket.Width = 1500
      B_Ticket.Height = 720

      ' linha 5 de botões
      B_Limpa.Left = 1540
      B_Limpa.Top = 4120
      B_Limpa.Width = 1500
      B_Limpa.Height = 720
      
      ' Dados Sub total / Desc. / Total
      lblSubTotal.Left = 11600
      lblSubTotal.Top = 4270
      lblSubTotal.BackColor = &HFFA324
      lblSubTotal.ForeColor = &HFFFFFF
      txtSubTotal.Left = 12500
      txtSubTotal.Top = 4270
      
      lblDescSubTotal.Left = 11600
      lblDescSubTotal.Top = 4620
      lblDescSubTotal.BackColor = &HFFA324
      lblDescSubTotal.ForeColor = &HFFFFFF
      txtDescSubTotal.Left = 12500
      txtDescSubTotal.Top = 4620
      
      Label4.Left = 11600
      Label4.Top = 4970
      Label4.BackColor = &HFFA324
      Label4.ForeColor = &HFFFFFF
      L_Tot_Pagar.Left = 12500
      L_Tot_Pagar.Top = 4970
      
      'lblComanda.Left = 12200
      'lblComanda.Top = 310
      'lblComanda.ForeColor = &HFFFFFF
      'txtComanda.Visible = True
      'txtComanda.Left = 12200
      'txtComanda.Top = 550
      'txtComanda.Width = 2100
      'txtComanda.BackColor = &HC0FFFF
      
      ' Status
      Efetivada.Left = 9000
      Efetivada.Top = 4280
      
      lbl_retornoEnvioNFCe.Width = 2950
      lbl_retornoEnvioNFCe.Caption = "Imprimindo NFC-e"
      lbl_retornoEnvioNFCe.Left = 8500
      lbl_retornoEnvioNFCe.Top = 4900
      
      Movimentação_Desfeita.Left = 9000
      Movimentação_Desfeita.Top = 4900
      
'''      frm_produtoSemPrecoNaGrade.Left = 7400
'''      frm_produtoSemPrecoNaGrade.Top = 4280
      frm_produtoSemPrecoNaGrade.Width = 4150
      frm_produtoSemPrecoNaGrade.Height = 1100
      Label14.Left = 1
      Label14.Top = 1
      cmd_fecharFrameProdutoSemPrecoNaGrade.Top = 730
      cmd_fecharFrameProdutoSemPrecoNaGrade.Left = 1300
      cmd_fecharFrameProdutoSemPrecoNaGrade.Height = 350
      
      
      ' botão achar venda
      cmd_acharVenda.Left = 100
      cmd_acharVenda.Top = 4550
      cmd_acharVenda.Width = 1360
      cmd_acharVenda.Height = 720
      
      ' botao consulta tabelas de precos
      cmd_tabelaDePrecos.BackColor = 12648447
      cmd_tabelaDePrecos.Left = 1500
      cmd_tabelaDePrecos.Top = 4550
      cmd_tabelaDePrecos.Width = 1360
      cmd_tabelaDePrecos.Height = 720
      
      cmd_pesquisaAlfa.Visible = True
      cmd_pesquisaAlfa.Left = 2900
      cmd_pesquisaAlfa.Top = 4550
      cmd_pesquisaAlfa.Width = 1360
      cmd_pesquisaAlfa.Height = 720
      
      cmd_carne.Visible = True
      cmd_carne.Left = 4300
      cmd_carne.Top = 4550
      cmd_carne.Width = 1360
      cmd_carne.Height = 720
      
      cmd_carneComRecibo.Visible = True
      cmd_carneComRecibo.Left = 5700
      cmd_carneComRecibo.Top = 4550
      cmd_carneComRecibo.Width = 1360
      cmd_carneComRecibo.Height = 720
      
      cmd_opcoes.Visible = True
      cmd_opcoes.Left = 7100
      cmd_opcoes.Top = 4550
      cmd_opcoes.Width = 1360
      cmd_opcoes.Height = 720
    
   
      ' Altura Tela pesquisaProdutosAlfa = 4220
      ' Altura Tela vendaRapidaCheckOut  = 5400
      ' Total de                         = 9620
      
      If Screen.Height > 9620 Then
          Me.Top = ((Screen.Height - 9620) / 2) + 4300
          Me.Left = (Screen.Width - Me.Width) / 2
          Me.Show
      Else
          Me.Top = 3000
          Me.Left = (Screen.Width - Me.Width) / 2
          Me.Show
      End If

  End If
  
  
'''  '06/02/2006 - mpdea
'''  'Tratamentos para a tela de Venda Rápida (CheckOut)
'''  If g_frmVendaRapida Is frmVendaRap2_CheckOut Then
'''    If Me.WindowState <> vbMinimized Then
'''      DoEvents
'''
'''      sngGridWidth = Grade1.Width
'''
'''      With Grade1.Columns
'''        .Item("Código").Width = sngGridWidth * 0.1978!
'''        .Item("Qtde").Width = sngGridWidth * 0.0686!
'''        .Item("Nome").Width = sngGridWidth * 0.3593!
'''        .Item("Preço").Width = sngGridWidth * 0.1117!
'''        .Item("Desc.%").Width = sngGridWidth * 0.0821!
'''        .Item("Total").Width = sngGridWidth * 0.1265!
'''      End With
'''
'''      'Força a correta exibição do form após redimensionamento
'''      Me.Show
'''    End If
'''  End If
 
  Exit Sub
 
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

'17/01/2006 - mpdea
'Incluído tratamento de erro
Private Sub Form_Unload(Cancel As Integer)
  
  On Error GoTo ErrHandler
  
  
  '-------------------------------------------------------------------
  '29/05/2003 - mpdea
  'Atualizado
  '
  '05/08/2002 - mpdea
  'Objeto TrafficLight
  If m_blnWorkTrafficLight Then
    Set TrafficLight = Nothing
  End If
  '-------------------------------------------------------------------
  
  If gParticipaProgramaFidelidade = 1 Then
      '1-SIM PARTICIPA;
      '0-NÃO PARTICIPA Empresa/filial;
      
      If gClienteEntregouResgatePontos = True Then
          gClienteEntregouResgatePontos = False
          gSaldoCdGuidResgate = 0
          gCdGuidResgate = ""
          gCdClienteCdGuidResgate = 0
          gNmClienteCdGuidResgate = ""
      End If

  End If
  
  
  Unload frmRecebimento
  Set frmRecebimento = Nothing


  rsParametros.Close
  rsPreços.Close
  rsProdutos.Close
  rsOp_Saída.Close
  rsFuncionarios.Close
  rsUsuarios.Close
  rsCliFor.Close
  rsGrade.Close
  rsSaidas.Close
  rsSaída_Prod.Close
  rsSaída_Cheques.Close
  rsSaída_Parcelas.Close
  rsTabelas.Close
  rsCotacoes.Close
  rsEstoque.Close
  rsContas_Receber.Close
  rsEstados.Close
  rsCartoes.Close
  rsLog.Close
  
  '05/04/2010 - Andrea
  rsProdutoCFOP.Close
 
  Set rsParametros = Nothing
  Set rsPreços = Nothing
  Set rsProdutos = Nothing
  Set rsOp_Saída = Nothing
  Set rsFuncionarios = Nothing
  Set rsUsuarios = Nothing
  Set rsCliFor = Nothing
  Set rsGrade = Nothing
  Set rsSaidas = Nothing
  Set rsSaída_Prod = Nothing
  Set rsSaída_Cheques = Nothing
  Set rsSaída_Parcelas = Nothing
  Set rsTabelas = Nothing
  Set rsCotacoes = Nothing
  Set rsEstoque = Nothing
  Set rsContas_Receber = Nothing
  Set rsEstados = Nothing
  Set rsCartoes = Nothing
  Set rsLog = Nothing
 
  Set rsProdutoCFOP = Nothing
  
  '17/01/2006 - mpdea
  'Restaura tela principal do Quick Store
  If g_frmVendaRapida Is frmVendaRap2_CheckOut Then
    frmMain.WindowState = vbMaximized
  End If
  'Desassocia a tela de Venda Rápida
  Set g_frmVendaRapida = Nothing
  
  
  Exit Sub
 
ErrHandler:
  'Desassocia a tela de Venda Rápida
  Set g_frmVendaRapida = Nothing
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
 
End Sub

Private Sub Grade1_AfterColUpdate(ByVal ColIndex As Integer)
  Call Calcula_Linha
End Sub

Private Sub Grade1_AfterUpdate(RtnDispErrMsg As Integer)

  Call Recalcula
'  L_Preço.Caption = ""
'  L_Estoque.Caption = ""
 
End Sub

Public Sub Grade1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  Dim Aux2 As Variant
  Dim Cód As String
  Dim Cód1 As String
  Dim Cód_Str As String
  Dim Valor As Single
  Dim Valor_Int As Long
  Dim Aux_Str As String
  Dim Aux_Str2 As String
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
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  '10/11/2004 - Daniel
  Dim strUF As String
  
  '05/04/2010 - Andrea
  'Dim Aux_Produto As String

  '08/03/2007 - Anderson
  'Inclusão de código para resolver problema ao digitar um código do fornecedor igual ao código do produto
  Dim rstProdutos As Recordset
  Dim strSQL      As String
  
  '10/02/2010 - mpdea
  'Flag para quantidade inicializada (padrão 1)
  Dim bln_qtde_inicializada As Boolean
  
  
  Call StatusMsg("")
  
  B_Recebe.Enabled = False
  
  Aux = Trim(Grade1.Columns(ColIndex).Text)

  If ColIndex = 0 Then 'Código
  
    Grade1.Columns(0).Text = UCase(Grade1.Columns(0).Text)
    
    If IsNull(Aux) Or Aux = "" Or Aux = "0" Then
       Grade1.Columns(1).Text = 0
       Grade1.Columns(2).Text = ""
       Grade1.Columns(3).Text = 0
       Grade1.Columns(4).Text = 0
       Grade1.Columns(5).Text = 0
       Grade1.Columns(6).Text = 0 'ICM
       Grade1.Columns(7).Text = 0 'IPI
       '05/04/2010 - Andrea
       'Registro do CFOP por produto
       Grade1.Columns(8).Text = "" 'CFOP
       Exit Sub
    End If
  
    '---------------------------------------------------------------------------------------------
    '06/05/2005 - Daniel
    '
    'Implementação.: Trabalhar com o código para fornecedor cadastrado na tela de produtos.
    '                Impacto: Ao entrar com o código para o fornecedor no campo código do produto
    '                o sistema deverá trazer o código do produto que estiver amarrado nele
    'Solicitação...: Cristiano Pavinato - PSI RS
    If m_blnUsaCodFornec Then
      Dim strCodParaFornec As String

      '-----------------------
      '08/03/2007 - Anderson
      'Inclusão de código para resolver problema ao digitar um código do fornecedor igual ao código do produto
      strSQL = "SELECT Código, [Código do Fornecedor] FROM Produtos WHERE Código = '" & Aux & "'"
      
      Set rstProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
      
      If rstProdutos.RecordCount = 0 Then
        strCodParaFornec = Aux
      Else
        strCodParaFornec = rstProdutos("Código do Fornecedor")
      End If
      
      Set rstProdutos = Nothing
  
      If Not (strCodParaFornec = "0" Or strCodParaFornec = "") Then
        'strCodParaFornec = Aux
        Aux = g_strBuscarCodProd(strCodParaFornec)
        Grade1.Columns(0).Text = Aux
        
        '07/12/2006 - Anderson
        'Alterado pois causando problemas quando o código do produto fornecedor era igual ao código do produto
        'Se não existir nenhum produto amarrado
        'If Aux = strCodParaFornec Then Exit Sub
        If Aux = "" Then
          Aux = strCodParaFornec
          Exit Sub
        End If
      End If
    End If
    '---------------------------------------------------------------------------------------------
  
    Aux_Str = Aux
    '26/05/2004 - Daniel
    'Tratamento para 0 'zero' a esquerda
    If Not gbZeroEsquerda Then
      Aux = Retira_Zeros(Aux_Str)
    End If
    Grade1.Columns(0).Text = Aux
  
    Rem Rotina para verificar se é de balança
    Balança = False
    If ActiveBar1.Tools("miOpEtiquetas").Checked = True Then
      Aux_Str = Aux
      If Len(Aux_Str) >= 12 Then
        Aux_Str = Left$(Aux_Str, 1)
        If Aux_Str = "2" Then 'é produto pesável
          Balança = True
          Comp_Prod = rsParametros("Qtde Balança")
          If Comp_Prod = 3 Then Início_Preço = 5
          If Comp_Prod = 3 Then Tam_Preço = 8
          If Comp_Prod = 4 Then Início_Preço = 6
          If Comp_Prod = 4 Then Tam_Preço = 7
          If Comp_Prod = 5 Then Início_Preço = 7
          If Comp_Prod = 5 Then Tam_Preço = 6
          If Comp_Prod = 6 Then Início_Preço = 8
          If Comp_Prod = 6 Then Tam_Preço = 5
          Aux_Str = Aux
          Aux = Mid(Aux, 2, Comp_Prod)
          '26/05/2004 - Daniel
          'Tratamento para 0 'zero' a esquerda
          If Not gbZeroEsquerda Then
            Aux = Retira_Zeros(Trim(CStr(Aux)))
          End If
          Cód = Aux
'          If Cód = 0 Then
'             Exit Sub
'          End If
          Grade1.Columns(0).Text = Aux
          Preço_Balança = Val(Mid(Aux_Str, Início_Preço, Tam_Preço))
          Preço_Balança = Preço_Balança / 100
          'Exit Sub
        End If
      End If
    End If
    
    Cód = Trim(CStr(Aux))
    Tamanho = 0
    Cor = 0
    Edição = 0
    
    Cód_Str = Trim(Cód)
    
    Call Acha_Produto(Cód_Str, Cód, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)
    If Aux_Erro <> 0 Then
      If Aux_Erro = 1 Then DisplayMsg "Produto não encontrado."
      If Aux_Erro = 2 Then DisplayMsg "Produto com grade, digite tamanho e cor."
      If Aux_Erro = 3 Then DisplayMsg "Produto com edição, digite a edição."
      Cancel = True
      Exit Sub
    End If

    '05/04/2010 - Andrea
    'Alteração para o registro do CFOp por produto
    m_CodOper = rsParametros("VR Código Operação")
    If m_CodOper <> 0 Then
      rsProdutoCFOP.Index = "PrimaryKey"
      rsProdutoCFOP.Seek "=", Cód, m_CodOper
      If rsProdutoCFOP.NoMatch Then
        rsOp_Saída.Index = "Código"
        rsOp_Saída.Seek "=", m_CodOper
        If Not rsOp_Saída.NoMatch Then
          '15/03/2008 - mpdea
          'Corrigido RT-13 ao ler o código fiscal como nulo
          Grade1.Columns("CFOP").Text = rsOp_Saída("Código Fiscal") & ""
        End If
      Else
        Grade1.Columns("CFOP").Text = rsProdutoCFOP("CFOP") & ""
      End If
    End If


    Cód = Trim(Cód)
    
    If Balança = False Then
      If Grade1.Columns(1).Text = 0 Then
        Grade1.Columns(1).Text = 1
        '10/02/2010 - mpdea
        'Campo quantidade inicializado com valor padrão
        bln_qtde_inicializada = True
      End If
    End If
    
    rsProdutos.Index = "Código"
    rsProdutos.Seek "=", Cód
    If rsProdutos.NoMatch Then Exit Sub
    
    If rsProdutos.Fields("Desativado") Then
      MsgBox "Produto Inativo, verifique !", vbCritical, "Quick Store"
      'Grade1.Columns(0).Text = ""    27/12/2004 - Daniel
      'Grade1.Columns(1).Text = "0"   Antigo código
      
      '27/12/2004 - Daniel
      'BUG: Estava travando o usuário não permitindo
      'que ele saísse do campo Qtde
      'Notificação: De Mais Presentes (Nazareno)
      Grade1.Columns(1).Text = "0" 'Qtde
      Grade1.Columns(0).Text = "0" 'Código
      Cancel = True
      '---------------------------------------------
      
      Exit Sub
    End If
    
    
    '30/12/2003 - mpdea
    'Zera o desconto
    Grade1.Columns(4).Text = 0
    
    
    '-------------------------------------------------------------------------------
    '08/11/2002 - mpdea
    'Adicionado verificação de erro ao encontrar estoque
    '07/10/2002 - mpdea
    'Comentado dupla verificação de estoque
    'Adicionado verificação para produto que não controla estoque
    '09/10/2002 - mpdea
    'Adicionado verificação da operação (Movimenta estoque)
    ''''''''''''''''
    Estoque = Acha_Estoque(gnCodFilial, CStr(Cód), Tamanho, Cor, Edição, Aux_Erro)
     L_Estoque.Caption = "Estoque=" + CStr(Estoque)
    'Adicionada essa clausula para verificar quando o estoque, mesmo quando o foco está no
    'campo Código
      If Not rsParametros.Fields("Venda Sem Estoque").Value And _
         rsProdutos.Fields("Estoque").Value And _
         rsOp_Saída.Fields("Estoque").Value Then
         
'        If Verifica_Estoque = True Then
        If Aux_Erro = 0 Then
          If CDbl(Grade1.Columns(1).Text) > Estoque Then
            If Estoque <> -999999 Then
              '26/08/2004 - Daniel
              'Criado validação para verificar se o usuário possui permissão
              'para enchergar o estoque ou não
                
              '30/01/2007 - Anderson - Alterado para que a permissão de visualizar estoque funcione para diversos clientes
              'If Not m_blnPermitido And m_blnTendresse Then   'Não permitido
              
              '10/02/2010 - mpdea
              'Zera quantidade quando o produto for fracionado, a quantidade for inicializada automaticamente (padrão 1),
              'possuir estoque maior do que 0 e inferior a 1
              'Resolve questões para vendas de produtos fracionados que possuem estoque como 0,8
              If gbIsFrac(Cód, 0) And bln_qtde_inicializada And Estoque > 0 And Estoque < 1 Then
                Grade1.Columns(1).Text = "0"
              Else
                If Not m_blnPermitido Then   'Não permitido
                  DisplayMsg "Quantidade superior ao estoque."
                Else
                  DisplayMsg "Quantidade superior ao estoque. Estoque atual : " + Format(Estoque, "#########0")
                End If
                
                If CDbl(Grade1.Columns(1).Text) <> 0 Then Cancel = True
                Exit Sub
              End If
            End If
          End If
        Else
          If Aux_Erro = 1 Then
            DisplayMsg "Produto com estoque não inicializado."
          Else
            DisplayMsg "Erro [" & Aux_Erro & "] ao encontrar estoque do produto."
          End If
          Cancel = True
          Exit Sub
        End If
'        End If
      End If
    ''''''''''''''''
    '-------------------------------------------------------------------------------

    Grade1.Columns(2).Text = rsProdutos("Nome") & ""
    
    '------------------------------------------------------
    '23/05/2006 - mpdea
    'Comentado restrição de isenção de IPI para a Embalavi
    'É utilizado configuração do cadastro de clientes
    '
    '07/05/2004 - Daniel
    'Personalização Embalavi
    'Exatamente neste ponto que temos em mãos
    'o percentual do IPI do produto
    'Tratamento através da função IsencaoIPI para
    'verificação se suspende ou não a taxa de IPI conforme
    'o cliente e não o produto
'    If m_blnEmbalavi Then
'      If Len(Nome_Cliente.Caption) > 0 Then
'        If IsencaoIPI(CLng(Combo_Cliente.Text)) Then 'Cliente É Isento de IPI


'''        If m_blnIsentoIPI Then
'''          Grade1.Columns(7).Text = "0"
'''        Else
'''          Grade1.Columns(7).Text = rsProdutos("Percentual IPI")
'''        End If
        
        ' ==============================================================
        ' Tratar IPI
        If rsParametros("CodigoRegimeTributario") <> 1 Then
            If m_blnIsentoIPI Then
                Grade1.Columns(7).Text = "0"
            Else
                '''Grade1.Columns(7).Text = rsProdutos("Percentual IPI") 'saida
                Grade1.Columns(7).Text = "0"
            End If
        End If
        
        
'      Else 'Len...
'        Grade1.Columns(7).Text = rsProdutos("Percentual IPI")
'      End If
'
'    Else 'Não Embalavi
'      Grade1.Columns(7).Text = rsProdutos("Percentual IPI")
'    End If
    '------------------------------------------------------
    
    
    'Mostra ICM do Estado
    If Estado = "" Then
      Grade1.Columns(6).Text = rsProdutos("Percentual ICM Saida") & ""
    End If
    If Estado <> "" Then
      rsEstados.Index = "Estado"
      rsEstados.Seek "=", Estado
      If rsEstados.NoMatch Then Grade1.Columns(6).Text = rsProdutos("Percentual ICM Saida") & ""
      If Not rsEstados.NoMatch Then
         If rsEstados("ICM") = -1 Then
            Grade1.Columns(6).Text = rsProdutos("Percentual ICM Saida") & ""
         Else
         
             '10/11/2004 - Daniel
             'Tratamento do ICM solicitado pela Teknika
             If Not m_blnTeknika Then 'Demais clientes
               
                If m_blnEmbalavi Then
                
                  If Len(Combo_Cliente.Text) > 0 Then 'Está preenchido
                    If PessoaFisica(Combo_Cliente.Text) Then
                      Grade1.Columns(6).Text = rsProdutos("Percentual ICM Saida").Value & ""
                    Else
                      Grade1.Columns(6).Text = rsEstados("ICM").Value
                    End If
                    
                  Else 'Não terá como verificar sem saber o cliente
                    Grade1.Columns(6).Text = rsEstados("ICM").Value
                  End If
                
                Else 'Demais clientes
                  Grade1.Columns(6).Text = rsEstados("ICM").Value
                End If
             
             Else
        
                If IE_Isento(strUF) Then 'ISENTO = TRUE
                  If strUF = "PR" Then
                    Grade1.Columns(6).Text = rsProdutos("Percentual ICM Saida") & ""
                  Else
                    Grade1.Columns(6).Text = rsProdutos("Percentual ICM Saida") & ""
                  End If
                Else 'ISENTO = FALSE
                  If strUF = "PR" Then
                    Grade1.Columns(6).Text = rsProdutos("Percentual ICM Saida") & ""
                  Else
                    Grade1.Columns(6).Text = rsEstados("ICM")
                  End If
                End If
        
             End If
         
         End If
      End If
    End If
    
    If rsProdutos("Tipo ICM") = "I" Then
      Grade1.Columns(6).Text = "0"
    End If
    
    '<<<<
    With Grade1
      .Columns("Base_ICM").Text = 0
      .Columns("Valor_ICM").Text = 0
      .Columns("Valor_Base_Unit").Text = 0
      .Columns("Redução_ICM").Text = 0
      .Columns("Tipo_ICM").Text = rsProdutos("Tipo ICM") & ""
      Select Case rsProdutos("Tipo ICM")
        Case "I"
          .Columns("ICM").Text = "0"
        Case "R" 'ICM Retido
          If rsProdutos("Base Cálculo") <> 0 Then
            .Columns("Valor_Base_Unit").Text = rsProdutos("Base Cálculo")
          End If
          If rsProdutos("Redução ICM") <> 0 Then
            .Columns("Redução_ICM").Text = rsProdutos("Redução ICM")
          End If
        Case "Z" 'ICM Reduzido
          If rsProdutos("Base Cálculo") <> 0 Then
            .Columns("Valor_Base_Unit").Text = rsProdutos("Base Cálculo")
          End If
          If rsProdutos("Redução ICM") <> 0 Then
            .Columns("Redução_ICM").Text = rsProdutos("Redução ICM")
          End If
      End Select
    End With
    '>>>>
    
    
    ' *********************************************
    ' AJUSTE ABRIL/19 PARA TRATAMENTO DE VALOR ACATADO NA TELA DE PESQUISA DE PRODUTO
    If gTabelaPrecoAcatadaTelaPesquisaProduto <> "" Then
        rsPreços.Index = "Tabela"
        rsPreços.Seek "=", gTabelaPrecoAcatadaTelaPesquisaProduto, Cód
    Else
        rsPreços.Index = "Tabela"
        rsPreços.Seek "=", Combo_Preço.Text, Cód
    End If
    
    gTabelaPrecoAcatadaTelaPesquisaProduto = ""
        
    'Acha preço
''''''''''        rsPreços.Index = "Tabela"
''''''''''        rsPreços.Seek "=", Combo_Preço.Text, Cód
    If rsPreços.NoMatch Then
       Grade1.Columns(3).Text = 0
    End If
    If Not rsPreços.NoMatch Then
           
        '----------------------------------------------------------------------------
        '05/02/2004 - mpdea
        'Verifica permissão de desconto no produto
        If rsProdutos.Fields("DontAllowDesc").Value Then
            '05/05/2004 - Daniel
            'Personalização Embalavi
            'Tratamento de Máscara
            If g_bln5CasasDecimais Then
                Aux_Preço = (Format((rsPreços("Preço")), "##,###,##0.00000")) * ((100 - (rsProdutos("Desconto"))) / 100)
                '30/04/2007 - Anderson - Implementação de 3 casas decimais
            ElseIf g_bln3CasasDecimais Then
                Aux_Preço = (Format((rsPreços("Preço")), "##,###,##0.000")) * ((100 - (rsProdutos("Desconto"))) / 100)
            Else
                Aux_Preço = rsPreços("Preço") * ((100 - (rsProdutos("Desconto"))) / 100)
            End If
        Else
            '05/05/2004 - Daniel
            'Personalização Embalavi
            'Tratamento de Máscara
            If g_bln5CasasDecimais Then
                Aux_Preço = (Format((rsPreços("Preço")), "##,###,##0.00000")) * ((100 - (rsProdutos("Desconto") + Desconto_Cli)) / 100)
              '30/04/2007 - Anderson - Implementação de 3 casas decimais
            ElseIf g_bln3CasasDecimais Then
                Aux_Preço = (Format((rsPreços("Preço")), "##,###,##0.000")) * ((100 - (rsProdutos("Desconto") + Desconto_Cli)) / 100)
            Else
                Aux_Preço = rsPreços("Preço") * ((100 - (rsProdutos("Desconto") + Desconto_Cli)) / 100)
            End If
        End If
        '----------------------------------------------------------------------------
      
        If rsProdutos("Moeda") <> 1 Then
            rsCotacoes.Index = "Moeda"
            rsCotacoes.Seek "<=", rsProdutos("Moeda"), Data_Atual
            If Not rsCotacoes.NoMatch Then
                If rsCotacoes("Moeda") = rsProdutos("Moeda") Then
                    '05/05/2004 - Daniel
                    'Personalização Embalavi
                    'Tratamento de Máscara
                    If g_bln5CasasDecimais Then
                      Aux_Preço = (Format(Aux_Preço, "##,###,##0.00000")) * rsCotacoes("Cotação")
                    '30/04/2007 - Anderson - Implementação de 3 casas decimais
                    ElseIf g_bln3CasasDecimais Then
                      Aux_Preço = (Format(Aux_Preço, "##,###,##0.000")) * rsCotacoes("Cotação")
                    Else
                      Aux_Preço = Aux_Preço * rsCotacoes("Cotação")
                    End If
                End If
            End If
        End If
        '05/05/2004 - Daniel
        'Personalização Embalavi
        'Tratamento de Máscara
        If g_bln5CasasDecimais Then
            Grade1.Columns(3).Text = Format(Aux_Preço, "##,###,##0.00000")
            '30/04/2007 - Anderson - Implementação de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
            Grade1.Columns(3).Text = Format(Aux_Preço, "##,###,##0.000")
        Else
            Grade1.Columns(3).Text = Format(Aux_Preço, FORMAT_VALUE)
        End If
    End If
'''    End If
    
    
    If Balança = True Then
      '09/06/2004 - Daniel
      'Tratamento para evitar problema no arredondamento conforme queixa da SMQ
      'Grade1.Columns(1).Text = Format(Preço_Balança / rsPreços("Preço"), FORMAT_VALUE) <= Código antigo
      Dim strMascara As String
      
      strMascara = "#0." & String(Comp_Prod, "0")
      
      Grade1.Columns(1).Text = Format(Preço_Balança / rsPreços("Preço"), strMascara)
    End If
    
    Calcula_Linha
    
    L_Preço.Caption = "Preço=" + CStr(Format(Aux_Preço, FORMAT_VALUE))
    
    Estoque = Acha_Estoque(gnCodFilial, CStr(Cód), Tamanho, Cor, Edição, 0)
    L_Estoque.Caption = "Estoque=" + CStr(Estoque)
    
  End If
  
  
  Rem QTDE
  If ColIndex = 1 Then 'Qtde
    Rem Acha o produto
    Aux2 = Grade1.Columns(0).Text
    Cód = Aux2
    Tamanho = 0
    Cor = 0
    Estoque = -999999
    
    Cód_Str = Trim(CStr(Aux2))
    
    Call Acha_Produto(Cód_Str, Cód, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)
    If Aux_Erro <> 0 Then GoTo Cont_Qtde
    
    Cód = Trim(Cód)

    Estoque = Acha_Estoque(gnCodFilial, CStr(Cód), Tamanho, Cor, Edição, Aux_Erro)
  
  
  
    'Verifica se Qtde é decimal
Cont_Qtde:
    If IsNull(Aux) Then
       Grade1.Columns(1).Text = 0
       Calcula_Linha
       Exit Sub
    End If
    If Aux = "" Then
       Grade1.Columns(1).Text = 0
       Calcula_Linha
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
    
    
    '08/11/2002 - mpdea
    'Adicionado verificação de erro ao encontrar estoque
    '07/10/2002 - mpdea
    'Modificado verificação de estoque para acesso direto ao recordset
    'Adicionado verificação para produto que não controla estoque
    If Not rsParametros.Fields("Venda Sem Estoque").Value And _
       rsProdutos.Fields("Estoque").Value Then
'    If Verifica_Estoque = True Then
      If Aux_Erro = 0 Then
        If CDbl(Aux) > Estoque Then
          If Estoque <> -999999 Then
            '26/08/2004 - Daniel
            'Criado validação para verificar se o usuário possui permissão
            'para enchergar o estoque ou não
            
            '30/01/2007 - Anderson - Alterado para que a permissão de visualizar estoque funcione para diversos clientes
            'If Not m_blnPermitido And m_blnTendresse Then   'Não permitido
            If Not m_blnPermitido Then   'Não permitido
              DisplayMsg "Quantidade superior ao estoque."
            Else
              DisplayMsg "Quantidade superior ao estoque. Estoque atual : " + Format(Estoque, "#########0")
            End If
            
            If CDbl(Aux) <> 0 Then Cancel = True
            Exit Sub
          End If
        End If
      Else
        If Aux_Erro = 1 Then
          DisplayMsg "Produto com estoque não inicializado."
        Else
          DisplayMsg "Erro [" & Aux_Erro & "] ao encontrar estoque do produto."
        End If
        Cancel = True
        Exit Sub
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
      Calcula_Linha
      Exit Sub
    End If
    
    Aux = Grade1.Columns(0).Text
    'Acha produto
    If IsNull(Aux) Or Aux = "" Then Exit Sub
    
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
    Else
      Grade1.Columns(1).Text = Format(Valor, "#0." & String(rsProdutos("QtdeCasasDecimais").Value, "0"))
    End If


    Calcula_Linha
  End If
    
    
  '----------------------------------------------------------'
  ' Data: 27/09/2002                                         '
  ' Nome: Maikel Cordeiro                                    '
  ' Descrição: Se o usuário apagasse o nome do produto e     '
  '            gravasse, quando achasse a venda... o produto '
  '            não aparecia na lista                         '
  '----------[ Prevenção para erro com o nome vazio]---------'
    If ColIndex = 2 Then    'Nome
      If Len(Trim(Grade1.Columns(ColIndex).Text)) <= 0 Then
        Grade1.Columns(0).Text = ""
        Grade1.Columns(1).Text = ""
        Grade1.Columns(2).Text = ""
        Grade1.Columns(3).Text = ""
        Grade1.Columns(4).Text = ""
      Else
        rsProdutos.Index = "Código"
        rsProdutos.Seek "=", Grade1.Columns(0).Text
        If rsProdutos.NoMatch Then Exit Sub
        Grade1.Columns(2).Text = rsProdutos("Nome") & ""
      End If
    End If
  '----------[ Prevenção para erro com o nome vazio]---------'
  
  If ColIndex = 3 Then  'Preço
    If IsNull(Aux) Then
       Grade1.Columns(3).Text = 0
       Calcula_Linha
       Exit Sub
    End If
    
    If Aux = "" Then
       Grade1.Columns(3).Text = 0
       Calcula_Linha
       Exit Sub
    End If
    
    If Not IsNumeric(Aux) Then
      DisplayMsg "Preço incorreto."
      Cancel = 1
      Exit Sub
    End If
    If CDbl(Aux) < 0 Then
      DisplayMsg "Preço não pode ser menor que 0."
      Cancel = 1
      Exit Sub
    End If
    If CDbl(Aux) > 9999999 Then
       DisplayMsg "Preço incorreto, máximo é 9.999.999"
       Cancel = 1
       Exit Sub
    End If
    
    '29/10/2007 - Anderson
    'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
    If g_bolLucroMinimoClasse Then
      If Not PermiteDescontoMargemLucro(Grade1.Columns("Código").Text, Grade1.Columns("Desc.%").Text, Grade1.Columns("Qtde").Text, Grade1.Columns("Preço").Text) And Not m_bolLucroMinimoPermitido Then
        DisplayMsg "Preço unitário não permitido para este produto. Esta operação é permitada apenas com a senha do gerente."
        Cancel = True
        Exit Sub
      End If
    End If
    
    Calcula_Linha
 End If

 
  If ColIndex = 4 Then
    
    
    '29/12/2003 - mpdea
    'Verifica a existência do produto
    If Grade1.Columns("Código").Text = "" Then
      DisplayMsg "Escolha o produto primeiro."
      Grade1.Columns(ColIndex).Text = 0
      Exit Sub
    End If
    
    
    If Not IsNumeric(Aux) Then
      DisplayMsg "Desconto incorreto."
      Cancel = 1
      Exit Sub
    End If
    
    If CDbl(Aux) < 0 Then
      DisplayMsg "Desconto não pode ser menor que 0."
      Cancel = 1
      Exit Sub
    End If
    
    If CDbl(Aux) > 100 Then
      DisplayMsg "Desconto incorreto, máximo é 100.00%"
      Cancel = 1
      Exit Sub
    End If
    
    
    '----------------------------------------------------------------------------
    '29/12/2003 - mpdea
    'Verifica permissão de desconto no produto
    Call Acha_Produto(Grade1.Columns("Código").Text, Cód, 0, 0, 0, 0, Aux_Erro)
    If Aux_Erro <> 0 Then
      If Aux_Erro = 1 Then DisplayMsg "Produto não encontrado."
      If Aux_Erro = 2 Then DisplayMsg "Produto com grade, digite tamanho e cor."
      If Aux_Erro = 3 Then DisplayMsg "Produto com edição, digite a edição."
      Cancel = True
      Exit Sub
    End If
    
    rsProdutos.Index = "Código"
    rsProdutos.Seek "=", Cód
    If rsProdutos.NoMatch Then Exit Sub
    
    '19/10/2007 - Anderson
    'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
    If g_bolLucroMinimoClasse Then
      If Not PermiteDescontoMargemLucro(Grade1.Columns("Código").Text, Grade1.Columns("Desc.%").Text, Grade1.Columns("Qtde").Text, Grade1.Columns("Preço").Text) And Not m_bolLucroMinimoPermitido Then
        DisplayMsg "Desconto não permitido para este produto. Para continuar esta operação é necessária a senha do gerente."
        Cancel = True
        Exit Sub
      End If
    End If
    
    If rsProdutos.Fields("DontAllowDesc").Value Then
      MsgBox "Produto não permite desconto.", vbExclamation, "Atenção"
      Grade1.Columns(ColIndex).Text = 0
      Exit Sub
    End If
    '------------------------------------------------------------------------------
    
    
    Calcula_Linha
  End If

End Sub

Private Sub Grade1_BeforeUpdate(Cancel As Integer)
  Dim Aux As Variant
  '19/04/2013-Alexandre Afornali
  'Criado tratamento do erro ao apagar um produto da venda rapida
  Aux = Grade1.Columns(0).Text
  
  If IsNull(Aux) Or Aux = "" Or Aux = "0" Then
    Grade1.Columns(1).Text = 0
    Grade1.Columns(2).Text = "-"
    Grade1.Columns(3).Text = 0
    Grade1.Columns(4).Text = 0
    Grade1.Columns(5).Text = 0
    Grade1.Columns(6).Text = 0 'ICM
    Grade1.Columns(7).Text = 0 'IPI
    Grade1.Columns(8).Text = "" 'CFOP
    Grade1.Columns("Base_ICM").Text = "0"
    Grade1.Columns("Valor_ICM").Text = "0"
    Grade1.Columns("Valor_Base_Unit").Text = "0"
    Grade1.Columns("Redução_ICM").Text = "0"
    Grade1.Columns("Tipo_ICM").Text = ""
    Exit Sub
  End If

  '21/12/2012 - mpdea
  'Verifica quantidade inválida e não permite prosseguir com quantidade zerada
  Aux = Trim(Grade1.Columns(1).Text)

  If IsNull(Aux) Then
     DisplayMsg "Quantidade inválida."
     Cancel = True
     Exit Sub
  End If
  
  If Aux = "" Then
     DisplayMsg "Quantidade inválida."
     Cancel = True
     Exit Sub
  End If
  
  If Not IsNumeric(Aux) Then
     DisplayMsg "Quantidade inválida."
     Cancel = True
     Exit Sub
  End If
  
  If CDbl(Aux) <= 0 Then
     DisplayMsg "Quantidade inválida."
     Cancel = True
     Exit Sub
  End If
  
  If CDbl(Aux) > 9999999 Then
     DisplayMsg "Quantidade inválida."
     Cancel = True
     Exit Sub
  End If
    
End Sub


Private Sub Grade1_GotFocus()
  Grade1.Col = 0
  
  '18/08/2008 - mpdea
  'Reativado código para selecionar o conteúdo quando obtém o foco
  'Solicitante: Patrício (Technomax)
  '
  ''13/06/2007 - Anderson
  ''Retirado comando para evitar que o cursor se posicione no campo desconto quando recebe o foco

  'Comentado em Junho/2019
'''SendKeys "{Home}+{End}" 'Exatamente aqui está selecionando o conteúdo
End Sub

Private Sub Grade1_InitColumnProps()
'  Grade1.Columns("Código").DropDownHwnd = DropDown1.hwnd

  Select Case rsParametros("PesquisaCodigoENome_VR")
    Case 0
      Grade1.Columns("Código").DropDownHwnd = DropDown1.hwnd
      Grade1.Columns("Nome").Locked = True
    Case -1
      Grade1.Columns("Código").DropDownHwnd = DropDown2.hwnd
      Grade1.Columns("Nome").DropDownHwnd = DropDown1.hwnd
      Grade1.Columns("Nome").Locked = False
      
      If rsParametros("VROrdenacaoCombo") Then
        Data6.RecordSource = "SELECT Produtos.Nome, Produtos.Código FROM Produtos WHERE Produtos.Código <> '0' AND Produtos.[Desativado]=False ORDER BY Produtos.[Código Ordenação]"
      Else
        Data6.RecordSource = "SELECT Produtos.Nome, Produtos.Código FROM Produtos WHERE Produtos.Código <> '0' AND Produtos.[Desativado]=False ORDER BY Produtos.Código"
      End If
  End Select
End Sub

Private Sub Grade1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF6 Then
    Call Form_KeyDown(KeyCode, Shift)
  End If
End Sub

'23/10/2002 - mpdea
'Criado o controle de chamadas aos controles DropDown1 e DropDown2,
'antes só era feito para DropDown1
Private Sub Grade1_KeyPress(KeyAscii As Integer)
  Dim objDropDown As SSDBDropDown
  
  If rsParametros.Fields("PesquisaCodigoENome_VR").Value Then
    Set objDropDown = DropDown2
  Else
    Set objDropDown = DropDown1
  End If
  
  If Len(Grade1.Columns("Código").Text) > 0 Then
    If Asc(Grade1.Columns("Código").Text) = 13 Then Grade1.Columns("Código").Text = ""
  End If
  
  With objDropDown
    If Grade1.Col = 0 Then
      If .DroppedDown And .Name = "DropDown1" Then
        .DataFieldList = "Nome"
      End If
      
      If KeyAscii = vbKeyReturn Then
        If ActiveBar1.Tools("miOpLeitorOtico").Checked And Not .DroppedDown Then
          With Grade1.Columns("Código")
            If .Text <> "" And .Text <> "0" Then
              Grade1.Columns("Qtde").Text = 1
              'Grade1.SetFocus '07/10/2004 - Daniel - Case: Sucupira
              SendKeys "{DOWN}{HOME}", True
              
              '07/10/2004 - Daniel - Case: Sucupira
              If Grade1.Columns("Código").Text = "" Then
                Grade1.Col = 0
                Grade1.SetFocus '21/10/2004 - Forçar o foco
                Grade1.Col = 0  'ADICIONADA ESTA LINHA....FOCO SEM O ZERO 2019 AGOSTO
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
            
          End With
        End If
      End If
      
    End If
  End With

' If Grade1.Col = 0 Then
'   DropDown1.DataFieldToDisplay = "Nome"
'   If KeyAscii = 13 Then  'enter
'     If ActiveBar1.Tools("miOpLeitorOtico").Checked = True Then
'       Grade1.Columns(1).Text = 1 'qtde
'       SendKeys "{DOWN}"
'     End If
'   End If
' End If

End Sub

Private Sub Grade1_LostFocus()
  L_Estoque.Caption = ""
  If Grade1.RowChanged Then
    Grade1.Update
  End If
End Sub


Private Sub Grade1_Scroll(Cancel As Integer)
'  MsgBox "Scroll"
'  Calcula_Linha
End Sub

Private Sub Grade1_UnboundAddData(ByVal RowBuf As ssRowBuffer, NewRowBookmark As Variant)
  Dim nLinha As Integer
  
  nLinha = Grade1.Row
  
  With Tabe(nLinha)
    .Código = UCase(Grade1.Columns("Código").Text)
    .Nome = Grade1.Columns("Nome").Text
    .Preço_Final = CDbl(Grade1.Columns("Total").Text)
    .Qtde = CDbl(Grade1.Columns("Qtde").Text)
    '05/05/2004 - Daniel
    'Personalização Embalavi
    'Tratamento de Máscara
    If g_bln5CasasDecimais Then
      .Preço = CDbl(Format((Grade1.Columns("Preço").Text), "##,###,##0.00000"))
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Preço = CDbl(Format((Grade1.Columns("Preço").Text), "##,###,##0.000"))
    Else
      .Preço = CDbl(Grade1.Columns("Preço").Text)
    End If
    
    .Desconto = CDbl(Grade1.Columns("Desc.%").Text)
    .ICM = CDbl(Grade1.Columns("ICM").Text)
    .IPI = CDbl(Grade1.Columns("IPI").Text)
    .Base_ICM = CDbl(Grade1.Columns("Base_ICM").Text)
    .Valor_ICM = CDbl(Grade1.Columns("Valor_ICM").Text)
    .Valor_Base_Unit = CDbl(Grade1.Columns("Valor_Base_Unit").Text)
    .Redução_ICM = CDbl(Grade1.Columns("Redução_ICM").Text)
    .Tipo_ICM = Grade1.Columns("Tipo_ICM").Text
    
    '05/04/2010 - Andrea
    'Alteração para o Registro de SCFOP por produto
    .CFOP_Produto = Grade1.Columns("CFOP").Text

  End With
End Sub

Private Sub Grade1_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim nX As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      nX = Grade1.Rows
    Else
      nX = 0
    End If
  Else
    nX = StartLocation
  End If
  NewLocation = nX + NumberOfRowsToMove
End Sub

Private Sub Grade1_UnboundReadData(ByVal RowBuf As ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Dim nPos As Integer
  Dim nX As Integer
  Dim nCount As Integer
  
  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      nPos = Grade1.Rows
    Else
      nPos = 0
    End If
  Else
    nPos = StartLocation
    If ReadPriorRows Then
      nPos = nPos - 1
    Else
      nPos = nPos + 1
    End If
  End If

  With RowBuf
    For nX = 0 To .RowCount - 1
      If nPos < 0 Or nPos >= Grade1.Rows Then
        Exit For
      Else
        .Value(nX, 0) = Tabe(nPos).Código
        .Value(nX, 1) = Tabe(nPos).Qtde
        .Value(nX, 2) = Tabe(nPos).Nome
        .Value(nX, 3) = Tabe(nPos).Preço
        .Value(nX, 4) = Tabe(nPos).Desconto
        .Value(nX, 5) = Tabe(nPos).Preço_Final
        .Value(nX, 6) = Tabe(nPos).ICM
        .Value(nX, 7) = Tabe(nPos).IPI
        .Value(nX, 8) = Tabe(nPos).Base_ICM
        .Value(nX, 9) = Tabe(nPos).Valor_ICM
        .Value(nX, 10) = Tabe(nPos).Valor_Base_Unit
        .Value(nX, 11) = Tabe(nPos).Redução_ICM
        .Value(nX, 12) = Tabe(nPos).Tipo_ICM
        
        '05/04/2010 - Andrea
        .Value(nX, 13) = Tabe(nPos).CFOP_Produto

        .Bookmark(nX) = nPos
        If ReadPriorRows Then
          nPos = nPos - 1
        Else
          nPos = nPos + 1
        End If
        nCount = nCount + 1
      End If
    Next nX
    .RowCount = nCount
  End With

End Sub
  
Public Sub Recalcula()
'  Dim i As Integer
'  Dim Tot_Prod As Double
'  Dim Tot_Desc As Double
'  Dim Tot_IPI As Double
'  Dim Tot_Pagar As Double
'  Dim Valor_Desc As Double
'  Dim Valor_IPI As Double
'  Dim Temp As Double
'  Dim Base_ICM As Double
'  Dim Valor_ICM As Double
'  Dim Num_Prod As Integer
'
'  Dim ICM(200) As Double
'
'  Num_Prod = 0
'
'  For i = 0 To Linhas_Grade
'    If Tabe(i).Código <> "" Then 'Faz somente os preenchidos
'      Temp = Tabe(i).Preço * Tabe(i).Qtde
'      Temp = Format(Temp, "#0.00")
'      Tot_Prod = Tot_Prod + Temp
'      Valor_Desc = (Temp * Tabe(i).Desconto / 100)
'      Tot_Desc = Tot_Desc + Valor_Desc
'      Valor_IPI = Temp * Tabe(i).IPI / 100
'      If Calcula_IPI = False Then Valor_IPI = 0
'      Tot_IPI = Tot_IPI + Valor_IPI
'      'Num_Prod = Num_Prod + Tabe(I).Qtde
'      If Tabe(i).ICM <> 0 Then
'        If Calcula_ICM = True Then
'          ICM(Tabe(i).ICM) = ICM(Tabe(i).ICM) + Temp
'        End If
'      End If
'    End If
'  Next i
'
'  For i = 1 To 199
'    If ICM(i) <> 0 Then
'      If Calcula_ICM = True Then
'        Base_ICM = Base_ICM + ICM(i)
'        Valor_ICM = Valor_ICM + Format((ICM(i) * i / 100), "##########0.00")
'      End If
'    End If
'  Next i
'
'  Tot_Pagar = Tot_Prod - Tot_Desc + Tot_IPI
'
'  L_Tot_Prod.Text = Format(Tot_Prod, "###,###,##0.00")
'  Total_Produtos = Tot_Prod
'  'txtDescSubTotal.Text = Format(Tot_Desc, "###,###,##0.00")
'  L_Tot_IPI.Text = Format(Tot_IPI, "###,###,##0.00")
'  Total_IPI = Tot_IPI
'
'  L_Tot_Pagar.Text = Format(Tot_Pagar, "###,###,##0.00")
'  Total_Pagar = CDbl(Format(Tot_Pagar, "###########0.00"))
'  'L_Base_ICM = Format(Base_ICM, "###,###,##0.00")
'  Total_Base_ICM = Base_ICM
'  'L_Valor_ICM = Format(Valor_ICM, "###,###,##0.00")
'  Total_ICM = Valor_ICM
'  L_Receber.Text = Total_Pagar - Total_Recebido
  Dim nX As Integer
  
  Dim Qtde As Double
  Dim Preço As Double
  Dim Desconto As Double
  Dim Valor_Desconto As Double
  Dim IPI As Double
'  Dim Valor_IPI As Double
  Dim Preço_Total As Double
  Dim Preço_Final As Double
  Dim Preço_Final2 As Double
  
  Dim Tot_Prod As Double
  Dim Tot_Desc As Double
  Dim Tot_IPI As Double
  Dim Tot_Pagar As Double
  Dim Valor_Desc As Double
  Dim Valor_IPI As Double
  Dim temp As Double
  Dim Base_ICM As Double
  Dim Valor_ICM As Double
  Dim Base_ICM_Subs As Double
  Dim Valor_ICM_Subs As Double
  Dim Valor_ISS As Double
  Dim ICM(200, 2) As Double
  Dim nTbValIPI(200) As Currency
  Dim sCodProd As String
  '10/11/2004 - Daniel
  Dim strUF As String
  Dim Vpreco As String
  Dim nQtdeTotal As Single
    
  Tot_Desc = 0#
  Tot_Prod = 0#
  gnPesoLiquido = 0#
  gnPesoBruto = 0#
  
  For nX = 0 To (Linhas_Grade - 1)
    sCodProd = gsHandleNull(Tabe(nX).Código)
    If sCodProd <> "0" Then  'Faz somente os preenchidos
      
      Qtde = Tabe(nX).Qtde
      
      'Calcula Quantidade total de itens no grid
      nQtdeTotal = nQtdeTotal + Qtde
      
      '05/05/2004 - Daniel
      'Personalização Embalavi
      'Tratamento de Máscara
      If g_bln5CasasDecimais Then
        Preço = Format((Tabe(nX).Preço), "##,###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Preço = Format((Tabe(nX).Preço), "##,###,##0.000")
      Else
        Preço = Tabe(nX).Preço
      End If
      
      Desconto = Tabe(nX).Desconto
      
      '------------------------------------------------------
      '23/05/2006 - mpdea
      'Comentado restrição de isenção de IPI para a Embalavi
      'É utilizado configuração do cadastro de clientes
      '
      '07/05/2004 - Daniel
      'Personalização Embalavi
      'Exatamente neste ponto que temos em mãos
      'o percentual do IPI do produto
      'Tratamento através da função IsencaoIPI para
      'verificação se suspende ou não a taxa de IPI conforme
      'o cliente e não o produto
'      If m_blnEmbalavi Then
'        If Len(Nome_Cliente.Caption) > 0 Then
'          If IsencaoIPI(CLng(Combo_Cliente.Text)) Then 'Cliente É Isento de IPI
          If m_blnIsentoIPI Then
            IPI = 0
          Else
            IPI = Tabe(nX).IPI
          End If
'        Else 'Len...
'          IPI = Tabe(nX).IPI
'        End If
'
'      Else 'Não Embalavi
'        IPI = Tabe(nX).IPI
'      End If
      '------------------------------------------------------
      
      Preço_Total = Format(Qtde * Preço, "#0.00")
      Vpreco = Format(Preço_Total, "##,###,##0.00")
      Tabe(nX).Preço_Final = Vpreco
      
      Valor_Desconto = Format(Preço_Total * Desconto / 100#, "#0.00")
      Preço_Final = Format((Preço_Total - Valor_Desconto), "#0.00")
      Valor_IPI = Preço_Final * IPI / 100#
      Valor_IPI = Format(Valor_IPI, "#0.00")
      If Not Calcula_IPI Then
        Valor_IPI = 0
      End If
      
      Preço_Final2 = Format((Preço_Final + Valor_IPI), "#0.00")
      Vpreco = Format(Preço_Final2, "##,###,##0.00")
      Tabe(nX).Preço_Final = Vpreco
  
      With Tabe(nX)
        'Calculo do ICM
        If .Tipo_ICM = "N" Then
          'ICM Normal
          If gbBaseICMSomadoIPI = True Then
            .Base_ICM = Preço_Final2
            .Valor_ICM = Preço_Final2 * CSng(gsHandleNull(.ICM & "")) / 100
          Else
            .Base_ICM = Preço_Final
            .Valor_ICM = Preço_Final * CSng(gsHandleNull(.ICM & "")) / 100
          End If
        ElseIf .Tipo_ICM = "R" Then
          'ICM Retido
          If CDbl(.Valor_Base_Unit) <> 0 Then
            'Base Fixa
            .Base_ICM = CDbl(.Qtde) * CDbl(.Valor_Base_Unit)
            .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
          End If
          If CDbl(.Redução_ICM) <> 0 Then
            'Base Reduzida
            .Base_ICM = Preço_Final * CDbl(.Redução_ICM) / 100
            .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
          End If
        ElseIf .Tipo_ICM = "Z" Then
        
          '10/11/2004 - Daniel
          'Personalização para o cliente Teknika
          'Tratamento do ICM
          If Not m_blnTeknika Then
        
            'ICM Reduzido
            If CDbl(.Valor_Base_Unit) <> 0 Then
              'Base Fixa
              .Base_ICM = CDbl(.Qtde) * CDbl(.Valor_Base_Unit)
              .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
            End If
            If CDbl(.Redução_ICM) <> 0 Then
              'Base Reduzida
              .Base_ICM = Preço_Final * CDbl(.Redução_ICM) / 100
              .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
            End If
          
          Else 'Teknika
          
            'ICM Reduzido
            If CDbl(.Valor_Base_Unit) <> 0 Then
              'Base Fixa
              .Base_ICM = CDbl(.Qtde) * CDbl(.Valor_Base_Unit)
              .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
            End If
          
            '10/11/2004 - Daniel
            'Tratamento para base reduzida
            'Chamamos a Function IE_Isento para verificação
            If IE_Isento(strUF) Then 'ISENTO = TRUE
            
              .Base_ICM = Preço_Final
              .Valor_ICM = Preço_Final * CSng(gsHandleNull(.ICM & "")) / 100
          
            Else 'ISENTO = FALSE
            
              If strUF = "PR" Then
                If CDbl(.Redução_ICM) <> 0 Then
                  'Base Reduzida
                  .Base_ICM = Preço_Final * CDbl(.Redução_ICM) / 100 'CDbl(.Redução_ICM) / 100 = 66,66
                  .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
                End If
              Else
                  .Base_ICM = Preço_Final
                  .Valor_ICM = Preço_Final * CSng(gsHandleNull(.ICM & "")) / 100
              End If
            
            End If
          
          End If 'Else da Teknika
          
        End If
      End With
            
'      gnPesoLiquido = gnPesoLiquido + Tabe(nX).PesoLiquido * Tabe(nX).Qtde
'      gnPesoBruto = gnPesoBruto + Tabe(nX).PesoBruto * Tabe(nX).Qtde
      
      '05/05/2004 - Daniel
      'Personalização Embalavi
      'Tratamento de Máscara
      If g_bln5CasasDecimais Then
        temp = Format((Tabe(nX).Preço * Tabe(nX).Qtde), "##,###,##0.00000")
        temp = Format(temp, "##,###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        temp = Format((Tabe(nX).Preço * Tabe(nX).Qtde), "##,###,##0.000")
        temp = Format(temp, "##,###,##0.000")
      Else
        temp = Tabe(nX).Preço * Tabe(nX).Qtde
        temp = Format(temp, "#0.00")
      End If
      
      Tot_Prod = Format((Tot_Prod + temp), "##,###,##0.00")
      Valor_Desc = temp * Tabe(nX).Desconto / 100#
      Valor_Desc = Format(Valor_Desc, "#0.00")
      Tot_Desc = Tot_Desc + Valor_Desc
      temp = temp - Valor_Desc
      Valor_IPI = temp * Tabe(nX).IPI / 100#
      Valor_IPI = Format(Valor_IPI, "#0.00")
      If Calcula_IPI = False Then
        Valor_IPI = 0
      End If
      Tot_IPI = Tot_IPI + Valor_IPI
     
      If Calcula_ICM Then
        If Tabe(nX).Valor_Base_Unit <> 0 Or Tabe(nX).Redução_ICM <> 0 Then
          If Tabe(nX).Tipo_ICM = "R" Then
            Base_ICM_Subs = Base_ICM_Subs + Tabe(nX).Base_ICM
            Valor_ICM_Subs = Valor_ICM_Subs + Tabe(nX).Valor_ICM
          End If
        End If
      End If
     
      If Tabe(nX).ICM <> 0 Then
        If Calcula_ICM Then
          'ICM(Tabe(nX).ICM) = ICM(Tabe(nX).ICM) + Temp
          'If Tabe(nX).Valor_Base_Unit = 0 And Tabe(nX).Redução_ICM = 0 Then
          If Tabe(nX).Tipo_ICM = "Z" Or Tabe(nX).Tipo_ICM = "N" Then
            ICM(Tabe(nX).ICM, 1) = ICM(Tabe(nX).ICM, 1) + Tabe(nX).Base_ICM
            ICM(Tabe(nX).ICM, 2) = ICM(Tabe(nX).ICM, 2) + Tabe(nX).Valor_ICM
          End If
        End If
      End If
     
    End If
  Next nX
  
  'Quantidade deve ser single (conforme estrutura da base de dados)
  lblQtdeTotal.Caption = nQtdeTotal
  
  For nX = 1 To 199
    If ICM(nX, 1) <> 0 Then
      If Calcula_ICM Then
        Base_ICM = Base_ICM + ICM(nX, 1)
        Valor_ICM = Valor_ICM + ICM(nX, 2)
      End If
    End If
  Next nX
  
'  Total_Desconto = Tot_Desc
  'Alterado para manter o total de desconto no Total Geral
  Total_Desconto = Tot_Desc + gcDescInTotal
  
  Tot_Pagar = Tot_Prod + Tot_IPI

  
  '23/09/2002 - mpdea
  'Adicionado o Desconto no SubTotal
  Tot_Pagar = Format(Tot_Pagar - Total_Desconto, FORMAT_VALUE)

  '11/11/2008 - mpdea
  'Soma o ICMS Retido ao total da nota
  If m_blnSomaIcmsRetidoTotalNota Then
    Tot_Pagar = Format(Tot_Pagar + Valor_ICM_Subs, FORMAT_VALUE)
  End If

  '06/11/2007 - Anderson
  'Verifica se deve somar os produtos no total da nota
  'L_Tot_Prod.Text = Format(Tot_Prod, "###,###,##0.00")
  If rsOp_Saída("SomarProdutosTotalNota") Then
    L_Tot_Prod.Text = Format(Tot_Prod, "###,###,##0.00")
  Else
    L_Tot_Prod.Text = Format(0, "###,###,##0.00")
  End If
  Total_Produtos = Tot_Prod
  'txtDescSubTotal.Text = Format(Tot_Desc, "###,###,##0.00")
  L_Tot_IPI.Text = Format(Tot_IPI, "###,###,##0.00")
  Total_IPI = Tot_IPI

  '06/11/2007 - Anderson
  'Verifica se deve somar os produtos no total da nota
  '24/09/2002 - mpdea
  'Desconto no SubTotal
  'txtSubTotal.Text = Format(Tot_Pagar + mcurDescontoSubTotal, FORMAT_VALUE)
  'txtDescSubTotal.Text = Format(mcurDescontoSubTotal, FORMAT_VALUE)
  '************
  '************ junho 2019 COMENTEI ESTE CODIGO AQUI
'''  If rsOp_Saída("SomarProdutosTotalNota") Then
'''    txtSubTotal.Text = Format(Tot_Pagar + Total_Desconto, FORMAT_VALUE)
'''    txtDescSubTotal.Text = Format(Total_Desconto, FORMAT_VALUE)
'''  Else
'''    txtSubTotal.Text = Format(0, FORMAT_VALUE)
'''    txtDescSubTotal.Text = Format(0, FORMAT_VALUE)
'''  End If
  
  ' e fiz assim
  If rsOp_Saída("SomarProdutosTotalNota") Then
    txtSubTotal.Text = Format(Tot_Pagar + Total_Desconto, FORMAT_VALUE)
    
    If (mcurDescontoSubTotal > 0 And Total_Desconto = 0) Or (Total_Pagar + mcurDescontoSubTotal + Total_Desconto = txtSubTotal.Text) Then
        txtDescSubTotal.Text = Format(mcurDescontoSubTotal + Total_Desconto, FORMAT_VALUE)
        Tot_Pagar = txtSubTotal.Text - (mcurDescontoSubTotal + Total_Desconto)
    Else
        txtDescSubTotal.Text = Format(Total_Desconto, FORMAT_VALUE)
    End If
  Else
    txtSubTotal.Text = Format(0, FORMAT_VALUE)
    txtDescSubTotal.Text = Format(0, FORMAT_VALUE)
  End If
  '*****************
  '*****************

  '06/11/2007 - Anderson
  'Verifica se deve somar os produtos no total da nota
  'L_Tot_Pagar.Text = Format(Tot_Pagar, "###,###,##0.00")
  If rsOp_Saída("SomarProdutosTotalNota") Then
    L_Tot_Pagar.Text = Format(Tot_Pagar, "###,###,##0.00")
  Else
    L_Tot_Pagar.Text = Format(0, "###,###,##0.00")
  End If
  Total_Pagar = Format(Tot_Pagar, "###########0.00")
  'L_Base_ICM = Format(Base_ICM, "###,###,##0.00")
  'L_Valor_ICM = Format(Valor_ICM, "###,###,##0.00")
  
  '23/05/2006 - mpdea
  'Centralizado verificação do uso de Diferimento
  '
  '07/05/2004 - Daniel
  'Case: Embalavi
  'Verificação de Diferimento sobre o ICM
  'quando for Embalavi. A Base de ICM será
  'reduzida 33% e deste valor extrairemos 18%
  'que será o Valor do ICM
  '17/05/2004 - Daniel
  'Atualizamos os valores de cálculo (percentuais)
  'através da busca na tabela Diferimento
  'If m_blnEmbalavi Then
  If g_blnDiferimento Then
  
    If Len(Nome_Cliente.Caption) > 0 Then
    
        If Diferimento(CLng(Combo_Cliente.Text)) Then 'Há Diferimento
          'Tratamento para Diferimento
          Dim dblTotal As Double 'Antiga dblTrintaETres
          Dim dblBase  As Double 'Antiga dblDezoito
          Dim rstDiferimento As Recordset '17/05/2004
          Dim dblTotalTable  As Double
          Dim dblBaseTable   As Double
          
          Set rstDiferimento = db.OpenRecordset("SELECT Total, Base FROM Diferimento WHERE Filial = " & gnCodFilial, dbOpenDynaset)

          With rstDiferimento
            If Not (.BOF And .EOF) Then
              If Not IsNumeric(.Fields("Total").Value) Or Not IsNumeric(.Fields("Base").Value) Then
                dblTotalTable = 0
                dblBaseTable = 0
              Else
                dblTotalTable = .Fields("Total").Value
                dblBaseTable = .Fields("Base").Value
              End If
            Else
              dblTotalTable = 0
              dblBaseTable = 0
            End If
            .Close
          End With
          
          Set rstDiferimento = Nothing
          
          
          dblTotal = Format(((Base_ICM * dblTotalTable) / 100), "##,###,##0.00")
          Total_Base_ICM = Base_ICM - dblTotal
          
          dblBase = Format(((Total_Base_ICM * dblBaseTable) / 100), "##,###,##0.00")
          Total_ICM = dblBase
        
        Else
          Total_Base_ICM = Base_ICM
          Total_ICM = Valor_ICM
        End If
    
    Else 'Início free
      Total_Base_ICM = Base_ICM
      Total_ICM = Valor_ICM
    End If
  
  Else 'Não Embalavi
    Total_Base_ICM = Base_ICM
    Total_ICM = Valor_ICM
  End If
  
  '06/11/2007 - Anderson
  'Verifica se deve somar os produtos no total da nota
  'L_Receber.Text = Format(Abs(Total_Pagar - Total_Recebido), FORMAT_VALUE)
  If rsOp_Saída("SomarProdutosTotalNota") Then
    L_Receber.Text = Format(Abs(Total_Pagar - Total_Recebido), FORMAT_VALUE)
  Else
    L_Receber.Text = Format(0, FORMAT_VALUE)
  End If

End Sub

Private Sub Grade1_UnboundWriteData(ByVal RowBuf As ssRowBuffer, WriteLocation As Variant)
 Dim nLinha As Integer
 Dim iIndice As Integer
 Dim sAUX As String
 
 nLinha = WriteLocation

  With Tabe(nLinha)
    .Código = UCase(Grade1.Columns("Código").Text)
    .Nome = Grade1.Columns("Nome").Text
    .Preço_Final = CDbl(Grade1.Columns("Total").Text)
    
    ' Pilatti
    sAUX = .Preço_Final
    iIndice = InStr(1, sAUX, ",")
    If iIndice = 0 Then
      .Preço_Final = sAUX + ",00"
    ElseIf iIndice + 1 = Len(sAUX) Then
      .Preço_Final = sAUX + "0"
    End If
    
    .Qtde = CDbl(Grade1.Columns("Qtde").Text)
    '05/05/2004 - Daniel
    'Personalização Embalavi
    'Tratamento de Máscara
    If g_bln5CasasDecimais Then
      .Preço = CDbl(Format((Grade1.Columns("Preço").Text), "##,###,##0.00000"))
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Preço = CDbl(Format((Grade1.Columns("Preço").Text), "##,###,##0.000"))
    Else
      .Preço = CDbl(Grade1.Columns("Preço").Text)
    End If
    
    .Desconto = CDbl(Grade1.Columns("Desc.%").Text)
    .ICM = CDbl(Grade1.Columns("ICM").Text)
    .IPI = CDbl(Grade1.Columns("IPI").Text)
    .Base_ICM = CDbl(Grade1.Columns("Base_ICM").Text)
    .Valor_ICM = CDbl(Grade1.Columns("Valor_ICM").Text)
    .Valor_Base_Unit = CDbl(Grade1.Columns("Valor_Base_Unit").Text)
    .Redução_ICM = CDbl(Grade1.Columns("Redução_ICM").Text)
    .Tipo_ICM = Grade1.Columns("Tipo_ICM").Text
    
    '05/04/2010 - Andrea
    'Alteração para o Registro de CFOP por Produto
    .CFOP_Produto = Grade1.Columns("CFOP").Text

  End With
End Sub


Private Sub Lançar_Débito_Click()
  Recalcula_Recebido
End Sub

Private Sub FindVenda()
  lbl_retornoEnvioNFCe.Visible = False

  frmVendasHoje.Show vbModal
End Sub

Private Sub EmiteBoletos()
 Dim Nome_Arq As String
 Dim Aux_Contador As Long
 Dim Mensa As String
 Dim Resp As Integer
 Dim Impressos As Integer
 Dim F As Form
 Dim rstContasReceber As Recordset '05/06/2007 - Anderson
 Dim bolErroNossoNumero As Boolean '05/06/2007 - Anderson
 
 Dim strNossoNumero As String '17/05/2007 - Anderson
  
 Call StatusMsg("")

 If IsNull(Num_Registro) Then
   DisplayMsg "Encontre ou grave uma venda antes."
   Exit Sub
 End If
 
 
 Rem mostra tela para escolha de configuração
  Set F = New frmObsDoc
  F.Caption = "Impressão de Boletos"
  F.gsFileExt = ".CBB"
  F.Show vbModal
  Set F = Nothing
 If gsRetornoDoc <> "OK" Then
'   DisplayMsg "Impressão cancelada."
   Exit Sub
 End If
  
 Nome_Arq = gsConfigPath & gsDocFileName & ".CBB"
 If Dir(Nome_Arq) = "" Then
   DisplayMsg "Arquivo """ & Nome_Arq & """ não encontrado."
   Exit Sub
 End If

 rsContas_Receber.Index = "Sequência"
 Impressos = 0
 Aux_Contador = 0
Lp1:
 rsContas_Receber.Seek ">", gnCodFilial, "R", rsSaidas("Sequência"), Aux_Contador
 If rsContas_Receber.NoMatch Then GoTo Fim
 If rsContas_Receber("Filial") <> gnCodFilial Then GoTo Fim
 If rsContas_Receber("Sequência") <> rsSaidas("Sequência") Then GoTo Fim
 If rsContas_Receber("Tipo") <> "R" Then GoTo Fim
 
 Aux_Contador = rsContas_Receber("Contador")
 
 If rsContas_Receber("Tipo Parcelamento") <> "B" Then GoTo Lp1
  
  
 If rsContas_Receber("Impresso") = True Then
   Mensa = "O boleto com vencimento em " + CStr(rsContas_Receber("Vencimento"))
   Mensa = Mensa + " no valor de R$ " + Format(rsContas_Receber("Valor"), "###,###,##0.00")
   Mensa = Mensa + " já foi impresso. DESEJA A REIMPRESSÃO ?"
 Else
   Mensa = "Deseja imprimir o boleto com vencimento em " + CStr(rsContas_Receber("Vencimento"))
   Mensa = Mensa + " no valor de R$ " + Format(rsContas_Receber("Valor"), "###,###,##0.00")
   Mensa = Mensa + " ?"
 End If
   
 Resp = MsgBox(Mensa, vbOKCancel, "Atenção")
 If Resp = vbCancel Then GoTo Lp1
 
  '16/05/2007 - Anderson
  'Se Número de série Agrotama, informar nosso número para boletos pré-impressos
  If g_blnInformarNossoNumero And strNossoNumero = "" Then
    Do
      strNossoNumero = InputBox("Informe o Nosso Número para a impressão do boleto.", "Impressão de Boletos")
      If strNossoNumero = "" Then
        Exit Sub
      End If
      If Not IsNumeric(strNossoNumero) Then
        MsgBox "O valor digitado não é válido!", vbExclamation, "Impressão de Boletos"
      End If
    Loop Until IsNumeric(strNossoNumero)
  End If
  
  '05/06/2007 - Anderson
  'Verifica se o Nosso Número já foi emitido em outro boleto para evitar duplicidade.
  'Solicitado pelo cliente Agrotama
  If g_blnInformarNossoNumero Then
  
    'Abre registro para evitar duplicidade em nosso número
    Set rstContasReceber = db.OpenRecordset("SELECT CNAB_NossoNumero, Filial, Cliente, Vendedor, Sequência, Nota, [Data Emissão], Vencimento, Valor FROM [Contas a Receber] Where CNAB_NossoNumero='" & strNossoNumero & "'")
    
    'Informa que não existe problemas com Nosso Numero
    bolErroNossoNumero = False
    
    'Verifica se existe Nosso número no banco de dados
    If Not rstContasReceber.EOF Then
      MsgBox "Já existe um título com o Nosso Número: " & strNossoNumero & " informado em outro boleto." & Chr(13) & _
             "Favor verificar o título com os dados abaixo: " & Chr(13) & Chr(13) & _
             "Nosso Número: " & rstContasReceber("CNAB_NossoNumero") & Chr(13) & _
             "Filial: " & rstContasReceber("Filial") & Chr(13) & _
             "Cliente: " & rstContasReceber("Cliente") & Chr(13) & _
             "Vendedor: " & rstContasReceber("Vendedor") & Chr(13) & _
             "Sequência: " & rstContasReceber("Sequência") & Chr(13) & _
             "Nota: " & rstContasReceber("Nota") & Chr(13) & _
             "Data Emissão: " & rstContasReceber("Data Emissão") & Chr(13) & _
             "Vencimento: " & rstContasReceber("Vencimento") & Chr(13) & _
             "Valor: " & rstContasReceber("Valor"), vbOKOnly + vbInformation, "Impressão de Boletos"
             
      'Informa que existe um título com o mesmo Nosso Numero
      bolErroNossoNumero = True
    End If
  
    'Fecha tabela de contas a receber
    rstContasReceber.Close
    Set rstContasReceber = Nothing
    
    'Se houver duplicidade em Nosso Número, o sistema encerra.
    If bolErroNossoNumero Then
      GoTo Fim
    End If
  
  End If
 
 Resp = Imprime_Boleto("R", rsContas_Receber("Filial"), rsContas_Receber("Vencimento"), rsContas_Receber("Contador"), Nome_Arq)
 
 If Resp <> 0 Then
   MsgBox "Houve o erro  " + str(Resp) + " ao emitir o boleto."
 Else
   Impressos = Impressos + 1
   rsContas_Receber.Edit
     rsContas_Receber("Impresso") = True
     rsContas_Receber("Data Alteração") = Format(Date, "dd/mm/yyyy")
     '16/05/2007 - Anderson
     'Se Número de série Agrotama, informar nosso número para boletos pré-impressos
     If CheckSerialCaseMod("QS73070-894") Then
       rsContas_Receber("CNAB_NossoNumero") = Right(String(11, "0") & strNossoNumero, 11)
       rsContas_Receber("CNAB_DigitoVerificador") = GetDigitoVerificador_NossoNumero(strNossoNumero, Bradesco)
       rsContas_Receber("CNAB_Carteira") = "9"
     End If
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
        "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequência") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
        "frmVendaRap2_EmiteBoletos", _
        "Contas a Receber", g_strArquivoSystemLog
      End If
   rsContas_Receber.Update
 End If
 
 
 GoTo Lp1
 
Fim:
  DisplayMsg "Foram impressos " + str(Impressos) + " boleto(s)."
  

End Sub

Private Sub CadastraCliente()

  Call StatusMsg("")
  
  '29/08/2003 - mpdea
  'Verificação incluída em Form_Load
'  If rsUsuarios("Clientes") = False Then
'    Beep
'    DisplayMsg "Este usuário não tem permissão para cadastrar novos clientes."
'    Exit Sub
'  End If
   
  Call StatusMsg("Aguarde ...")
  Me.MousePointer = vbHourglass
  frmCliFor.Show
  Call StatusMsg("")
  Me.MousePointer = vbDefault
End Sub


Private Sub EmiteCarnesNOVOS(Optional pImpressaoDireta As Integer, Optional pNomeCliente As String)
On Error GoTo Erro:
  Dim Resp As String
  Dim strNomeArq As String
  Dim sDeclaracao As String
  Dim sDeclaracaoTotalVenda As String
  Dim sDeclaracaoDataVenda As String
  Dim sDeclaracaoVencParc1 As String
  Dim sDeclaracaoValorParc As String
  Dim sDeclaracaoNumParc As String
 

  ' pImpressaoDireta = 0  - NÃO CHAMOU PELO ICONE IMPRESSAO_DIRETA
  ' pImpressaoDireta = 1  - CHAMOU PELO ICONE IMPRESSAO_DIRETA
  ' pImpressaoDireta = 2  - CHAMOU PELO ICONE IMPRESSAO_DIRETA_COM_DECLARACAO

  If pImpressaoDireta = 1 Then
      CrystalReport1.WindowShowPrintSetupBtn = True
      CrystalReport1.WindowState = crptMaximized
      CrystalReport1.Destination = crptToPrinter
      
      strNomeArq = gsReportPath & "carne02_todasParcelas_46Colunas.rpt"
  ElseIf pImpressaoDireta = 2 Then
      Dim rsContasReceber As Recordset
      Dim itotalNumParc As Integer
      Dim sValorParcelas As String
      Dim sVencimentoParcela1 As String
      
      Set rsContasReceber = db.OpenRecordset(" SELECT * FROM [Contas a Receber] " & _
                                       " WHERE Filial = " & gnCodFilial & _
                                       " AND Sequência=" & rsSaidas("Sequência") & _
                                       " Order by Contador", dbOpenSnapshot)
      If rsContasReceber.RecordCount > 0 Then
          itotalNumParc = rsContasReceber.RecordCount
          sValorParcelas = Format(rsContasReceber.Fields("Valor").Value, "###,###,##0.00")
          sVencimentoParcela1 = rsContasReceber.Fields("Vencimento").Value
      End If
      rsContasReceber.Close
      Set rsContasReceber = Nothing
      
      CrystalReport1.WindowShowPrintSetupBtn = True
      CrystalReport1.WindowState = crptMaximized
      CrystalReport1.Destination = IIf(False, crptToWindow, crptToPrinter)
      
      strNomeArq = gsReportPath & "carne02_todasParcelas_comRecibo_46Col.rpt"
  
      '''sDeclaracao = "Eu " & pNomeCliente
      '''sDeclaracao = sDeclaracao + " declaro que comprei na empresa "
      sDeclaracao = "declaro que comprei na empresa "
      sDeclaracao = sDeclaracao + gsNomeFilial
      sDeclaracaoTotalVenda = "Total da compra  R$ " + Format(rsSaidas("Total"), "###,###,##0.00")
      sDeclaracaoDataVenda = "Data da compra   " + CStr(rsSaidas("Data"))
      sDeclaracaoNumParc = "Parcelado em     x" + CStr(itotalNumParc)
      sDeclaracaoVencParc1 = "1ª parcela vence " + sVencimentoParcela1
      sDeclaracaoValorParc = "Cada parcela     R$ " + Format(sValorParcelas, "###,###,##0.00")
  Else
      Resp = InputBox("Impressão em modelo:" & vbCrLf & vbCrLf & "     1 - TICKET         [40 colunas]" & vbCrLf & vbCrLf & "     2 - RELATÓRIO [110 colunas]", "Qual o modelo de impressão?", "2")
      If Not IsNumeric(Resp) Then
          DisplayMsg "Opção de impressão inválida!"
          Exit Sub
      Else
          If Resp <> "1" And Resp <> "2" Then
              DisplayMsg "Opção de impressão inválida!"
              Exit Sub
          End If
      End If

      If Resp = "2" Then
          CrystalReport1.Destination = 0
          
          strNomeArq = gsReportPath & "carne02.rpt"
      Else
          CrystalReport1.WindowShowPrintSetupBtn = True
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Destination = IIf(False, crptToWindow, crptToPrinter)
          
          strNomeArq = gsReportPath & "carne02_todasParcelas_46Colunas.rpt"
      End If
  End If
  
  If Dir(strNomeArq) = "" Then
    DisplayMsg "Arquivo """ & strNomeArq & """ não encontrado."
    Exit Sub
  End If

  CrystalReport1.DataFiles(0) = gsQuickDBFileName
  CrystalReport1.ReportFileName = strNomeArq
  CrystalReport1.ParameterFields(5) = "pSequencia;" & rsSaidas("Sequência") & ";true"
  CrystalReport1.ParameterFields(6) = "pFilial;" & gnCodFilial & ";true"
  
  Dim sEmpresaNome As String
  Dim sEmpresaRuaComNumero As String
  Dim sEmpresaCidadeEstado As String
  Dim sEmpresaFone As String
  Dim sEmpresaCep As String
  
  If Len(gsNomeFilial) > 30 Then
      sEmpresaNome = Mid(gsNomeFilial, 1, 30)
  Else
      sEmpresaNome = gsNomeFilial
  End If
  
  If Len(gsFilialEndereco) > 30 Then
      sEmpresaRuaComNumero = Mid(gsFilialEndereco, 1, 30)
  Else
      sEmpresaRuaComNumero = gsFilialEndereco
  End If

  If Len(gsFilialCidadeEstado) > 30 Then
      sEmpresaCidadeEstado = Mid(gsFilialCidadeEstado, 1, 30)
  Else
      sEmpresaCidadeEstado = gsFilialCidadeEstado
  End If

  If Len(gsFilialFone) > 14 Then
      sEmpresaFone = Mid(gsFilialFone, 1, 14)
  Else
      sEmpresaFone = gsFilialFone
  End If

  If Len(gsFilialCep) > 11 Then
      gsFilialCep = Mid(gsFilialFone, 1, 11)
  Else
      sEmpresaCep = gsFilialCep
  End If
  
  CrystalReport1.ParameterFields(0) = "pEmpresa;" & sEmpresaNome & ";true"
  CrystalReport1.ParameterFields(4) = "pEmpresaEnderecoRua;" & sEmpresaRuaComNumero & ";true"
  CrystalReport1.ParameterFields(2) = "pEmpresaEnderecoCidadeEstado;" & sEmpresaCidadeEstado & ";true"
  CrystalReport1.ParameterFields(3) = "pEmpresaEnderecoFone;" & sEmpresaFone & ";true"
  CrystalReport1.ParameterFields(1) = "pEmpresaEnderecoCep;" & "Cep " & sEmpresaCep & ";true"
  
  If pImpressaoDireta = 2 Then
      CrystalReport1.ParameterFields(7) = "pDeclaracao;" & sDeclaracao & ";true"
      'CrystalReport1.ParameterFields(8) = "pDeclaracaoTitulo;" & "-  Recibo de entrega de carnê    -" & ";true"
      CrystalReport1.ParameterFields(8) = "pDeclaracaoTitulo;" & "-      Declaração de Dívida      -" & ";true"
      CrystalReport1.ParameterFields(9) = "pDeclaracaoTotalValor;" & sDeclaracaoTotalVenda & ";true"
      CrystalReport1.ParameterFields(10) = "pDeclaracaoData;" & sDeclaracaoDataVenda & ";true"
      CrystalReport1.ParameterFields(11) = "pDeclaracaoNmParc;" & sDeclaracaoNumParc & ";true"
      CrystalReport1.ParameterFields(12) = "pDeclaracaoParc1;" & sDeclaracaoValorParc & ";true"
      CrystalReport1.ParameterFields(13) = "pDeclaracaoDataVenc;" & sDeclaracaoVencParc1 & ";true"
  End If
  
  CrystalReport1.WindowState = crptMaximized
  
  Call SetPrinterName("REL", CrystalReport1)

  CrystalReport1.Action = 1

  Exit Sub
Erro:
  MsgBox "Erro tentando gerar Carnês. Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub EmiteCarnes()
 Dim Nome_Arq As String
 Dim Aux_Contador As Long
 Dim Mensa As String
 Dim Resp As Integer
 Dim Impressos As Integer
 Dim F As Form
  
 Call StatusMsg("")

 If IsNull(Num_Registro) Then
   DisplayMsg "Encontre ou grave uma venda antes."
   Exit Sub
 End If
 
  '27/09/2007 - Anderson
  'Implementação da impressão de carnê com código de barras
  'Solicitado por: Naativa
  If g_bolCarneCodigoBarras Then
    Set F = New frmImprimeCarneCodigoBarrasConfirmar
    F.Caption = "Impressão de Carnês"
    F.intFilial = gnCodFilial
    F.lngSeq = rsSaidas("Sequência")
    F.Show vbModal
    Exit Sub
  End If
 
 Rem mostra tela para escolha de configuração
  Set F = New frmObsDoc
  F.Caption = "Impressão de Carnês"
  F.gsFileExt = ".CCA"
  F.Show vbModal
  Set F = Nothing
 If gsRetornoDoc <> "OK" Then
'   DisplayMsg "Impressão cancelada."
   Exit Sub
 End If
  
 Nome_Arq = gsConfigPath & gsDocFileName & ".CCA"
 If Dir(Nome_Arq) = "" Then
   DisplayMsg "Arquivo """ & Nome_Arq & """ não encontrado."
   Exit Sub
 End If

 rsContas_Receber.Index = "Sequência"
 Impressos = 0
 Aux_Contador = 0
 
Lp1:
 rsContas_Receber.Seek ">", gnCodFilial, "R", rsSaidas("Sequência"), Aux_Contador
 If rsContas_Receber.NoMatch Then GoTo Fim
 If rsContas_Receber("Filial") <> gnCodFilial Then GoTo Fim
 If rsContas_Receber("Sequência") <> rsSaidas("Sequência") Then GoTo Fim
 If rsContas_Receber("Tipo") <> "R" Then GoTo Fim
 
 Aux_Contador = rsContas_Receber("Contador")
 
 If rsContas_Receber("Tipo Parcelamento") <> "T" Then GoTo Lp1
  
  
 If rsContas_Receber("Impresso") = True Then
   Mensa = "O carnê com vencimento em " + CStr(rsContas_Receber("Vencimento"))
   Mensa = Mensa + " no valor de R$ " + Format(rsContas_Receber("Valor"), "###,###,##0.00")
   Mensa = Mensa + " já foi impresso. DESEJA A REIMPRESSÃO ?"
 Else
   Mensa = "Deseja imprimir o carnê com vencimento em " + CStr(rsContas_Receber("Vencimento"))
   Mensa = Mensa + " no valor de R$ " + Format(rsContas_Receber("Valor"), "###,###,##0.00")
   Mensa = Mensa + " ?"
 End If
   
 Resp = MsgBox(Mensa, vbOKCancel, "Atenção")
 If Resp = vbCancel Then GoTo Lp1
 
 Resp = Imprime_Carnê("R", rsContas_Receber("Filial"), rsContas_Receber("Vencimento"), rsContas_Receber("Contador"), Nome_Arq)
 
 If Resp <> 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Voce deseja continuar com a impressão?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Exit Sub
    End If
 Else
    Impressos = Impressos + 1
    rsContas_Receber.Edit
    rsContas_Receber("Impresso") = True
    rsContas_Receber("Data Alteração") = Format(Date, "dd/mm/yyyy")
    rsContas_Receber.Update
 End If
 
 GoTo Lp1
 
Fim:
  DisplayMsg "Foram impressos " + str(Impressos) + " carnê(s)."
  
  '18/05/2005 - Daniel
  'Adicionado Exit Sub para encerrar a rotina
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Imprimir documento."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  SetPrinterName "REL"
  Exit Sub

  
End Sub

Private Sub ConsultaProduto()
  Dim F As Form
  
  '26/08/2004 - Daniel
  'Criado validação para verificar se o usuário possui permissão
  'para enchergar o estoque ou não
  Call EnchergarEstoque
  
  If m_blnPermitido Then
    Call StatusMsg("Aguarde....")
    '31/08/2006 - Anderson
    'Implementação de pesquisa avançada na tela de consulta do produto
    'Set F = New frmConsultaProd
    'F.Show vbModal
  
    If gbMostrarTelaPesquisaProdutoTipoFoto = True Then
        frmConsultaProd.Show
    Else
        frmPesquisaProduto.Show
    End If
  
  'Set F = Nothing
    Call StatusMsg("")
  Else
    MsgBox "Usuário não possui permissão para visualizar o estoque do produto.", vbExclamation, "Atenção"
  End If
  
End Sub

Private Sub UndoMovimento()
  Dim nRet As Integer
  Dim nMoviment As Long
  Dim CancelaNFCe As New clsNFCe
  
  lbl_retornoEnvioNFCe.Visible = False
  
  Call StatusMsg("")
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre uma saída antes."
  ElseIf Not rsSaidas("Efetivada") Then
    DisplayMsg "Esta operação não foi efetivada."
  ElseIf rsSaidas("Movimentação Desfeita") Then
    '30/01/2017 Jean
    'Adicionado tratamento para poder pedir o cancelamento de NFCe mesmo com a venda já desfeita
    If (rsSaidas("NFCe") > 0) And (rsSaidas("Nota Cancelada") = False) Then
        Call StatusMsg("Aguarde, cancelando nota...")
        CancelaNFCe.CancelaNFCe (Número.Text)
        
        If gsRetornoDoc = "OK" Then
            rsSaidas.Edit
            rsSaidas("Nota Cancelada") = True
            rsSaidas.Update
            
            DisplayMsg "Movimentação desfeita e Cupom Fiscal cancelado"
        Else
            DisplayMsg "Movimentação desfeita, porém o Cupom Fiscal NÃO foi cancelado devido INCONSISTÊNCIA."
            Exit Sub
        End If
    Else
        If rsSaidas("Nota Cancelada") = True Then
            DisplayMsg "Esta Movimentação possivelmente já foi desfeita ANTERIORMENTE, pois o Cupom Fiscal JÁ ESTA cancelado. Para maiores detalhes, já até a tela de Saídas e consulte esta venda por lá."
            Exit Sub
        End If
    End If
    
  ElseIf rsSaidas("Nota Cancelada") Then
    DisplayMsg "A Nota Fiscal desta movimentação já foi cancelada."
    
  ElseIf rsSaidas("Data") < CDate(Data_Atual) Then
    DisplayMsg "ATENÇÃO" & Chr(13) & Chr(13) & "Esta movimentação NÃO foi feita hoje e " & _
      "por isso NÃO PODE SER DESFEITA." & Chr(13) & Chr(13) & "Se desejar ajuste o " & _
      "estoque dos produtos e contas a receber manualmente."
  Else
    '30/01/2017 Jean
    'Alterado tratamento para poder cancelar a NFCe e desfazer a movimentação
    If rsSaidas("NFCe") > 0 Then
      
      Call StatusMsg("Aguarde, cancelando nota...")
      CancelaNFCe.CancelaNFCe (Número.Text)
      If gsRetornoDoc = "OK" Then
        rsSaidas.Edit
        rsSaidas("Nota Cancelada") = True
        rsSaidas.Update
        
        DisplayMsg "Pedido de Cancelamento de NFCe feito com sucesso"
        
      ElseIf MsgBox("Ocorreu um erro ao tentar cancelar a NFCe desta venda" & _
          " Deseja desfazer a movimentação?" _
          , vbExclamation + vbYesNo, "Atenção") = vbNo Then
          Exit Sub
      End If
    End If
    If rsSaidas("Nota Impressa") <> 0 Then
      If MsgBox("Esta movimentação já teve a Nota Fiscal impressa." & _
        " Deseja desfazer a movimentação e cancelar a Nota Fiscal?" _
        , vbExclamation + vbYesNo, "Atenção") = vbNo Then
        Exit Sub
      End If
    End If
    'If rsSaidas("Cupom Fiscal Impresso") = True Then
      'MsgBox ("Esta movimentação já teve o cupom fical impresso, por isso não pode ser desfeita")
      'Exit Sub
    'End If
    'Senha do gerente
    If Not frmGerente.gbSenhaGerente Then
      Exit Sub
    End If
    nMoviment = Val(Número.Text)
    
    ws.BeginTrans
    nRet = Desefetiva_Saída(gnCodFilial, nMoviment)
    If nRet = 0 Then
      ws.CommitTrans
    Else
      ws.Rollback
      DisplayMsg "Erro nº" & CStr(nRet) & " ao desfazer movimentação de saída."
      Exit Sub
    End If
    
    If rsSaidas("Nota Impressa") = 0 Then
      Call StatusMsg("Aguarde...")
      '30/01/2017 Jean
      'Comentei a parte do código responsavél pela exclusão da venda e de suas dependencias ao desfazer a movimentação a pedido do Mauro
      'Apaga movimentação de Saídas
'      Call EraseTypeMoviment(tmSaidas, gnCodFilial, nMoviment)
'      Call EraseTypeMoviment(tmSaidasProdutos, gnCodFilial, nMoviment)
'      Call EraseTypeMoviment(tmSaidasServicos, gnCodFilial, nMoviment)
'      Call EraseTypeMoviment(tmMovimentoCheques, gnCodFilial, nMoviment)
'      Call EraseTypeMoviment(tmMovimentoParcelas, gnCodFilial, nMoviment)
'      Número.Text = ""
'      Num_Registro = Null
      rsSaidas.Edit
      rsSaidas("Movimentação Desfeita") = True
      rsSaidas.Update
'      Efetivada.Visible = False
    Else
      rsSaidas.Edit
      rsSaidas("Nota Cancelada") = True
      rsSaidas("Movimentação Desfeita") = True
      rsSaidas.Update
    End If
    Call StatusMsg("")
    DisplayMsg "Operação desfeita."
  End If

End Sub

Private Sub EmiteFatura()
  Dim Resp As String
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Grave ou encontre uma venda antes."
    Exit Sub
  End If
  
  Resp = InputBox("Qual a data de vencimento ?", "Vencimento da Fatura", Date + 30)
  If Not IsDate(Resp) Then
    DisplayMsg "Data inválida."
    Exit Sub
  End If
  
  With frmEmiteFatura
    .Tipo.Caption = "F"
    .L_Encontrar.Caption = "NÃO"
    .L_Nota.Caption = rsSaidas("Nota Impressa") & ""
    .L_Fatura.Caption = ""
    .L_Valor.Caption = rsSaidas("Total")
    .L_Vencimento.Caption = Resp
    .L_Cliente.Caption = rsSaidas("Cliente")
    .lblDataEmissao.Caption = Data_Atual
    .lblSequencia.Caption = rsSaidas("Sequência")
    .Caption = "Impressão de Fatura"
    .Show vbModal
  End With
End Sub

Private Sub EmiteFaturaParcelados()
  Dim Resp As String
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Grave ou encontre uma venda antes."
    Exit Sub
  End If
  
  With frmEmiteFatura
    .Tipo.Caption = "F"
    .L_Encontrar.Caption = "NÃO"
    .L_Nota.Caption = rsSaidas("Nota Impressa") & ""
    .L_Fatura.Caption = ""
    .L_Valor.Caption = rsSaidas("Total")
    .L_Vencimento.Caption = Resp
    .L_Cliente.Caption = rsSaidas("Cliente")
    .lblDataEmissao.Caption = Data_Atual
    .lblSequencia.Caption = rsSaidas("Sequência")
    .Caption = "Impressão de Fatura"
    .optTotalParcela.Value = True
    .Show vbModal
  End With
End Sub


Private Sub InfoCliente()
 
  Call StatusMsg("")
  
  If IsNull(Nome_Cliente.Caption) Then Exit Sub
  If Nome_Cliente.Caption = "" Then Exit Sub
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Combo_Cliente.Text = "" Then Exit Sub
  
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) = 0 Then Exit Sub
  
  gsCodCliente = Combo_Cliente.Text
  
  '20/04/2006 - mpdea
  'Modificado exibição do form de informações do cliente
  'para poder ser acessível de diversas maneiras (Ex.: VR CheckOut)
  If g_frmVendaRapida Is frmVendaRap2_CheckOut Then
    frmInformacoes.Show , Me
  Else
    frmInformacoes.Show , frmMain
  End If
 
End Sub

Private Sub EmiteRecibo()
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Grave ou encontre uma venda antes."
    Exit Sub
  End If
  
  With frmEmiteFatura
    .Tipo.Caption = "R"
    .L_Encontrar.Caption = "NÃO"
    .L_Nota.Caption = rsSaidas("Nota Impressa") & ""
    .L_Fatura.Caption = ""
    .L_Valor.Caption = rsSaidas("Total")
    .L_Vencimento.Caption = ""
    .L_Cliente.Caption = rsSaidas("Cliente")
    .Caption = "Impressão de Recibo"
    .Show vbModal
  End With
End Sub


Private Sub Num_Cartão_GotFocus()
 Num_Cartão.SelStart = 0
 Num_Cartão.SelLength = Len(Num_Cartão.Text)
End Sub

Private Sub Número_CloseUp()
  '29/05/2003 - mpdea
  'Atualiza informações na tela sobre a movimentação
  Call Número_LostFocus
End Sub

Private Sub Número_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub SSResizer1_BeforeResize(ByVal Control As Variant, ByVal ResizeOptions As ActiveResizer.SSResizeOptions)

End Sub


Private Sub Observacao_KeyPress(KeyAscii As Integer)
      Select Case KeyAscii
      Case 8
         KeyAscii = KeyAscii
      Case 13
         KeyAscii = KeyAscii
      Case 32
         KeyAscii = KeyAscii
      Case 44 To 46 ' , - .
         KeyAscii = KeyAscii
      Case 48 To 57 'Numbers
         KeyAscii = KeyAscii
      Case 65 To 90 'Upper case letters
         KeyAscii = KeyAscii
      Case 97 To 122 'Lower case letters
         KeyAscii = KeyAscii
      Case Else 'Discard anything else
         KeyAscii = 0
      End Select
End Sub

Private Sub TrafficLight_StatusMessage(ByVal Message As String)
  Call StatusMsg(Message)
End Sub

'24/01/2006 - mpdea
'Incluído tratamento de erro para evitar Overflow
Private Sub Número_LostFocus()

On Error GoTo TratarErro
  
  Dim lngRet As Long

  If Número.Text <> "" Then
    Call IsDataType(dtLong, Número.Text, lngRet)
    '12/09/2013 - Jean
    'Incluído tratamento para evitar erro se não trabalha com comanda
    If (rsParametros("TrabalharComComanda").Value = 1) Then
      Call BuscarComanda
    End If
    If lngRet > 0 Then
      Call Mostra_Mov(lngRet)
    Else
      DisplayMsg "Número inválido para sequência de movimentação."
    End If
  End If
'  If IsNumeric(Número.Text) Then
'    Mostra_Mov Número.Text
'  End If

  Exit Sub

TratarErro:
  MsgBox "Por favor, tente novamente.", vbInformation, "Atenção"
End Sub

Private Sub txtComanda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 9 Or KeyAscii = 10 Then
    btnComandaVendas.Visible = False
    txtComanda.Width = 1785

    If Trim(txtComanda.Text) <> "" Then
      frmComanda.Comanda = Trim(txtComanda.Text)
      If frmComanda.ComandaOK Then
        If frmComanda.Total > 0 Then
          If frmComanda.Sequencia > 0 Then
            Mostra_Mov frmComanda.Sequencia
          Else
            txtComanda.Width = 1395
            btnComandaVendas.Visible = True
          End If
        End If
      End If
    End If
  ElseIf ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  ElseIf KeyAscii <> 8 And KeyAscii <> 127 Then
    KeyAscii = 0
  End If
  Refresh
End Sub

Public Sub SearchRecord_peloNumComanda(ByVal Num As Long)
  Dim lngSequencia As Long
  Dim sSQlComanda As String
  Dim rsSaidasComandas As Recordset

  On Error GoTo ErrHandler
  
'  If Not IsNumeric(Num) Then
'      DisplayMsg "Número de comanda para pesquisa inválida."
'      Exit Sub
'  End If

  If Num > 0 Then
      sSQlComanda = "SELECT CodSaida FROM SaidasComandas WHERE CodComanda = '" & Num & "'"
      sSQlComanda = sSQlComanda & " And Filial = " & gnCodFilial & ""
  Else
      DisplayMsg "Número de sequência para pesquisa inválida."
      Exit Sub
  End If

  Set rsSaidasComandas = db.OpenRecordset(sSQlComanda, dbOpenDynaset)
  
  If Not (rsSaidasComandas.EOF And rsSaidasComandas.BOF) Then
      rsSaidasComandas.MoveFirst
      Num = rsSaidasComandas.Fields("CodSaida").Value
      rsSaidasComandas.Close
      Set rsSaidasComandas = Nothing
  Else
      rsSaidasComandas.Close
      Set rsSaidasComandas = Nothing
      DisplayMsg "Não existe uma venda relacionada a este número de comanda."
      Exit Sub
  End If

  Mostra_Mov Num
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub txtComanda_LostFocus()
  txtComanda_KeyPress (13)
End Sub

Private Sub Val_Cartão_GotFocus()
  Val_Cartão.SelStart = 0
  Val_Cartão.SelLength = Len(Val_Cartão.Text)
End Sub

Private Sub Val_Cartão_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Val_Cartão_LostFocus()
  Call Recalcula_Recebido
End Sub

Private Sub Val_Cheque_GotFocus(Index As Integer)
  Val_Cheque(Index).SelStart = 0
  Val_Cheque(Index).SelLength = Len(Val_Cheque(Index).Text)
End Sub

Private Sub Val_Cheque_KeyPress(Index As Integer, KeyAscii As Integer)
 If Not IsDate(Bom_Para(Index).Text) Then
   KeyAscii = 0
   Exit Sub
 End If
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Val_Cheque_LostFocus(Index As Integer)
  If IsNumeric(Val_Cheque(Index).Text) Then
    If CDbl(Val_Cheque(Index).Text) Then
       Call StatusMsg("")
    End If
  End If
  Call Recalcula_Recebido
End Sub

Private Sub Val_Parc_GotFocus(Index As Integer)
  Val_Parc(Index).SelStart = 0
  Val_Parc(Index).SelLength = Len(Val_Parc(Index).Text)
End Sub

Private Sub Val_Parc_KeyPress(Index As Integer, KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Val_Parc_LostFocus(Index As Integer)
  Call Recalcula_Recebido
End Sub

Private Sub Vale_GotFocus()
  Vale.SelStart = 0
  Vale.SelLength = Len(Vale.Text)
End Sub

Private Sub Vale_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Vale_LostFocus()
  Call Recalcula_Recebido
End Sub

Private Sub RecalculaPrecos()
'  Dim nRow As Long
'  Dim sCodProd As String
'  Dim bm As Variant
'  Screen.MousePointer = vbHourglass
'  Call StatusMsg("Refazendo tabela..."
'  rsPreços.Index = "Tabela"
'
'  For nRow = 0 To Linhas_Grade - 1
'    sCodProd = gsHandleNull(Tabe(nRow).Código)
'    If sCodProd <> "0" Then
'      rsPreços.Seek "=", Combo_Preço.Text, sCodProd
'      If rsPreços.NoMatch Then
'        Tabe(nRow).Preço = 0#
'      Else
'        Tabe(nRow).Preço = rsPreços("Preço")
'      End If
'    Else
'      Tabe(nRow).Preço = 0#
'    End If
'    Call Calcula_Linha_Tabe(nRow)
'  Next nRow
'
'  Grade1.MoveLast
'  Grade1.MoveFirst
'
  Call Recalcula
'
'  Screen.MousePointer = vbDefault
'  Call StatusMsg("")

End Sub

Private Sub RecalculaPesos()
  Dim sCodProd As String
  Dim sCod As String
  Dim nRow As Long
'  Dim bm As Variant
  Dim Aux_Produto As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim nQtde As Single
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Refazendo tabela...")
  
  Grade1.Update
  Grade1.MoveFirst
  gnPesoLiquido = 0#
  gnPesoBruto = 0#
  
'  For nRow = 0 To (Linhas_Grade - 1)
  For nRow = LBound(Tabe) To UBound(Tabe)
'    bm = Grade1.GetBookmark(nRow)
'    Grade1.Bookmark = bm
    sCodProd = gsHandleNull(Tabe(nRow).Código)
    If sCodProd <> "0" Then  'Faz somente os preenchidos
'      nQtde = Grade1.Columns(1).CellText(bm)
      nQtde = Tabe(nRow).Qtde
      sCod = sCodProd
      Call Acha_Produto(sCod, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Tipo, Aux_Erro)
      sCod = Aux_Produto
      If Aux_Erro = 0 Then
        rsProdutos2.FindFirst "Código = '" & Aux_Produto & "'"
        gnPesoLiquido = gnPesoLiquido + nQtde * gsHandleNull(rsProdutos2("PesoLiquido"))
        gnPesoBruto = gnPesoBruto + nQtde * gsHandleNull(rsProdutos2("PesoBruto"))
      End If
    End If
  Next nRow
  
  Grade1.MoveFirst
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
End Sub

'24/10/2002 - mpdea
'Corrigido verificação de estoque para produtos com grade/edição
'
'09/10/2002 - mpdea
'Função que verifica o estoque dos produtos
Private Function mblnCheckStock() As Boolean
  Dim intRow As Integer
  Dim intX As Integer
  Dim strCodProdPrincipal As String
  Dim intTamanho As Integer
  Dim intCor As Integer
  Dim lngEdicao As Long
  Dim intErr As Integer
  Dim intCountItem As Integer
  Dim blnNewItem As Boolean
  Dim dblEstoque As Double
  Dim blnStockInsufficient As Boolean

  'Tabela para controle dos produtos a serem verificados o estoque
  intCountItem = 0
  ReDim typCheckStock(intCountItem) As CheckStock

  For intRow = LBound(Tabe) To UBound(Tabe)
    With Tabe(intRow)

      If .Código <> "" And .Qtde <> 0 And .Nome <> "" Then

        Call Acha_Produto(.Código, strCodProdPrincipal, 0, 0, 0, 0, intErr)

        Select Case intErr
          Case 1
            DisplayMsg "Produto [" & .Código & "] não existe."
          Case 2
            DisplayMsg "Produto [" & .Código & "] com grade sem tamanho e cor informados."
          Case 3
            DisplayMsg "Produto [" & .Código & "] com edição sem edição informada."
        End Select
        If intErr <> 0 Then Exit Function
        
        rsProdutos.Index = "Código"
        rsProdutos.Seek "=", strCodProdPrincipal
        If rsProdutos.NoMatch Then
          DisplayMsg "Produto [" & .Código & "] não existe."
          Exit Function
        End If

        'Verifica se o produto possui estoque controlado
        If rsProdutos.Fields("Estoque").Value Then

          blnNewItem = True
          For intX = LBound(typCheckStock) To UBound(typCheckStock)
            'Item já listado
            If typCheckStock(intX).strCode = .Código Then
              typCheckStock(intX).dblQuantity = _
              typCheckStock(intX).dblQuantity + .Qtde
              blnNewItem = False
              Exit For
            End If
          Next intX

          'Novo item (agrupa)
          If blnNewItem Then
            ReDim Preserve typCheckStock(intCountItem) As CheckStock
            typCheckStock(intCountItem).strCode = .Código
            typCheckStock(intCountItem).dblQuantity = .Qtde
            intCountItem = intCountItem + 1
            intX = intCountItem
          End If

        End If

      End If

    End With
  Next intRow

  'Fim da criação da lista

  'Início da verificação do estoque
  For intX = LBound(typCheckStock) To UBound(typCheckStock)

    With typCheckStock(intX)

      '12/11/2002 - mpdea
      'Adicionado validação dos itens no array
      If .strCode <> "" And .dblQuantity <> 0 Then
      
        Call Acha_Produto(.strCode, strCodProdPrincipal, intTamanho, intCor, _
                          lngEdicao, 0, 0)
  
        'Estoque atual
        dblEstoque = -999999
        dblEstoque = Acha_Estoque(gnCodFilial, strCodProdPrincipal, intTamanho, intCor, lngEdicao, intErr)
        
        .dblStock = dblEstoque
        
        If intErr = 0 And dblEstoque <> -999999 Then
          'Se o estoque atual for superior a quantidade a ser movimentada
          'ativa flag de estoque insuficiente
          If .dblQuantity > dblEstoque Then
            .blnStockInsufficient = True
            blnStockInsufficient = True
          End If
        Else
          '20/11/2002 - mpdea
          'Adicionado descrição do erro 1
          If intErr = 1 Then
            DisplayMsg "Produto [" & .strCode & "] com estoque não inicializado."
          Else
            DisplayMsg "Erro [" & intErr & "] ao encontrar estoque do produto."
          End If
          Exit Function
        End If
    
      End If
    
    End With

  Next intX
  
  'Exibe os produtos com estoque insuficiente
  If blnStockInsufficient Then
    '31/01/2007 - Anderson - Alterado para que a permissão que impede o usuário de ver a quantidade em estoque funcione
    If Not m_blnPermitido Then   'Não permitido
      DisplayMsg "Quantidade superior ao estoque."
    Else
      frmCheckStock.ShowStockInsufficient
    End If
    Exit Function
  End If
  
  'Todos os produtos possuem estoque para a movimentação
  mblnCheckStock = True

End Function

'26/12/2002 - mpdea
'Exibe tela para inserir produto obtendo a quantidade
'através da digitação do valor total
Private Sub IncluirProdutoXValor()
  Dim strTablePrice As String
  
  strTablePrice = Combo_Preço
  
  If strTablePrice = "" Then
    DisplayMsg "Tabela de preços não configurada."
    Call SelectAllText(Combo_Preço, True)
    Exit Sub
  End If
  
  If gbCheckTabPreco(strTablePrice) Then
    Call frmVendaValorXQtde.Start(strTablePrice, CBool(Calcula_IPI))
  Else
    DisplayMsg "Tabela de preços incorreta."
    Call SelectAllText(Combo_Preço, True)
  End If
End Sub


'06/08/2003 - maikel
'             Função que analisa o crédito do cliente para o recebimento simplificado
Private Function AnalisaCreditoCliente() As Boolean
  Dim lngCodigoCliente        As Long
  Dim intCheque               As Integer
  Dim intParcelamento         As Integer
  Dim blnRecebimentoFaturado  As Boolean
  Dim dblValorFaturado        As Double
  Dim rstCliente              As Recordset
  Dim dblLimiteCredito        As Double
  Dim dblValorRecebidoPrazo   As Double
  
  If Len(Trim(Nome_Cliente.Caption)) <= 0 Then
    AnalisaCreditoCliente = False
    Exit Function
  End If
  
  lngCodigoCliente = CLng(Combo_Cliente.Text)
  
  Call StatusMsg("Analisando o crédito do cliente . . . ")
  
  blnRecebimentoFaturado = False
  dblValorFaturado = 0
  
  For intCheque = Bom_Para.LBound To Bom_Para.UBound
    If IsNumeric(Val_Cheque(intCheque).Text) Then dblValorFaturado = dblValorFaturado + CDbl(Val_Cheque(intCheque).Text)
  
    If IsDate(Bom_Para(intCheque)) Then
      If CDate(Bom_Para(intCheque)) > CDate(Data_Atual) Then
        If Not blnRecebimentoFaturado Then
          blnRecebimentoFaturado = True
        End If
      End If
    End If
  Next intCheque
  
  
  For intParcelamento = Data_Parc.LBound To Data_Parc.UBound
    If IsNumeric(Val_Cheque(intParcelamento).Text) Then dblValorFaturado = dblValorFaturado + CDbl(Val_Parc(intParcelamento).Text)
  
    If IsDate(Data_Parc(intParcelamento)) Then
      If CDate(Data_Parc(intParcelamento)) > CDate(Data_Atual) Then
        If Not blnRecebimentoFaturado Then
          blnRecebimentoFaturado = True
        End If
      End If
    End If
  Next intParcelamento
  
  If Not blnRecebimentoFaturado Then
    blnRecebimentoFaturado = Lançar_Débito.Value
  End If
  
  
  
  If blnRecebimentoFaturado Then
    Set rstCliente = db.OpenRecordset(" SELECT Faturado, [Limite Crédito] FROM Cli_For " & _
                                       " WHERE Código = " & lngCodigoCliente, dbOpenSnapshot)
    
    With rstCliente
      If Not (.BOF And .EOF) Then
        If .Fields("Limite Crédito") = 0 Then
          AnalisaCreditoCliente = True
        Else
          If (Not .Fields("Faturado")) And (dblValorFaturado > 0) Then
            MsgBox "O cliente ao qual você está fazendo recebimento não pode fazer compra faturada. Para mudar essa opção entre no cadastro de clientes e marque a opção [Compra a Prazo]", vbCritical, "Quick Store"
            AnalisaCreditoCliente = False
          Else
            dblLimiteCredito = (.Fields("Limite Crédito").Value - Pega_Limite_Usado(lngCodigoCliente))
            
            If (dblValorFaturado > dblLimiteCredito) Then
              MsgBox "O cliente ao qual você está fazendo o recebimento tem R$ " & _
                     Format(dblLimiteCredito, FORMAT_VALUE) & " de saldo para novas compras. O recebimento parcelado é de R$ " & _
                     Format(dblValorFaturado, FORMAT_VALUE) & ". Não é possivel continuar !! ", vbCritical, "Quick Store"
              
              AnalisaCreditoCliente = False
            Else
              AnalisaCreditoCliente = True
            End If
          End If
        End If
      End If
      
      If Not rstCliente Is Nothing Then .Close
      Set rstCliente = Nothing
    End With
  Else
    AnalisaCreditoCliente = True
  End If

  Call StatusMsg("")
End Function

'11-12/08/2003 - mpdea
'Habilita e desabilita os controles abaixo conforme parâmetro
'Usado para impedir alterações durante operações
Private Sub EnableControls(ByVal blnEnabled As Boolean)
  L_Tot_Pagar.SetFocus
  fraButtonRecebeSimples.Enabled = blnEnabled
  fraButtons.Enabled = blnEnabled
End Sub

'23/05/2006 - mpdea
'Comentado função abaixo devido otimizado na verificação de cliente isento de IPI
'
'Private Function IsencaoIPI(ByVal CodCliente As Long) As Boolean
'  '07/05/2004 - Daniel
'  'Case: Embalavi
'  'Esta função tem a finalidade de verificar na tabela Cli_For se o
'  'Cli_For é Isento de IPI
'  Dim rstCliFor As Recordset
'
'  Set rstCliFor = db.OpenRecordset("SELECT IsentoIPI FROM Cli_For WHERE Código = " & CodCliente, dbOpenDynaset)
'
'  With rstCliFor
'    If Not (.BOF And .EOF) Then
'      IsencaoIPI = .Fields("IsentoIPI").Value
'    End If
'    .Close
'  End With
'
'  Set rstCliFor = Nothing
'
'End Function

Private Function Diferimento(ByVal CodCliente As Long) As Boolean
  '07/05/2004 - Daniel
  'Case: Embalavi
  'Esta função tem a finalidade de veridicar na tabela Cli_For se o
  'estado do Cli_For é PR e se é pessoa jurídica
  Dim rstCliFor As Recordset
  
  Set rstCliFor = db.OpenRecordset("SELECT Estado, Física_Jurídica FROM Cli_For WHERE Código = " & CodCliente, dbOpenDynaset)

  With rstCliFor
    If Not (.BOF And .EOF) Then
      If .Fields("Estado").Value = "PR" And .Fields("Física_Jurídica").Value = "J" Then
        Diferimento = True
      Else
        Diferimento = False
      End If
    End If
    .Close
  End With
  
  Set rstCliFor = Nothing

End Function

Private Sub EnchergarEstoque()
  '26/08/2004 - Daniel
  'Criado validação para verificar se o usuário possui permissão
  'para enchergar o estoque ou não
  Dim rstFuncionarios As Recordset
  
  Set rstFuncionarios = db.OpenRecordset("SELECT Código, VRVisualizarEstoque FROM Funcionários WHERE Código = " & gnUserCode, dbOpenDynaset)
  
  With rstFuncionarios
   If Not (.BOF And .EOF) Then
     .MoveFirst
     
     m_blnPermitido = .Fields("VRVisualizarEstoque").Value
   End If
   .Close
  End With
  
  Set rstFuncionarios = Nothing

End Sub

Private Function IE_Isento(ByRef Estado As String) As Boolean
  '09/11/2004 - Daniel
  'Verificação da I.E. retorna o estado do Cliente
  Dim rstCliFor As Recordset
  Dim strSQL    As String
  Dim strIE     As String
  
  If Len(Nome_Cliente.Caption) <= 0 Then Exit Function
  
  strSQL = "SELECT Inscrição, Estado FROM Cli_For "
  strSQL = strSQL & " WHERE Código = " & CLng(Combo_Cliente.Text)

  Set rstCliFor = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstCliFor
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      strIE = .Fields("Inscrição").Value & ""
      strIE = UCase(strIE)
      
      If strIE = "ISENTO" Or strIE = "" Then IE_Isento = True
      
      Estado = Trim(.Fields("Estado").Value)
      
    End If
    .Close
  End With
  
  Set rstCliFor = Nothing

End Function

Private Function PessoaFisica(ByVal CodCliente As Long) As Boolean
  'Function criada em 07/12/2004 por Daniel
  'Finalidade: Atender as necessidades da Embalavi
  Dim rstCliFor As Recordset
  Dim strSQL    As String
  
  strSQL = "SELECT Física_Jurídica FROM Cli_For "
  strSQL = strSQL & " WHERE Código = " & CodCliente
    
  Set rstCliFor = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstCliFor
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .Fields("Física_Jurídica").Value = "F" Then PessoaFisica = True
      
    End If
    .Close
  End With
  
  Set rstCliFor = Nothing
  
End Function

Private Function GetFabricante(ByVal CodProdu As String) As String
  '29/03/2005 - Daniel
  'Case: Elétrica Leal
  'Exibição da coluna Fabricante
  Dim rstProdutos As Recordset
  
  GetFabricante = ""
  
  Set rstProdutos = db.OpenRecordset("SELECT Fabricante FROM Produtos WHERE Código = '" & CodProdu & "'", dbOpenDynaset)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      GetFabricante = .Fields("Fabricante").Value & ""
    End If
    .Close
  End With
  
  Set rstProdutos = Nothing

End Function

Private Sub GetLimiteCliente(ByVal lngCliente As Long, ByRef dblLimite As Double)
  '02/05/2005 - Daniel
  Dim rstCliente As Recordset
  
  dblLimite = 0
  
  Set rstCliente = db.OpenRecordset("SELECT [Limite Crédito] FROM Cli_For WHERE Código = " & lngCliente, dbOpenDynaset)
  
  With rstCliente
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      dblLimite = Format(CDbl("0" & .Fields("Limite Crédito").Value), FORMAT_VALUE)
    End If
    .Close
  End With

  Set rstCliente = Nothing

End Sub

Private Function ValidarDesconto() As Boolean
  '02/06/2005 - Daniel
  'Criado rotina para verificar se existe um ou mais
  'produtos que não permitem desconto
  Dim intI               As Integer
  Dim strArray(1 To 300) As String
  Dim intAuxi            As Integer
  Dim strMsgCabe         As String
  Dim strMsg             As String
  Dim intX               As Integer
  
  On Error GoTo TratarErro
  
  For intI = 0 To (Grade1.Rows - 1)
    If Tabe(intI).Código <> "0" And Tabe(intI).Código <> "" Then
      If blnProdNaoPermiteDesconto(Tabe(intI).Código) Then
        intAuxi = intAuxi + 1
        strArray(intAuxi) = (Tabe(intI).Código)
      End If
    End If
  Next intI
  
  If strArray(1) <> "" Then
    
    If strArray(2) <> "" Then
      strMsgCabe = "Para estes Produtos não são permitidos Descontos: " & vbCrLf & vbCrLf
    Else
      strMsgCabe = "Para este Produto não é permitido Desconto: " & vbCrLf & vbCrLf
    End If
    
    For intX = 1 To intAuxi
      strMsg = strMsg & strArray(intX) & " "
    Next intX
    
    MsgBox strMsgCabe & strMsg, vbExclamation, "Atenção"
    
    If Not frmGerente.gbSenhaGerente Then
      ValidarDesconto = True
      Grade1.MoveLast
      Grade1.MoveFirst
      Exit Function
    End If
    
  End If
  
  Grade1.MoveLast
  Grade1.MoveFirst
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function

Private Function blnProdNaoPermiteDesconto(ByVal CodProd As String) As Boolean
  '02/06/2005 - Daniel
  'Criado rotina para verificar se o produto
  'permite ou não o desconto
  Dim rstCodGrade As Recordset
  Dim rstProdutos As Recordset
  Dim strSQL      As String
  Dim srtProduto  As String
  
  On Error GoTo TratarErro
  
  blnProdNaoPermiteDesconto = False
  srtProduto = ""
  
  strSQL = "SELECT [Código Original] FROM [Códigos da Grade] WHERE Código = '" & CodProd & "'"
  
  Set rstCodGrade = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstCodGrade.RecordCount <> 0 Then 'Utiliza grade
    With rstCodGrade
      If Not (.BOF And .EOF) Then
        .MoveFirst
        srtProduto = .Fields("Código Original").Value & ""
      End If
    End With
  End If
  
  rstCodGrade.Close
  Set rstCodGrade = Nothing
  
  If srtProduto <> "" Then
    strSQL = ""
    strSQL = "SELECT DontAllowDesc FROM Produtos WHERE Código = '" & srtProduto & "'"
  Else
    strSQL = ""
    strSQL = "SELECT DontAllowDesc FROM Produtos WHERE Código = '" & CodProd & "'"
  End If
  
  Set rstProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      blnProdNaoPermiteDesconto = .Fields("DontAllowDesc").Value
    End If
    .Close
  End With
  
  Set rstProdutos = Nothing
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Atenção"
  
End Function

Private Function UserSemPermissao(ByVal intCodFunc As Integer) As Boolean
  '01/06/2005 - Daniel
  Dim rstFunc As Recordset
  
  On Error GoTo TratarErro
  
  Set rstFunc = db.OpenRecordset("SELECT AllowDescProd FROM Funcionários WHERE Código = " & intCodFunc, dbOpenDynaset)
  
  With rstFunc
    If Not (.BOF And .EOF) Then
      .MoveFirst
      UserSemPermissao = .Fields("AllowDescProd").Value
    End If
    .Close
  End With
  
  Set rstFunc = Nothing
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function

Private Function UpdateTotalNCM()
'  rsSaidas.Edit
'  Dim totalNCM As Double 'Total em R$ de imposto pago na movimentação
'  Dim Valor_Aprox_Impostos As Double
'  Dim rsAliquotas As Recordset 'Tabela que filtra todos os produtos da sequencia
'  Dim rsProdutos3 As Recordset 'Tabela que filtra produto por produto da movimentação
'  totalNCM = 0#
'  Set rsAliquotas = db.OpenRecordset("SELECT [Código],[Preço Final],[Valor_Aprox_Impostos] FROM [Saídas - Produtos] WHERE [Sequência] = " & Número.Text, dbOpenDynaset)
'  rsAliquotas.MoveFirst
'  While Not rsAliquotas.EOF
'    Set rsProdutos3 = db.OpenRecordset("SELECT [AliqNCM] FROM [Produtos] WHERE [Código] = '" & rsAliquotas("Código") & "'", dbOpenDynaset)
'    rsProdutos3.MoveFirst
'    If (rsProdutos3("AliqNCM") <> "" Or rsProdutos3("AliqNCM") = 0) Then
'      totalNCM = totalNCM + (rsProdutos3("AliqNCM") * rsAliquotas("Preço Final") / 100)
'      rsAliquotas.Edit
'      rsAliquotas("Valor_Aprox_Impostos") = Format((rsProdutos3("AliqNCM") * rsAliquotas("Preço Final") / 100), "##,###,##0.00")
'      rsAliquotas.Update
'    Else
'      'MsgBox "O produto " & rsAliquotas("Código") & " não possui aliquota de NCM", vbExclamation
'    End If
'    rsAliquotas.MoveNext
'  Wend
'  rsSaidas("TotalNCM") = totalNCM
'  rsSaidas("TotalNCM") = Format(rsSaidas("TotalNCM"), "##,###,##0.00")
'  rsSaidas.Update
rsSaidas.Edit
  Dim totalNCM As Double 'Total em R$ de imposto pago na movimentação
  Dim Valor_Aprox_Impostos As Double
  Dim rsAliquotas As Recordset 'Tabela que filtra todos os produtos da sequencia
  Dim rsProdutos3 As Recordset 'Tabela que filtra produto por produto da movimentação
  totalNCM = 0#
  Set rsAliquotas = db.OpenRecordset("SELECT [Código Sem Grade],[Preço Final],[Valor_Aprox_Impostos],[MotivoDesoneracaoICMS] FROM [Saídas - Produtos] WHERE [Sequência] = " & Número.Text, dbOpenDynaset)
  On Error GoTo UpdateExit
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
    
    rsProdutos3.Close
  Wend
  rsSaidas("TotalNCM") = totalNCM
  'rsSaidas("TotalNCM") = FormataValorTexto(rsSaidas("TotalNCM"), 2)
  rsSaidas.Update
  
  rsAliquotas.Close
  Set rsAliquotas = Nothing
  
UpdateExit:
End Function

Private Function UpdateTotalNCM_2(ByVal sCodProduto As String)
On Error GoTo UpdateExit
  
  Dim Valor_Aprox_Impostos As Double
  Dim rsProdutos3 As Recordset 'Tabela que filtra produto por produto da movimentação
  
  Set rsProdutos3 = db.OpenRecordset("SELECT [AliqNCM],[MotivoDesoneracaoICMS] FROM [Produtos] WHERE [Código] = '" & sCodProduto & "'") ', dbOpenDynaset)
  rsProdutos3.MoveFirst
  If (rsProdutos3("AliqNCM") <> "" Or rsProdutos3("AliqNCM") = 0) Then
      Valor_Aprox_Impostos = (rsProdutos3("AliqNCM") * rsSaída_Prod("Preço Final") / 100)
      Valor_Aprox_Impostos = FormatNumber(Valor_Aprox_Impostos, 2)
      totalNCM_2 = totalNCM_2 + (rsProdutos3("AliqNCM") * rsSaída_Prod("Preço Final") / 100)
      totalNCM_2 = FormatNumber(totalNCM_2, 2)
      
      rsSaída_Prod("Valor_Aprox_Impostos") = Valor_Aprox_Impostos
      rsSaída_Prod("MotivoDesoneracaoICMS") = rsProdutos3("MotivoDesoneracaoICMS")
  Else
      rsSaída_Prod("MotivoDesoneracaoICMS") = rsProdutos3("MotivoDesoneracaoICMS")
  End If
    
  rsProdutos3.Close
  
UpdateExit:
End Function


Private Function GravarComanda()
  '05/07/2013-Alexandre Afornali
  'Case DiskEmbalagens
  Dim valida As Boolean
  valida = True
  If (txtComanda.Text <> "") Then
    Dim rsComandas As Recordset
    Set rsComandas = db.OpenRecordset("SaidasComandas")
    
    If Not (rsComandas.EOF And rsComandas.BOF) Then
        rsComandas.MoveFirst
    End If
    
    While Not rsComandas.EOF
      If (rsComandas("CodSaida") = Número.Text And rsComandas("Filial") = gnCodFilial) Then
        rsComandas.Edit
        rsComandas("CodComanda") = txtComanda.Text
        rsComandas.Update
        rsComandas.MoveLast
        valida = False
      End If
      rsComandas.MoveNext
    Wend
    If (valida = True) Then
      rsComandas.AddNew
      rsComandas("CodComanda") = txtComanda.Text
      rsComandas("CodSaida") = Número.Text
      rsComandas("Filial") = gnCodFilial
      rsComandas.Update
    End If
  End If
End Function

Private Function BuscarComanda()
On Error GoTo Erro

  Dim rsComandas As Recordset
  Set rsComandas = db.OpenRecordset("SaidasComandas")
  rsComandas.MoveFirst
  While Not rsComandas.EOF
    If (rsComandas("CodSaida").Value = Número.Text) And (rsComandas("Filial").Value = gnCodFilial) Then
      txtComanda.Text = rsComandas("CodComanda")
      rsComandas.MoveLast
      rsComandas.MoveNext
    Else
      rsComandas.MoveNext
    End If
  Wend

  Exit Function
Erro:
  MsgBox "Erro em BuscarComanda " + Err.Number + " " + Err.Description, vbInformation, "Atenção"
End Function

Public Function CarregaComanda()
On Error GoTo Erro
  txtComanda.Text = ""
  
  Dim rsComandas As Recordset
  Set rsComandas = db.OpenRecordset("SaidasComandas")
  Dim countrs As Long
  countrs = 0
  
  While Not rsComandas.EOF
    countrs = countrs + 1
    rsComandas.MoveNext
  Wend
  
  If (countrs > 0) Then
    rsComandas.MoveFirst
  End If
  
  While Not rsComandas.EOF
    If (rsComandas("CodSaida") = Número And rsComandas("Filial") = gnCodFilial) Then
      txtComanda.Text = rsComandas("CodComanda")
      rsComandas.MoveLast
    End If
    rsComandas.MoveNext
  Wend

  Exit Function
Erro:
  MsgBox "Erro na busca do código da comanda " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Function

Private Function Retorno_PDV()
  Dim GestoBD As Database
  Dim Cfisc_Pgto As Recordset
  Dim TipoRecebimpgto As Recordset
  Dim DocumentoFiscal As Recordset
  Dim QuickBD As Database
  Dim Caixa As Recordset
  Dim CaixaAnterior As Recordset
  Dim Resumo_Diário_Financeiro As Recordset
  Dim Resumo_Diário As Recordset
  Dim Contas_Receber As Recordset
  Dim produtos As Recordset
  Dim cad_prod As Recordset
  Dim Estoque_Final As Recordset
  Dim Estoque As Recordset
  Dim Estoque_Anterior As Recordset
  Dim Cfisc_Base As Recordset
  If frmParametros.VerificaPAF = True Then
    'Atualiza Financeiro vindo do PAF
    Set rsParametros = db.OpenRecordset("Parâmetros Filial")
    
    Set GestoBD = OpenDatabase(rsParametros("BancoPDV").Value & "\Gesto.mde", False, False)
    Set DocumentoFiscal = GestoBD.OpenRecordset("Select * from DocumentoFiscal where Num_Docto = " & Número.Text & "")
    If DocumentoFiscal.EOF Then
      MsgBox "Cupom não encontrado, favor verificar"
      Exit Function
    End If
    Set Cfisc_Pgto = GestoBD.OpenRecordset("Select * From Cfisc_Pgto where FIS_NRO = " & DocumentoFiscal("Num_Docto_Fiscal") & "")
    Set TipoRecebimpgto = GestoBD.OpenRecordset("Select * From TipoRecebimpgto Where Cint(cod_Pdv) = '" & Cfisc_Pgto("Tipo_Pagto") & "'")
    Set Cfisc_Base = GestoBD.OpenRecordset("Select * From Cfisc_Base Where FIS_NRO = " & Cfisc_Pgto("FIS_NRO") & "")
    Cfisc_Base.Edit
    Cfisc_Base("Importado_Retaguarda") = True
    Cfisc_Base.Update
    'Cfisc_Base = Nothing
  
    Set Caixa = db.OpenRecordset("Select * from Caixa where Filial = " & gnCodFilial & " and Data = #" & Data_Atual & "# order by Ordem")
    If Caixa.EOF Then
      Caixa.AddNew
      Set CaixaAnterior = db.OpenRecordset("Select * from Caixa where Filial = " & gnCodFilial & " order by Data, Ordem")
      CaixaAnterior.MoveLast
      Caixa!Filial = gnCodFilial
      Caixa!Data = Data_Atual
      Caixa!Caixa = 1
      Caixa!Ordem = 1
      Caixa!Funcionário = Combo_Vendedor.Text
      Caixa!Hora = Format(Time, "hh:mm:ss")
      Caixa("Saldo Anterior") = 0
      Caixa!Dinheiro = CaixaAnterior("Total Dinheiro")
      Caixa!Cheques = CaixaAnterior("Total Cheques")
      Caixa("Cheques Pré") = CaixaAnterior("Total Cheques Pré")
      Caixa!Cartões = CaixaAnterior("Total Cartões")
      Caixa!Vales = CaixaAnterior("Total Vales")
      Caixa!Parcelamento = CaixaAnterior("Total Parcelamento")
      Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro")
      Caixa("Total Cheques") = CaixaAnterior("Total Cheques")
      Caixa("Total Cheques Pré") = CaixaAnterior("Total Cheques Pré")
      Caixa("Total Cartões") = CaixaAnterior("Total Cartões")
      Caixa("Total Vales") = CaixaAnterior("Total Vales")
      Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
      Caixa!Final = CaixaAnterior("Final")
      Caixa!Descrição = "Início do dia"
      Caixa.Update
    End If
    Set CaixaAnterior = db.OpenRecordset("Select * from Caixa where Filial = " & gnCodFilial & " and Data = #" & Data_Atual & "# order by Ordem")
    CaixaAnterior.MoveLast
    Select Case TipoRecebimpgto("id")
      Case 1
        Caixa.AddNew
        Caixa!Filial = gnCodFilial
        Caixa!Data = Data_Atual
        Caixa!Caixa = 1
        Caixa!Ordem = CaixaAnterior("Ordem") + 1
        Caixa!Funcionário = Combo_Vendedor.Text
        Caixa!Hora = Format(Time, "hh:mm:ss")
        Caixa("Saldo Anterior") = CaixaAnterior("Final")
        Caixa!Dinheiro = Cfisc_Pgto("Valor_Pagto")
        Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Cheques = 0
        Caixa("Cheques Pré") = 0
        Caixa!Cartões = 0
        Caixa!Vales = 0
        Caixa!Parcelamento = 0
        Caixa("Total Cheques") = CaixaAnterior("Total Cheques")
        Caixa("Total Cheques Pré") = CaixaAnterior("Total Cheques Pré")
        Caixa("Total Cartões") = CaixaAnterior("Total Cartões")
        Caixa("Total Vales") = CaixaAnterior("Total Vales")
        Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
        Caixa!Final = CaixaAnterior("Final") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Descrição = "Saída nr. " & Número.Text
        Caixa.Update
        rsSaidas.Edit
        rsSaidas("Recebe - Dinheiro") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas.Update
      Case 2
        Set Contas_Receber = db.OpenRecordset("Select * from [Contas a Receber] where Sequência = " & Número.Text & "")
        Caixa.AddNew
        Caixa!Filial = gnCodFilial
        Caixa!Data = Data_Atual
        Caixa!Caixa = 1
        Caixa!Ordem = CaixaAnterior("Ordem") + 1
        Caixa!Funcionário = Combo_Vendedor.Text
        Caixa!Hora = Format(Time, "hh:mm:ss")
        Caixa("Saldo Anterior") = CaixaAnterior("Final")
        Caixa!Dinheiro = 0
        Caixa!Cheques = Cfisc_Pgto("Valor_Pagto")
        Caixa("Total Cheques") = CaixaAnterior("Total Cheques") + Cfisc_Pgto("Valor_Pagto")
        Caixa("Cheques Pré") = 0
        Caixa!Cartões = 0
        Caixa!Vales = 0
        Caixa!Parcelamento = 0
        Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro")
        Caixa("Total Cheques Pré") = CaixaAnterior("Total Cheques Pré")
        Caixa("Total Cartões") = CaixaAnterior("Total Cartões")
        Caixa("Total Vales") = CaixaAnterior("Total Vales")
        Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
        Caixa!Final = CaixaAnterior("Final") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Descrição = "Saída nr. " & Número.Text
        Caixa.Update
        rsSaidas.Edit
        rsSaidas("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas.Update
        Contas_Receber.AddNew
        Contas_Receber("Filial") = gnCodFilial
        Contas_Receber("Cliente") = Combo_Cliente.Text
        Contas_Receber!Sequência = Número.Text
        Contas_Receber!Tipo = "C"
        Contas_Receber("Vencimento") = Data_Atual
        Contas_Receber!Valor = Cfisc_Pgto("Valor_Pagto")
        Contas_Receber("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        Contas_Receber("Data Recebimento") = Data_Atual
        Contas_Receber("Vendedor") = Combo_Vendedor.Text
        Contas_Receber!Processado = True
        Contas_Receber.Update
      Case 3
        Set Contas_Receber = db.OpenRecordset("Select * from [Contas a Receber] where Sequência = " & Número.Text & "")
        Caixa.AddNew
        Caixa!Filial = gnCodFilial
        Caixa!Data = Data_Atual
        Caixa!Caixa = 1
        Caixa!Ordem = CaixaAnterior("Ordem") + 1
        Caixa!Funcionário = Combo_Vendedor.Text
        Caixa!Hora = Format(Time, "hh:mm:ss")
        Caixa("Saldo Anterior") = CaixaAnterior("Final")
        Caixa!Dinheiro = 0
        Caixa!Cheques = 0
        Caixa("Cheques Pré") = Cfisc_Pgto("Valor_Pagto")
        Caixa("Total Cheques Pré") = CaixaAnterior("Total Cheques Pré") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Cartões = 0
        Caixa!Vales = 0
        Caixa!Parcelamento = 0
        Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro")
        Caixa("Total Cheques") = CaixaAnterior("Total Cheques")
        Caixa("Total Cartões") = CaixaAnterior("Total Cartões")
        Caixa("Total Vales") = CaixaAnterior("Total Vales")
        Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
        Caixa!Final = CaixaAnterior("Final") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Descrição = "Saída nr. " & Número.Text
        Caixa.Update
        rsSaidas.Edit
        rsSaidas("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas.Update
        Contas_Receber.AddNew
        Contas_Receber("Filial") = gnCodFilial
        Contas_Receber("Cliente") = Combo_Cliente.Text
        Contas_Receber!Sequência = Número.Text
        Contas_Receber!Tipo = "C"
        Contas_Receber("Vencimento") = Data_Atual
        Contas_Receber!Valor = Cfisc_Pgto("Valor_Pagto")
        Contas_Receber("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        Contas_Receber("Data Recebimento") = Data_Atual
        Contas_Receber("Vendedor") = Combo_Vendedor.Text
        Contas_Receber!Processado = True
        Contas_Receber.Update
      Case 5, 8, 9
        Caixa.AddNew
        Caixa!Filial = gnCodFilial
        Caixa!Data = Data_Atual
        Caixa!Caixa = 1
        Caixa!Ordem = CaixaAnterior("Ordem") + 1
        Caixa!Funcionário = Combo_Vendedor.Text
        Caixa!Hora = Format(Time, "hh:mm:ss")
        Caixa("Saldo Anterior") = CaixaAnterior("Final")
        Caixa!Dinheiro = 0
        Caixa!Cheques = 0
        Caixa("Cheques Pré") = 0
        Caixa!Cartões = Cfisc_Pgto("Valor_Pagto")
        Caixa("Total Cartões") = CaixaAnterior("Total Cartões") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Vales = 0
        Caixa!Parcelamento = 0
        Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro")
        Caixa("Total Cheques") = CaixaAnterior("Total Cheques")
        Caixa("Total Cheques Pré") = CaixaAnterior("Total Cheques Pré")
        Caixa("Total Vales") = CaixaAnterior("Total Vales")
        Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
        Caixa!Final = CaixaAnterior("Final") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Descrição = "Saída nr. " & Número.Text
        Caixa.Update
        rsSaidas.Edit
        rsSaidas("Recebe - Cartão") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas.Update
    End Select
    Set Resumo_Diário = db.OpenRecordset("Select * from [Resumo Diário] where Filial = " & gnCodFilial & " and Data = #" & Data_Atual & "#")
    If Resumo_Diário.EOF Then
      Resumo_Diário.AddNew
      Resumo_Diário!Filial = gnCodFilial
      Resumo_Diário!Data = Data_Atual
      Resumo_Diário("Valor Vendas") = L_Tot_Pagar.Text
      Resumo_Diário.Update
    Else
      Resumo_Diário.Edit
      Resumo_Diário!Filial = gnCodFilial
      Resumo_Diário!Data = Data_Atual
      Resumo_Diário("Valor Vendas") = Resumo_Diário("Valor Vendas") + L_Tot_Pagar.Text
      Resumo_Diário.Update
    End If
    Set Resumo_Diário_Financeiro = db.OpenRecordset("Select * from [Resumo Diário] where Filial = " & gnCodFilial & " and Data = #" & Data_Atual & "#")
    If Resumo_Diário_Financeiro.EOF Then
      Resumo_Diário_Financeiro.AddNew
      Resumo_Diário_Financeiro!Filial = gnCodFilial
      Resumo_Diário_Financeiro!Data = Data_Atual
      Resumo_Diário_Financeiro("Valor Vendas") = L_Tot_Pagar.Text
      Resumo_Diário_Financeiro.Update
    Else
      Resumo_Diário_Financeiro.Edit
      Resumo_Diário_Financeiro!Filial = gnCodFilial
      Resumo_Diário_Financeiro!Data = Data_Atual
      Resumo_Diário_Financeiro("Valor Vendas") = Resumo_Diário("Valor Vendas") + L_Tot_Pagar.Text
      Resumo_Diário_Financeiro.Update
    End If
    'Atualiza estoque PAF
    Set produtos = db.OpenRecordset("Select * from [Saídas - Produtos] where Filial = " & gnCodFilial & " and Sequência = " & Número.Text & "")
    Do Until produtos.EOF
      Set cad_prod = db.OpenRecordset("Select * from Produtos where Código = '" & produtos("Código sem Grade") & "'")
      If cad_prod("Tipo") = "N" Then
        Set Estoque_Final = db.OpenRecordset("Select * From [Estoque Final] where Filial = " & gnCodFilial & " and Produto = '" & produtos("Código sem Grade") & "'")
        If Estoque_Final.EOF Then
          MsgBox "O produto " & cad_prod("DESCRICAO") & " esta com estoque não inicializado. Não foi possivel dar baixa no estoque"
        Else
          Estoque_Final.Edit
          Estoque_Final("Estoque Atual") = Estoque_Final("Estoque Atual") - produtos("Qtde")
          Estoque_Final("Última Data") = Data_Atual
          Estoque_Final.Update
        End If
        Set Estoque_Anterior = db.OpenRecordset("Select * From Estoque where Filial = " & gnCodFilial & " and Produto = '" & produtos("Código sem Grade") & "' order by data")
        Estoque_Anterior.MoveLast
        Set Estoque = db.OpenRecordset("Select * From Estoque where Filial = " & gnCodFilial & " and Produto = '" & produtos("Código sem Grade") & "' And Data = #" & Data_Atual & "#")
        If Estoque.EOF Then
          Estoque.AddNew
          Estoque!Filial = gnCodFilial
          Estoque!Data = Data_Atual
          Estoque!Produto = produtos("Código sem Grade")
          Estoque!Tamanho = 0
          Estoque!Cor = 0
          Estoque!Edição = 0
          Estoque!Classe = cad_prod("Classe")
          Estoque("Sub Classe") = cad_prod("Sub Classe")
          Estoque("Estoque Anterior") = Estoque_Anterior("Estoque Final")
          Estoque!Vendas = produtos("Qtde")
          Estoque("Valor Vendas") = produtos("Preço Final")
          Estoque.Update
        Else
          Estoque.Edit
          Estoque("Vendas") = Estoque("Vendas") + produtos("Qtde")
          Estoque("Valor Vendas") = Estoque("Valor Vendas") + produtos("Preço Final")
          Estoque("Estoque Final") = Estoque("Estoque Final") - produtos("Qtde")
          Estoque.Update
        End If
        'atualiza estoque de produto com grade PAF
      ElseIf cad_prod("Tipo") = "G" Then
          Tamanho = 0
          Cor = 0
          Edicao = 0
          Tipo = 1
          Erro = 0
        modFuncoes.Acha_Produto produtos("Código"), produtos("Código"), Tamanho, Cor, Edicao, Tipo, Erro
        Set Estoque_Final = db.OpenRecordset("Select * From [Estoque Final] where Filial = " & gnCodFilial & " and Produto = '" & produtos("Código sem Grade") & "' AND Cor = " & Cor & " And Tamanho = " & Tamanho & "")
        If Estoque_Final.EOF Then
          MsgBox "O produto " & cad_prod("DESCRICAO") & " esta com estoque não inicializado. Não foi possivel dar baixa no estoque"
        Else
          Estoque_Final.Edit
          Estoque_Final("Estoque Atual") = Estoque_Final("Estoque Atual") - produtos("Qtde")
          Estoque_Final("Última Data") = Data_Atual
          Estoque_Final.Update
        End If
        Set Estoque_Anterior = db.OpenRecordset("Select * From Estoque where Filial = " & gnCodFilial & " and Produto = '" & produtos("Código sem Grade") & "' AND Cor = " & Cor & " And Tamanho = " & Tamanho & " order by data")
        Estoque_Anterior.MoveLast
        Set Estoque = db.OpenRecordset("Select * From Estoque where Filial = " & gnCodFilial & " and Produto = '" & produtos("Código sem Grade") & "' AND Cor = " & Cor & " And Tamanho = " & Tamanho & "")
        If Estoque.EOF Then
          Estoque.AddNew
          Estoque!Filial = gnCodFilial
          Estoque!Data = Data_Atual
          Estoque!Produto = produtos("Código sem Grade")
          Estoque!Tamanho = Left(Right(produtos("Código"), 3), 3)
          Estoque!Cor = Right(produtos("Código"), 3)
          Estoque!Edição = 0
          Estoque!Classe = cad_prod("Classe")
          Estoque("Sub Classe") = cad_prod("Sub Classe")
          Estoque("Estoque Anterior") = Estoque_Anterior("Estoque Final")
          Estoque!Vendas = produtos("Qtde")
          Estoque("Valor Vendas") = produtos("Preço Final")
          Estoque.Update
        Else
          Estoque.Edit
          Estoque("Vendas") = Estoque("Vendas") + produtos("Qtde")
          Estoque("Valor Vendas") = Estoque("Valor Vendas") + produtos("Preço Final")
          Estoque("Estoque Final") = Estoque("Estoque Final") - produtos("Qtde")
          Estoque.Update
        End If
      End If
      produtos.MoveNext
    Loop
    End If
    'marca venda como efetivada
    rsSaidas.Edit
    rsSaidas("Efetivada") = True
    rsSaidas("Recebimento") = True
    rsSaidas("Cupom Fiscal Impresso") = True
    rsSaidas.Update
    Efetivada.Visible = True
    
End Function
