VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmInformaConta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"InformaContagem.frx":0000
   ClientHeight    =   8265
   ClientLeft      =   390
   ClientTop       =   495
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1030
   Icon            =   "InformaContagem.frx":00E8
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8265
   ScaleWidth      =   13710
   Begin VB.Frame frm_acertaEstoque 
      Caption         =   "Acertar Estoque "
      Height          =   1185
      Left            =   60
      TabIndex        =   22
      Top             =   7020
      Width           =   13575
      Begin VB.CommandButton B_Acerta 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Acertar"
         Height          =   465
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Width           =   13275
      End
      Begin VB.OptionButton O_Consertar 
         Caption         =   "&Somente produtos com a coluna ""Consertar"" marcada"
         Height          =   405
         Left            =   1845
         TabIndex        =   24
         Top             =   195
         Value           =   -1  'True
         Width           =   4305
      End
      Begin VB.OptionButton O_Todos 
         Caption         =   "S&omente produtos com a coluna ""Diferença"" diferente de zero"
         Height          =   375
         Left            =   7020
         TabIndex        =   23
         Top             =   210
         Width           =   4815
      End
   End
   Begin VB.Frame frm_ajustaEstoque 
      Caption         =   "Ajustar Estoque"
      Height          =   5235
      Left            =   60
      TabIndex        =   16
      Top             =   1710
      Width           =   13575
      Begin VB.TextBox txt_quantidade 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F5DA33&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3360
         TabIndex        =   38
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txt_NomeProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5DA33&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   32
         Top             =   240
         Width           =   3045
      End
      Begin VB.OptionButton opt_ajusteComLeitor 
         Appearance      =   0  'Flat
         Caption         =   "Modo com Leitor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   31
         Top             =   690
         Width           =   1815
      End
      Begin VB.OptionButton opt_ajustePadrão 
         Appearance      =   0  'Flat
         Caption         =   "Modo Padrão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   315
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txt_CodProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5DA33&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   28
         Top             =   615
         Width           =   3045
      End
      Begin VB.CommandButton B_Monta 
         BackColor       =   &H00F5DA33&
         Caption         =   "&Listar os produtos para ajuste de estoque"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1110
         Width           =   6285
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1455
         Left            =   11940
         TabIndex        =   18
         Top             =   120
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   2566
         _StockProps     =   14
         Caption         =   "Ordeção produto"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox O_Classe 
            Appearance      =   0  'Flat
            Caption         =   "Classe"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   29
            Top             =   960
            Width           =   795
         End
         Begin VB.OptionButton O_Código 
            Appearance      =   0  'Flat
            Caption         =   "Por código"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   20
            Top             =   300
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton O_Nome 
            Appearance      =   0  'Flat
            Caption         =   "Por nome"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
      End
      Begin SSDataWidgets_B.SSDBGrid Grade1 
         Bindings        =   "InformaContagem.frx":4EA42
         Height          =   3525
         Left            =   135
         TabIndex        =   21
         Top             =   1635
         Width           =   13305
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
         AllowDelete     =   -1  'True
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         ForeColorEven   =   0
         BackColorOdd    =   16112179
         RowHeight       =   450
         ExtraHeight     =   26
         Columns(0).Width=   3200
         UseDefaults     =   0   'False
         _ExtentX        =   23469
         _ExtentY        =   6218
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
         Bindings        =   "InformaContagem.frx":4EA56
         DataSource      =   "Data1"
         Height          =   345
         Left            =   7050
         TabIndex        =   33
         ToolTipText     =   "Use 0 para todas"
         Top             =   240
         Width           =   885
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
         Columns(0).Width=   6244
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1455
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1561
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   16112179
      End
      Begin VB.Line Line6 
         X1              =   10740
         X2              =   10650
         Y1              =   1500
         Y2              =   1350
      End
      Begin VB.Line Line5 
         X1              =   10740
         X2              =   10830
         Y1              =   1500
         Y2              =   1350
      End
      Begin VB.Line Line4 
         X1              =   9870
         X2              =   9780
         Y1              =   1500
         Y2              =   1350
      End
      Begin VB.Line Line3 
         X1              =   9870
         X2              =   9960
         Y1              =   1500
         Y2              =   1350
      End
      Begin VB.Line Line1 
         X1              =   10740
         X2              =   10740
         Y1              =   810
         Y2              =   1470
      End
      Begin VB.Line Line2 
         X1              =   9870
         X2              =   9870
         Y1              =   810
         Y2              =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Contagem do inventário"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   10800
         TabIndex        =   40
         Top             =   780
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Atualmente em estoque no QuickStore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7800
         TabIndex        =   39
         Top             =   780
         Width           =   2040
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
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
         Left            =   2385
         TabIndex        =   37
         Top             =   330
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Nome_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   8010
         TabIndex        =   36
         Top             =   240
         Width           =   3885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   6540
         TabIndex        =   35
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   2085
         TabIndex        =   34
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Produto"
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
         Left            =   2055
         TabIndex        =   27
         Top             =   690
         Width           =   1275
      End
   End
   Begin VB.Frame frm_contagemEstoque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   13575
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Relatório de estoque atual"
         Height          =   795
         Left            =   11535
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Chamar o Rel. de Contagem de Estoque"
         Top             =   240
         Width           =   1890
      End
      Begin VB.CommandButton bt_gerarContagemEstoque 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Inicializar processo para ajuste de estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1110
         Width           =   13275
      End
      Begin VB.Frame Frame1 
         Caption         =   "Saída"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   285
         TabIndex        =   9
         Top             =   2460
         Width           =   3465
         Begin VB.OptionButton B_Vídeo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Vídeo"
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
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton B_Impressora 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Impressora"
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
            Left            =   1260
            TabIndex        =   10
            Top             =   210
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ordem"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3825
         TabIndex        =   6
         Top             =   2460
         Width           =   3465
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Código"
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
            Left            =   300
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Nome"
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
            Left            =   1560
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Opções"
         Height          =   915
         Left            =   6120
         TabIndex        =   2
         Top             =   120
         Width           =   5325
         Begin VB.CheckBox O_Zero 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Não considerar produtos com estoque zero"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   210
            TabIndex        =   5
            Top             =   585
            Width           =   3645
         End
         Begin VB.CheckBox O_Inativos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Considerar produtos inativos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2670
            TabIndex        =   4
            Top             =   285
            Width           =   2550
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Separar por classe"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   3
            Top             =   270
            Width           =   2235
         End
      End
      Begin SSDataWidgets_B.SSDBCombo Combo 
         Bindings        =   "InformaContagem.frx":4EA6A
         DataSource      =   "Data2"
         Height          =   345
         Left            =   735
         TabIndex        =   12
         Top             =   240
         Width           =   885
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
         Columns(0).Width=   8467
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1561
         Columns(1).Caption=   "Filial"
         Columns(1).Name =   "Filial"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Filial"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         _ExtentX        =   1561
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
      End
      Begin VB.Label Nome_Combo 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1665
         TabIndex        =   14
         Top             =   240
         Width           =   4275
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Classe"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.CommandButton B_Iguala 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Igualar"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13500
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Fazer a Contagem Igual ao Estoque"
      Top             =   1740
      Visible         =   0   'False
      Width           =   2340
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
      Left            =   13440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   2310
      Visible         =   0   'False
      Width           =   1695
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
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2700
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmInformaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsParametros1 As Recordset
Dim rsProdutos1 As Recordset
Dim rsEstoque_Final1 As Recordset
Dim TB2_Contagem1 As Recordset
Dim rsClasses1 As Recordset
Dim rsSub_Classes1 As Recordset

Dim rsContagem As Recordset
Dim rsProdutos2 As Recordset
Dim rsEstoque2  As Recordset


Dim rsPreços As Recordset
Dim Rec_Contas As Recordset
Dim rsClasses As Recordset
Dim rsSub_Classes As Recordset
Dim TB2_Contagem As Recordset

Private Sub B_Acerta_Click()
  Dim Resposta       As Integer
  Dim Código         As String
  Dim Tamanho        As Integer
  Dim Cor            As Integer
  Dim Conta          As Long
  Dim Criar_Registro As Integer
  Dim Estoque_Final  As Single
  Dim Mes_Atual      As Integer
  Dim Ano_Atual      As Integer
  
  gbAcertaGrade = False
  
  Call StatusMsg("")
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Este processo não poderá ser desfeito, deseja prosseguir?"
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    DisplayMsg "Estoque não foi atualizado."
    Exit Sub
  End If
  
  On Error GoTo ErrTrans
  
  Screen.MousePointer = vbHourglass
  
  Call ws.BeginTrans
  
  Código = ""
  Tamanho = 0
  Cor = 0
  Conta = 0
  rsProdutos2.Index = "Código"
  rsContagem.Index = "Código"

Lp1:
  If gbAcertaGrade = True Then
    rsContagem.Seek ">", Código, Tamanho, Cor
  Else
    rsContagem.Seek ">", Código
  End If
  
  If rsContagem.NoMatch Then GoTo Fim_Lp
  Código = rsContagem("Código")
  
  If gbAcertaGrade = True Then
    Tamanho = rsContagem("Tamanho")
    Cor = rsContagem("Cor")
  End If
  
  'Verifica se a filial de origem é a mesma que está logado
  If rsContagem("Empresa") <> gnCodFilial Then GoTo Lp1
  
  If rsContagem("Diferença") = 0 Then GoTo Lp1
  
  If O_Consertar.Value = True Then
    If rsContagem("Consertar") = False Then GoTo Lp1
  End If
  
  rsProdutos2.Seek "=", rsContagem("Código")
  If rsProdutos2.NoMatch Then GoTo Lp1
  
  Conta = Conta + 1
  
  Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
  
  Rem Acha Último Estoque deste produto
  Criar_Registro = False
  Estoque_Final = 0
  rsEstoque2.Index = "Produto"
  rsEstoque2.Seek "=", rsContagem("Empresa"), Data_Atual, rsContagem("Código"), Tamanho, Cor, 0
  
  If Not rsEstoque2.NoMatch Then
    Estoque_Final = rsEstoque2("Estoque Final")
  End If
  
  If rsEstoque2.NoMatch Then
    
    rsEstoque2.Index = "Data"
    rsEstoque2.Seek "<", rsContagem("Empresa"), rsContagem("Código"), Tamanho, Cor, 0, Data_Atual
    If rsEstoque2.NoMatch Then Criar_Registro = True
    If Not rsEstoque2.NoMatch Then
      If rsEstoque2("Filial") = rsContagem("Empresa") And rsEstoque2("Produto") = rsContagem("Código") And rsEstoque2("Tamanho") = 0 And rsEstoque2("Cor") = 0 And rsEstoque2("Edição") = 0 Then
        Criar_Registro = True
        Estoque_Final = rsEstoque2("Estoque Final")
      End If
    End If
  
    rsEstoque2.AddNew
    rsEstoque2("Filial") = rsContagem("Empresa")
    rsEstoque2("Data") = Data_Atual
    rsEstoque2("Produto") = rsContagem("Código")
    rsEstoque2("Tamanho") = Tamanho
    rsEstoque2("Cor") = Cor
    rsEstoque2("Edição") = 0
    rsEstoque2("Classe") = rsProdutos("Classe")
    rsEstoque2("Sub Classe") = rsProdutos("Sub Classe")
    rsEstoque2("Estoque Anterior") = Estoque_Final
    rsEstoque2.Update
    
    rsEstoque2.Index = "Produto"
    rsEstoque2.Seek "=", rsContagem("Empresa"), Data_Atual, rsContagem("Código"), Tamanho, Cor, 0
  
  End If
  
  'Verifica se a real diferença está correta
  If rsContagem("Qtde Estoque") <> Estoque_Final Then
    With rsContagem
      .Edit
      .Fields("Qtde Estoque") = Estoque_Final
      .Fields("Diferença") = .Fields("Digitado") - Estoque_Final
      .Update
    End With
    If gbAcertaGrade = True Then
      rsContagem.Seek "=", Código, Tamanho, Cor
    Else
      rsContagem.Seek "=", Código
    End If
  End If
  
  Rem neste ponto esta com o registro de estoque
  Rem no buffer, agora soma com os valores da movimentação
  rsEstoque2.Edit
  If rsContagem("Diferença") < 0 Then
    rsEstoque2("Ajuste Saída") = rsEstoque2("Ajuste Saída") + Abs(rsContagem("Diferença"))
  End If
  
  If rsContagem("Diferença") > 0 Then
    rsEstoque2("Ajuste Entra") = rsEstoque2("Ajuste Entra") + Abs(rsContagem("Diferença"))
  End If
  
  Estoque_Final = rsEstoque2("Estoque Anterior") - rsEstoque2("Vendas") + rsEstoque2("Compras")
  Estoque_Final = Estoque_Final - rsEstoque2("Transf Saída") + rsEstoque2("Transf Entra")
  Estoque_Final = Estoque_Final - rsEstoque2("Ajuste Saída") + rsEstoque2("Ajuste Entra")
  Estoque_Final = Estoque_Final - rsEstoque2("Grátis Saída") + rsEstoque2("Grátis Entra")
  Estoque_Final = Estoque_Final - rsEstoque2("Empre Saída") + rsEstoque2("Empre Entra")
  Estoque_Final = Estoque_Final - rsEstoque2("Quebras") + rsEstoque2("Devolução")
  
  If rsProdutos2("Estoque") = False Then
    Estoque_Final = 0
  End If
  
  rsEstoque2("Estoque Final") = Estoque_Final
  rsEstoque2.Update
  
  If gbAcertaGrade Then
    Call Grava_Estoque_Final(rsContagem("Empresa"), rsProdutos2("Código"), Tamanho, Cor, 0, Estoque_Final, CDate(Data_Atual))
  Else
    Call Grava_Estoque_Final(rsContagem("Empresa"), rsProdutos2("Código"), 0, 0, 0, Estoque_Final, CDate(Data_Atual))
  End If
  
  rsContagem.Edit
  rsContagem("Diferença") = 0
  rsContagem("Qtde Estoque") = rsContagem("Digitado")
  rsContagem("Consertar") = False
  rsContagem.Update
  
  GoTo Lp1
  
Fim_Lp:

  '---[ Gera Log do usuário ]---'
      g_GravaLog Data_Atual, "Acerto de Estoque, DQ(" & Data_Atual & "), DW(" & Date & "),Funcionário: " & _
                            gnUserCode & " - " & gsUserName, "ACERTO ESTOQUE"
  '---[ Gera Log do usuário ]---'
  
  
  dbTemp.Execute "Delete * From Contagem"
  
  Call ws.CommitTrans
  Screen.MousePointer = vbDefault
  DisplayMsg "Fim de processo. Registros atualizados : " + str(Conta)
  Exit Sub
  
ErrTrans:
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Acertar Estoque."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly & vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  On Error Resume Next
  Call ws.Rollback
  Exit Sub
End Sub

Private Sub B_Iguala_Click()
' Dim Código As String
'
'  Call StatusMsg("Aguarde...")
'  DoEvents
'
'  TB2_Contagem.Index = "Código"
'  Código = 0
'Lp1:
'  TB2_Contagem.Seek ">", Código
'  If TB2_Contagem.NoMatch Then GoTo Fim_Loop
'
'  Código = TB2_Contagem("Código")
'
'  'Verifica se a filial de origem
'  If TB2_Contagem("Empresa") <> gnCodFilial Then GoTo Lp1
'
'  TB2_Contagem.Edit
'    TB2_Contagem("Digitado") = TB2_Contagem("Qtde Estoque")
'    TB2_Contagem("Diferença") = 0
'  TB2_Contagem.Update
'
'  GoTo Lp1
'
'Fim_Loop:
'  Call StatusMsg("")
'
'  B_Monta_Click


'Novo Código
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde...")
  dbTemp.Execute "UPDATE Contagem SET Digitado = [Qtde Estoque], Diferença = 0 WHERE Empresa = " & gnCodFilial, dbFailOnError
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  Call B_Monta_Click
  
End Sub

Private Sub B_Monta_Click()
 Dim sSql As String
 Dim i As Integer
 Dim Classe As Integer
 
  Call StatusMsg("")
  
 On Error GoTo Processa_Erro
 
  Classe = 0
  If Nome_Classe.Caption <> "" Then Classe = Val(Combo_Classe.Text)
  
  If Trim(txt_CodProduto.Text) <> "" And Trim(txt_NomeProduto.Text) <> "" Then
      MsgBox "Preencha o campo 'Código do Produto' ou 'Parte do Nome'. Não pode usar junto os dois na pesquisa.", vbInformation, "Atenção"
      Exit Sub
  End If
 
  sSql = "SELECT Código, Nome, Classe, [Qtde Estoque], Digitado, Diferença, Consertar FROM [Contagem]"
  
  'Verifica se a filial de origem
  sSql = sSql & " WHERE Empresa = " & gnCodFilial
  
  If Classe <> 0 Then
      sSql = sSql & " AND Classe = " + str(Classe)
  End If
  
  If Trim(txt_NomeProduto.Text) <> "" Then
      sSql = sSql & " AND Nome like '*" & Trim(txt_NomeProduto.Text) & "*' "
  End If
  
  If Trim(txt_CodProduto.Text) <> "" Then
      sSql = sSql & " AND Código = '" & Trim(txt_CodProduto.Text) & "' "
  End If
  
  If O_Classe.Value = 1 Then
    If O_Código.Value Then
      sSql = sSql + " ORDER BY Classe, [Código Ordenação]"
    Else
      sSql = sSql + " ORDER BY Classe, Nome"
    End If
  Else
    If O_Código.Value Then
      sSql = sSql + " ORDER BY [Código Ordenação]"
    Else
      sSql = sSql + " ORDER BY Nome"
    End If
  End If
  
  Set Rec_Contas = dbTemp.OpenRecordset(sSql, dbOpenDynaset)

  Grade1.DataMode = 1

  Set Data2.Recordset = Rec_Contas


  Grade1.Visible = False
  
  Grade1.DataMode = 0
  
  Grade1.ReBind
 
    Grade1.Columns(0).Width = 2300
    Grade1.Columns(0).Locked = True
    
    Grade1.Columns(1).Width = 6100
    Grade1.Columns(1).Locked = True
    
    Grade1.Columns(2).NumberFormat = "#####0"
    Grade1.Columns(2).Width = 750
    Grade1.Columns(2).Locked = True
    
    Grade1.Columns(3).Locked = True
    Grade1.Columns(3).Width = 750
    
    Grade1.Columns(5).Locked = True
    
    Grade1.Columns(6).Style = 2
       
    
  Grade1.Visible = True
  Call StatusMsg("")
  
  If opt_ajusteComLeitor.Value = True Then
      AtualizaEstoqueDoProdutoNaGrade
  End If

 
  Exit Sub
Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Altera Preços - Montar")
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

'
Private Sub AtualizaEstoqueDoProdutoNaGrade()
On Error GoTo Erro

  Dim Aux As Variant
  Dim Erro As Integer
  Dim Est, Dig, Dif As Double
  
  Dim nQtdeCasaDec As Integer
  Dim bFrac As Boolean
  
  Erro = False
  Aux = Grade1.Columns(4).Text
  If IsNull(Aux) Then Erro = True
  If Erro = False Then If Not IsNumeric(Aux) Then Erro = True
  If Erro = False Then If Abs(CDbl(Aux)) > 999999999 Then Erro = True
  If Erro = True Then
      DisplayMsg "Digite um valor."
      Exit Sub
  End If
    
    
  If IsNull(txt_quantidade.Text) Then Erro = True
  If Erro = False Then If Not IsNumeric(txt_quantidade.Text) Then Erro = True
  If Erro = False Then If Abs(CDbl(txt_quantidade.Text)) > 999999999 Then Erro = True
  If Erro = True Then
      DisplayMsg "Digite uma quantidade válida."
      txt_quantidade.SetFocus
      Exit Sub
  End If
    
  If gbIsFrac(Grade1.Columns(0).Text, nQtdeCasaDec) Then
    bFrac = True
    Grade1.Columns(4).Text = Round(CDbl(Aux) + CDbl(txt_quantidade.Text), nQtdeCasaDec)   'Format(CDbl(Aux), "#0.000")
  Else
    Grade1.Columns(4).Text = Format(CDbl(Aux) + CDbl(txt_quantidade.Text), "#0")
  End If
    
  Est = Grade1.Columns(3).Text
  Dig = CDbl(Aux) + CDbl(txt_quantidade.Text)
  Dif = Abs(CDbl(Est) - Dig)
    
  If CDbl(Est) > CDbl(Dig) Then Dif = Dif * -1
  
  If bFrac Then
    Grade1.Columns(5).Text = Round(Dif, nQtdeCasaDec)  'Format(Dif, "#0.000")
  Else
    Grade1.Columns(5).Text = Format(Dif, "#0")
  End If
  
  Grade1.Columns(6).Text = vbChecked
  Grade1.Update
  
  txt_quantidade.Text = "1"
  txt_CodProduto.Text = ""
      
  Exit Sub
Erro:
  MsgBox "Erro na atualização do número estoque na grade " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
      
End Sub
'



Private Sub bt_gerarContagemEstoque_Click()
 Dim Termina As Integer
 Dim Val2 As Integer
 Dim Erro As Integer
 Dim Str1 As String
 Dim Str2 As String
 Dim Str3 As String
 Dim Str_Data1 As String
 Dim Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 Dim Produto As String
 Dim Tamanho As Integer
 Dim Cor As Integer
 Dim sSql As String
 Dim Estoque As Double
 Dim Aux_Data As Variant
 Dim Aux_Classe As Integer
 Dim Aux_Sub As Integer
 Dim Aux_Produto As Double
 Dim Nome_Cla As String
 Dim Nome_Sub As String
  
 Call StatusMsg("")

 Rem apaga pesquisa anterior desta filial do arquivo temporario
 Call StatusMsg("Aguarde, preparando arquivo temporário ...")

 sSql = "Delete * From Contagem where Empresa = " & gnCodFilial
 dbTemp.Execute sSql

 Call StatusMsg("")
 
 Rem Le estoque e joga no temporário
 rsProdutos1.Index = "Código"
 rsEstoque_Final1.Index = "Produto"
 Termina = False
 Produto = ""
 Call StatusMsg("Aguarde, contando estoque.")

 rsClasses1.Index = "Código"
 rsSub_Classes1.Index = "Código"

LP1S:
  rsProdutos1.Seek ">", Produto
  If rsProdutos1.NoMatch Then GoTo Imprime
  Produto = rsProdutos1("Código")
  
  '14/01/2005 - Daniel
  'Em algumas bases de dados o campo Produtos.Código está
  'aparecendo com caracteres incorretos tais como ...
  '
  'Case: São Francisco Móveis e Eletro. de Olinda - PE
  If Len(Produto) > 20 Then Produto = "0"
  '-------------------------------------------------------
  
  If Produto = "0" Then GoTo LP1S
  
  If Nome_Combo.Caption <> "" Then
    If rsProdutos1("Classe") <> Val(Combo.Text) Then GoTo LP1S
  End If
  
  If rsProdutos1("Desativado") = True And O_Inativos.Value = 0 Then GoTo LP1S

  If rsProdutos1("Tipo") <> "N" Then GoTo LP1S

  'If rsProdutos("Fracionado") = True Then GoTo LP1S
     
  Estoque = 0
  '''rsEstoque_Final1.Seek "=", Val(Combo.Text), Produto, 0, 0, 0
  rsEstoque_Final1.Seek "=", gnCodFilial, Produto, 0, 0, 0
  If Not rsEstoque_Final1.NoMatch Then Estoque = rsEstoque_Final1("Estoque Atual")
  
  If O_Zero.Value = 1 Then
    If Estoque = 0 Then GoTo LP1S
  End If

  Call StatusMsg("Aguarde, gravando arquivo temporário, produto " + (Produto))

  rsClasses1.Seek "=", rsProdutos("Classe")
  If rsClasses1.NoMatch Then
     Nome_Cla = "Classe não cadastrada"
  Else
     Nome_Cla = rsClasses1("Nome")
  End If
  
  rsSub_Classes1.Seek "=", rsProdutos1("Sub Classe")
  If rsSub_Classes1.NoMatch Then
    Nome_Sub = "Subclasse não cadastrada"
  Else
    Nome_Sub = rsSub_Classes1("Nome")
  End If
  
  
  TB2_Contagem1.AddNew
     TB2_Contagem1("Código") = Produto
     TB2_Contagem1("Código Ordenação") = rsProdutos1("Código Ordenação")
     TB2_Contagem1("Nome") = rsProdutos1("Nome")
     TB2_Contagem1("Classe") = rsProdutos1("Classe")
     TB2_Contagem1("Nome Classe") = Nome_Cla
     TB2_Contagem1("Sub Classe") = rsProdutos1("Sub Classe")
     TB2_Contagem1("Nome Sub") = Nome_Sub
     TB2_Contagem1("Unidade") = rsProdutos1("Unidade Venda")
     TB2_Contagem1("Fracionado") = rsProdutos1("Fracionado")
     TB2_Contagem1("Qtde Estoque") = Estoque
     TB2_Contagem1("Empresa") = gnCodFilial
  TB2_Contagem1.Update
  
  GoTo LP1S
  
Imprime:

 MsgBox "Processo iniciado com sucesso", vbInformation, "Estoque"
 Exit Sub
 
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

Private Sub Command1_Click()
  '14/05/2005 - Daniel
  'Otimizando a chamada da tela
  frmRelContagem.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
      'Call objHelp.Show(strfile, "QuickStore10Help")
      Call objHelp.Show(strfile, "QuickStore10Help", 10001) 'O número 10001 está relacionado com o arquivo estoque.htm
      Set objHelp = Nothing
  End If
End Sub

Private Sub Form_Load()
 
  Call CenterForm(Me)
  
  ' ======================================================================
  ' Tratando o frm_contagemEstoque
  Set rsParametros1 = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsProdutos1 = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsEstoque_Final1 = db.OpenRecordset("Estoque Final", , dbReadOnly)
  Set rsClasses1 = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSub_Classes1 = db.OpenRecordset("Sub Classes", , dbReadOnly)
  
  Set TB2_Contagem1 = dbTemp.OpenRecordset("Contagem")
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  ' ======================================================================
  
  
  ' ======================================================================
  ' Tratando o frm_acertaEstoque
    Dim sSql As String
    Dim sCaption As String
  
    If gbAcertaGrade Then
      sSql = "Contagem Grade"
      sCaption = "(Produtos com Grade)"
    Else
      sSql = "Contagem"
      sCaption = ""
    End If
    '''Me.Caption = "Acerta Estoque " & sCaption
    Set rsContagem = dbTemp.OpenRecordset(sSql)
    Set rsProdutos2 = db.OpenRecordset("Produtos", , dbReadOnly)
    Set rsEstoque2 = db.OpenRecordset("Estoque")
  ' ======================================================================

  
  
  
  
  
 Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
 Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
 Set rsSub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)
 
 Data1.DatabaseName = gsQuickDBFileName
 
 Set TB2_Contagem = dbTemp.OpenRecordset("Contagem")
 
End Sub

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(1).Text
  Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
  Dim Aux As Variant
  
  Nome_Combo.Caption = ""
  Aux = Combo.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsClasses.Index = "Código"
  rsClasses.Seek "=", Val(Aux)
  If rsClasses.NoMatch Then Exit Sub
  
  Nome_Combo.Caption = rsClasses("Nome")

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  ' ========================================================
  ' Objetos do frm_contagemEstoque
  rsParametros1.Close
  rsProdutos1.Close
  rsEstoque_Final1.Close
  rsClasses1.Close
  rsSub_Classes1.Close
  Set rsParametros1 = Nothing
  Set rsProdutos1 = Nothing
  Set rsEstoque_Final1 = Nothing
  Set rsClasses1 = Nothing
  Set rsSub_Classes1 = Nothing
  ' ========================================================
  
  ' ========================================================
  ' Objetos do frm_acertaEstoque
  rsContagem.Close
  rsProdutos2.Close
  rsEstoque2.Close
  Set rsContagem = Nothing
  Set rsProdutos2 = Nothing
  Set rsEstoque2 = Nothing
  ' ========================================================

End Sub

Private Sub Grade1_AfterDelete(RtnDispErrMsg As Integer)
  Grade1.Scroll 0, -32767
  Grade1.Scroll 0, 32767
End Sub

Private Sub Grade1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  Dim Erro As Integer
  Dim Est, Dig, Dif As Double
  
  Dim nQtdeCasaDec As Integer
  Dim bFrac As Boolean
  
  If ColIndex = 4 Then
    Erro = False
    Aux = Grade1.Columns(4).Text
    If IsNull(Aux) Then Erro = True
    If Erro = False Then If Not IsNumeric(Aux) Then Erro = True
    If Erro = False Then If Abs(CDbl(Aux)) > 999999999 Then Erro = True
  '  If Erro = False Then If CDbl(Aux) < 0 Then Erro = True
    If Erro = True Then
      DisplayMsg "Digite um valor."
      Cancel = True
      Exit Sub
    End If
    
    If gbIsFrac(Grade1.Columns(0).Text, nQtdeCasaDec) Then
      bFrac = True
      Grade1.Columns(4).Text = Round(CDbl(Aux), nQtdeCasaDec)  'Format(CDbl(Aux), "#0.000")
    Else
      Grade1.Columns(4).Text = Format(CDbl(Aux), "#0")
    End If
    
    Est = Grade1.Columns(3).Text
    Dig = Aux
    Dif = Abs(CDbl(Est) - Dig)
    
    If CDbl(Est) > CDbl(Dig) Then Dif = Dif * -1
   ' If Est < 0 And Dig < 0 And CDbl(Dig) < CDbl(Est) Then Dif = Dif * -1
    
    If bFrac Then
      Grade1.Columns(5).Text = Round(Dif, nQtdeCasaDec)  'Format(Dif, "#0.000")
    Else
      Grade1.Columns(5).Text = Format(Dif, "#0")
    End If
    
  End If
      
End Sub

Private Sub Grade1_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  If Not bGridBeforeDelete() Then
    Cancel = True
  End If
End Sub

Private Sub Grade1_BeforeUpdate(Cancel As Integer)
' Exit Sub
' On Error GoTo Deu_Erro
' Grade1.Columns(4).Text = Format(Date, "dd/mm/yyyy")
'Deu_Erro:
' On Error GoTo 0
End Sub

Private Sub Grade1_LostFocus()
  With Grade1
    If .RowChanged Then
      .Update
    End If
  End With
End Sub

Private Sub opt_ajusteComLeitor_Click()
    If opt_ajusteComLeitor.Value = True Then
        Label3.Visible = True
        txt_quantidade.Visible = True
        txt_quantidade.Text = "1"
        txt_CodProduto.Text = ""
'        Label7.Visible = True
'        Label8.Visible = True
'        Line1.Visible = True
'        Line2.Visible = True
'        Line3.Visible = True
'        Line4.Visible = True
'        Line5.Visible = True
'        Line6.Visible = True
        
        Label2.Visible = False
        txt_NomeProduto.Text = ""
        txt_NomeProduto.Visible = False
        Label4.Visible = False
        Combo_Classe.Text = ""
        Combo_Classe.Visible = False
        Nome_Classe.Caption = ""
        Nome_Classe.Visible = False
        SSFrame1.Visible = False
        B_Monta.Visible = False
        Grade1.Visible = False
        
        txt_CodProduto.SetFocus
    End If
End Sub

Private Sub opt_ajustePadrão_Click()
    If opt_ajustePadrão.Value = True Then
        Label2.Visible = True
        txt_NomeProduto.Visible = True
        Label4.Visible = True
        Combo_Classe.Visible = True
        Nome_Classe.Visible = True
        SSFrame1.Visible = True
        B_Monta.Visible = True
    
        txt_CodProduto.Text = ""
        Label3.Visible = False
'        Label7.Visible = false
'        Label8.Visible = false
        txt_quantidade.Text = ""
        txt_quantidade.Visible = False
        Grade1.Visible = False
'        Line1.Visible = False
'        Line2.Visible = False
'        Line3.Visible = False
'        Line4.Visible = False
'        Line5.Visible = False
'        Line6.Visible = False
    End If
End Sub

Private Sub txt_CodProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim(txt_CodProduto)) > 0 Then  'Tecla Enter
        B_Monta_Click
    End If
End Sub
