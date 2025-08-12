VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManContasPagar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Realizar baixa de Contas a Pagar"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   375
   ClientWidth     =   15090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1300
   Icon            =   "ManContasPagar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8220
   ScaleWidth      =   15090
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   75
      TabIndex        =   65
      Text            =   "Valor Total"
      Top             =   4950
      Width           =   885
   End
   Begin VB.Data datFilial 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
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
      Height          =   315
      Left            =   14880
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   90
      Visible         =   0   'False
      Width           =   1575
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "ManContasPagar.frx":4E95A
      DataSource      =   "datFilial"
      Height          =   330
      Left            =   540
      TabIndex        =   7
      Top             =   67
      Width           =   780
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
      Columns.Count   =   2
      Columns(0).Width=   6350
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1005
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1376
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   12648447
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
      Left            =   14880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Centros de Custo] WHERE Ativo ORDER BY Nome"
      Top             =   1230
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vencimento"
      Height          =   630
      Left            =   75
      TabIndex        =   54
      Top             =   420
      Width           =   5790
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
         Left            =   4800
         Picture         =   "ManContasPagar.frx":4E972
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   187
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
         Left            =   2190
         Picture         =   "ManContasPagar.frx":4F254
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   187
         Width           =   465
      End
      Begin MSMask.MaskEdBox Vcto_FInal 
         Height          =   315
         Left            =   3540
         TabIndex        =   1
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox Vcto_Inicial 
         Height          =   315
         Left            =   930
         TabIndex        =   0
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label7 
         Caption         =   "Inicial"
         Height          =   255
         Left            =   465
         TabIndex        =   56
         Top             =   255
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Final"
         Height          =   255
         Left            =   3135
         TabIndex        =   55
         Top             =   255
         Width           =   390
      End
   End
   Begin VB.Data Data5 
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
      Height          =   345
      Left            =   14880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Caixas"
      Top             =   840
      Visible         =   0   'False
      Width           =   2115
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
      Left            =   14880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   450
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   14880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1620
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   14880
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   2010
      Visible         =   0   'False
      Width           =   1815
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Bindings        =   "ManContasPagar.frx":4FB36
      Height          =   3255
      Left            =   75
      TabIndex        =   11
      Top             =   1560
      Width           =   14940
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
      AllowUpdate     =   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   12648384
      RowHeight       =   423
      ExtraHeight     =   53
      Columns(0).Width=   3200
      _ExtentX        =   26352
      _ExtentY        =   5741
      _StockProps     =   79
      Caption         =   "Contas"
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
   Begin VB.CommandButton B_Baixa 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Alterar / Baixar"
      Height          =   430
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4860
      Width           =   12120
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   9840
      TabIndex        =   34
      Top             =   420
      Width           =   5145
      Begin VB.OptionButton O_Vencimento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Por data de vencimento"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   450
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton O_Cliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Fornecedor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3030
         TabIndex        =   6
         Top             =   285
         Width           =   1170
      End
   End
   Begin VB.CommandButton B_Monta 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Pesquisar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   14940
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contas"
      Height          =   630
      Left            =   5910
      TabIndex        =   33
      Top             =   420
      Width           =   3855
      Begin VB.OptionButton O_Todas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton O_Recebidas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Já pagas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2670
         TabIndex        =   3
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton O_Receber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "A Pagar"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1500
         TabIndex        =   2
         Top             =   270
         Width           =   960
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "ManContasPagar.frx":4FB4A
      DataSource      =   "Data1"
      Height          =   330
      Left            =   7305
      TabIndex        =   8
      Top             =   67
      Width           =   1380
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
      Columns.Count   =   2
      Columns(0).Width=   8625
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1376
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2434
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Frame Frame_Modo 
      BackColor       =   &H00FFA324&
      Caption         =   "FORMA DE PAGAMENTO"
      Height          =   2310
      Left            =   9690
      TabIndex        =   47
      Top             =   5820
      Visible         =   0   'False
      Width           =   5340
      Begin SSDataWidgets_B.SSDBCombo Combo_Caixa 
         Bindings        =   "ManContasPagar.frx":4FB5E
         DataSource      =   "Data5"
         Height          =   330
         Left            =   2145
         TabIndex        =   28
         Top             =   937
         Width           =   750
         DataFieldList   =   "Descrição"
         ListAutoPosition=   0   'False
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
         Columns(0).Width=   7832
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1773
         Columns(1).Caption=   "Caixa"
         Columns(1).Name =   "Caixa"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Caixa"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         _ExtentX        =   1323
         _ExtentY        =   582
         _StockProps     =   93
         BackColor       =   12648447
         Enabled         =   0   'False
      End
      Begin VB.TextBox Num_Cheque 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   32
         Top             =   1837
         Width           =   1560
      End
      Begin MSMask.MaskEdBox Cheque_Bom 
         Height          =   330
         Left            =   3960
         TabIndex        =   31
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1837
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
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
      Begin VB.OptionButton O_Não_determinado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Não determinado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   225
         TabIndex        =   25
         Top             =   300
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton O_caixa_d 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Caixa dinheiro"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   225
         TabIndex        =   27
         Top             =   975
         Width           =   1380
      End
      Begin VB.OptionButton O_Conta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Conta Corrente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   225
         TabIndex        =   29
         Top             =   1455
         Width           =   1455
      End
      Begin VB.OptionButton O_Caixa_C 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Caixa cheque"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   225
         TabIndex        =   26
         Top             =   630
         Width           =   1575
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
         Bindings        =   "ManContasPagar.frx":4FB72
         DataSource      =   "Data4"
         Height          =   330
         Left            =   2145
         TabIndex        =   30
         Top             =   1417
         Width           =   750
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
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   6482
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Conta"
         Columns(1).Name =   "Conta"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Conta"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2011
         Columns(2).Caption=   "Código"
         Columns(2).Name =   "Código"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Código"
         Columns(2).DataType=   2
         Columns(2).FieldLen=   256
         _ExtentX        =   1323
         _ExtentY        =   582
         _StockProps     =   93
         BackColor       =   12648447
         Enabled         =   0   'False
      End
      Begin VB.Label Nome_Caixa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2925
         TabIndex        =   53
         Top             =   937
         Width           =   2310
      End
      Begin VB.Label Label_Caixa 
         BackColor       =   &H00FFA324&
         Caption         =   "Caixa"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1695
         TabIndex        =   52
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFA324&
         Caption         =   "Num. Cheque"
         Height          =   225
         Left            =   255
         TabIndex        =   51
         Top             =   1890
         Width           =   1080
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFA324&
         Caption         =   "Bom para"
         Height          =   225
         Left            =   3165
         TabIndex        =   49
         Top             =   1890
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   135
         X2              =   5250
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label Nome_Conta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2925
         TabIndex        =   48
         Top             =   1417
         Width           =   2310
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Centro 
      Bindings        =   "ManContasPagar.frx":4FB86
      DataSource      =   "Data3"
      Height          =   330
      Left            =   4020
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   7673
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1482
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1561
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.CommandButton B_Dia 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Em &Dia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7710
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CommandButton B_Confirma 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7710
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.TextBox Descrição 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3570
      MaxLength       =   30
      TabIndex        =   16
      Top             =   6525
      Visible         =   0   'False
      Width           =   5190
   End
   Begin VB.CommandButton B_Cancela 
      BackColor       =   &H00F7F7F7&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7710
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CommandButton B_Imprime_Cheque 
      BackColor       =   &H00F7F7F7&
      Caption         =   "&Imprimir Cheque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7710
      Visible         =   0   'False
      Width           =   2235
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   330
      Left            =   1170
      TabIndex        =   12
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Data_Pagto 
      Height          =   330
      Left            =   4020
      TabIndex        =   19
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   7305
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   330
      Left            =   1170
      TabIndex        =   13
      Top             =   6150
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Desconto 
      Height          =   330
      Left            =   1170
      TabIndex        =   15
      Top             =   6525
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Acréscimo 
      Height          =   330
      Left            =   1170
      TabIndex        =   17
      Top             =   6930
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Valor_Pago 
      Height          =   330
      Left            =   1170
      TabIndex        =   18
      Top             =   7305
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Nota 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   9435
      MaxLength       =   15
      TabIndex        =   20
      Top             =   5370
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lbl_valorTotalGrade 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   375
      Left            =   1140
      TabIndex        =   66
      Top             =   4890
      Width           =   1635
   End
   Begin VB.Label lblFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1350
      TabIndex        =   62
      Top             =   60
      Width           =   4515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      Height          =   195
      Left            =   75
      TabIndex        =   61
      Top             =   135
      Width           =   300
   End
   Begin VB.Label Sequência 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
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
      Height          =   330
      Left            =   7260
      TabIndex        =   60
      Top             =   5355
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Baixa 
      Caption         =   "Sequência"
      Height          =   255
      Index           =   7
      Left            =   6345
      TabIndex        =   59
      Top             =   5385
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Nome_Fornecedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8760
      TabIndex        =   58
      Top             =   67
      Width           =   6255
   End
   Begin VB.Label Label4 
      Caption         =   "Fornecedor"
      Height          =   255
      Left            =   6390
      TabIndex        =   57
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Baixa 
      Caption         =   "Fornecedor"
      Height          =   255
      Index           =   8
      Left            =   75
      TabIndex        =   45
      Top             =   5385
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1140
      TabIndex        =   44
      Top             =   5355
      Visible         =   0   'False
      Width           =   5070
   End
   Begin VB.Label Nome_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4950
      TabIndex        =   35
      Top             =   5760
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.Label Baixa 
      Caption         =   "Vencimento"
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   46
      Top             =   5805
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Baixa 
      Caption         =   "Valor"
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   43
      Top             =   6195
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Baixa 
      Caption         =   "Desconto"
      Height          =   255
      Index           =   2
      Left            =   75
      TabIndex        =   42
      Top             =   6570
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Baixa 
      Caption         =   "Acréscimo"
      Height          =   255
      Index           =   3
      Left            =   75
      TabIndex        =   41
      Top             =   6975
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Baixa 
      Caption         =   "Valor Pago"
      Height          =   255
      Index           =   4
      Left            =   75
      TabIndex        =   40
      Top             =   7350
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Baixa 
      Caption         =   "Nota"
      Height          =   255
      Index           =   6
      Left            =   8910
      TabIndex        =   39
      Top             =   5415
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Baixa 
      Caption         =   "Data Pagamento"
      Height          =   255
      Index           =   5
      Left            =   2775
      TabIndex        =   37
      Top             =   7350
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label_Vários 
      Caption         =   "Baixa de várias contas. Digite a data de pagamento. O valor pago será assumido como o valor da conta."
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
      Height          =   480
      Left            =   10740
      TabIndex        =   50
      Top             =   5250
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label Baixa 
      Caption         =   "Descrição"
      Height          =   255
      Index           =   10
      Left            =   2775
      TabIndex        =   38
      Top             =   6570
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Baixa 
      Caption         =   "Centro de Custo"
      Height          =   255
      Index           =   9
      Left            =   2775
      TabIndex        =   36
      Top             =   5805
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "frmManContasPagar"
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

Private rsCliFor As Recordset
Private Rec_Contas As Recordset
Private rsCentros As Recordset
Private rsCaixas As Recordset
Private rsCaixa As Recordset
Private rsContas_Pagar As Recordset
Private rsLançamentos As Recordset
Private rsContas As Recordset
'25/04/2005 - Daniel
'Otimizado rotina para abrir a tela de lançamentos de contas
'com a conta desejada a partir do duplo click
'
'Solicitante: Consultor Carlos (Petrópolis - RJ)
Public g_blnFind  As Boolean
Public g_strQuery As String

Sub Arruma_Caixa()

 Dim Sem_Caixa As Boolean
 Dim Ordem As Long
 Dim Tot_Dinheiro As Double
 Dim Tot_Cheques As Double
 Dim Tot_Pré As Double
 Dim Tot_Cartões As Double
 Dim Tot_Vales As Double
 Dim Saldo_Ant As Double
 Dim Tot_Parcela As Double

  'Arruma caixa
  
    Sem_Caixa = False
    rsCaixa.Index = "Data"
    rsCaixa.Seek ">", gnCodFilial, Val(Combo_Caixa.Text), Data_Atual, 0
    If rsCaixa.NoMatch Then Sem_Caixa = True
    If Sem_Caixa = False Then If gnCodFilial <> rsCaixa("Filial") Then Sem_Caixa = True
    If Sem_Caixa = False Then If Data_Atual <> rsCaixa("Data") Then Sem_Caixa = True
    If Sem_Caixa = False Then If Val(Combo_Caixa.Text) <> rsCaixa("Caixa") Then Sem_Caixa = True
  
    If Sem_Caixa = True Then
       'Acha o último dia
       Sem_Caixa = False
       rsCaixa.Seek "<", gnCodFilial, Val(Combo_Caixa.Text), Data_Atual, , 0
       If rsCaixa.NoMatch Then Sem_Caixa = True
       If Sem_Caixa = False Then If rsCaixa("Filial") <> gnCodFilial Then Sem_Caixa = True
       If Sem_Caixa = False Then If rsCaixa("Caixa") <> Val(Combo_Caixa.Text) Then Sem_Caixa = True
     
       If Sem_Caixa = True Then  'Caixa zerado
         rsCaixa.AddNew
           rsCaixa("Filial") = gnCodFilial
           rsCaixa("Data") = Data_Atual
           rsCaixa("Caixa") = Val(Combo_Caixa.Text)
           rsCaixa("Ordem") = 1
           rsCaixa("Descrição") = "Início do dia"
           rsCaixa("Hora") = Format(Time, "hh:mm:ss")
         rsCaixa.Update
       End If
       
       If Sem_Caixa = False Then  'Pegar caixa de outro dia
         rsCaixa.Index = "Data"
         rsCaixa.Seek "<", gnCodFilial, Val(Combo_Caixa.Text), Data_Atual, 9999#
         If rsCaixa.NoMatch Then
           Exit Sub
         End If
       
         If rsCaixa("Filial") <> gnCodFilial Then Exit Sub
     '    If rsCaixa("Data") <> Data_Atual Then Exit Sub
  
         Ordem = 1
  
         Tot_Dinheiro = rsCaixa("Total Dinheiro")
         Tot_Cheques = rsCaixa("Total Cheques")
         Tot_Pré = rsCaixa("Total Cheques Pré")
         Tot_Cartões = rsCaixa("Total Cartões")
         Tot_Vales = rsCaixa("Total Vales")
         Tot_Parcela = rsCaixa("Total Parcelamento")
         Saldo_Ant = rsCaixa("Final")
         rsCaixa.AddNew
          rsCaixa("Filial") = gnCodFilial
          rsCaixa("Data") = Data_Atual
          rsCaixa("Caixa") = Val(Combo_Caixa.Text)
          rsCaixa("Ordem") = 1
          rsCaixa("Descrição") = "Início do dia"
          rsCaixa("Dinheiro") = Tot_Dinheiro
          rsCaixa("Total Dinheiro") = Tot_Dinheiro
          rsCaixa("Cheques") = Tot_Cheques
          rsCaixa("Total Cheques") = Tot_Cheques
          rsCaixa("Cheques Pré") = Tot_Pré
          rsCaixa("Total Cheques Pré") = Tot_Pré
          rsCaixa("Cartões") = Tot_Cartões
          rsCaixa("Total Cartões") = Tot_Cartões
          rsCaixa("Vales") = Tot_Vales
          rsCaixa("Total Vales") = Tot_Vales
          rsCaixa("Parcelamento") = Tot_Parcela
          rsCaixa("Total Parcelamento") = Tot_Parcela
          rsCaixa("Saldo Anterior") = Saldo_Ant
          rsCaixa("Final") = Saldo_Ant
        rsCaixa.Update
      End If
    End If
       
       
    ' Acha o último caixa
    rsCaixa.Index = "Data"
    rsCaixa.Seek "<", gnCodFilial, Val(Combo_Caixa.Text), Data_Atual, 9999#
    If rsCaixa.NoMatch Then
      Exit Sub
    End If
       
    If rsCaixa("Filial") <> gnCodFilial Then Exit Sub
    If rsCaixa("Data") <> Data_Atual Then Exit Sub
    If rsCaixa("Caixa") <> Val(Combo_Caixa.Text) Then Exit Sub
  
    Ordem = rsCaixa("Ordem") + 1
  
    Tot_Dinheiro = rsCaixa("Total Dinheiro")
    Tot_Cheques = rsCaixa("Total Cheques")
    Tot_Pré = rsCaixa("Total Cheques Pré")
    Tot_Cartões = rsCaixa("Total Cartões")
    Tot_Vales = rsCaixa("Total Vales")
    Tot_Parcela = rsCaixa("Total Parcelamento")
    Saldo_Ant = rsCaixa("Final")
     
    rsCaixa.AddNew
      rsCaixa("Filial") = gnCodFilial
      rsCaixa("Data") = Data_Atual
      rsCaixa("Caixa") = Val(Combo_Caixa.Text)
      rsCaixa("Ordem") = Ordem
      rsCaixa("Descrição") = Left("Conta paga - " + Nome_Cliente.Caption, 30)
        
      rsCaixa("Dinheiro") = 0
      If O_caixa_d.Value = True Then rsCaixa("Dinheiro") = -CDbl(Valor_Pago.Text)
      rsCaixa("Total Dinheiro") = Tot_Dinheiro + rsCaixa("Dinheiro")
              
      rsCaixa("Cheques") = 0
      If O_Caixa_C.Value = True Then rsCaixa("Cheques") = -CDbl(Valor_Pago.Text)
      rsCaixa("Total Cheques") = Tot_Cheques + rsCaixa("Cheques")
        
      rsCaixa("Cheques Pré") = 0
     ' If O_Caixa_P.Value = True Then rsCaixa("Cheques Pré") = -CDbl(Valor_Pago.Text)
      rsCaixa("Total Cheques Pré") = Tot_Pré
        
      rsCaixa("Cartões") = 0
      rsCaixa("Total Cartões") = Tot_Cartões
      rsCaixa("Vales") = 0
      rsCaixa("Total Vales") = Tot_Vales
      rsCaixa("Parcelamento") = 0
      rsCaixa("Total Parcelamento") = 0
      rsCaixa("Saldo Anterior") = Saldo_Ant
      rsCaixa("Final") = Saldo_Ant + rsCaixa("Dinheiro") + rsCaixa("Cheques")
      rsCaixa("Final") = rsCaixa("Final") + rsCaixa("Cheques Pré")
    rsCaixa.Update

End Sub

Sub Arruma_Conta()

  Dim Saldo_Ant As Double

   'Atualiza Dinheiro na conta, se for o caso
   If CDbl(Valor_Pago.Text) <> 0 And O_Conta.Value = True Then
     rsLançamentos.Index = "Conta"
     rsLançamentos.Seek "<", Val(Combo_Conta.Text), CDate(Data_Atual), 99999999#
     If rsLançamentos.NoMatch Then Saldo_Ant = 0
     If Not rsLançamentos.NoMatch Then
        Saldo_Ant = 0
        If rsLançamentos("Conta") = Val(Combo_Conta.Text) Then Saldo_Ant = rsLançamentos("Saldo Atual")
     End If
     
     rsLançamentos.AddNew
       rsLançamentos("Conta") = Val(Combo_Conta.Text)
       rsLançamentos("Data") = Cheque_Bom.Text
       rsLançamentos("Descrição") = Left("Conta paga - " + Nome_Cliente.Caption, 40)
       rsLançamentos("Saldo Anterior") = Saldo_Ant
       rsLançamentos("Débito") = CDbl(Valor_Pago.Text)
       rsLançamentos("Cheque") = Num_Cheque.Text
       rsLançamentos("Saldo Atual") = Saldo_Ant - CDbl(Valor_Pago.Text)
     rsLançamentos.Update
   End If
   
End Sub

Private Sub Acréscimo_LostFocus()
  '30/05/2005 - Daniel
  'Cálculo do Valor Pago automático
  'Solicitante: Pedágio
  'Dim dblAcrescimo As Double
  
  'dblAcrescimo = Format(CDbl(0 & (Acréscimo.Text)), FORMAT_VALUE)
  
  'Valor_Pago.Text = Format((Valor.Text) + dblAcrescimo - CDbl(0 & (Desconto.Text)), FORMAT_VALUE)
  
  ' Calcula valores dos camos Desconto, Acréscimo e Valor_Pago (22/06/2022 - Pablo)
  Call CalculaValores("Acréscimo")
End Sub

Private Sub B_Baixa_Click()
On Error GoTo Erro

  Dim Linha As Long
  Dim i As Integer
  Dim Pagas As Integer
  Dim book As Variant
  Dim Valor_Contas As Double
  
  Call StatusMsg("")
 
  If Grade1.SelBookmarks.Count < 1 Then
    DisplayMsg "Para selecionar uma conta, clique na coluna cor CINZA da esquerda."
    Exit Sub
  End If
  
  Pagas = False
  Valor_Contas = 0
  For i = 0 To (Grade1.SelBookmarks.Count - 1)
    book = Grade1.SelBookmarks(i)
    If Grade1.Columns("Pagto").CellValue(book) <> 0 Then
        Pagas = True
    End If
    Valor_Contas = Valor_Contas + Grade1.Columns("Valor").CellValue(book)
  Next i
  If Pagas = True Then
    DisplayMsg "Uma ou mais contas selecionadas já foram pagas e não podem ser baixadas. Caso deseje use a tela de lançamentos para alterá-las."
    Exit Sub
  End If
 
  'Grade1.Height = 2175
 
  Call StatusMsg("Aguarde...")
  
  Vencimento.Visible = True
  Valor.Visible = True
  Desconto.Visible = True
  Acréscimo.Visible = True
  Valor_Pago.Visible = True
  Data_Pagto.Visible = True
  Nota.Visible = True
  ' Sequência.Visible = True
  Nome_Cliente.Visible = True
  Combo_Centro.Visible = True
  Nome_Centro.Visible = True
  Descrição.Visible = True
  
   
  B_Cancela.Visible = True
  B_Dia.Visible = True
  B_Confirma.Visible = True
  B_Imprime_Cheque.Visible = True
  
  O_Não_determinado.Value = True
  Cheque_Bom.Mask = ""
  Cheque_Bom.Text = ""
  Cheque_Bom.Mask = "##/##/####"
  Num_Cheque.Text = ""
 
 
  If Grade1.SelBookmarks.Count <> 1 Then
      Valor.Text = Valor_Contas
      Valor.Enabled = False
      Vencimento.Enabled = False
      Desconto.Enabled = False
      Acréscimo.Enabled = False
      Valor_Pago.Text = Valor
      Valor_Pago.Enabled = False
      Nota.Enabled = False
    '   Sequência.Enabled = False
      Nome_Cliente.Enabled = False
      Combo_Centro.Enabled = False
      Label_Vários.Visible = True
      Descrição.Enabled = False
  Else
      Valor.Enabled = True
      Vencimento.Enabled = True
      Desconto.Enabled = True
      Acréscimo.Enabled = True
      Valor_Pago.Enabled = True
      Nota.Enabled = True
    '   Sequência.Enabled = True
      Nome_Cliente.Enabled = True
      Combo_Centro.Enabled = True
      Label_Vários.Visible = False
      Descrição.Enabled = True
  End If
 
 
  For i = 0 To 10
      Baixa(i).Visible = True
  Next i
  Baixa(7).Visible = False
 
  If Grade1.SelBookmarks.Count = 1 Then
     book = Grade1.SelBookmarks(0)
     'Vencimento.Text = Format((Grade1.Columns("Vencimento").CellValue(book)), "dd/mm/yyyy")
     Vencimento.Text = gsFormatDate(Grade1.Columns("Vencimento").CellValue(book))
     Valor.Text = Grade1.Columns("Valor").CellValue(book)
     Desconto.Text = Grade1.Columns("Desconto").CellValue(book)
     Acréscimo.Text = Grade1.Columns("Acréscimo").CellValue(book)
     Nota.Text = Grade1.Columns("Nota").CellValue(book)
  '   Sequência.Caption = Grade1.Columns("Sequência").CellValue(book)
     Combo_Centro.Text = Grade1.Columns("C.Custo").CellValue(book)
     Nome_Cliente.Caption = Grade1.Columns("Nome").CellValue(book)
     Descrição.Text = Grade1.Columns("Descrição").CellValue(book)
     Valor_Pago.Text = ""
     Data_Pagto.Mask = ""
     Data_Pagto.Text = ""
     Data_Pagto.Mask = "##/##/####"
  End If
 
  B_Monta.Enabled = False
  B_Baixa.Enabled = False
  Frame_Modo.Visible = True
  Grade1.Enabled = False
 
  Data_Pagto.Text = Format(Data_Atual, "dd/MM/yyyy")
  
  '22/06/2022 (Pablo) - atualiza o nome do centro de custo no campo
  Call Combo_Centro_LostFocus
 
  Call StatusMsg("")
 
  Exit Sub
Erro:
  MsgBox "Erro em carregar baixa de conta " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
 
End Sub


Private Sub B_Cancela_Click()
  Dim i As Integer
  
  For i = 0 To 10
   Baixa(i).Visible = False
  Next i
  
  Vencimento.Visible = False
  Valor.Visible = False
  Desconto.Visible = False
  Acréscimo.Visible = False
  Valor_Pago.Visible = False
  Data_Pagto.Visible = False
  Nota.Visible = False
'  Sequência.Visible = False
  Nome_Cliente.Visible = False
  Combo_Centro.Visible = False
  Nome_Centro.Visible = False
  Descrição.Visible = False
  Label_Vários.Visible = False
  
  B_Cancela.Visible = False
  B_Dia.Visible = False
  B_Confirma.Visible = False
  B_Imprime_Cheque.Visible = False
  
  Frame_Modo.Visible = False
  
  B_Monta.Enabled = True
  B_Baixa.Enabled = True
  
  Grade1.Enabled = True
  'Grade1.Height = 5100
 
End Sub

Private Sub B_Confirma_Click()
  Dim Resposta As Integer
  Dim Erro As Integer
  Dim i As Integer
  Dim book As Variant
  Dim Valor_Aux As Double
  Dim sData As String
  Dim nContador As Long
  Dim bOk As Boolean

  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked  As Integer
  
  Call StatusMsg("")

  On Error GoTo Trata_Erro:

  If IsNull(Valor_Pago.Text) Then Valor_Pago.Text = 0
  If Valor_Pago.Text = "" Then Valor_Pago.Text = 0

  If CDbl(Valor_Pago.Text) = 0 And Grade1.SelBookmarks.Count = 1 Then
     Data_Pagto.Mask = ""
     Data_Pagto.Text = ""
     Data_Pagto.Mask = "##/##/####"
     GoTo Cont1
  End If
  
  If O_Não_determinado.Value = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja realizar a operação sem indicar a origem do dinheiro?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
'      DisplayMsg "Operação não realizada."
      Exit Sub
    End If
  Else
    If O_Conta.Value = False Then
      If Nome_Caixa.Caption = "" Then
        DisplayMsg "Escolha o caixa de onde está saindo o dinheiro / cheque."
        Combo_Caixa.SetFocus
        Exit Sub
      End If
    End If
    If O_Conta.Value = True Then
      If Nome_Conta.Caption = "" Then
        DisplayMsg "Escolha a conta."
        Exit Sub
      End If
      
      If Cheque_Bom.Text = "  /  /    " Then
        Cheque_Bom.Text = Data_Atual
      End If
      
      If Not IsDate(Cheque_Bom.Text) Then
        DisplayMsg "Digite a data do cheque."
        Exit Sub
      End If
      
      If IsNull(Num_Cheque.Text) Then
        Num_Cheque.Text = ""
        'DisplayMsg "Digite o número do cheque."
        'Exit Sub
      End If
    End If
    
    If O_Caixa_C.Value = True Then
      If Not IsDate(Cheque_Bom.Text) Then
        DisplayMsg "Digite a data do cheque."
        Exit Sub
      End If
      
      If IsNull(Num_Cheque.Text) Then
        DisplayMsg "Digite o número do cheque."
        Exit Sub
      End If
    End If
  End If
  
    
  If IsNull(Acréscimo.Text) Or Acréscimo.Text = "" Then Acréscimo.Text = 0
  If IsNull(Desconto.Text) Or Desconto.Text = "" Then Desconto.Text = 0
   
  
  Valor_Aux = CDbl(Valor.Text) + CDbl(Acréscimo.Text) - CDbl(Desconto.Text)
  If Abs(Valor_Aux - CDbl(Valor_Pago.Text)) > 0.001 Then
    DisplayMsg "Valor pago incorreto, verifique."
    Exit Sub
  End If

  Erro = False
  If Not IsDate(Data_Pagto.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data de pagamento incorreta, verifique."
    Exit Sub
  End If
  
  rsContas_Pagar.Requery
  
  ws.BeginTrans
  blnInTransaction = True
  
  
  If CDbl(Valor_Pago.Text) <> 0 And O_Não_determinado.Value = False And O_Conta.Value = False Then Arruma_Caixa
  
  Call Arruma_Conta
  
  If Grade1.SelBookmarks.Count > 1 Then
    For i = 0 To (Grade1.SelBookmarks.Count - 1)
      book = Grade1.SelBookmarks(i)
      Grade1.Bookmark = book
      nContador = Grade1.Columns("Contador").CellValue(book)
      With rsContas_Pagar
        '.Requery
        
        .FindFirst "Contador = " & nContador
        If Not .NoMatch Then
          .LockEdits = True
          .Edit
          ![Valor Pago] = !Valor
          !Pagamento = Data_Pagto.Text
          ![Data Alteração] = Data_Atual
          .Update
          bOk = True
        End If
      End With
    Next i
    
    ws.CommitTrans
    blnInTransaction = False
    
    If bOk Then
      B_Cancela_Click
      B_Monta_Click
    End If
    
    Exit Sub
  End If
  
Cont1:
  
  book = Grade1.SelBookmarks(0)
  Grade1.Bookmark = book
  nContador = Grade1.Columns("Contador").Value
'  Grade1.Columns("Vencimento").Value = CDate(Vencimento.Text)
'  Grade1.Columns("Valor").Value = CDbl(gsHandleNull(Valor.Text))
'  Grade1.Columns("Desconto").Value = CDbl(gsHandleNull(Desconto.Text))
'  Grade1.Columns("Acréscimo").Value = CDbl(gsHandleNull(Acréscimo.Text))
'  Grade1.Columns("Pago").Value = CDbl(gsHandleNull(Valor_Pago.Text))
'  If IsDate(Data_Pagto.Text) Then Grade1.Columns("Pagamento").Value = Data_Pagto.Text
'  Grade1.Columns("Nota").Value = Nota.Text
'  If Nome_Centro.Caption <> "" Then Grade1.Columns("C.Custo").Value = Val(Combo_Centro.Text)
'  Grade1.Columns("Descrição").Value = Descrição.Text
'  Grade1.Update
  
  With rsContas_Pagar
    If Not blnInTransaction Then
      ws.BeginTrans
      blnInTransaction = True
    End If
    
    .FindFirst "Contador = " & nContador
    If Not .NoMatch Then
      .LockEdits = True
      .Edit
      .Fields("Vencimento") = CDate(Vencimento.Text)
      .Fields("Valor") = CDbl(gsHandleNull(Valor.Text))
      .Fields("Desconto").Value = CDbl(gsHandleNull(Desconto.Text))
      .Fields("Acréscimo").Value = CDbl(gsHandleNull(Acréscimo.Text))
      .Fields("Valor Pago").Value = CDbl(gsHandleNull(Valor_Pago.Text))
      If IsDate(Data_Pagto.Text) Then
        .Fields("Pagamento").Value = Data_Pagto.Text
      End If
      .Fields("Nota").Value = Nota.Text & ""
      .Fields("Centro de Custo").Value = Val(Combo_Centro.Text)
      .Fields("Descrição") = Descrição.Text
      .Fields("Data Alteração") = Data_Atual
      '--------------------------------------------------------
      '27/04/2005 - Daniel
      'Adição do campo OrigemDinheiro
      'Solicitação: Bem Me Quer
      .Fields("OrigemDinheiro").Value = GetOrigemDinheiro & ""
      '--------------------------------------------------------
      .Update
    End If
    
    ws.CommitTrans
    blnInTransaction = False
  End With
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & Sequência.Caption & " Forn:" & Combo_Fornecedor.Text & " vOr:" & Valor.Text & " vPg:" & Valor_Pago.Text & " Dsc:" & Descrição.Text, 80) & "', 'CNT_PAG: baixar')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
  Valor_Pago.Text = ""
  Data_Pagto.Mask = ""
  Data_Pagto.Text = ""
  Data_Pagto.Mask = "##/##/####"
  
  B_Cancela_Click
  B_Monta_Click
  
  'Grade1.Height = 5100
  
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

Private Sub B_Dia_Click()
  Valor_Pago.Text = Valor.Text
  Data_Pagto.Mask = ""
'  Data_Pagto.Text = Vencimento.Text
  Data_Pagto.Text = Format(Data_Atual, "dd/MM/yyyy")
  Data_Pagto.Mask = "##/##/####"
  Desconto.Text = 0
  Acréscimo.Text = 0
End Sub

Private Sub B_Imprime_Cheque_Click()
  With frmImprimeCheque2
    .Favorecido.Text = Nome_Cliente.Caption
    If gsHandleNull(Valor_Pago.Text) <> "0" Then
      .Valor.Text = Valor_Pago.Text
    Else
      .Valor.Text = Valor.Text
    End If
    .Show
  End With
End Sub

Private Sub B_Monta_Click()
On Error GoTo Erro
  Dim sSql              As String
  Dim i                 As Integer
  Dim Data_Ini          As String
  Dim Data_Fim          As String
  Dim sFilial           As String
  Dim dvalorTotalGrade  As Double
  
  Call StatusMsg("")
  
  lbl_valorTotalGrade.Caption = "0,00"
  
  If Not IsDate(Vcto_Inicial.Text) Then
    DisplayMsg "Data vencimento inicial incorreta."
    Vcto_Inicial.SetFocus
    Exit Sub
  End If
  Data_Ini = gsGetInvDate(Vcto_Inicial.Text)
  
  If Not IsDate(Vcto_Final.Text) Then
    DisplayMsg "Data vencimento final incorreta."
    Vcto_Final.SetFocus
    Exit Sub
  End If
  Data_Fim = gsGetInvDate(Vcto_Final.Text)
  
  If CDate(Vcto_Final.Text) < CDate(Vcto_Inicial.Text) Then
    DisplayMsg "Data vencimento inicial deve ser menor ou igual a data vencimento final."
    Vcto_Inicial.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(cboFilial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If
  
  'Verifica a filial
  cboFilial_LostFocus
  If lblFilial.Caption = "" Then
    sFilial = "<> 0"
  Else
    sFilial = "= " & cboFilial
  End If
  
  ' sSql = "SELECT Filial, Valor, Vencimento, [Contas a Pagar].Desconto, Acréscimo, [Valor Pago], Pagamento,  Nota, [Centro de Custo], Fornecedor, Cli_For.Nome, Sequência, Descrição, Contador FROM [Contas a Pagar]"
  sSql = "SELECT Filial, Valor, Vencimento, [Contas a Pagar].Desconto, Acréscimo, [Valor Pago], Pagamento,  Nota, [Centro de Custo], Fornecedor, Cli_For.Nome, Sequência, Descrição, Contador FROM [Contas a Pagar]"
  sSql = sSql + " INNER JOIN Cli_For ON ([Contas a Pagar].Fornecedor = Cli_For.Código) "
  sSql = sSql + " WHERE Filial " & sFilial & " AND Vencimento >= " + Data_Ini
  sSql = sSql + " And Vencimento <= " + Data_Fim
  
  If Nome_Fornecedor.Caption <> "" Then
    sSql = sSql + " And Fornecedor = " + Combo_Fornecedor.Text
  End If
  
  If O_Receber = True Then sSql = sSql + " AND [Valor Pago] = 0"
  If O_Recebidas = True Then sSql = sSql + " AND [Valor Pago] <> 0"
  
  If O_Cliente.Value = True Then
    sSql = sSql + " ORDER BY Fornecedor;"
  Else
    If O_Vencimento.Value = True Then
      sSql = sSql + " ORDER BY Vencimento;"
    End If
  End If
  
  Call StatusMsg("Aguarde, montando tabela...")
  
  Grade1.DataMode = ssDataModeAddItem
  Grade1.RemoveAll
  Grade1.Refresh
  
  Set Rec_Contas = db.OpenRecordset(sSql, dbOpenDynaset)
  If Rec_Contas.EOF Then
    DisplayMsg "Nenhuma conta encontrada segundo os critérios solicitados."
    Exit Sub
  Else
    dvalorTotalGrade = 0
    Rec_Contas.MoveFirst
    While Not Rec_Contas.EOF
      dvalorTotalGrade = dvalorTotalGrade + Rec_Contas.Fields("Valor")
      Rec_Contas.MoveNext
    Wend
    lbl_valorTotalGrade.Caption = FormatNumber(dvalorTotalGrade, 2)
  End If
  
  Rec_Contas.MoveLast
  Rec_Contas.MoveFirst
  
  Grade1.DataMode = ssDataModeUnbound
  
  Set Data2.Recordset = Rec_Contas
  
  'Grade1.Visible = False
  
  'Colunas
  '0 Empresa
  '1 Vencimento
  '2 Cliente
  '3 Cli_For.Nome
  '4 Nota
  '5 Fatura
  '6 Sequência FROM [Contas a Receber]
  '7 Vencimento
  '8 Valor
  '9 [Contas a Receber].Desconto
  '10 Acréscimo
  '11 [Valor Pago]
  '12 Pagamento
  '13 Contador
  
  Grade1.Visible = False
  Grade1.DataMode = 0
  Grade1.ReBind
  'Grade1.Groups.Add 0
  Grade1.LevelCount = 1
  
'  Grade1.Columns(0).Visible = False
''  Grade1.Columns(1).NumberFormat = "###,##0.00"
  Grade1.Columns(13).Visible = False
  
'  Grade1.Columns(0).Level = 0
'  Grade1.Columns(1).Level = 0
'  Grade1.Columns(2).Level = 0
'  Grade1.Columns(3).Level = 0
'  Grade1.Columns(4).Level = 0
'  Grade1.Columns(5).Level = 0
'  Grade1.Columns(6).Level = 0
'
'  Grade1.Columns(7).Level = 0
'  Grade1.Columns(8).Level = 0
'  Grade1.Columns(9).Level = 0
'  Grade1.Columns(10).Level = 0
'  Grade1.Columns(11).Level = 0
'  Grade1.Columns(12).Level = 0
'  Grade1.Columns(13).Level = 0
  
  
''  Grade1.Columns(3).NumberFormat = "##,###,##0.00"
''  Grade1.Columns(4).NumberFormat = "##,###,##0.00"
''  Grade1.Columns(5).NumberFormat = "##,###,##0.00"
  
  Grade1.Columns("Filial").Width = 500
  
  Grade1.Columns(1).Width = 880
  Grade1.Columns(2).Width = 990
  Grade1.Columns(2).Caption = "Vencimento"
  
  Grade1.Columns(3).Width = 880
  Grade1.Columns(3).Caption = "Desconto"
  
  Grade1.Columns(4).Width = 880
  Grade1.Columns(4).Caption = "Acréscimo"
  
  Grade1.Columns(5).Width = 1020
  Grade1.Columns(5).Caption = "Valor Pago"
  
  Grade1.Columns(6).Width = 950
  Grade1.Columns(6).Caption = "Pagto"
  
  Grade1.Columns(7).Width = 880
  
  Grade1.Columns(8).Width = 880
  Grade1.Columns(8).Caption = "C.Custo"
  
  Grade1.Columns(9).Width = 880
  Grade1.Columns(9).Caption = "Forn"
  Grade1.Columns(10).Width = 2340
  
  Grade1.Columns(11).Width = 920
  Grade1.Columns(11).Caption = "Sequência"
  
  Grade1.Columns(12).Width = 2340
  
  Grade1.Visible = True
  Grade1.Caption = "Contas"
  Call StatusMsg("")
  
  Exit Sub
Erro:
  MsgBox "Erro na pesquisa de contas a pagar " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub Baixa_Click(Index As Integer)
  If Index = 4 Then
    If IsNull(Valor.Text) Then Exit Sub
    If Valor.Text = "" Then Exit Sub
    If Not IsNumeric(Valor.Text) Then Exit Sub
    If CDbl(Valor.Text) <= 0 Then Exit Sub
  
    If IsNull(Desconto.Text) Then Desconto.Text = 0
    If Desconto.Text = "" Then Desconto.Text = 0
    If Not IsNumeric(Desconto.Text) Then Desconto.Text = 0
    If CDbl(Desconto.Text) < 0 Then Desconto.Text = 0
  
    If IsNull(Acréscimo.Text) Then Acréscimo.Text = 0
    If Acréscimo.Text = "" Then Acréscimo.Text = 0
    If Not IsNumeric(Acréscimo.Text) Then Acréscimo.Text = 0
    If CDbl(Acréscimo.Text) < 0 Then Acréscimo.Text = 0
  
    Valor_Pago.Text = CDbl(Valor.Text) - CDbl(Desconto.Text) + CDbl(Acréscimo.Text)
  End If
End Sub

Private Sub cboFilial_CloseUp()
  lblFilial.Caption = cboFilial.Columns("Nome").Text
  cboFilial.Text = cboFilial.Columns("Filial").Text
End Sub

Private Sub cboFilial_KeyPress(KeyAscii As Integer)
  If Not cboFilial.DroppedDown Then
    KeyAscii = gnLimitKeyPress(cboFilial, 2, KeyAscii, True)
  End If
End Sub

Private Sub cboFilial_LostFocus()
  lblFilial.Caption = gsGetNameFilial(Val(cboFilial.Text))
End Sub

Private Sub Cheque_Bom_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Cheque_Bom.Text = frmCalendario.gsDateCalender(Cheque_Bom.Text)
  End Select
End Sub

Private Sub Cheque_Bom_LostFocus()
  Cheque_Bom.Text = Ajusta_Data(Cheque_Bom.Text)
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Vcto_Final.Text = frmCalendario.gsDateCalender(Vcto_Final.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Vcto_Inicial.Text = frmCalendario.gsDateCalender(Vcto_Inicial.Text)
End Sub

Private Sub Combo_Caixa_CloseUp()
  Combo_Caixa.Text = Combo_Caixa.Columns(1).Text
  Combo_Caixa_LostFocus
End Sub

Private Sub Combo_Caixa_LostFocus()
  Nome_Caixa.Caption = ""
  If IsNull(Combo_Caixa.Text) Then Exit Sub
  If Combo_Caixa.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Caixa.Text) Then Exit Sub
  If Val(Combo_Caixa.Text) < 1 Then Exit Sub
  If Val(Combo_Caixa.Text) > 99 Then Exit Sub
  
  rsCaixas.Index = "Caixa"
  rsCaixas.Seek "=", Val(Combo_Caixa.Text)
  If rsCaixas.NoMatch Then Exit Sub
  Nome_Caixa.Caption = rsCaixas("Descrição") & ""

End Sub

Private Sub Combo_Centro_CloseUp()
  Combo_Centro.Text = Combo_Centro.Columns(1).Text
  Combo_Centro_LostFocus
End Sub

Private Sub Combo_Centro_LostFocus()
  Nome_Centro.Caption = ""
  
  If IsNull(Combo_Centro.Text) Then Exit Sub
  If Combo_Centro.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Centro.Text) Then Exit Sub
  If Val(Combo_Centro.Text) < 1 Then Exit Sub
  If Val(Combo_Centro.Text) > 9999 Then
    Combo_Centro.Text = 0
    Exit Sub
  End If
  
  rsCentros.Index = "Código"
  rsCentros.Seek "=", Val(Combo_Centro.Text)
  If rsCentros.NoMatch Then Exit Sub
  Nome_Centro.Caption = rsCentros("Nome") & ""
  
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
 '28/11/2006 - Anderson
 'Alteração do número de contas bancárias de 99 para 255
 'Solicitado por: 2227883 - SANTA FÉ DO ARAGUAIA PREFEITURA MUNICIPAL
 If Val(Combo_Conta.Text) < 1 Or Val(Combo_Conta.Text) > 255 Then Exit Sub
 
 rsContas.Index = "Código"
 rsContas.Seek "=", Val(Combo_Conta.Text)
 
 If rsContas.NoMatch Then Exit Sub
 
 Nome_Conta.Caption = rsContas("Descrição") & ""
 
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
  Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_LostFocus()
 Nome_Fornecedor.Caption = ""
 If IsNull(Combo_Fornecedor.Text) Then Exit Sub
 If Combo_Fornecedor.Text = "" Then Exit Sub
 If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
 If Val(Combo_Fornecedor.Text) < 0 Or Val(Combo_Fornecedor.Text) > 99999999 Then Exit Sub
 
 rsCliFor.Index = "Código"
 rsCliFor.Seek "=", Val(Combo_Fornecedor.Text)
 If rsCliFor.NoMatch Then Exit Sub
 Nome_Fornecedor.Caption = rsCliFor("Nome")
End Sub

Private Sub Data_Pagto_LostFocus()
  Data_Pagto.Text = Ajusta_Data(Data_Pagto.Text)
End Sub

Private Sub Data_Pagto_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Pagto.Text = frmCalendario.gsDateCalender(Data_Pagto.Text)
  End Select
End Sub

Private Sub Desconto_LostFocus()
  '30/05/2005 - Daniel
  'Cálculo do Valor Pago automático
  'Solicitante: Pedágio
  'Dim dblDesconto As Double

  'dblDesconto = Format(0 & (Desconto.Text), FORMAT_VALUE)
  
  'Valor_Pago.Text = Format(CDbl(Valor.Text) - dblDesconto + CDbl(0 & (Acréscimo.Text)), FORMAT_VALUE)

  ' Calcula valores dos camos Desconto, Acréscimo e Valor_Pago (22/06/2022 - Pablo)
  Call CalculaValores("Desconto")
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsCentros = db.OpenRecordset("Centros de Custo", , dbReadOnly)
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsContas_Pagar = db.OpenRecordset("Contas a Pagar", dbOpenDynaset)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários")
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  'Grade1.Height = 5100  'Posição inicial
  
  If gbCaixas = False Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
    Combo_Caixa.Enabled = False
  End If
  
  Call GetSettings
  
  cboFilial.Text = gnCodFilial
  cboFilial_LostFocus
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Call SaveSetting("QuickStore", "CPMan", "Data1", Vcto_Inicial.Text)
  Call SaveSetting("QuickStore", "CPMan", "Data2", Vcto_Final.Text)
  
  rsCliFor.Close
  rsCentros.Close
  rsCaixas.Close
  rsCaixa.Close
  rsContas_Pagar.Close
  rsLançamentos.Close
  rsContas.Close
  
  Set rsCliFor = Nothing
  Set rsCentros = Nothing
  Set rsCaixas = Nothing
  Set rsCaixa = Nothing
  Set rsContas_Pagar = Nothing
  Set rsLançamentos = Nothing
  Set rsContas = Nothing
End Sub

Private Sub GetSettings()
  Vcto_Final.Text = GetSetting("QuickStore", "CPMan", "Data2", CDate(Date))
  Vcto_Inicial.Text = GetSetting("QuickStore", "CPMan", "Data1", CDate(Date))
End Sub

Private Sub Grade1_AfterDelete(RtnDispErrMsg As Integer)
  Grade1.Scroll 0, -32767
  Grade1.Scroll 0, 32767
End Sub

Private Sub Grade1_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  If Not bGridBeforeDelete() Then
    Cancel = True
  End If
End Sub

Private Sub Grade1_DblClick()
  '25/04/2005 - Daniel
  'Otimizado rotina para abrir a tela de lançamentos de contas
  'com a conta selecionada através do duplo click
  '
  'Solicitante: Consultor Carlos (Petrópolis - RJ)
  
  On Error GoTo TratarErro
  
  Me.g_blnFind = True
  
  g_strQuery = ""
  g_strQuery = "SELECT * FROM [Contas a Pagar] "
  g_strQuery = g_strQuery & " WHERE Filial = " & CByte(Grade1.Columns(0).Text)
  g_strQuery = g_strQuery & " AND Vencimento = #" & Format(CDate(Grade1.Columns(2).Text), "MM/DD/YYYY") & "#"
  g_strQuery = g_strQuery & " AND Fornecedor = " & CLng(Grade1.Columns(9).Text)
  g_strQuery = g_strQuery & " AND Contador = " & CLng(Grade1.Columns(13).Text)
  
  
  '01/02/2023 - Mauro
  'Liberado da senha do gerente para edição de contas pagas a pedido da MaréMansa
  'Aguarda estudo de processos de log para evitar fraudes.
  'If Not frmGerente.gbSenhaGerente Then
  '  Exit Sub
  'End If
  
  frmLancaCPagar.Show
  
  Exit Sub
  
TratarErro:
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub Grade1_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  Dim nValueSelected As Double
  Dim nX As Integer
  Dim vBook As Variant
  Dim sConta As String
  Dim sSelec As String
 
  If Grade1.SelBookmarks.Count = 0 Then
   Grade1.Caption = "Contas: nenhuma selecionada"
   Exit Sub
  End If
  
  nValueSelected = 0#
  For nX = 0 To (Grade1.SelBookmarks.Count - 1)
    vBook = Grade1.SelBookmarks(nX)
    nValueSelected = nValueSelected + Grade1.Columns("Valor").CellValue(vBook)
  Next nX
    
  If Grade1.SelBookmarks.Count = 1 Then
    sConta = "Conta: "
    sSelec = " selecionada"
  Else
    sConta = "Contas: "
    sSelec = " selecionadas"
  End If
  
  Grade1.Caption = sConta & CStr(Grade1.SelBookmarks.Count) & sSelec & ", valor " + Format((CStr(nValueSelected)), "Currency")

End Sub


Private Sub O_Caixa_C_Click()
  Label_Caixa.Enabled = True
  Combo_Caixa.Enabled = True
  Nome_Caixa.Enabled = True
End Sub

Private Sub O_caixa_d_Click()
  Label_Caixa.Enabled = True
  Combo_Caixa.Enabled = True
  Nome_Caixa.Enabled = True
  
  If gbCaixas = False Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
    Combo_Caixa.Enabled = False
  End If
End Sub

Private Sub O_Conta_Click()
  Combo_Conta.Enabled = True
  Nome_Conta.Enabled = True
End Sub

Private Sub O_Não_determinado_Click()
  Label_Caixa.Enabled = False
  Combo_Caixa.Enabled = False
  Nome_Caixa.Enabled = False
End Sub

Private Sub Valor_Pago_LostFocus()
  ' Calcula valores dos camos Desconto, Acréscimo e Valor_Pago (22/06/2022 - Pablo)
  Call CalculaValores("Valor_Pago")
End Sub

Private Sub Vcto_Final_GotFocus()
  With Vcto_Final
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub Vcto_Final_LostFocus()
  Vcto_Final.Text = Ajusta_Data(Vcto_Final.Text)
End Sub

Private Sub Vcto_Final_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Vcto_Final.Text = frmCalendario.gsDateCalender(Vcto_Final.Text)
  End Select
End Sub

Private Sub Vcto_Inicial_GotFocus()
  With Vcto_Inicial
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub Vcto_Inicial_LostFocus()
  Vcto_Inicial.Text = Ajusta_Data(Vcto_Inicial.Text)
End Sub

Private Sub Vcto_Inicial_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Vcto_Inicial.Text = frmCalendario.gsDateCalender(Vcto_Inicial.Text)
  End Select
End Sub

Private Sub Vencimento_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Vencimento.Text = frmCalendario.gsDateCalender(Vencimento.Text)
  End Select
End Sub

Private Sub Vencimento_LostFocus()
  Vencimento.Text = Ajusta_Data(Vencimento.Text)
End Sub

Private Function GetOrigemDinheiro() As String
  '27/04/2005 - Daniel
  'Solicitante: Bem Me Quer
  GetOrigemDinheiro = ""
  
  If O_Não_determinado.Value Then GetOrigemDinheiro = "Não determinado origem"
  If O_Caixa_C.Value Then GetOrigemDinheiro = Left("Caixa - cheque" & " (" & Nome_Caixa.Caption & ")", 30)
  If O_caixa_d.Value Then GetOrigemDinheiro = Left("Caixa - dinheiro" & " (" & Nome_Caixa.Caption & ")", 30)
  If O_Conta.Value Then GetOrigemDinheiro = Left("Conta Corrente" & " (" & Nome_Conta.Caption & ")", 30)
End Function

' criado em: 22/06/2022
' autor: Pablo Verçosa Silva
' descrição: calcula os campos Desconto, Acréscimo e Valor_Pago conforme edição
Private Sub CalculaValores(ByVal pStart As String)
  Dim nValor As Double
  Dim nDesconto As Double
  Dim nAcrescimo As Double
  Dim nValor_Pago As Double
  
  nValor = IIf(IsNumeric(Valor.Text), IIf(CDbl(gsHandleNull(Valor.Text)) > 0, CDbl(gsHandleNull(Valor.Text)), 0), 0)
  nDesconto = IIf(IsNumeric(Desconto.Text), IIf(CDbl(gsHandleNull(Desconto.Text)) > 0, CDbl(gsHandleNull(Desconto.Text)), 0), 0)
  nAcrescimo = IIf(IsNumeric(Acréscimo.Text), IIf(CDbl(gsHandleNull(Acréscimo.Text)) > 0, CDbl(gsHandleNull(Acréscimo.Text)), 0), 0)
  nValor_Pago = IIf(IsNumeric(Valor_Pago.Text), IIf(CDbl(gsHandleNull(Valor_Pago.Text)) > 0, CDbl(gsHandleNull(Valor_Pago.Text)), 0), 0)
  
  If StrComp(pStart, Desconto.Name, 1) = 0 Or StrComp(pStart, Acréscimo.Name, 1) = 0 Then nValor_Pago = nValor - nDesconto + nAcrescimo
  If StrComp(pStart, Valor_Pago.Name, 1) = 0 Then
    If nValor > nValor_Pago Then
      nDesconto = nValor - nValor_Pago
      nAcrescimo = 0
    ElseIf nValor < nValor_Pago Then
      nDesconto = 0
      nAcrescimo = nValor_Pago - nValor
    ElseIf nValor = nValor_Pago Then
      nDesconto = 0
      nAcrescimo = 0
    End If
  End If
  
  Desconto.Text = Format(nDesconto, "##,###,##0.00")
  Acréscimo.Text = Format(nAcrescimo, "##,###,##0.00")
  Valor_Pago.Text = Format(nValor_Pago, "##,###,##0.00")
End Sub

