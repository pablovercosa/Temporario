VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmManContasReceber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Realizar baixa de Contas a Receber"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   375
   ClientWidth     =   16965
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
   Icon            =   "ManContasReceber.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8265
   ScaleWidth      =   16965
   Begin VB.CommandButton cmd_acataUsuarioLogadoComoOperador 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6000
      Picture         =   "ManContasReceber.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Libera para alterar a data de pagamento"
      Top             =   6960
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Frame Frame5 
      Caption         =   "Localizar por Código de Barras"
      Height          =   705
      Left            =   0
      TabIndex        =   90
      Top             =   480
      Width           =   2805
      Begin VB.TextBox txtCodigoBarras 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   90
         TabIndex        =   91
         Top             =   240
         Width           =   2625
      End
   End
   Begin VB.CommandButton cmd_imprimeCarne 
      BackColor       =   &H00C0FFFF&
      Caption         =   "imprime Carne"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   14490
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   4890
      Width           =   2475
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2730
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   120
      TabIndex        =   88
      Text            =   "Valor Total"
      Top             =   4980
      Width           =   885
   End
   Begin VB.TextBox txt_valorTotalGrade 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1050
      TabIndex        =   87
      Top             =   4920
      Width           =   1245
   End
   Begin VB.TextBox txt_lembreteCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   13560
      MultiLine       =   -1  'True
      TabIndex        =   86
      Top             =   450
      Width           =   3375
   End
   Begin VB.Frame frm_troco 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fazer troco?"
      Height          =   2895
      Left            =   8910
      TabIndex        =   79
      Top             =   5310
      Visible         =   0   'False
      Width           =   1515
      Begin VB.TextBox txt_troco 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   150
         TabIndex        =   83
         Top             =   2070
         Width           =   1155
      End
      Begin VB.CommandButton cmd_troco 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ver troco"
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1380
         Width           =   1215
      End
      Begin VB.TextBox txt_cliPagou 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   150
         TabIndex        =   80
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Troco de..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   285
         TabIndex        =   84
         Top             =   1770
         Width           =   1005
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cliente pagou..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   81
         Top             =   390
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipos de Parcelamento"
      Height          =   810
      Left            =   11400
      TabIndex        =   71
      Top             =   420
      Width           =   2130
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1110
         TabIndex        =   75
         Top             =   510
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton O_Banco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Banco"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   210
         Width           =   795
      End
      Begin VB.OptionButton O_Carteira 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Carteira"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1110
         TabIndex        =   73
         Top             =   210
         Width           =   945
      End
      Begin VB.OptionButton O_Carnet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Carnet"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   510
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      ForeColor       =   &H00FF0000&
      Height          =   810
      Left            =   7560
      TabIndex        =   68
      Top             =   420
      Width           =   1425
      Begin VB.OptionButton O_Vencimento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Vencimento"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   70
         Top             =   240
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton O_Cliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Cliente"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   69
         Top             =   510
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipos de Contas"
      Height          =   795
      Left            =   9030
      TabIndex        =   64
      Top             =   420
      Width           =   2340
      Begin VB.OptionButton O_Todas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1470
         TabIndex        =   67
         Top             =   210
         Width           =   795
      End
      Begin VB.OptionButton O_Recebidas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Recebidas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   66
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton O_Receber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "A R&eceber"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   65
         Top             =   210
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txtQtdeImprimir 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6210
      TabIndex        =   16
      Text            =   "1"
      Top             =   6525
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdDesconto 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Calc. Descon&to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7350
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.TextBox txtSeq 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3345
      TabIndex        =   15
      Top             =   6525
      Visible         =   0   'False
      Width           =   855
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
      Left            =   -1410
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   1110
      Visible         =   0   'False
      Width           =   1575
   End
   Begin SSDataWidgets_B.SSDBGrid grdCR 
      Height          =   3135
      Left            =   30
      TabIndex        =   6
      Top             =   1710
      Width           =   16935
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
      Col.Count       =   17
      AllowUpdate     =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      BackColorOdd    =   12648384
      RowHeight       =   423
      ExtraHeight     =   185
      Columns.Count   =   17
      Columns(0).Width=   767
      Columns(0).Caption=   "Filial"
      Columns(0).Name =   "Filial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
      Columns(1).Caption=   "Valor"
      Columns(1).Name =   "Valor"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1720
      Columns(2).Caption=   "Vcto"
      Columns(2).Name =   "Vcto"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   10
      Columns(2).Mask =   "##/##/####"
      Columns(2).PromptInclude=   -1  'True
      Columns(2).PromptChar=   32
      Columns(3).Width=   1455
      Columns(3).Caption=   "Desc"
      Columns(3).Name =   "Desc"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1376
      Columns(4).Caption=   "Acresc"
      Columns(4).Name =   "Acresc"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1905
      Columns(5).Caption=   "Val Receb"
      Columns(5).Name =   "Val Receb"
      Columns(5).Alignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1799
      Columns(6).Caption=   "Data Receb"
      Columns(6).Name =   "Data Receb"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   10
      Columns(6).Mask =   "##/##/####"
      Columns(6).PromptInclude=   -1  'True
      Columns(6).PromptChar=   32
      Columns(7).Width=   1640
      Columns(7).Caption=   "Nota"
      Columns(7).Name =   "Nota"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1508
      Columns(8).Caption=   "Cliente"
      Columns(8).Name =   "Cliente"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3016
      Columns(9).Caption=   "Nome"
      Columns(9).Name =   "Nome"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1191
      Columns(10).Caption=   "Seq"
      Columns(10).Name=   "Seq"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   4022
      Columns(11).Caption=   "Descrição"
      Columns(11).Name=   "Descricao"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   1535
      Columns(12).Caption=   "ID"
      Columns(12).Name=   "Contador"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "Tipo Parcelamento"
      Columns(13).Name=   "Tipo Parcelamento"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   1614
      Columns(14).Caption=   "DiasAtraso"
      Columns(14).Name=   "DiasAtraso"
      Columns(14).Alignment=   2
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   1931
      Columns(15).Caption=   "ValAtraso"
      Columns(15).Name=   "ValAtraso"
      Columns(15).Alignment=   1
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   1640
      Columns(16).Caption=   "Pendência"
      Columns(16).Name=   "Pendencia"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      _ExtentX        =   29871
      _ExtentY        =   5530
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
   Begin VB.Frame Frame3 
      Caption         =   "Vencimento"
      Height          =   750
      Left            =   2850
      TabIndex        =   52
      Top             =   450
      Width           =   4665
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
         Height          =   420
         Left            =   4110
         Picture         =   "ManContasReceber.frx":4EEE4
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   210
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
         Height          =   420
         Left            =   1830
         Picture         =   "ManContasReceber.frx":4F7C6
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   210
         Width           =   465
      End
      Begin MSMask.MaskEdBox Vcto_Final 
         Height          =   315
         Left            =   2850
         TabIndex        =   1
         Top             =   270
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vcto_Inicial 
         Height          =   315
         Left            =   570
         TabIndex        =   0
         Top             =   270
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Final"
         Height          =   255
         Left            =   2445
         TabIndex        =   54
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Inicial"
         Height          =   255
         Left            =   90
         TabIndex        =   53
         Top             =   300
         Width           =   585
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
      Left            =   6300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Caixas"
      Top             =   8550
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
      Left            =   4470
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   8550
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
      Left            =   3270
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   8550
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton B_Baixa 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Alterar / Baixar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2370
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4890
      Width           =   11160
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
      Height          =   420
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Width           =   16920
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "ManContasReceber.frx":500A8
      DataSource      =   "Data1"
      Height          =   330
      Left            =   8190
      TabIndex        =   3
      Top             =   75
      Width           =   1170
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
      BackColorOdd    =   16777152
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
      _ExtentX        =   2064
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   330
      Left            =   1035
      TabIndex        =   7
      Top             =   5730
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Data_Pagto 
      Height          =   330
      Left            =   7545
      TabIndex        =   13
      Top             =   6945
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   330
      Left            =   1035
      TabIndex        =   8
      Top             =   6090
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
      Left            =   1035
      TabIndex        =   10
      Top             =   6450
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
      Left            =   1035
      TabIndex        =   11
      Top             =   6945
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
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
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Valor_Pago 
      Height          =   330
      Left            =   4185
      TabIndex        =   12
      Top             =   6915
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      ForeColor       =   0
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
   Begin VB.CommandButton B_Dia 
      BackColor       =   &H00C0FFFF&
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
      Height          =   400
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7350
      Visible         =   0   'False
      Width           =   2835
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
      Height          =   400
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7800
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Descrição 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3870
      MaxLength       =   30
      TabIndex        =   9
      Top             =   5745
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.TextBox Nota 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   7725
      MaxLength       =   15
      TabIndex        =   14
      Top             =   6525
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame_Modo 
      BackColor       =   &H00FFA324&
      Caption         =   "Forma de recebimento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   10470
      TabIndex        =   42
      Top             =   5295
      Visible         =   0   'False
      Width           =   6435
      Begin VB.OptionButton O_Caixa_Cartao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Cartão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   150
         TabIndex        =   78
         Top             =   780
         Width           =   1785
      End
      Begin VB.ComboBox cboTipoEmiss 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "ManContasReceber.frx":500BC
         Left            =   2790
         List            =   "ManContasReceber.frx":500BE
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2010
         Width           =   1200
      End
      Begin VB.CommandButton cmdEmiss 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Imprimir Tipo..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2430
         Width           =   6210
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Caixa 
         Bindings        =   "ManContasReceber.frx":500C0
         DataSource      =   "Data5"
         Height          =   315
         Left            =   2835
         TabIndex        =   25
         Top             =   705
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
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
         Enabled         =   0   'False
      End
      Begin VB.TextBox Num_Cheque 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5145
         MaxLength       =   10
         TabIndex        =   29
         Top             =   2025
         Width           =   1170
      End
      Begin MSMask.MaskEdBox Cheque_Bom 
         Height          =   315
         Left            =   1095
         TabIndex        =   28
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   2025
         Width           =   1170
         _ExtentX        =   2064
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton O_Não_determinado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Indeterminado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   150
         TabIndex        =   22
         Top             =   1020
         Width           =   1905
      End
      Begin VB.OptionButton O_caixa_d 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Dinheiro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   150
         TabIndex        =   24
         Top             =   270
         Value           =   -1  'True
         Width           =   1965
      End
      Begin VB.OptionButton O_Conta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Conta Corrente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   26
         Top             =   1395
         Width           =   1845
      End
      Begin VB.OptionButton O_Caixa_C 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Cheque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   510
         Width           =   1875
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
         Bindings        =   "ManContasReceber.frx":500D4
         DataSource      =   "Data4"
         Height          =   315
         Left            =   150
         TabIndex        =   27
         Top             =   1650
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
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   5371
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2487
         Columns(1).Caption=   "Conta"
         Columns(1).Name =   "Conta"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Conta"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2037
         Columns(2).Caption=   "Código"
         Columns(2).Name =   "Código"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Código"
         Columns(2).DataType=   2
         Columns(2).FieldLen=   256
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
         Enabled         =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "Tipo"
         Height          =   195
         Left            =   2370
         TabIndex        =   57
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label Nome_Caixa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3630
         TabIndex        =   48
         Top             =   705
         Width           =   2655
      End
      Begin VB.Label Label_Caixa 
         BackColor       =   &H00FFA324&
         Caption         =   "Caixa"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2355
         TabIndex        =   47
         Top             =   750
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "Num. Cheque"
         Height          =   195
         Left            =   4050
         TabIndex        =   46
         Top             =   2070
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "Bom para"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   2070
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   120
         X2              =   6240
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label Nome_conta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   900
         TabIndex        =   43
         Top             =   1650
         Width           =   3075
      End
   End
   Begin VB.CommandButton B_Cancela 
      BackColor       =   &H00FFFFFF&
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
      Height          =   400
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7800
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton B_Calc_Juros 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Calcular &Juros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7350
      Visible         =   0   'False
      Width           =   2835
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "ManContasReceber.frx":500E8
      DataSource      =   "datFilial"
      Height          =   330
      Left            =   750
      TabIndex        =   2
      Top             =   82
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
      BackColorOdd    =   16777152
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
      _ExtentX        =   1852
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Label10 
      Caption         =   "Evento pendente"
      Height          =   225
      Left            =   15660
      TabIndex        =   85
      Top             =   180
      Width           =   1245
   End
   Begin VB.Label lblQtde 
      AutoSize        =   -1  'True
      Caption         =   "Qtde para Impressão"
      Height          =   195
      Left            =   4650
      TabIndex        =   63
      Top             =   6555
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label lblSeq 
      AutoSize        =   -1  'True
      Caption         =   "Sequência"
      Height          =   195
      Left            =   2580
      TabIndex        =   62
      Top             =   6555
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      Height          =   195
      Left            =   90
      TabIndex        =   61
      Top             =   135
      Width           =   300
   End
   Begin VB.Label lblFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1830
      TabIndex        =   60
      Top             =   75
      Width           =   5700
   End
   Begin VB.Label Sequência 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2100
      TabIndex        =   59
      Top             =   5370
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Baixa 
      AutoSize        =   -1  'True
      Caption         =   "Sequência"
      Height          =   195
      Index           =   7
      Left            =   1230
      TabIndex        =   58
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Nome_Fornecedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9360
      TabIndex        =   56
      Top             =   75
      Width           =   5640
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente"
      Height          =   225
      Left            =   7590
      TabIndex        =   55
      Top             =   135
      Width           =   585
   End
   Begin VB.Label L_Tipo_Parc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "9"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   780
      TabIndex        =   51
      Top             =   5340
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label L_Descrição 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "8"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   420
      TabIndex        =   50
      Top             =   5340
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label L_Cliente 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "7"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      TabIndex        =   49
      Top             =   5340
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label_Vários 
      Alignment       =   2  'Center
      Caption         =   "Baixa de várias contas. Digite a data de recebimento. O valor recebido será assumido como o valor da conta."
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   2430
      TabIndex        =   45
      Top             =   6090
      Visible         =   0   'False
      Width           =   6405
   End
   Begin VB.Label Baixa 
      Caption         =   "Vencimento"
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   41
      Top             =   5760
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Baixa 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   8
      Left            =   3090
      TabIndex        =   40
      Top             =   5475
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3870
      TabIndex        =   39
      Top             =   5385
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Label Baixa 
      Caption         =   "Valor"
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   38
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Baixa 
      Caption         =   "Desconto"
      Height          =   255
      Index           =   2
      Left            =   75
      TabIndex        =   37
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Baixa 
      Caption         =   "Acréscimo"
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
      Index           =   3
      Left            =   75
      TabIndex        =   36
      Top             =   6990
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Baixa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFA324&
      Caption         =   "Valor Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   3045
      TabIndex        =   35
      Top             =   6960
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Baixa 
      Caption         =   "Nota"
      Height          =   225
      Index           =   6
      Left            =   7290
      TabIndex        =   34
      Top             =   6570
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Baixa 
      Caption         =   "Descrição"
      Height          =   255
      Index           =   9
      Left            =   3090
      TabIndex        =   33
      Top             =   5805
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Baixa 
      Caption         =   "Data Pagto"
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
      Index           =   5
      Left            =   6555
      TabIndex        =   32
      Top             =   6990
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFA324&
      Height          =   435
      Left            =   3000
      TabIndex        =   93
      Top             =   6870
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmManContasReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private numParcelaPaga(12) As Double
Private totalNumParcelaPaga As Integer


Private rsCliFor As Recordset
Private Rec_Contas As Recordset
Private rsCaixas As Recordset
Private rsCaixa As Recordset
Private rsContas_Receber As Recordset
Private rsLançamentos As Recordset
Private rsContas As Recordset
Private rsParametros As Recordset

'Variáveis que armazenarão o vencimento
'e o valor de [Contas a Receber]
Private m_dteVencimento As Date
Private m_dblValor As Double

'05/01/2004 - Daniel
'Alteração do juro (acréscimo) com a senha
'do gerente
'Case: F. Linhares
Private m_dblJurosAlterado As Double

'11/03/2004 - Daniel
'Ao invés de chamar a função CheckSerialCaseMod
'n vezes, foi criado esta variável para flag
'de verificação se é ou não F. Linhares
'verificamos já no form load
Private m_blnFLinhares As Boolean
'-------------------------------------------

Private Const GRID_FULL_HEIGHT As Single = 4700
Private Const GRID_MIN_HEIGHT As Single = 1775

'12/12/2003 - mpdea
'Flag indicando a solicitação de confirmação de impressão
'do comprovante no processo de baixa
'14/01/2004 - Daniel, Alterada para Public
Public g_blnConfirmarImpressao As Boolean
'29/06/2004 - Daniel
'Esta var será utilizada caso haja baixa parcial, nela estará a
'data do real vencimento como sugestão para o usuário confirmar ou não
Private m_datVencimento As Date
'17/08/2004 - Daniel
'Variável criada para monitorar impressões de
'boletos diretamente evitando n clicks
'Case Modelo: De Mais Presentes (Loja do Nazareno - RJ)
'Aberto também para o F. Linhares
Private m_blnImprimirDireto As Boolean
'27/08/2004 - Daniel
'Nesta var modular será guardada o boleto default para
'o cliente De Mais Presentes (Loja do Nazareno - RJ)
Private m_strBoletoDefault  As String
'22/04/2005 - Daniel
'Otimizado rotina para abrir a tela de lançamentos de contas
'com a conta desejada a partir do duplo click
'
'Solicitante: Consultor Carlos (Petrópolis - RJ)
Public g_blnFind  As Boolean
Public g_strQuery As String

Private Sub Arruma_Caixa()
  Dim Sem_Caixa As Boolean
  Dim Ordem As Long
  Dim Tot_Dinheiro As Double
  Dim Tot_Cheques As Double
  Dim Tot_Pré As Double
  Dim Tot_Cartões As Double
  Dim Tot_Vales As Double
  Dim Saldo_Ant As Double
  Dim Tot_Parcela As Double

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
        With rsCaixa
          .AddNew
          .Fields("Filial") = gnCodFilial
          .Fields("Data") = Data_Atual
          .Fields("Caixa") = Val(Combo_Caixa.Text)
          .Fields("Ordem") = 1
          .Fields("Descrição") = "Início do dia"
          .Fields("Hora") = Format(Time, "hh:mm:ss")
          .Update
        End With
     End If
     
     If Sem_Caixa = False Then  'Pegar caixa de outro dia
       rsCaixa.Index = "Data"
       rsCaixa.Seek "<", gnCodFilial, Val(Combo_Caixa.Text), Data_Atual, 9999#
       If rsCaixa.NoMatch Then
         Exit Sub
       End If
     
       If rsCaixa("Filial") <> gnCodFilial Then Exit Sub
   '    If .Fields("Data") <> Data_Atual Then Exit Sub
  
       Ordem = 1
  
       With rsCaixa
        Tot_Dinheiro = .Fields("Total Dinheiro")
        Tot_Cheques = .Fields("Total Cheques")
        Tot_Pré = .Fields("Total Cheques Pré")
        Tot_Cartões = .Fields("Total Cartões")
        Tot_Vales = .Fields("Total Vales")
        Tot_Parcela = .Fields("Total Parcelamento")
        Saldo_Ant = .Fields("Final")
        .AddNew
        .Fields("Filial") = gnCodFilial
        .Fields("Caixa") = Val(Combo_Caixa.Text)
        .Fields("Data") = Data_Atual
        .Fields("Hora") = Format(Time, "hh:mm:ss")
        .Fields("Ordem") = 1
        .Fields("Descrição") = "Início do dia"
        .Fields("Dinheiro") = Tot_Dinheiro
        .Fields("Total Dinheiro") = Tot_Dinheiro
        .Fields("Cheques") = Tot_Cheques
        .Fields("Total Cheques") = Tot_Cheques
        .Fields("Cheques Pré") = Tot_Pré
        .Fields("Total Cheques Pré") = Tot_Pré
        .Fields("Cartões") = Tot_Cartões
        .Fields("Total Cartões") = Tot_Cartões
        .Fields("Vales") = Tot_Vales
        .Fields("Total Vales") = Tot_Vales
        .Fields("Parcelamento") = Tot_Parcela
        .Fields("Total Parcelamento") = Tot_Parcela
        .Fields("Saldo Anterior") = Saldo_Ant
        .Fields("Final") = Saldo_Ant
        .Update
      End With
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
  
  With rsCaixa
    Tot_Dinheiro = .Fields("Total Dinheiro")
    Tot_Cheques = .Fields("Total Cheques")
    Tot_Pré = .Fields("Total Cheques Pré")
    Tot_Cartões = .Fields("Total Cartões")
    Tot_Vales = .Fields("Total Vales")
    Tot_Parcela = .Fields("Total Parcelamento")
    Saldo_Ant = .Fields("Final")
    .AddNew
    .Fields("Filial") = gnCodFilial
    .Fields("Data") = Data_Atual
    
    '06/05/2004 - mpdea
    'Incluído Hora e Funcionário
    .Fields("Hora").Value = Format(Time, "hh:mm:ss")
    .Fields("Funcionário").Value = gnUserCode
    
    .Fields("Caixa") = Val(Combo_Caixa.Text)
    .Fields("Ordem") = Ordem
    .Fields("Descrição") = Left("Conta recebida - " + Nome_Cliente.Caption, 55)
      
    .Fields("Dinheiro") = 0
    If O_caixa_d.Value = True Then .Fields("Dinheiro") = CDbl(Valor_Pago.Text)
    .Fields("Total Dinheiro") = Tot_Dinheiro + .Fields("Dinheiro")
            
    .Fields("Cheques") = 0
    If O_Caixa_C.Value = True Then .Fields("Cheques") = CDbl(Valor_Pago.Text)
    .Fields("Total Cheques") = Tot_Cheques + .Fields("Cheques")
      
    .Fields("Cheques Pré") = 0
   ' If O_Caixa_P.Value = True Then .Fields("Cheques Pré") = -CDbl(Valor_Pago.Text)
    .Fields("Total Cheques Pré") = Tot_Pré
      
    .Fields("Cartões") = 0
    If O_Caixa_Cartao.Value = True Then
        .Fields("Cartões") = CDbl(Valor_Pago.Text)
    End If
    .Fields("Total Cartões") = Tot_Cartões + .Fields("Cartões")
    
    .Fields("Vales") = 0
    .Fields("Total Vales") = Tot_Vales
    .Fields("Parcelamento") = 0
    .Fields("Total Parcelamento") = 0
    .Fields("Saldo Anterior") = Saldo_Ant
    .Fields("Final") = Saldo_Ant + .Fields("Dinheiro") + .Fields("Cheques")
    
    '16/02/2019 - Fevereiro/2019 ----- Inclui '+ .Fields("Cartões")' na linha abaixo
    .Fields("Final") = .Fields("Final") + .Fields("Cheques Pré") + .Fields("Cartões")
    .Update
  End With

End Sub

'13/02/2003 - mpdea
'Código revisado
Private Sub Arruma_Conta()
  Dim Saldo_Ant As Double
  
  'Atualiza Dinheiro na conta, se for o caso
  If CDbl(Valor_Pago.Text) <> 0 And O_Conta.Value Then
    With rsLançamentos
      .Index = "Conta"
      .Seek "<", Val(Combo_Conta.Text), CDate(Data_Atual), 99999999#
      Saldo_Ant = 0
      If Not .NoMatch Then
        If .Fields("Conta").Value = Val(Combo_Conta.Text) Then
          Saldo_Ant = .Fields("Saldo Atual").Value
        End If
      End If
    
      .AddNew
      .Fields("Conta").Value = Val(Combo_Conta.Text)
      .Fields("Data").Value = Cheque_Bom.Text
      .Fields("Descrição").Value = Left("Conta recebida - " + Nome_Cliente.Caption, 40)
      .Fields("Saldo Anterior").Value = Saldo_Ant
      .Fields("Crédito").Value = CDbl(Valor_Pago.Text)
      .Fields("Cheque").Value = Num_Cheque.Text
      .Fields("Saldo Atual").Value = Saldo_Ant + CDbl(Valor_Pago.Text)
      .Update
    End With
  End If

End Sub

Private Sub Acréscimo_LostFocus()
On Error GoTo Erro
  '30/05/2005 - Daniel
  'Cálculo do Valor Pago automático
  'Solicitante: Pedágio
  Dim dblAcrescimo As Double
  
  dblAcrescimo = Format(CDbl(0 & (Acréscimo.Text)), FORMAT_VALUE)
  
  Valor_Pago.Text = Format((Valor.Text) + dblAcrescimo - CDbl(0 & (Desconto.Text)), FORMAT_VALUE)

  Exit Sub
Erro:
  MsgBox "Inconsistência " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub B_Baixa_Click()
 Dim Linha As Long
 Dim i As Integer
 Dim Pagas As Integer
 Dim book As Variant
 Dim Valor_Contas As Double
 
 Call StatusMsg("")
 
 If grdCR.SelBookmarks.Count < 1 Then
   DisplayMsg "Selecione um registro na grade."
   Exit Sub
 End If
 
 Pagas = False
 Valor_Contas = 0
 For i = 0 To (grdCR.SelBookmarks.Count - 1)
   book = grdCR.SelBookmarks(i)
   If grdCR.Columns("Val Receb").CellValue(book) <> 0 Then Pagas = True
   Valor_Contas = Valor_Contas + grdCR.Columns("Valor").CellValue(book)
 Next i
 
 '
 If Pagas = True Then
   DisplayMsg "Uma ou mais contas selecionadas já foram recebidas e não podem ser baixadas. Caso deseje use a tela de lançamentos para alterá-las."
   Exit Sub
 End If
 
 Call StatusMsg("Aguarde...")
 
  
  '12/12/2003 - mpdea
  'Solicitar confirmação na baixa
  g_blnConfirmarImpressao = True
 
 cmd_acataUsuarioLogadoComoOperador.Visible = True
 Vencimento.Visible = True
 Valor.Visible = True
 Desconto.Visible = True
 Acréscimo.Visible = True
 Valor_Pago.Visible = True
 Label7.Visible = True
 Data_Pagto.Visible = True
 '16/01/2004 - Daniel
 'Carregar a data atual
 Data_Pagto.Text = Format(Date, "dd/mm/yyyy")
 Nota.Visible = True
 '23/11/2004 - Daniel
 'Tratamento para o objeto txtQtdeImprimir
 'Case: Nazareno
 If m_blnImprimirDireto Then
  lblQtde.Visible = True
  txtQtdeImprimir.Visible = True
  txtQtdeImprimir.Text = "1"
 Else
  lblQtde.Visible = False
  txtQtdeImprimir.Visible = False
 End If
 
 '--------------------------------------------
 '13/01/2004 - Daniel
 '
 'Verifica alteração personalizada
 '
 'QS37818-990 = F. Linhares
 'Mostrar o número da sequência além
 'do número da nota fiscal
 '
''' txtSeq.Visible = m_blnFLinhares
 lblSeq.Visible = m_blnFLinhares
 '--------------------------------------------
   
 'Sequência.Visible = True
 Nome_Cliente.Visible = True
 
 '--------------------------------------------
 '05/01/2004 - Daniel
 '
 'Verifica alteração personalizada
 '
 'QS37818-990 = F. Linhares
 'Cálculo de Juros automático sem clicar no
 'B_Calc_Juros
 B_Calc_Juros.Visible = True
 '17/08/2004 - Daniel
 cmdDesconto.Visible = True
 
 If m_blnFLinhares Then
  B_Calc_Juros.Enabled = False
 Else
  B_Calc_Juros.Enabled = True
 End If
  
 
 Descrição.Visible = True
 B_Cancela.Visible = True
 '17/08/2004 - Daniel
 cmdDesconto.Visible = True
 B_Dia.Visible = True
 B_Confirma.Visible = True
 
 Cheque_Bom.Mask = ""
 Cheque_Bom.Text = ""
 Cheque_Bom.Mask = "##/##/####"
 Num_Cheque.Text = ""
 
 L_Cliente.Caption = ""
 L_Descrição.Caption = ""
 L_Tipo_Parc.Caption = ""
 
 
 If grdCR.SelBookmarks.Count <> 1 Then
   Valor.Text = Valor_Contas
   Valor.Enabled = False
   Vencimento.Enabled = False
   Desconto.Enabled = False
   Acréscimo.Enabled = False
   Valor_Pago.Text = Valor
   Valor_Pago.Enabled = False
   Nota.Enabled = False
   
   '13/01/2004 - Daniel
   txtSeq.Enabled = False
   lblSeq.Enabled = False
   '---------------------------------
   
   'Sequência.Enabled = False
   Nome_Cliente.Enabled = False
   Label_Vários.Visible = True
   Descrição.Enabled = False
   Nome_Cliente.Caption = ""
   Descrição.Text = ""
 Else
   Valor.Enabled = True
   Vencimento.Enabled = True
   Desconto.Enabled = True
   Acréscimo.Enabled = True
   Valor_Pago.Enabled = True
   
   '--------------------------------------------
   '13/01/2004 - Daniel
   '
   'Verifica alteração personalizada
   '
   'QS37818-990 = F. Linhares
   'Mostrar o número da sequência além
   'do número da nota fiscal
   '
   If m_blnFLinhares Then
       Nota.Enabled = False
       '''txtSeq.Enabled = False
       lblSeq.Visible = True
   Else
      Nota.Enabled = True
      '''txtSeq.Enabled = False
      lblSeq.Visible = False
   End If
   '--------------------------------------------
   
   'Sequência.Enabled = True
   Nome_Cliente.Enabled = True
   Label_Vários.Visible = False
   Descrição.Enabled = True
 End If
 
 
 For i = 0 To 9
   Baixa(i).Visible = True
 Next i
 Baixa(7).Visible = False
 
 If grdCR.SelBookmarks.Count = 1 Then
   book = grdCR.SelBookmarks(0)
   'Vencimento.Text = Format((grdCR.Columns("Vcto").CellValue(book)), "dd/mm/yyyy")
   Vencimento.Text = gsFormatDate(grdCR.Columns("Vcto").CellValue(book))
   Valor.Text = grdCR.Columns("Valor").CellValue(book)
   Desconto.Text = grdCR.Columns("Desc").CellValue(book)
   Acréscimo.Text = grdCR.Columns("Acresc").CellValue(book)
   
   '--------------------------------------------
   '13/01/2004 - Daniel
   '
   'Verifica alteração personalizada
   '
   'QS37818-990 = F. Linhares
   'Mostrar o número da sequência além
   'do número da nota fiscal
   '
'''   If m_blnFLinhares Then
       txtSeq.Text = grdCR.Columns("Seq").CellValue(book)
'''   End If
   '--------------------------------------------
   
   Nota.Text = grdCR.Columns("Nota").CellValue(book)
   Sequência.Caption = grdCR.Columns("Seq").CellValue(book)
   Nome_Cliente.Caption = grdCR.Columns("Nome").CellValue(book)
   Descrição.Text = grdCR.Columns("Descricao").CellValue(book)
   L_Cliente.Caption = grdCR.Columns("Cliente").CellValue(book)
   L_Descrição.Caption = grdCR.Columns("Descricao").CellValue(book)
   L_Tipo_Parc.Caption = grdCR.Columns("Tipo Parcelamento").CellValue(book)
   
 ElseIf grdCR.SelBookmarks.Count > 1 Then
    Nome_Cliente.Caption = grdCR.Columns("Nome").CellText(0)
    Descrição.Text = "Juntou " & grdCR.SelBookmarks.Count & " parcelas"
 End If
 
 B_Monta.Enabled = False
 B_Baixa.Enabled = False
 Frame_Modo.Visible = True
 grdCR.Enabled = False
 
 txt_cliPagou.Text = ""
 txt_troco.Text = ""
 frm_troco.Visible = True
 
 Call StatusMsg("")
 
  '25/09/2007 - Anderson
  'Otimizar o pagamento de parcelas através de código de barras no carnê
  'grdCR.Height = GRID_MIN_HEIGHT
''  grdCR.Top = IIf(g_bolCarneCodigoBarras, 2080, 1600)
''  grdCR.Height = IIf(g_bolCarneCodigoBarras, GRID_MIN_HEIGHT - 285, GRID_MIN_HEIGHT)

End Sub


Private Sub B_Calc_Juros_Click()
  Dim Valor_Aux As Double
  Dim Juros As Double
  Dim Erro As Integer
  Dim Dias As Integer
  Dim i As Integer
  Dim book As Variant
  Dim DiasMultaAux As Long
  Dim JurosMulta As Double
  Dim JurosTaxaMulta As Double
  
  If IsDate(Data_Pagto.Text) Then
     B_Calc_Juros.Enabled = False
  End If
  Call StatusMsg("")
  
  If Not IsDate(Vencimento.Text) And grdCR.SelBookmarks.Count = 1 Then
    DisplayMsg "Data de vencimento incorreta, verfique."
    If Vencimento.Enabled = False Then Exit Sub
    Vencimento.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Pagto.Text) Then
    DisplayMsg "Digite a data de recebimento, para que os juros sejam calculados."
    Data_Pagto.SetFocus
    Exit Sub
  End If
  
  
  Valor.Text = gsHandleNull(Valor.Text)
  If Not IsNumeric(Valor.Text) Then
    DisplayMsg "Valor incorreto, verifique."
    Valor.SetFocus
    Exit Sub
  End If
  
  Valor_Aux = CDbl(Valor.Text)
  
  
  '20/06/2003 - mpdea
  'Corrigido soma do acréscimo no cálculo dos juros
'  If IsNumeric(Acréscimo.Text) Then
'    Valor_Aux = Valor_Aux + CDbl(Acréscimo.Text)
'  End If
  
  If grdCR.SelBookmarks.Count = 1 Then
      Dias = CDate(Data_Pagto.Text) - CDate(Vencimento.Text)
      If Dias = 0 Then
         DisplayMsg "Recebimento em dia, sem juros a calcular."
         Exit Sub
      End If
  End If
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then Exit Sub
  
  If Not IsNull(rsParametros("TaxaMultaParcelaVencida")) Then
      JurosTaxaMulta = rsParametros("TaxaMultaParcelaVencida")
  End If
  
  If grdCR.SelBookmarks.Count > 1 Then
      Juros = 0
      'Calcular Juros para VARIAS parcelas
      For i = 0 To (grdCR.SelBookmarks.Count - 1)
          book = grdCR.SelBookmarks(i)
          Dias = CDate(Data_Pagto.Text) - CDate(grdCR.Columns("Vcto").CellValue(book))
          Valor_Aux = grdCR.Columns("Valor").CellValue(book)
      
          ' Verificar se aplica multa
          
          If Not IsNull(rsParametros("CobrarMultaAposVencimentoParcela")) Then
              If rsParametros("CobrarMultaAposVencimentoParcela") = True Then
                  DiasMultaAux = rsParametros("MultaDiasAposParcelaVencida")
                  If Dias > DiasMultaAux Then
                      JurosMulta = (Valor_Aux * JurosTaxaMulta) / CDbl(100)
                  Else
                      JurosMulta = 0
                  End If
              End If
          End If
      
          If Dias > 0 Then
              Juros = Juros + JurosMulta + (Valor_Aux * rsParametros("Juros") / CDbl(30) * CDbl(Dias) / CDbl(100))
          End If
      Next i
      Juros = Format(Juros, "#############.00")
      Acréscimo.Text = Juros
  Else
      'Calcular Juros para apenas UMA parcela
          
      If Not IsNull(rsParametros("CobrarMultaAposVencimentoParcela")) Then
          If rsParametros("CobrarMultaAposVencimentoParcela") = True Then
              DiasMultaAux = rsParametros("MultaDiasAposParcelaVencida")
              If Dias > DiasMultaAux Then
                  JurosMulta = (Valor_Aux * JurosTaxaMulta) / CDbl(100)
              Else
                  JurosMulta = 0
              End If
          End If
      End If
      
      If Dias > 0 Then
          Juros = JurosMulta + (Valor_Aux * rsParametros("Juros") / CDbl(30) * CDbl(Dias) / CDbl(100))
      End If
      Juros = Format(Juros, "#############.00")
      Acréscimo.Text = Juros
  End If
  
  '--------------------------------------------
  '05/01/2004 - Daniel
  '
  'Verifica alteração personalizada
  '
  'QS37818-990 = F. Linhares
  'Solicitar a senha do gerente para alterar
  'o juro (acréscimo)
  If m_blnFLinhares Then
    m_dblJurosAlterado = Juros
  End If
  
  Baixa_DblClick (4)
  
  Call StatusMsg("")
  
End Sub

Private Sub Calc_Juros_NaGrid(dJuros As Double, valorParcela As String, sDataVencimento As String, ByRef totalDiasAtrasado As String, ByRef valorTotalDiasAtrasado As String)
On Error GoTo Erro
  Dim Valor_Aux As Double
  Dim Juros As Double
  Dim Erro As Integer
  Dim Dias As Integer
  Dim sDataAtual As String
  Dim JurosTaxaMulta As Double
  Dim JurosMulta As Double
  Dim DiasMultaAux As Long
  
  Valor_Aux = CDbl(FormatNumber(valorParcela, 2))
  
  sDataAtual = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  Dias = CDate(sDataAtual) - CDate(sDataVencimento)
  'Dias = CDate(Now) - CDate(sDataVencimento)
  
  If Dias <= 0 Then
      totalDiasAtrasado = 0
      valorTotalDiasAtrasado = "0.00"
      Exit Sub
  Else
      totalDiasAtrasado = Dias
  End If

  ' Verificar se aplica multa
  If Not IsNull(rsParametros("TaxaMultaParcelaVencida")) Then
      JurosTaxaMulta = rsParametros("TaxaMultaParcelaVencida")
  End If

  JurosMulta = 0
  If Not IsNull(rsParametros("CobrarMultaAposVencimentoParcela")) Then
      If rsParametros("CobrarMultaAposVencimentoParcela") = True Then
            DiasMultaAux = rsParametros("MultaDiasAposParcelaVencida")
            If Dias > DiasMultaAux Then
                JurosMulta = (Valor_Aux * JurosTaxaMulta) / CDbl(100)
            Else
                JurosMulta = 0
            End If
      End If
  End If
  
  If Dias > 0 Then
      Juros = JurosMulta + (Valor_Aux * dJuros / CDbl(30) * CDbl(Dias) / CDbl(100))
  End If
  
  Juros = Format(Juros, "#############.00")
  
  valorTotalDiasAtrasado = Juros + Valor_Aux
  
  Exit Sub
Erro:
  MsgBox "Erro na função de calculo de Juros para os dias em atraso da parcela"

End Sub


Private Sub B_Cancela_Click()
  
  Dim i As Integer
  
  For i = 0 To 9
   Baixa(i).Visible = False
  Next i
  
  Data_Pagto.Enabled = False
  cmd_acataUsuarioLogadoComoOperador.Visible = False
  Vencimento.Visible = False
  Valor.Visible = False
  Desconto.Visible = False
  Acréscimo.Visible = False
  Valor_Pago.Visible = False
  Label7.Visible = False
  Data_Pagto.Visible = False
  
  '13/01/2004 - Daniel
  txtSeq.Visible = False
  lblSeq.Visible = False
  '-----------------------------
  
  Nota.Visible = False
  'Sequência.Visible = False
  Nome_Cliente.Visible = False
  Descrição.Visible = False
  Label_Vários.Visible = False
  B_Calc_Juros.Visible = False
  '23/11/2004 - Daniel
  'Tratamento para o objeto txtQtdeImprimir
  'Case: Nazareno
  If m_blnImprimirDireto Then
    txtQtdeImprimir.Visible = False
    lblQtde.Visible = False
  End If
  
  B_Cancela.Visible = False
  '17/08/2004 - Daniel
  cmdDesconto.Visible = False
  B_Dia.Visible = False
  B_Confirma.Visible = False
  
  Frame_Modo.Visible = False
  
  B_Monta.Enabled = True
  B_Baixa.Enabled = True
  
  Valor_Pago.Text = ""
  Data_Pagto.Mask = ""
  Data_Pagto.Text = ""
  Data_Pagto.Mask = "##/##/####"
  
  grdCR.Enabled = True
  
  frm_troco.Visible = False

  
  '25/09/2007 - Anderson
  'Otimizar o pagamento de parcelas através de código de barras no carnê
  'grdCR.Height = GRID_FULL_HEIGHT
''  grdCR.Top = IIf(g_bolCarneCodigoBarras, 2080, 1600)
''  grdCR.Height = IIf(g_bolCarneCodigoBarras, GRID_FULL_HEIGHT - 285, GRID_FULL_HEIGHT)
  
End Sub

'04/05/2004 - mpdea
'Incluído tratamento de erro
Private Sub B_Confirma_Click()
  Dim Resposta As Integer
  Dim Erro As Integer
  Dim i As Integer
  Dim book As Variant
  Dim Parcial As Boolean
  Dim Diferença As Double
  Dim Data_Prox As Variant
  Dim sData As String
  Dim nContador As Long
  Dim bOk As Boolean
  Dim nVendedor As Integer
  Dim rsCrVendedor As Recordset
  Dim sSql As String
  '13/01/2004 - Daniel
  'Mostrar Seq ao invés da nota
  'Case: F. Linhares
  Dim blnMostrarSeq As Boolean
  '15/01/2004 - mpdea
  Dim enuRetVbMsgBoxResult As VbMsgBoxResult
  
  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked As Integer
  Dim bEmiteCarnesNOVOS As Boolean
  Dim iQualCarne40Col As Integer
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")
  
  B_Confirma.Enabled = False

  '--------------------------------------------
  '29/12/2003 - Daniel
  '
  'Verifica alteração personalizada
  '
  'QS37818-990 = F. Linhares
  'Solicitar a senha do gerente para dar baixa
  'em uma parcela se a anterior não houver sido
  'quitada
  If m_blnFLinhares Then
  
    If m_blnExisteParcelaPendente Then
      If MsgBox("Há parcela(s) pendente(s), deseja continuar?", vbExclamation + vbYesNo, "Atenção") = vbYes Then
          If Not frmGerente.gbSenhaGerente Then
            B_Confirma.Enabled = True
            Exit Sub
          End If
      Else
          B_Confirma.Enabled = True
          Exit Sub
      End If
    End If
    
    '05/01/2004
    'Solicitar a senha do gerente para efetuar
    'baixas com datas ou valores diferentes dos
    'previstos
    If m_blnValoresDiffPrevisto Then
      If MsgBox("Data de Vencimento ou Valor(es) diferente(s) do previsto, deseja continuar?", vbExclamation + vbYesNo, "Atenção") = vbYes Then
        If Not frmGerente.gbSenhaGerente Then
            B_Confirma.Enabled = True
            Exit Sub
        End If
      Else
        B_Confirma.Enabled = True
        Exit Sub
      End If
    End If
    
'    '05/01/2004
'    'Solicitar a senha do gerente para alterar
'    'o juro (acréscimo)
'    '02/09/2004 - Daniel - Conforme solicitação do F. Linhares
'    'para valores acrescentados não precisamos solicitar a senha do gerente
'    If IsNumeric(Acréscimo.Text) Then
'
'      If Acréscimo.Text <> m_dblJurosAlterado Then
'        If MsgBox("Acréscimo diferente do previsto, deseja continuar?", vbExclamation + vbYesNo, "Atenção") = vbYes Then
'          If Not frmGerente.gbSenhaGerente Then Exit Sub
'        Else
'          Exit Sub
'        End If
'      End If
'
'    End If
    
    blnMostrarSeq = True 'Mostrará a Seq ao invés da nota
    
  End If 'If CheckSerialCaseMod
  '--------------------------------------------
  
  '14/01/2004 - Daniel
  'Tratamento da impressão do comprovante
  'Case F. Linhares
  If m_blnFLinhares Then
  
    Dim rstFuncionarios As Recordset
    
    Set rstFuncionarios = db.OpenRecordset(" SELECT ImprimirTicket FROM Funcionários WHERE Código=" & gnUserCode, dbOpenDynaset)
  
    With rstFuncionarios
      If Not (.BOF And .EOF) Then
        If .Fields("ImprimirTicket").Value = True Then
          g_blnConfirmarImpressao = True
        Else
          g_blnConfirmarImpressao = False
        End If
      End If
      .Close
    End With
    
    Set rstFuncionarios = Nothing
    
    '02/09/2004 - Daniel
    'Aproveitado a realidade do cliente De Mais Presentes (Nazareno - RJ)
    'para o F. Linhares
    If m_blnImprimirDireto Then
      If Not BoletoDefault Then
          B_Confirma.Enabled = True
          Exit Sub
      End If
      'Caso não saia, confirmaremos
      g_blnConfirmarImpressao = True
    End If
  
  Else 'Demais clientes
      
    '27/08/2004 - Daniel
    'Adicionado a realidade para o cliente De Mais Presentes (Nazareno - RJ)
    'que precisa de impressão do boleto automaticamente sem 'n' click's
    If m_blnImprimirDireto Then
      If Not BoletoDefault Then
          B_Confirma.Enabled = True
          Exit Sub
      End If
      'Caso não saia, confirmaremos
      g_blnConfirmarImpressao = True
    Else
      
      
      '
      Dim Resp As String

      Resp = InputBox("Se deseja imprimir comprovante, escolha o modelo:" & vbCrLf & vbCrLf & "     0 - NÃO QUERO COMPROVANTE" & vbCrLf & vbCrLf & "     1 - COMPROVANTE RELATÓRIO" & vbCrLf & vbCrLf & "     2 - COMPROVANTE CARNÊ (todas parcelas)" & vbCrLf & vbCrLf & "     3 - COMPROVANTE CARNÊ (só parcela paga agora)", "Qual o modelo de impressão?", "1")
      If Not IsNumeric(Resp) Then
        'DisplayMsg "Opção de impressão inválida!"
        Resp = "0"
      End If

      Dim strNomeArq As String
      g_blnConfirmarImpressao = False
      bEmiteCarnesNOVOS = False
      
      If Resp = "0" Then
          g_blnConfirmarImpressao = False
      ElseIf Resp = "1" Then
          g_blnConfirmarImpressao = True
      Else
          bEmiteCarnesNOVOS = True
          
          If Resp = "3" Then
              iQualCarne40Col = 3
          Else
              iQualCarne40Col = 2
          End If
      End If
      '
      
'''      enuRetVbMsgBoxResult = MsgBox("Deseja imprimir comprovante?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Atenção")
'''
'''      Select Case enuRetVbMsgBoxResult
'''        Case vbYes
'''          g_blnConfirmarImpressao = True
'''        Case vbNo
'''          g_blnConfirmarImpressao = False
'''        Case vbCancel
'''          Exit Sub
'''      End Select
    
    End If
  
  End If
  '--------------------------------------------
  
  
'  '12/12/2003 - mpdea
'  'Solicitar confirmação na baixa
'  If g_blnConfirmarImpressao Then
'    If MsgBox("O comprovante não foi impresso. Deseja continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção") = vbNo Then
'      Exit Sub
'    End If
'  End If
  
  
'  '17/07/2003 - Maikel
'  '     Agora pergunta para o usuário se ele quer que imprima o recibo, pois muitos usuários esqueciam de imprimir recibo e depois de confirmado o sistema limpa a tela
'  If MsgBox("Deseja imprimir recibo ?", vbQuestion + vbYesNo, "Quick Store") = vbYes Then
'    cboTipoEmiss.Text = "Recibo"
'    cmdEmiss_Click
'  End If
  
  
  Valor_Pago.Text = gsHandleNull(Valor_Pago.Text)
  
  If CDbl(Valor_Pago.Text) = 0 And grdCR.SelBookmarks.Count = 1 Then
     Data_Pagto.Mask = ""
     Data_Pagto.Text = ""
     Data_Pagto.Mask = "##/##/####"
     GoTo Cont1
  End If
  

  If O_Não_determinado.Value = True Then
    Resposta = MsgBox("Deseja baixar a(s) conta(s) sem indicar o destino do dinheiro ? ", vbQuestion + vbYesNo)
    If Resposta = vbNo Then
'      DisplayMsg "Contas não baixadas."
      B_Confirma.Enabled = True
      Exit Sub
    End If
  Else
    If Nome_Caixa.Caption = "" And O_Conta.Value = False Then
      DisplayMsg "Escolha o caixa de destino do dinheiro / cheque."
      If Combo_Caixa.Enabled = True Then Combo_Caixa.SetFocus
      B_Confirma.Enabled = True
      Exit Sub
    End If
  End If
  
    
  If O_Conta.Value = True Then
    If Nome_Conta.Caption = "" Then
      DisplayMsg "Informe a conta."
      Combo_Conta.SetFocus
      B_Confirma.Enabled = True
      Exit Sub
    End If
    
    If Cheque_Bom.Text = "  /  /    " Then
        Cheque_Bom.Text = Data_Atual
    End If
    
    If Not IsDate(Cheque_Bom.Text) Then
      DisplayMsg "Data do cheque inválida, verifique."
      Cheque_Bom.SetFocus
      B_Confirma.Enabled = True
      Exit Sub
    End If
    If IsNull(Num_Cheque.Text) Then Num_Cheque.Text = ""
  End If
  
  If O_Caixa_C.Value = True Then
    If Not IsDate(Cheque_Bom.Text) Then
        DisplayMsg "Data do cheque inválida, verifique."
        Cheque_Bom.SetFocus
        B_Confirma.Enabled = True
        Exit Sub
    End If
    If IsNull(Num_Cheque.Text) Then Num_Cheque.Text = ""
  End If
    
  If IsNull(Acréscimo.Text) Or Acréscimo.Text = "" Then Acréscimo.Text = 0
  If IsNull(Desconto.Text) Or Desconto.Text = "" Then Desconto.Text = 0
  
  Parcial = False
  If Label_Vários.Visible = False Then
    Diferença = CDbl(Valor.Text) + CDbl(Acréscimo.Text) - CDbl(Desconto.Text)
    If Format((Diferença - CDbl(Valor_Pago.Text)), "#########0.00") < 0 Then
      DisplayMsg "Valor pago incorreto, verifique."
      B_Confirma.Enabled = True
      Exit Sub
    End If
    
    Diferença = CDbl(Valor.Text) + CDbl(Acréscimo.Text) - CDbl(Desconto.Text)
    
    If Abs(Diferença - CDbl(Valor_Pago.Text)) > 0.001 Then
      Resposta = MsgBox("O valor digitado é menor que o valor da conta. Deseja fazer uma baixa parcial ?", vbYesNo + vbQuestion)
      If Resposta = vbNo Then
        DisplayMsg "Conta não baixada."
        B_Confirma.Enabled = True
        Exit Sub
      Else '29/06/2004 - Daniel - Se a resposta for yes populo a var m_datVencimento
        m_datVencimento = CDate(Vencimento.Text)
      End If
      
      Data_Prox = InputBox("Digite ou confirme a data de vencimento para a diferença.", "Vencimento", m_datVencimento) '29/06/2004 - Daniel - Adicionado o terceiro parâmetro
      If Not IsDate(Data_Prox) Then
        DisplayMsg "Data incorreta, conta não baixada."
        B_Confirma.Enabled = True
        Exit Sub
      End If
      If IsDate(Vencimento.Text) Then
        If CDate(Data_Prox) < CDate(Vencimento.Text) Then
          DisplayMsg "Data deve ser igual ou posterior a data do Título original."
          B_Confirma.Enabled = True
          Exit Sub
        End If
      Else
        If CDate(Data_Prox) < Date Then
          DisplayMsg "Data deve ser igual ou posterior a data atual."
          B_Confirma.Enabled = True
          Exit Sub
        End If
      End If
      
      Diferença = CDbl(Valor.Text) - CDbl(Valor_Pago.Text) + CDbl(Acréscimo.Text) - CDbl(Desconto.Text)
      Acréscimo.Text = 0
      Desconto.Text = 0
      Valor.Text = Valor_Pago.Text
      Parcial = True
    End If
  End If
  
  Erro = False
  If Not IsDate(Data_Pagto.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data de pagamento incorreta, verifique."
    Data_Pagto.SetFocus
    B_Confirma.Enabled = True
    Exit Sub
  End If
  
  
  rsContas_Receber.Requery
  
  '04/05/2004 - mpdea
  'Tratamento de transação e status
  Call StatusMsg("Aguarde...")
  Screen.MousePointer = vbHourglass
  'Inicia transação
  ws.BeginTrans
  blnInTransaction = True
  
  
  If CDbl(Valor_Pago.Text) <> 0 And O_Não_determinado.Value = False And O_Conta.Value = False Then
    Call Arruma_Caixa
  End If
  
  Call Arruma_Conta
  
  totalNumParcelaPaga = 0
  
  If grdCR.SelBookmarks.Count > 1 Then
    bOk = False
    For i = 0 To (grdCR.SelBookmarks.Count - 1)
      book = grdCR.SelBookmarks(i)
      grdCR.Bookmark = book
      nContador = grdCR.Columns("Contador").CellValue(book)
      With rsContas_Receber
'        .Requery
        .FindFirst "Contador = " & nContador
        '04/05/2004 - mpdea
        'Incluído tratamento caso a conta não seja localizada
        If .NoMatch Then
          Call StatusMsg("")
          Screen.MousePointer = vbDefault
          MsgBox "Erro ao localizar a conta para baixa.", vbCritical, "Erro"
          ws.Rollback
          blnInTransaction = False
          B_Confirma.Enabled = True
          Exit Sub
        Else
          .LockEdits = True
          .Edit

          'verificar se foi cobrado com acrescimo/juros
          If CDbl(Acréscimo.Text) > 0 Then
              'If CDbl(Valor.Text) + CDbl(Acréscimo.Text) = CDbl(Valor_Pago.Text) Then
                  If CDbl(grdCR.Columns("ValAtraso").CellValue(book)) = 0 Then
                      ' Esta parcela não esta atrasada
                      ![Valor Recebido] = CDbl(grdCR.Columns("Valor").CellValue(book))
                  Else
                      ![Valor Recebido] = CDbl(grdCR.Columns("ValAtraso").CellValue(book))
                      ![Acréscimo] = CDbl(grdCR.Columns("ValAtraso").CellValue(book)) - CDbl(grdCR.Columns("Valor").CellValue(book))
                  End If
              'End If
          Else
              ![Valor Recebido] = !Valor
          End If
          
          '![Valor Recebido] = !Valor
          ![Data Recebimento] = Data_Pagto.Text
          ![Data Alteração] = Data_Atual
          .Update
          
          If totalNumParcelaPaga < 10 Then
              numParcelaPaga(totalNumParcelaPaga) = nContador
              totalNumParcelaPaga = totalNumParcelaPaga + 1
          End If
          bOk = True
        End If
      End With
    Next i
    
    '05/05/2004 - mpdea
    'Finaliza transação
    ws.CommitTrans
    blnInTransaction = False
    Screen.MousePointer = vbDefault
    If bOk = True Then
      '-----------------------------------
      '15/01/2004 - Daniel
      '..no Exit Sub encerrará então:
      '
      '27/08/2004 - Daniel
      'Adicionado a realidade para o cliente De Mais Presentes (Nazareno - RJ)
      'que precisa de impressão do boleto automaticamente sem 'n' click's
      If m_blnImprimirDireto Then
        frmManContasReceber.cmdEmiss_Click
      Else 'Continua para os demais
      
        If g_blnConfirmarImpressao Then
          frmImpressaoComprovante.Show vbModal
        End If
      
      End If
      '-----------------------------------
      
      B_Cancela_Click
      B_Monta_Click
    End If
    
    
    If totalNumParcelaPaga > 1 Then
      If bEmiteCarnesNOVOS = True Then
          EmiteCarnesNOVOS False, txtSeq.Text, iQualCarne40Col
      End If
    End If
    
    B_Confirma.Enabled = True
    Exit Sub
  End If
  
Cont1:
  
  '05/05/2004 - mpdea
  'Tratamento de transação e status
  If Not blnInTransaction Then
    Call StatusMsg("Aguarde...")
    Screen.MousePointer = vbHourglass
    'Inicia transação
    ws.BeginTrans
    blnInTransaction = True
  End If
  
  
  book = grdCR.SelBookmarks(0)
  grdCR.Bookmark = book
  nContador = grdCR.Columns("Contador").Value
'  grdCR.Columns("Vencimento").Value = CDate(Vencimento.Text)
'  grdCR.Columns("Valor").Value = CDbl(Valor.Text)
'  grdCR.Columns("Desconto").Value = CDbl(Desconto.Text)
'  grdCR.Columns("Acréscimo").Value = CDbl(Acréscimo.Text)
'  grdCR.Columns("Valor Recebido").Value = CDbl(Valor_Pago.Text)
'  If IsDate(Data_Pagto.Text) Then grdCR.Columns("Data Recebimento").Value = Data_Pagto.Text
'  grdCR.Columns("Nota").Value = Nota.Text
'  grdCR.Columns("Descrição").Value = Descrição.Text
'  grdCR.Update

  numParcelaPaga(0) = nContador
  totalNumParcelaPaga = 1
  
  With rsContas_Receber
'    .Requery
    .FindFirst "Contador = " & nContador
    '04/05/2004 - mpdea
    'Incluído tratamento caso a conta não seja localizada
    If .NoMatch Then
      Call StatusMsg("")
      Screen.MousePointer = vbDefault
      MsgBox "Erro ao localizar a conta para baixa.", vbCritical, "Erro"
      ws.Rollback
      blnInTransaction = False
      B_Confirma.Enabled = True
      Exit Sub
    Else
      .LockEdits = True
      .Edit
      .Fields("Vencimento") = CDate(Vencimento.Text)
      .Fields("Valor") = CDbl(gsHandleNull(Valor.Text))
      .Fields("Desconto").Value = CDbl(gsHandleNull(Desconto.Text))
      .Fields("Acréscimo").Value = CDbl(gsHandleNull(Acréscimo.Text))
      .Fields("Valor Recebido").Value = CDbl(gsHandleNull(Valor_Pago.Text))
      If IsDate(Data_Pagto.Text) Then
        .Fields("Data Recebimento").Value = Data_Pagto.Text
      End If
      
      
      '--------------------------------------------
      '13/01/2004 - Daniel
      '
      'Verifica alteração personalizada
      '
      'QS37818-990 = F. Linhares
      'Mostrar o número da sequência ao invés
      'do número da nota fiscal
      '
      If blnMostrarSeq Then
          If txtSeq.Text = "" Then
              txtSeq.Text = "0"
          End If
          
          .Fields("Sequência").Value = txtSeq.Text & ""
      Else
          If Nota.Text = "" Then
              Nota.Text = "0"
          End If
          
          .Fields("Nota").Value = Nota.Text & ""
      End If
      
      blnMostrarSeq = False
      '--------------------------------------------
      
      .Fields("Descrição") = Descrição.Text
      .Fields("Data Alteração") = Data_Atual
      .Update
      
      
      sSql = "SELECT Vendedor FROM [Contas a Receber] WHERE Contador = " & nContador
      Set rsCrVendedor = db.OpenRecordset(sSql, dbOpenDynaset)
      If rsCrVendedor.RecordCount <> 0 Then
         nVendedor = rsCrVendedor("Vendedor")
      Else
         nVendedor = 0
      End If
      rsCrVendedor.Close
      Set rsCrVendedor = Nothing
      If Parcial = True Then
        .AddNew
        .Fields("Filial") = grdCR.Columns("Filial").Value
        .Fields("Cliente") = L_Cliente.Caption
        .Fields("Sequência") = Sequência.Caption
        .Fields("Tipo") = "R"
        .Fields("Tipo Parcelamento") = L_Tipo_Parc.Caption
        .Fields("Descrição") = L_Descrição.Caption
        .Fields("Data Emissão") = Data_Atual
        .Fields("Vencimento") = CDate(Data_Prox)
        .Fields("Nota").Value = Nota.Text & ""
        .Fields("Valor") = Diferença
        .Fields("Vendedor") = nVendedor
        .Fields("Data Alteração") = Data_Atual
        '10/09/2007 - Anderson
        'Gera arquivo log do sistema
        If g_bolSystemLog Then
          SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
          "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequência") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
          "frmManContasReceber_B_Confirma_Click", _
          "Contas a Receber", g_strArquivoSystemLog
        End If
        .Update
       End If
    End If
  End With
  
  
  '04/05/2004 - mpdea
  'Finaliza transação
  ws.CommitTrans
  blnInTransaction = False
  
  Screen.MousePointer = vbDefault
  
  If bEmiteCarnesNOVOS = True Then
      EmiteCarnesNOVOS False, txtSeq.Text, iQualCarne40Col
  End If
  
  
  '14/01/2004 - Daniel
  '
  '27/08/2004 - Daniel
  'Adicionado a realidade para o cliente De Mais Presentes (Nazareno - RJ)
  'que precisa de impressão do boleto automaticamente sem 'n' click's
  If Not m_blnImprimirDireto Then 'Continua normal para os demais clientes
  
    If g_blnConfirmarImpressao Then
      frmImpressaoComprovante.Show vbModal
    End If
  
  Else 'Nazareno
    Call ImpressaoDiretaBoleto
  End If
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Filial:" & gnCodFilial & " Seq:" & Sequência.Caption & " Cli: " & L_Cliente.Caption & " VrOrig:" & Valor.Text & " VrPago:" & Valor_Pago.Text & " Dsc:" & Descrição.Text, 80) & "', 'CNT_REC: baixar-alt')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
  Valor_Pago.Text = ""
  Data_Pagto.Mask = ""
  Data_Pagto.Text = ""
  Data_Pagto.Mask = "##/##/####"
  
  B_Cancela_Click
  
  '26/09/2007 - Anderson
  'Executa o código de barras ao invés de pesquisar
  'B_Monta_Click
  If g_bolCarneCodigoBarras Then txtCodigoBarras_KeyPress (13) Else B_Monta_Click
  
  '25/09/2007 - Anderson
  'Otimizar o pagamento de parcelas através de código de barras no carnê
  'grdCR.Height = GRID_FULL_HEIGHT
''  grdCR.Top = IIf(g_bolCarneCodigoBarras, 2080, 1600)
''  grdCR.Height = IIf(g_bolCarneCodigoBarras, GRID_FULL_HEIGHT - 285, GRID_FULL_HEIGHT)
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
   
  B_Confirma.Enabled = True

  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
  Select Case Err.Number
    Case 3022  'Duplicação de chave primaria
      If intRepeatUpdateLocked < 30 Then
        Call StatusMsg("Verificando registros (" & Err.Number & ")...")
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
          B_Confirma.Enabled = True
          Exit Sub
        End If
      End If
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
          B_Confirma.Enabled = True
          Exit Sub
        End If
      End If
    Case Else
      'Cancelamento da transação
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      MsgBox "Erro em Manutenção - Contas a receber: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  End Select
End Sub

Private Sub B_Dia_Click()
  Valor_Pago.Text = Valor.Text
'''  Data_Pagto.Text = Vencimento.Text
  Data_Pagto.Text = Format(Now, "dd/mm/yyyy")
  Desconto.Text = 0
  Acréscimo.Text = 0
End Sub

Private Sub B_Monta_Click()
 
  Call StatusMsg("")
  
  grdCR.SelBookmarks.RemoveAll
  grdCR.Caption = "Contas"
  
  If Not IsDate(Vcto_Inicial.Text) Then
     DisplayMsg "Vencimento Inicial incorreto."
     Vcto_Inicial.SetFocus
     Exit Sub
  End If
  
  If Not IsDate(Vcto_Final.Text) Then
     DisplayMsg "Vencimento Final incorreto."
     Vcto_Final.SetFocus
     Exit Sub
  End If
    
  If CDate(Vcto_Final.Text) < CDate(Vcto_Inicial.Text) Then
    DisplayMsg "Vencimento inicial deve ser menor ou igual ao vencimento final."
    Vcto_Inicial.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(cboFilial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If
  
  B_Baixa.Enabled = True
  
  Call StatusMsg("Aguarde, pesquisando...")
  DoEvents
  
  Call LoadGridCR
  
  Call StatusMsg("")
 
End Sub

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Sub LoadGridCR()
  Dim rsCR As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim Aux_Erro As Integer
  Dim sDescricao As String
  Dim sUnidVenda As String
  Dim sCod As String
  Dim sSql As String
  Dim Data_Ini As String
  Dim Data_Fim As String
  Dim sFilial As String
  Dim sTipoParcelamento As String
  Dim dvalorTotalGrade As Double
  Dim sDescricaoNaGrade As String
  Dim sPendencia As String
  
  dvalorTotalGrade = 0
    
  On Error GoTo ErrHandler
  
  If Not IsDate(Vcto_Inicial.Text) Then
    DisplayMsg "Vencimento inicial incorreto."
    Vcto_Inicial.SetFocus
    Exit Sub
  End If
  Data_Ini = gsGetInvDate(Vcto_Inicial.Text)
  
  If Not IsDate(Vcto_Final.Text) Then
    DisplayMsg "Vencimento final incorreto."
    Vcto_Final.SetFocus
    Exit Sub
  End If
  Data_Fim = gsGetInvDate(Vcto_Final.Text)
  
  'Verifica a filial
  cboFilial_LostFocus
  If lblFilial.Caption = "" Then
    sFilial = "<> 0"
  Else
    sFilial = "= " & cboFilial
  End If
  
  bAllow = grdCR.AllowAddNew
  grdCR.AllowAddNew = True
  grdCR.AllowUpdate = True
  
  sSql = "SELECT Filial, Valor, Vencimento, [Contas a Receber].Desconto as Desconto, Acréscimo as Acrescimo, "
  sSql = sSql & " [Valor Recebido], [Data Recebimento] , Nota, Cliente, Cli_For.Nome, Sequência as Seq, "
  sSql = sSql & " Descrição as Descricao, Contador, [Tipo Parcelamento], Pendencia "
  sSql = sSql & " FROM [Contas a Receber] "
  sSql = sSql + " INNER JOIN Cli_For ON ([Contas a Receber].Cliente = Cli_For.Código)"
  sSql = sSql + " WHERE Filial " & sFilial & " AND Vencimento >= " + Data_Ini
  sSql = sSql + " And Vencimento <= " + Data_Fim + " AND [Contas a Receber].Tipo = 'R'"
  
  '15/01/2008 - Celso
  'Filtrar na pesquisa por tipo de parcelamento
  sTipoParcelamento = ""
  If O_Banco = True Then sTipoParcelamento = "B"
  If O_Carteira = True Then sTipoParcelamento = "C"
  If O_Carnet = True Then sTipoParcelamento = "T"
  If sTipoParcelamento <> "" Then
     sSql = sSql + " And [Tipo Parcelamento] = '" + sTipoParcelamento + "'"
  End If
    
  If Nome_Fornecedor.Caption <> "" Then
    sSql = sSql + " And Cliente = " + Combo_Fornecedor.Text
  End If
  
  If O_Receber = True Then sSql = sSql + " AND [Valor Recebido] = 0"
  If O_Recebidas = True Then sSql = sSql + " AND [Valor Recebido] <> 0"
  
  If O_Cliente.Value = True Then sSql = sSql + " ORDER BY Cliente"
  If O_Vencimento.Value = True Then sSql = sSql + " ORDER BY Vencimento"
  
  Set rsCR = db.OpenRecordset(sSql, dbOpenDynaset)

  grdCR.RemoveAll
  grdCR.Redraw = False
  
  grdCR.Columns("Valor").NumberFormat = "##,###,##0.00"
  grdCR.Columns("Val Receb").NumberFormat = "##,###,##0.00"

  Dim totalDiasAtrasado As String
  Dim valorTotalDiasAtrasado As String
  Dim sAcrescimoGrade As String
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  
  If Not rsCR.EOF Then
    With rsCR
      .MoveFirst
      Do While Not .EOF
        
        If IsNull(.Fields("Data Recebimento")) Or .Fields("Data Recebimento") = "" Then
            'Calculo Juros na grid (se houver parcela atrasada)
            Calc_Juros_NaGrid rsParametros("Juros"), .Fields("Valor"), .Fields("Vencimento"), totalDiasAtrasado, valorTotalDiasAtrasado
        Else
            Dim tmpVencimento As String
            tmpVencimento = .Fields("Vencimento")
            
            Dim tmpPagamento As String
            tmpPagamento = .Fields("Data Recebimento")
            
            Dim tmpDias As Integer
            tmpDias = CDate(tmpPagamento) - CDate(tmpVencimento)
            
            totalDiasAtrasado = IIf(tmpDias > 0, tmpDias, "")
            valorTotalDiasAtrasado = ""
        End If
        
        dvalorTotalGrade = dvalorTotalGrade + .Fields("Valor")
        
        sDescricaoNaGrade = .Fields("Descricao")
        sDescricaoNaGrade = Replace(sDescricaoNaGrade, vbCrLf, "")
        sDescricaoNaGrade = Replace(sDescricaoNaGrade, vbTab, "")
        
        sAcrescimoGrade = FormataValorTexto(.Fields("Acrescimo"), 2)
        
        sPendencia = ""
        If .Fields("Pendencia") = -1 Then
            sPendencia = "TEM"
        End If
        
        sRecord = .Fields("Filial") & vbTab & _
          .Fields("Valor") & vbTab & _
          .Fields("Vencimento") & vbTab & _
          .Fields("Desconto") & vbTab & _
          sAcrescimoGrade & vbTab & _
          .Fields("Valor Recebido") & vbTab & _
          .Fields("Data Recebimento") & vbTab & _
          .Fields("Nota") & vbTab & _
          .Fields("Cliente") & vbTab & _
          .Fields("Nome") & vbTab & _
          .Fields("Seq") & vbTab & _
          sDescricaoNaGrade & vbTab & _
          .Fields("Contador") & vbTab & _
          .Fields("Tipo Parcelamento") & vbTab & _
          totalDiasAtrasado & vbTab & _
          valorTotalDiasAtrasado & vbTab & _
          sPendencia
'          .Fields("Vendedor")

        grdCR.AddItem sRecord
        
        .MoveNext
      Loop
      .MoveFirst
    End With
    grdCR.Scroll -99, -99
    grdCR.Redraw = True
  Else
'''    DisplayMsg "Nenhuma conta encontrada segundo os critérios fornecidos."
    grdCR.Redraw = True
  End If

  grdCR.AllowAddNew = bAllow
  grdCR.AllowUpdate = bAllow

  rsCR.Close
  Set rsCR = Nothing
  
  txt_valorTotalGrade.Text = FormatNumber(dvalorTotalGrade, 2)

  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao ler registros do Contas a Receber."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

'12/02/2003 - mpdea
'Comentado código não utilizado
'
'Private Sub WriteGridCR()
'  Dim sSql As String
'  Dim bm As Variant
'  Dim nRow As Long
'  Dim rsCR As Recordset
'
'  On Error GoTo ErrHandler
'
'  grdCR.Update
'
'  Call ws.BeginTrans
'
'  Set rsCR = db.OpenRecordset(sSql, dbOpenDynaset)
'
'  With rsCR
'
'    If Not .EOF Then
'      Do While Not .EOF
'        .Delete
'        .MoveNext
'      Loop
'    End If
'
'    For nRow = 0 To grdCR.Rows - 1
'      bm = grdCR.AddItemBookmark(nRow)
'      If Len(grdCR.Columns("Contato").CellText(bm)) > 0 Then
'        .AddNew
'        '.Fields("Cliente") = cboCodigo.Text
'        .Fields("Seqüência") = nRow + 1
'        .Fields("Contato") = grdCR.Columns("Contato").CellText(bm)
'        .Fields("Cargo") = grdCR.Columns("Cargo").CellText(bm)
'        .Fields("Dia Aniversário") = CInt(gsHandleNull(grdCR.Columns("DiaAniv").CellValue(bm) & ""))
'        .Fields("Mês Aniversário") = grdCR.Columns("MesAniv").CellValue(bm) & ""
'        .Fields("Ramal") = grdCR.Columns("Ramal").CellValue(bm) & ""
'        .Fields("email") = grdCR.Columns("e-mail").CellValue(bm) & ""
'        .Update
'      End If
'    Next nRow
'
'  End With
'
'  rsCR.Close
'  Set rsCR = Nothing
'
'  Call ws.CommitTrans
'  Exit Sub
'
'ErrHandler:
'  gsTitle = LoadResString(201)
'  gsMsg = "Erro ao Atualizar Contatos."
'  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
'  gnStyle = vbOKOnly + vbExclamation
'  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'  Exit Sub
'
'End Sub

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

Private Sub cmd_acataUsuarioLogadoComoOperador_Click()
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  Data_Pagto.Enabled = True
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Vcto_Final.Text = frmCalendario.gsDateCalender(Vcto_Final.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Vcto_Inicial.Text = frmCalendario.gsDateCalender(Vcto_Inicial.Text)
End Sub

Private Sub EmiteCarnesNOVOS(pBolPergunta As Boolean, pSequencia As String, pQual40Colunas As Integer)
On Error GoTo Erro:

  Dim Resp As String
  If pBolPergunta = True Then
    
      Resp = InputBox("Impressão em modelo:" & vbCrLf & vbCrLf & "     1 - TICKET         [40 colunas]" & vbCrLf & vbCrLf & "     2 - RELATÓRIO [110 colunas]", "Qual o modelo de impressão?", "1")
      If Not IsNumeric(Resp) Then
          DisplayMsg "Opção de impressão inválida!"
          Exit Sub
      Else
          If Resp <> "1" And Resp <> "2" Then
              DisplayMsg "Opção de impressão inválida!"
              Exit Sub
          End If
      End If

      Dim strNomeArq As String
    
      If Resp = "2" Then
          CrystalReport1.Destination = 0
      
          strNomeArq = gsReportPath & "carne02.rpt"
      Else
          CrystalReport1.WindowShowPrintSetupBtn = True
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Destination = IIf(False, crptToWindow, crptToPrinter)

          strNomeArq = gsReportPath & "carne02_todasParcelas_46Colunas.rpt"
      End If
  Else
      CrystalReport1.WindowShowPrintSetupBtn = True
      CrystalReport1.WindowState = crptMaximized
      CrystalReport1.Destination = IIf(False, crptToWindow, crptToPrinter)

      If pQual40Colunas = 3 Then
        strNomeArq = gsReportPath & "carne02_todasParcelas_46Colunas_parcelasPagasNoDia.rpt"
      Else
        strNomeArq = gsReportPath & "carne02_todasParcelas_46Colunas.rpt"
      End If
  End If
  
  If Dir(strNomeArq) = "" Then
    DisplayMsg "Arquivo """ & strNomeArq & """ não encontrado."
    Exit Sub
  End If
  
  CrystalReport1.DataFiles(0) = gsQuickDBFileName
  CrystalReport1.ReportFileName = strNomeArq
  
  If pSequencia = "" Then
      pSequencia = "0"
  End If
  
  CrystalReport1.ParameterFields(0) = "pSequencia;" & pSequencia & ";true"
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
  
  CrystalReport1.ParameterFields(1) = "pEmpresa;" & sEmpresaNome & ";true"
  CrystalReport1.ParameterFields(2) = "pEmpresaEnderecoRua;" & sEmpresaRuaComNumero & ";true"
  CrystalReport1.ParameterFields(3) = "pEmpresaEnderecoCidadeEstado;" & sEmpresaCidadeEstado & ";true"
  CrystalReport1.ParameterFields(4) = "pEmpresaEnderecoFone;" & sEmpresaFone & ";true"
  CrystalReport1.ParameterFields(5) = "pEmpresaEnderecoCep;" & "Cep " & sEmpresaCep & ";true"

  Dim i As Integer

  If pQual40Colunas = 3 Then
      If totalNumParcelaPaga > 0 Then
          For i = 0 To totalNumParcelaPaga - 1
              If i = 0 Then
                  CrystalReport1.ParameterFields(7) = "pParcela1;" & numParcelaPaga(i) & ";true"
              ElseIf i = 1 Then
                  CrystalReport1.ParameterFields(8) = "pParcela2;" & numParcelaPaga(i) & ";true"
              ElseIf i = 2 Then
                  CrystalReport1.ParameterFields(9) = "pParcela3;" & numParcelaPaga(i) & ";true"
              ElseIf i = 3 Then
                  CrystalReport1.ParameterFields(10) = "pParcela4;" & numParcelaPaga(i) & ";true"
              ElseIf i = 4 Then
                  CrystalReport1.ParameterFields(11) = "pParcela5;" & numParcelaPaga(i) & ";true"
              ElseIf i = 5 Then
                  CrystalReport1.ParameterFields(12) = "pParcela6;" & numParcelaPaga(i) & ";true"
              ElseIf i = 6 Then
                  CrystalReport1.ParameterFields(13) = "pParcela7;" & numParcelaPaga(i) & ";true"
              ElseIf i = 7 Then
                  CrystalReport1.ParameterFields(14) = "pParcela8;" & numParcelaPaga(i) & ";true"
              ElseIf i = 8 Then
                  CrystalReport1.ParameterFields(15) = "pParcela9;" & numParcelaPaga(i) & ";true"
              ElseIf i = 9 Then
                  CrystalReport1.ParameterFields(16) = "pParcela10;" & numParcelaPaga(i) & ";true"
              ElseIf i = 10 Then
                  CrystalReport1.ParameterFields(17) = "pParcela11;" & numParcelaPaga(i) & ";true"
              ElseIf i = 11 Then
                  CrystalReport1.ParameterFields(18) = "pParcela12;" & numParcelaPaga(i) & ";true"
              End If
          Next
          
          If totalNumParcelaPaga = 1 Then
              CrystalReport1.ParameterFields(8) = "pParcela2;0;true"
              CrystalReport1.ParameterFields(9) = "pParcela3;0;true"
              CrystalReport1.ParameterFields(10) = "pParcela4;0;true"
              CrystalReport1.ParameterFields(11) = "pParcela5;0;true"
              CrystalReport1.ParameterFields(12) = "pParcela6;0;true"
              CrystalReport1.ParameterFields(13) = "pParcela7;0;true"
              CrystalReport1.ParameterFields(14) = "pParcela8;0;true"
              CrystalReport1.ParameterFields(15) = "pParcela9;0;true"
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 2 Then
              CrystalReport1.ParameterFields(9) = "pParcela3;0;true"
              CrystalReport1.ParameterFields(10) = "pParcela4;0;true"
              CrystalReport1.ParameterFields(11) = "pParcela5;0;true"
              CrystalReport1.ParameterFields(12) = "pParcela6;0;true"
              CrystalReport1.ParameterFields(13) = "pParcela7;0;true"
              CrystalReport1.ParameterFields(14) = "pParcela8;0;true"
              CrystalReport1.ParameterFields(15) = "pParcela9;0;true"
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 3 Then
              CrystalReport1.ParameterFields(10) = "pParcela4;0;true"
              CrystalReport1.ParameterFields(11) = "pParcela5;0;true"
              CrystalReport1.ParameterFields(12) = "pParcela6;0;true"
              CrystalReport1.ParameterFields(13) = "pParcela7;0;true"
              CrystalReport1.ParameterFields(14) = "pParcela8;0;true"
              CrystalReport1.ParameterFields(15) = "pParcela9;0;true"
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 4 Then
              CrystalReport1.ParameterFields(11) = "pParcela5;0;true"
              CrystalReport1.ParameterFields(12) = "pParcela6;0;true"
              CrystalReport1.ParameterFields(13) = "pParcela7;0;true"
              CrystalReport1.ParameterFields(14) = "pParcela8;0;true"
              CrystalReport1.ParameterFields(15) = "pParcela9;0;true"
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 5 Then
              CrystalReport1.ParameterFields(12) = "pParcela6;0;true"
              CrystalReport1.ParameterFields(13) = "pParcela7;0;true"
              CrystalReport1.ParameterFields(14) = "pParcela8;0;true"
              CrystalReport1.ParameterFields(15) = "pParcela9;0;true"
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 6 Then
              CrystalReport1.ParameterFields(13) = "pParcela7;0;true"
              CrystalReport1.ParameterFields(14) = "pParcela8;0;true"
              CrystalReport1.ParameterFields(15) = "pParcela9;0;true"
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 7 Then
              CrystalReport1.ParameterFields(14) = "pParcela8;0;true"
              CrystalReport1.ParameterFields(15) = "pParcela9;0;true"
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 8 Then
              CrystalReport1.ParameterFields(15) = "pParcela9;0;true"
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 9 Then
              CrystalReport1.ParameterFields(16) = "pParcela10;0;true"
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 10 Then
              CrystalReport1.ParameterFields(17) = "pParcela11;0;true"
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          ElseIf totalNumParcelaPaga = 11 Then
              CrystalReport1.ParameterFields(18) = "pParcela12;0;true"
          End If
          totalNumParcelaPaga = 0
      End If
  End If
  
  CrystalReport1.WindowState = crptMaximized
  
  If Resp = "1" Then
    Call SetPrinterName("TICKET", CrystalReport1)
  Else
    Call SetPrinterName("REL", CrystalReport1)
  End If

  CrystalReport1.Action = 1
  
  CrystalReport1.Reset

  Exit Sub
Erro:
  MsgBox "Erro tentando gerar Carnês. Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_imprimeCarne_Click()
On Error GoTo Erro
  Dim book As Variant
  Call StatusMsg("")

  If grdCR.SelBookmarks.Count < 1 Then
    DisplayMsg "Selecione um registro na grade."
    Exit Sub
  End If
 
  If grdCR.SelBookmarks.Count = 1 Then
      book = grdCR.SelBookmarks(0)
     
      EmiteCarnesNOVOS True, grdCR.Columns("Seq").CellValue(book), 2
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro ao imprimir carne" & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub cmd_troco_Click()
On Error GoTo Erro
  Dim troco As Double
  
  If Valor_Pago.Text = "" Then
    MsgBox "Campo 'Valor Pago' precisa estar preenchido", vbInformation, "Atenção"
    Exit Sub
  End If

  If txt_cliPagou.Text = "" Then
    MsgBox ("Digite corretamente quanto o cliente pagou...Ex: 100,00")
    Exit Sub
  End If
  
  If CDbl(Valor_Pago.Text) >= CDbl(txt_cliPagou.Text) Then
      txt_troco.Text = "0,00"
      Exit Sub
  End If
  
  txt_cliPagou.Text = Replace(txt_cliPagou.Text, ".", ",")
  txt_troco.Text = CDbl(txt_cliPagou.Text) - CDbl(Valor_Pago.Text)

  Exit Sub
Erro:
  MsgBox "Erro !! Digite no campo 'Cliente pagou...'  neste formato 1000,00 (vírgula é o separador decimal)", vbInformation, "Atenção"
End Sub

Private Sub cmdDesconto_Click()
  '17/08/2004 - Daniel
  'Criado para atender inicialmente a necessidade do Nazareno - RJ
  'mas liberado para toda clientela
  Dim rstParametros   As Recordset
  Dim dblTaxaDesconto As Double
  Dim dblDivisao      As Double
  Dim intDiferenca    As Integer
  Dim dblResult       As Double
  Dim blnSair         As Boolean
  
  If Not IsDate(Vencimento.Text) Then
    MsgBox "Vencimento inválido, verifique.", vbExclamation, "Quick Store"
    Exit Sub
  End If
  
  If CDate(Vencimento.Text) <= Data_Atual Then
    MsgBox "Impossível calcular o Desconto, Vencimento menor ou igual a hoje.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  Set rstParametros = db.OpenRecordset("SELECT Filial, TaxaDesconto FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset)

  With rstParametros
    If Not (.BOF And .EOF) Then
      If .Fields("TaxaDesconto").Value = 0 Or Not IsNumeric(.Fields("TaxaDesconto").Value) Then
        blnSair = True
      Else
        dblTaxaDesconto = .Fields("TaxaDesconto").Value
      End If
    End If
    .Close
  End With

  Set rstParametros = Nothing
  
  If blnSair Then
    MsgBox "Cadastre em Parâmetros a Taxa de Desconto.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  dblDivisao = Format((dblTaxaDesconto / 30), "##,###,##0.00")

  intDiferenca = CInt(CDate(Vencimento.Text) - Data_Atual)
  
  dblResult = Format((intDiferenca * dblDivisao), "##,###,##0.00")

  Desconto.Text = dblResult
  Valor_Pago.Text = CDbl(Valor.Text) - CDbl(Desconto.Text)

End Sub

'29/12/2003 - mpdea
'Realiza a verificação do tipo de documento na emissão
'
'Private Sub cboTipoEmiss_Click()
'  If cboTipoEmiss.Text = "Fatura" Then
'    strTipoEmiss = "F"
'  Else
'    strTipoEmiss = "R"
'  End If
'End Sub

'12/12/2003 - mpdea
'Implementado impressão de Boleto
'14/01/2004 - Daniel
'Alterado de Private Sub cmdEmiss_Click() para
'Public Sub cmdEmiss_Click()
Public Sub cmdEmiss_Click()
  Dim nX As Integer
  Dim vBook As Variant
  '-- removido por Pablo em 20/06/2022
  'Dim cValue As Currency
  Dim nCodCliente As Long
  Dim strTipoEmiss As String
    
  '27/08/2004 - Daniel
  'Adicionado a realidade para o cliente De Mais Presentes (Nazareno - RJ)
  'que precisa de impressão do boleto automaticamente sem 'n' click's
  If m_blnImprimirDireto Then
    If grdCR.SelBookmarks.Count = 0 Then Exit Sub
    'Se não continuará...
    strTipoEmiss = "B"
  Else 'Para os demais clientes
    
    '14/01/2004 - Daniel
    'Flag proveniente do frmImpressaoComprovante
    'Antes havia apenas o seguinte código:
    'If grdCR.SelBookmarks.Count = 0 Then Exit Sub
    If frmImpressaoComprovante.g_blnFlag Then
      strTipoEmiss = frmImpressaoComprovante.g_strTipoEmiss
    Else
      If grdCR.SelBookmarks.Count = 0 Then Exit Sub
    End If
    '-----------------------------------------------------------
  
  End If
  
  '29/12/2003 - mpdea
  'Alterado estilo do objeto
  '
  '27/08/2004 - Daniel
  'Adicionado a realidade para o cliente De Mais Presentes (Nazareno - RJ)
  'que precisa de impressão do boleto automaticamente sem 'n' click's
  If Not m_blnImprimirDireto Then 'Demais clientes
    If Not frmImpressaoComprovante.g_blnFlag Then
      Select Case cboTipoEmiss.Text
        Case "Recibo"
          strTipoEmiss = "R"
        Case "Fatura"
          strTipoEmiss = "F"
        Case "Boleto"
          strTipoEmiss = "B"
      End Select
    End If
  End If
  
  
  '14/01/2004 - Daniel
  'g_blnFlag Voltando ao estado inicial
  frmImpressaoComprovante.g_blnFlag = False
  
  Select Case strTipoEmiss
    Case "B"
      '15/01/2004 - mpdea
      'Implementado a impressão com várias contas
      '
'      If grdCR.SelBookmarks.Count > 1 Then
'        DisplayMsg "Para a emissão de Boleto selecione apenas uma conta."
'        Exit Sub
'      End If
      
      If Not m_blnPrintBoleto Then Exit Sub
      
    Case "F", "R"
      If strTipoEmiss = "F" And grdCR.SelBookmarks.Count > 1 Then
        DisplayMsg "Para a emissão de Fatura selecione apenas uma conta."
        Exit Sub
      End If
      
      'Recibo para várias contas
      If strTipoEmiss = "R" And grdCR.SelBookmarks.Count > 1 Then
        'Verifica se as contas selecionadas pertencem ao mesmo cliente
        'e obtém o valor total
        For nX = 0 To (grdCR.SelBookmarks.Count - 1)
          vBook = grdCR.SelBookmarks(nX)
          If nX = 0 Then
            nCodCliente = grdCR.Columns("Cliente").CellValue(vBook)
          Else
            If nCodCliente <> grdCR.Columns("Cliente").CellValue(vBook) Then
              DisplayMsg "Para a emissão de Recibo de várias contas selecione apenas o mesmo cliente."
              Exit Sub
            End If
          End If
          'cValue = cValue + grdCR.Columns("Valor").CellValue(vBook) + _
          '  grdCR.Columns("Acresc").CellValue(vBook) - _
          '  grdCR.Columns("Desc").CellValue(vBook)
        Next nX
      End If
      
      If Not IsNumeric(Acréscimo.Text) Then Acréscimo.Text = 0
      If Not IsNumeric(Desconto.Text) Then Desconto.Text = 0
      
      nReciboVALOR = CDbl(Valor.Text)
      nReciboACRESCIMO = CDbl(Acréscimo.Text)
      nReciboDESCONTO = CDbl(Desconto.Text)
      
' autor: Pablo Verçosa Silva
' data: 18/06/2022
      If strTipoEmiss = "R" Then Call Monta_Texto_Imprime
      
      With frmEmiteFatura
        If grdCR.SelBookmarks.Count = 1 Then
          vBook = grdCR.SelBookmarks(0)
          .Transf1.Caption = grdCR.Columns("Vcto").CellValue(vBook)
          .Transf2.Caption = grdCR.Columns("Contador").CellValue(vBook)
          .lblCheckValue.Caption = "True"
        Else
          vBook = grdCR.SelBookmarks(0)
          .Transf1.Caption = grdCR.Columns("Vcto").CellValue(vBook)
          .Transf2.Caption = grdCR.Columns("Contador").CellValue(vBook)
          .lblCheckValue.Caption = "False"
          '-- removido por Pablo em 20/06/2022
          '.L_Valor.Caption = Format(cValue, FORMAT_VALUE)
          .L_Valor.Caption = Format(CDbl(Valor_Pago.Text), FORMAT_VALUE)
        End If
        If strTipoEmiss = "F" Then
          .Caption = "Emissão de Fatura"
        Else
          .Caption = "Emissão de Recibo"
        End If
        .L_Encontrar.Caption = "SIM"
        .Tipo.Caption = strTipoEmiss
        .optTotalParcela.Enabled = False
        .Show vbModal
      End With
      
  End Select

  '12/12/2003 - mpdea
  'Processo de impressão iniciado, não solicitar confirmação
  g_blnConfirmarImpressao = False
  
End Sub

' autor: Pablo Verçosa Silva
' data: 18/06/2022
' descrição: monta texto padrão do recibo
Private Sub Monta_Texto_Imprime()
  Dim g_acresc As Currency
  Dim g_descon As Currency
  Dim g_descri As String
  Dim Texto As String
  Dim texto_tamanho As Integer
  Dim n As Integer
  Dim book As Variant
  
  g_acresc = 0
  g_descon = 0
  g_descri = ""
  Texto = ""
  texto_tamanho = 0
  n = 0
  
  g_acresc = IIf(IsNumeric(Acréscimo.Text), CDbl(Acréscimo.Text), 0)
  g_descon = IIf(IsNumeric(Desconto.Text), CDbl(Desconto.Text), 0)
  
  For n = 0 To (grdCR.SelBookmarks.Count - 1)
    book = grdCR.SelBookmarks(n)
    If Not IsNull(grdCR.Columns("Descricao").CellValue(book)) Then
      If Trim(grdCR.Columns("Descricao").CellValue(book)) <> "" Then
        If Trim(g_descri) <> "" Then g_descri = g_descri & ", "
        g_descri = g_descri & "(" & CStr(grdCR.Columns("Seq").CellValue(book)) & " - " & Trim(grdCR.Columns("Descricao").CellValue(book)) & ")"
      End If
    End If
  Next n
    
  If g_acresc <> 0 Then Texto = Texto & "Acrésc.: R$" & Format(g_acresc, FORMAT_VALUE) & ". "
  If g_descon <> 0 Then Texto = Texto & "Desc.: R$" & Format(g_descon, FORMAT_VALUE) & ". "
  If Trim(g_descri) <> "" Then Texto = Texto & Trim(g_descri)
  
  texto_tamanho = Len(Texto)
  If texto_tamanho > 0 Then
    If texto_tamanho > 300 Then Texto = Left(Texto, 300)
    If texto_tamanho < 300 Then Texto = Left(Texto & Space(300), 300)
    For n = 0 To 4
      frmEmiteFatura.Texto_Recibo(n).Text = Trim(Mid(Texto, 1 + (n * 60), 60 * (n + 1)))
    Next n
  End If
End Sub

Private Sub Baixa_DblClick(Index As Integer)
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

Private Sub Cheque_Bom_LostFocus()
  Cheque_Bom.Text = Ajusta_Data(Cheque_Bom.Text)
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
 
 
  txt_lembreteCliente.Text = ""
  Dim rsLembreteEventosCliente As Recordset

  Set rsLembreteEventosCliente = db.OpenRecordset("SELECT * FROM [Contatos Efetuados] WHERE Cliente = " & Val(Combo_Fornecedor.Text) & " ORDER BY Data desc, Seqüência", dbOpenDynaset)

  'Se houver registro...mostrar apenas o ultimo
  If Not rsLembreteEventosCliente.EOF Then
      With rsLembreteEventosCliente
          If .Fields("Pendência") = True Then
              txt_lembreteCliente.Text = .Fields("Data") & ": " & .Fields("Descrição")
          End If
      End With
  End If
  rsLembreteEventosCliente.Close
  Set rsLembreteEventosCliente = Nothing
End Sub

Private Sub Data_Pagto_LostFocus()
  Data_Pagto.Text = Ajusta_Data(Data_Pagto.Text)
  
  '--------------------------------------------
  '05/01/2004 - Daniel
  '
  'Verifica alteração personalizada
  '
  'QS37818-990 = F. Linhares
  'Chama o Cálculo de Juros automáticamente
  'sem clicar no B_Calc_Juros
  '
  If m_blnFLinhares Then
    '10/03/2004 - Daniel
    'Verifica se os dois campos realmente são datas para evitar
    'erros quando realizarmos baixa de várias parcelas
    If IsDate(Vencimento.Text) And IsDate(Data_Pagto.Text) Then
      'Somente se estiver em atraso...
      If CDate(Vencimento.Text) < CDate(Data_Pagto.Text) Then
        'Solicitado para tornar-se Opcional o cálculo de juros
        If MsgBox("O pagamento está sendo recebido com atraso, deseja calcular juros?", vbExclamation + vbYesNo, "Atenção") = vbYes Then
          B_Calc_Juros_Click
        End If
      End If
    End If
  End If
  '--------------------------------------------------------
  
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
On Error GoTo Erro
  '30/05/2005 - Daniel
  'Cálculo do Valor Pago automático
  'Solicitante: Pedágio
  Dim dblDesconto As Double

  dblDesconto = Format(0 & (Desconto.Text), FORMAT_VALUE)
  
  Valor_Pago.Text = Format(CDbl(Valor.Text) - dblDesconto + CDbl(0 & (Acréscimo.Text)), FORMAT_VALUE)

  Exit Sub
Erro:
  MsgBox "Inconsistência " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber", dbOpenDynaset)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários")
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  '25/09/2007 - Anderson
  'Otimizar o pagamento de parcelas através de código de barras no carnê
  'grdCR.Height = GRID_FULL_HEIGHT  'Posição inicial
''  grdCR.Top = IIf(g_bolCarneCodigoBarras, 2080, 1600)
''  grdCR.Height = IIf(g_bolCarneCodigoBarras, GRID_FULL_HEIGHT - 285, GRID_FULL_HEIGHT)
  
  If gbCaixas = False Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
  End If
 
  cboTipoEmiss.Enabled = True
  cmdEmiss.Enabled = True
  'strTipoEmiss = "R"
  
  '11/03/2004 - Daniel
  'Var flag de verificação se é F. Linhares
  m_blnFLinhares = CheckSerialCaseMod("QS37818-990")
  '-------------------------------------------------
  
  '27/08/2004 - Daniel
  'Case: De Mais Presentes (Nazareno) QS31735-849
  'Aberto também para o cliente F. Linhares QS37818-990
  'Evitar n click's no momento da impressão do boleto
  m_blnImprimirDireto = CheckSerialCaseMod("QS31735-849", "QS37818-990")
  
  '26/12/2003 - mpdea
  'Adiciona os tipos de documentos para emissão
  With cboTipoEmiss
    .Clear
    .AddItem "Recibo"
    .AddItem "Fatura"
    '15/01/2004 - mpdea
    'Removido opção devido a mesma ser disponibilizada no momento
    'de confirmar a baixa da(s) conta(s)
    '.AddItem "Boleto"
  End With
  
  '12/12/2003 - mpdea
  'Alterado estilo do objeto
  'cboTipoEmiss.Text = "Recibo"
  cboTipoEmiss.ListIndex = 0 'Recibo
  
  Call GetSettings
  
  cboFilial.Text = gnCodFilial
  cboFilial_LostFocus
  
  '26/09/2007 - Anderson
  'Posiciona o cursor para digitação do código de barras
  'Solicitante: Naativa
  If Not g_bolCarneCodigoBarras Then
    txtCodigoBarras.Enabled = False
    txtCodigoBarras.Visible = False
  End If

  grdCR.StyleSets("Vermelho").ForeColor = RGB(217, 82, 84)
  grdCR.StyleSets("Preto").ForeColor = RGB(0, 0, 0)
  
  If gbCaixas = True Then
    Combo_Caixa.Enabled = True
    Label_Caixa.Enabled = True
  End If
  
  '15/08/2023 - Pablo
  'Deixa o TOTAL visível apenas se for acessado por um superusuário
  Text2.Visible = gbSuperUser
  txt_valorTotalGrade.Visible = gbSuperUser
  Text2.Refresh
  txt_valorTotalGrade.Refresh
End Sub

Private Sub Form_Paint()
  '15/08/2023 - Pablo
  'Deixa o TOTAL visível apenas se for acessado por um superusuário
  Text2.Visible = gbSuperUser
  txt_valorTotalGrade.Visible = gbSuperUser
  Text2.Refresh
  txt_valorTotalGrade.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Call SaveSetting("QuickStore", "CRMan", "Data1", Vcto_Inicial.Text)
  Call SaveSetting("QuickStore", "CRMan", "Data2", Vcto_Final.Text)
  
  rsCliFor.Close
  rsCaixas.Close
  rsCaixa.Close
  rsContas_Receber.Close
  rsLançamentos.Close
  rsContas.Close
  rsParametros.Close
  
  Set rsCliFor = Nothing
  Set rsCaixas = Nothing
  Set rsCaixa = Nothing
  Set rsContas_Receber = Nothing
  Set rsLançamentos = Nothing
  Set rsContas = Nothing
  Set rsParametros = Nothing
  

End Sub

Private Sub GetSettings()
  Vcto_Final.Text = GetSetting("QuickStore", "CRMan", "Data2", CDate(Date))
  Vcto_Inicial.Text = GetSetting("QuickStore", "CRMan", "Data1", CDate(Date))
End Sub

Private Sub grdCR_AfterDelete(RtnDispErrMsg As Integer)
  grdCR.Scroll 0, -32767
  grdCR.Scroll 0, 32767
End Sub

Private Sub grdCR_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  If Not bGridBeforeDelete() Then
    Cancel = True
  End If
End Sub

Private Sub grdCR_DblClick()
  '22/04/2005 - Daniel
  'Otimizado rotina para abrir a tela de lançamentos de contas
  'com a conta selecionada
  '
  'Solicitante: Consultor Carlos (Petrópolis - RJ)
  
  On Error GoTo TratarErro
  
  'If Not frmGerente.gbSenhaGerente Then
  '  Exit Sub
  'End If
  
  Me.g_blnFind = True
  
  g_strQuery = ""
  g_strQuery = "SELECT * FROM [Contas a Receber] "
  g_strQuery = g_strQuery & " WHERE Filial = " & CByte(grdCR.Columns(0).Text)
  g_strQuery = g_strQuery & " AND Vencimento = #" & Format(CDate(grdCR.Columns(2).Text), "MM/DD/YYYY") & "#"
  g_strQuery = g_strQuery & " AND Sequência = " & CLng(grdCR.Columns(10).Text)
  g_strQuery = g_strQuery & " AND Contador = " & CLng(grdCR.Columns(12).Text)
  
'''  If Not frmGerente.gbSenhaGerente Then
'''    Exit Sub
'''  End If
  
  frmLancaCReceber.Show
  
  Exit Sub
  
TratarErro:
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub grdCR_RowLoaded(ByVal Bookmark As Variant)
On Error GoTo Deu_Erro
 
  Dim DiasAtraso As String
  'Dim dataReceb As String
 
  If IsEmpty(Bookmark) Then Exit Sub
 
  DiasAtraso = grdCR.Columns("DiasAtraso").CellText(Bookmark)
  'dataReceb = grdCR.Columns("data Receb").CellText(Bookmark)
  
  'If DiasAtraso <> "0" And dataReceb = "" Then
  If IsNumeric(DiasAtraso) And CInt(DiasAtraso) > 0 Then
      grdCR.Columns(0).CellStyleSet "Vermelho"
      grdCR.Columns(1).CellStyleSet "Vermelho"
      grdCR.Columns(2).CellStyleSet "Vermelho"
      grdCR.Columns(3).CellStyleSet "Vermelho"
      grdCR.Columns(4).CellStyleSet "Vermelho"
      grdCR.Columns(5).CellStyleSet "Vermelho"
      grdCR.Columns(6).CellStyleSet "Vermelho"
      grdCR.Columns(7).CellStyleSet "Vermelho"
      grdCR.Columns(8).CellStyleSet "Vermelho"
      grdCR.Columns(9).CellStyleSet "Vermelho"
      grdCR.Columns(10).CellStyleSet "Vermelho"
      grdCR.Columns(11).CellStyleSet "Vermelho"
      grdCR.Columns(12).CellStyleSet "Vermelho"
      grdCR.Columns(13).CellStyleSet "Vermelho"
      grdCR.Columns(14).CellStyleSet "Vermelho"
      grdCR.Columns(15).CellStyleSet "Vermelho"
  End If

  Exit Sub
  
Deu_Erro:
  Exit Sub
End Sub

Private Sub grdCR_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  Dim Val_Selec As Double
  Dim i As Integer
  Dim book As Variant
  Dim sConta As String
  Dim sSelec As String
  
  If grdCR.SelBookmarks.Count = 0 Then
   grdCR.Caption = "Contas: nenhuma selecionada"
   cboTipoEmiss.Enabled = False
   cmdEmiss.Enabled = False
   Exit Sub
  End If
  
  cboTipoEmiss.Enabled = True
  cmdEmiss.Enabled = True
  
  Val_Selec = 0#
  For i = 0 To (grdCR.SelBookmarks.Count - 1)
    book = grdCR.SelBookmarks(i)
    Val_Selec = Val_Selec + grdCR.Columns("Valor").CellValue(book)
  Next i
  
  If grdCR.SelBookmarks.Count = 1 Then
    sConta = "Conta: "
    sSelec = " selecionada"
  Else
    sConta = "Contas: "
    sSelec = " selecionadas"
  End If
  
  grdCR.Caption = sConta & CStr(grdCR.SelBookmarks.Count) & sSelec & ", valor " + Format((CStr(Val_Selec)), "Currency")

End Sub


Private Sub Nota_KeyPress(KeyAscii As Integer)
 KeyAscii = Verifica_Tecla_Integer(KeyAscii)
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
  Cheque_Bom = gsFormatDate(Data_Atual)
End Sub

Private Sub O_Não_determinado_Click()
  If Not frmGerente.gbSenhaGerente Then
    O_caixa_d.Value = True
    Exit Sub
  End If
'  frmGerente.Show vbModal
'  If gsRetornoDoc <> "OK" Then
'    Exit Sub
'  End If
  Label_Caixa.Enabled = False
  Combo_Caixa.Enabled = False
  Nome_Caixa.Enabled = False
End Sub

Private Sub txtCodigoBarras_GotFocus()
  With txtCodigoBarras
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtCodigoBarras_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call LoadGridCRCodigoBarras
    Call txtCodigoBarras_GotFocus
  End If
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

 KeyAscii = 0
 
End Sub

Private Sub Vcto_Final_LostFocus()
  
  Vcto_Final.Text = Ajusta_Data(Vcto_Final.Text)
End Sub

Private Sub Vcto_Final_GotFocus()
  With Vcto_Final
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
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

Private Sub Vcto_Inicial_LostFocus()
  
  Vcto_Inicial.Text = Ajusta_Data(Vcto_Inicial.Text)
End Sub

Private Sub Vcto_Inicial_GotFocus()
  With Vcto_Inicial
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
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

'15/01/2004 - mpdea
'Alterado para permitir a impressão de vários boletos
'
'12/12/2003 - mpdea
'Impressão do boleto (adaptado da tela Lançamento de Contas a Receber)
Private Function m_blnPrintBoleto() As Boolean
  Dim Resp As Integer
  Dim Nome_Arq As String
  Dim F As Form
  
  Dim varBookmark As Variant
  Dim intFilial As Integer
  Dim dteVencimento As Date
  Dim lngContador As Long
  
  Dim intX As Integer
  
  '11/03/2004 - Daniel
  'Vars para auxiliar na na impressão da Qtde. de cópias do ticket
  Dim intCopias As Integer
  
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")

  Set F = New frmObsDoc
  F.Caption = "Impressão de Boletos"
  F.gsFileExt = ".CBB"
  F.Show vbModal
  Set F = Nothing

  If gsRetornoDoc <> "OK" Then Exit Function

  Nome_Arq = gsConfigPath & gsDocFileName & ".CBB"
  If Dir(Nome_Arq) = "" Then
    DisplayMsg "Arquivo """ & Nome_Arq & """ não encontrado."
    Exit Function
  End If

  '11/03/2004 - Daniel
  'Impressão de n cópias
  For intCopias = 1 To frmImpressaoComprovante.g_intCopias
  
      For intX = 0 To (grdCR.SelBookmarks.Count - 1)
        varBookmark = grdCR.SelBookmarks(intX)
        Call IsDataType(dtInteger, grdCR.Columns("Filial").CellValue(varBookmark), intFilial)
        Call IsDataType(dtDate, grdCR.Columns("Vcto").CellValue(varBookmark), dteVencimento)
        Call IsDataType(dtLong, grdCR.Columns("Contador").CellValue(varBookmark), lngContador)
      
        '15/01/2004 - mpdea
        'Aguarda 5 segundos antes da próxima impressão
        If intX > 0 Then Call WaitSeconds(5, False)
        
        Resp = Imprime_Boleto("R", intFilial, dteVencimento, lngContador, Nome_Arq)
      
        If Resp <> 0 Then
          gsTitle = LoadResString(201)
          gsMsg = "Houve o erro " + str(Resp) + " na emissão do boleto."
          gnStyle = vbOKOnly + vbExclamation
          gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          
          '15/01/2004 - mpdea
          'Sai da função em caso de erro
          Exit Function
          
        Else
          
          '12/12/2003 - mpdea
          'Atualiza o campo de boleto impresso
          db.Execute "UPDATE [Contas a Receber] SET Impresso = True WHERE Contador = " & lngContador, dbFailOnError
          '10/09/2007 - Anderson
          'Gera arquivo log do sistema
          If g_bolSystemLog Then
            SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
            "UPDATE [Contas a Receber] SET Impresso = True WHERE Contador = " & lngContador, _
            "frmManContasReceber_m_blnPrintBoleto", _
            "Contas a Receber", g_strArquivoSystemLog
          End If
        End If
        
      Next intX
  
  Next intCopias
  
  'Limpamos a variável
  frmImpressaoComprovante.g_intCopias = 0
  
  
  '15/01/2004 - mpdea
  'Movido mensagem
  gsTitle = LoadResString(201)
  gsMsg = "Boleto impresso."
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  
  m_blnPrintBoleto = True
  
  Exit Function
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Function

Private Function m_blnExisteParcelaPendente() As Boolean
  Dim rstCRParcelasPendentes As Recordset
  Dim lngSeq As Long
  Dim intParcela As Integer
  Dim strSQL As String
  Dim book As Variant
  Dim i As Long
  Dim lngContador As Long
  Dim blnSelecionouUm As Boolean
  
  Dim blnExisteParcPend As Boolean
  
  On Error GoTo ErrHandler

  If grdCR.SelBookmarks.Count = 1 Then 'Selecionou Um, start no processo de checagem de valores diferentes do previsto
    blnSelecionouUm = True
  End If

  If grdCR.SelBookmarks.Count >= 1 Then 'selecionou um ou mais
    
    rsContas_Receber.Requery 'para não executar n vezes o requery
    
    For i = 0 To (grdCR.SelBookmarks.Count - 1)
      book = grdCR.SelBookmarks(i)
      grdCR.Bookmark = book
      lngContador = grdCR.Columns("Contador").CellValue(book)
      
      With rsContas_Receber
        .FindFirst "Contador = " & lngContador

        'Call IsDataType(dtLong, .Fields("Sequência").Value, lngSeq)
        lngSeq = .Fields("Sequência").Value
        intParcela = .Fields("Parcela").Value
        
        If blnSelecionouUm Then 'influir em m_blnValoresDiffPrevisto
          m_dteVencimento = .Fields("Vencimento").Value
          m_dblValor = .Fields("Valor").Value
        End If
        
        If intParcela > 1 Then
        
          'Obtém somente as parcelas anteriores em aberto da movimentação
          strSQL = "SELECT Contador, Parcela FROM [Contas a Receber] "
          strSQL = strSQL & "WHERE Sequência = " & lngSeq & " "
          strSQL = strSQL & "AND Parcela <> 0 AND "
          strSQL = strSQL & "Parcela < " & intParcela
          strSQL = strSQL & " AND [Valor Recebido] = 0 ORDER BY Parcela"
          
          Set rstCRParcelasPendentes = db.OpenRecordset(strSQL, dbOpenSnapshot)
          With rstCRParcelasPendentes
            If Not (.BOF And .EOF) Then
              Do Until .EOF
                
                'Verifica se a parcela está em processo de baixa
                If Not m_blnParcelaEmBaixa(.Fields("Contador").Value) Then
                  
                  'Parcela Pendente!
                  blnExisteParcPend = True
                  Exit Do
                  
                End If
                
                .MoveNext
              Loop
            End If
            .Close
          End With
          Set rstCRParcelasPendentes = Nothing
          
          If blnExisteParcPend Then Exit For
          
        End If 'Fim da verificação de parcela
        
      End With
    Next i
  End If
  
  m_blnExisteParcelaPendente = blnExisteParcPend

  Exit Function

ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Function

'30/12/2003 - mpdea
'Verifica se a parcela está em processo de baixa através do seu contador
Private Function m_blnParcelaEmBaixa(ByVal lngContador As Long) As Boolean
  Dim lngX As Long
  Dim varBookmark As Variant
  Dim blnRet As Boolean
  
  
  With grdCR
    'Para a atualização de desenho do objeto
    .Redraw = False
    For lngX = 0 To (.SelBookmarks.Count - 1)
      varBookmark = .SelBookmarks(lngX)
      .Bookmark = varBookmark
      If lngContador = .Columns("Contador").CellValue(varBookmark) Then
        'Encontrou a parcela
        blnRet = True
        Exit For
      End If
    Next lngX
    'Retorna a atualização de desenho do objeto
    .Redraw = True
  End With
  
  'Retorna o resultado da pesquisa
  m_blnParcelaEmBaixa = blnRet
  
End Function

'05/01/2004 - Daniel
'Solicitar a senha do gerente para efetuar
'baixas com datas ou valores diferentes dos
'previstos
Private Function m_blnValoresDiffPrevisto() As Boolean
  Dim rstFuncionarios As Recordset
  Dim strSQL As String
  
  On Error GoTo ErrHandler
  
  strSQL = "SELECT SenhaConfirmarCRDiff FROM Funcionários WHERE Código = " & gnUserCode 'var global que carrega o código do usuário
  Set rstFuncionarios = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstFuncionarios
            
    If .Fields("SenhaConfirmarCRDiff").Value = True Then
    'Verificação
      If IsDate(Vencimento.Text) Then             '11/03/2004 - Verificamos se é Data
        If Vencimento.Text <> m_dteVencimento Then
          m_blnValoresDiffPrevisto = True
        End If
      End If
    
      If Valor.Text < m_dblValor Then             '02/09/2004 - Daniel - tirado o <> e colocado o <
        m_blnValoresDiffPrevisto = True
      End If
      
      If Valor_Pago.Text < m_dblValor Then        '02/09/2004 - Daniel - tirado o <> e colocado o <
        m_blnValoresDiffPrevisto = True
      End If
    
    End If

  End With
  
  rstFuncionarios.Close
  Set rstFuncionarios = Nothing

  Exit Function

ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

Private Function BoletoDefault() As Boolean
  '27/08/2004 - Daniel
  'Case: De Mais Presentes (Nazareno - RJ)
  'Nesta private, armazenaremos o boleto default para impressão direta
  Dim rstParametros As Recordset
  Dim strQuery      As String
  
  strQuery = "SELECT Filial, BoletoPadrao "
  strQuery = strQuery & " FROM [Parâmetros Filial] "
  strQuery = strQuery & " WHERE Filial = " & gnCodFilial
  
  Set rstParametros = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If Len(.Fields("BoletoPadrao").Value) <= 0 Then  'Vazio
        BoletoDefault = False
        MsgBox "Cadastre em Parâmetros da Filial um Boleto Padrão.", vbExclamation, "Impressão Automática de Boletos"
      Else
        BoletoDefault = True
        m_strBoletoDefault = .Fields("BoletoPadrao").Value & ".cbb"
      End If
    End If
    .Close
  End With
  
  Set rstParametros = Nothing
  
End Function

Private Function ImpressaoDiretaBoleto() As Boolean
  '02/09/2004 - Daniel
  'Impressão Direta de Boletos sem precisar dar 'n' click's
  'Case: Nazareno, F. Linhares
  Dim intCopias     As Integer
  Dim Resp          As Integer
  Dim Nome_Arq      As String
  Dim varBookmark   As Variant
  Dim intFilial     As Integer
  Dim dteVencimento As Date
  Dim lngContador   As Long
  Dim intX          As Integer
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")

  Nome_Arq = gsConfigPath & m_strBoletoDefault
  If Dir(Nome_Arq) = "" Then
    DisplayMsg "Arquivo """ & Nome_Arq & """ não encontrado."
    Exit Function
  End If
  
  '23/11/2004 - Daniel
  'Tratamento para a Qtde de impressão
  If Not IsNumeric(txtQtdeImprimir.Text) Then
    MsgBox "Qtde para a Impressão inválida, verifique.", vbExclamation, "Atenção"
    txtQtdeImprimir.SetFocus
    Exit Function
  End If
  
  For intCopias = 1 To CInt(txtQtdeImprimir.Text)
  
      For intX = 0 To (grdCR.SelBookmarks.Count - 1)
        varBookmark = grdCR.SelBookmarks(intX)
        Call IsDataType(dtInteger, grdCR.Columns("Filial").CellValue(varBookmark), intFilial)
        Call IsDataType(dtDate, grdCR.Columns("Vcto").CellValue(varBookmark), dteVencimento)
        Call IsDataType(dtLong, grdCR.Columns("Contador").CellValue(varBookmark), lngContador)
      
        '15/01/2004 - mpdea
        'Aguarda 5 segundos antes da próxima impressão
        If intX > 0 Then Call WaitSeconds(5, False)
        
        Resp = Imprime_Boleto("R", intFilial, dteVencimento, lngContador, Nome_Arq)
      
        If Resp <> 0 Then
          gsTitle = LoadResString(201)
          gsMsg = "Houve o erro " + str(Resp) + " na emissão do boleto."
          gnStyle = vbOKOnly + vbExclamation
          gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          
          '15/01/2004 - mpdea
          'Sai da função em caso de erro
          Exit Function
          
        Else
          
          '12/12/2003 - mpdea
          'Atualiza o campo de boleto impresso
          db.Execute "UPDATE [Contas a Receber] SET Impresso = True WHERE Contador = " & lngContador, dbFailOnError
          '10/09/2007 - Anderson
          'Gera arquivo log do sistema
          If g_bolSystemLog Then
            SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
            "UPDATE [Contas a Receber] SET Impresso = True WHERE Contador = " & lngContador, _
            "frmmanContasReceber_ImpressaoDiretaBoleto", _
            "Contas a Receber", g_strArquivoSystemLog
          End If

        End If
        
      Next intX
  
  Next intCopias
  
  Exit Function
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'26/09/2007 - Anderson
'Função Criada para automatizar o pagamento de parcelas por carnê de Código de barras
'Solicitante: Naativa
Private Sub LoadGridCRCodigoBarras()
  Dim rsCR As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sSql As String
  
  On Error GoTo ErrHandler
  
  Call B_Cancela_Click
  
  bAllow = grdCR.AllowAddNew
  grdCR.AllowAddNew = True
  grdCR.AllowUpdate = True
  
  sSql = "SELECT Filial, Valor, Vencimento, [Contas a Receber].Desconto as Desconto, Acréscimo as Acrescimo, "
  sSql = sSql & "[Valor Recebido], [Data Recebimento] , Nota, Cliente, Cli_For.Nome, Sequência as Seq, Descrição as Descricao, Contador, [Tipo Parcelamento] FROM [Contas a Receber]"
  sSql = sSql + " INNER JOIN Cli_For ON ([Contas a Receber].Cliente = Cli_For.Código)"
  sSql = sSql + " WHERE CarneCodigoBarras= '*" & txtCodigoBarras.Text & "*'"
  
  Set rsCR = db.OpenRecordset(sSql, dbOpenDynaset)

  grdCR.RemoveAll
  grdCR.Redraw = False
  
  grdCR.Columns("Valor").NumberFormat = "##,###,##0.00"
  grdCR.Columns("Val Receb").NumberFormat = "##,###,##0.00"
  
  If Not rsCR.EOF Then
    With rsCR
      .MoveFirst
      
      Vcto_Inicial.Text = .Fields("Vencimento")
      Vcto_Inicial_LostFocus
      Vcto_Final.Text = .Fields("Vencimento")
      Vcto_Final_LostFocus
      cboFilial.Text = .Fields("Filial")
      cboFilial_LostFocus
      Combo_Fornecedor.Text = .Fields("Cliente")
      Combo_Fornecedor_LostFocus
      O_Todas.Value = True
      
      Do While Not .EOF
        sRecord = .Fields("Filial") & vbTab & _
          .Fields("Valor") & vbTab & _
          .Fields("Vencimento") & vbTab & _
          .Fields("Desconto") & vbTab & _
          .Fields("Acrescimo") & vbTab & _
          .Fields("Valor Recebido") & vbTab & _
          .Fields("Data Recebimento") & vbTab & _
          .Fields("Nota") & vbTab & _
          .Fields("Cliente") & vbTab & _
          .Fields("Nome") & vbTab & _
          .Fields("Seq") & vbTab & _
          .Fields("Descricao") & vbTab & _
          .Fields("Contador") & vbTab & _
          .Fields("Tipo Parcelamento")
        grdCR.AddItem sRecord
        .MoveNext
      Loop
      .MoveFirst
    End With
    
    grdCR.Scroll -99, -99
    grdCR.Redraw = True
    grdCR.SelBookmarks.Add 0
    
    If rsCR.Fields("Valor Recebido").Value < rsCR.Fields("Valor").Value Then
      Call grdCR_SelChange(2, 0, 0)
      Call B_Baixa_Click
      Valor_Pago.SetFocus
    End If
  Else
'''    DisplayMsg "Nenhuma conta encontrada segundo os critérios fornecidos."
    grdCR.Redraw = True
  End If

  grdCR.AllowAddNew = bAllow
  grdCR.AllowUpdate = bAllow
  
  rsCR.Close
  Set rsCR = Nothing

  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao ler registros do Contas a Receber."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub


