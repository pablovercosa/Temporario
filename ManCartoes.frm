VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManCartoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Realizar baixa de Cartões"
   ClientHeight    =   7320
   ClientLeft      =   1530
   ClientTop       =   1785
   ClientWidth     =   14130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1590
   Icon            =   "ManCartoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7320
   ScaleWidth      =   14130
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Left            =   5970
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   7290
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   4980
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   90
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Administradora"
      Height          =   735
      Left            =   5445
      TabIndex        =   36
      Top             =   465
      Width           =   2970
      Begin SSDataWidgets_B.SSDBCombo cboAdm 
         Bindings        =   "ManCartoes.frx":4E95A
         Height          =   315
         Left            =   195
         TabIndex        =   2
         ToolTipText     =   "Escolha a Administradora ou deixe em branco para todas"
         Top             =   255
         Width           =   2655
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
         Columns(0).Width=   4339
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdCartoes 
      Height          =   3900
      Left            =   120
      TabIndex        =   11
      Top             =   1710
      Width           =   13935
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
      Col.Count       =   10
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   10
      Columns(0).Width=   873
      Columns(0).Caption=   "Filial"
      Columns(0).Name =   "Filial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1455
      Columns(1).Caption=   "Cód. Adm"
      Columns(1).Name =   "Codigo"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3149
      Columns(2).Caption=   "Nome"
      Columns(2).Name =   "Nome"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3043
      Columns(3).Caption=   "Número Cartão"
      Columns(3).Name =   "NumCartao"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1984
      Columns(4).Caption=   "Data Vcto"
      Columns(4).Name =   "DataVcto"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   10
      Columns(4).Mask =   "##/##/####"
      Columns(4).PromptInclude=   -1  'True
      Columns(4).PromptChar=   32
      Columns(5).Width=   2196
      Columns(5).Caption=   "Valor"
      Columns(5).Name =   "Valor"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2170
      Columns(6).Caption=   "Val. Desc"
      Columns(6).Name =   "ValDesc"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1667
      Columns(7).Caption=   "Depositado"
      Columns(7).Name =   "Depositado"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   11
      Columns(7).FieldLen=   256
      Columns(7).Style=   2
      Columns(8).Width=   2170
      Columns(8).Caption=   "Val Recebido"
      Columns(8).Name =   "ValReceb"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "Contador"
      Columns(9).Name =   "Contador"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   24580
      _ExtentY        =   6879
      _StockProps     =   79
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
   Begin VB.Frame Frame5 
      Caption         =   "Vencimento"
      Height          =   750
      Left            =   120
      TabIndex        =   33
      Top             =   450
      Width           =   5280
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
         Left            =   2070
         Picture         =   "ManCartoes.frx":4E96E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   210
         Width           =   465
      End
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
         Left            =   4590
         Picture         =   "ManCartoes.frx":4F250
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   210
         Width           =   465
      End
      Begin MSMask.MaskEdBox Vcto_Final 
         Height          =   315
         Left            =   3300
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   270
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
      Begin MSMask.MaskEdBox Vcto_Inicial 
         Height          =   315
         Left            =   765
         TabIndex        =   0
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   270
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
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2850
         TabIndex        =   35
         Top             =   300
         Width           =   390
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   285
         TabIndex        =   34
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seleção de Cartões"
      Height          =   1575
      Left            =   3690
      TabIndex        =   28
      Top             =   5640
      Width           =   10365
      Begin VB.CommandButton cmdBaixar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Baixar Seleção"
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
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   285
         Width           =   2085
      End
      Begin VB.CommandButton B_Cancela_Sel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Cancelar"
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Cancela atual seleção"
         Top             =   630
         Width           =   2085
      End
      Begin VB.CommandButton Seleciona_Tudo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Selecionar &Tudo"
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Seleciona Todos os Cartões"
         Top             =   165
         Width           =   2085
      End
      Begin VB.CommandButton B_Apaga 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Apagar Seleção"
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Apaga os cartões selecionados"
         Top             =   1095
         Width           =   2085
      End
      Begin VB.Frame Quadro1 
         Caption         =   "Depositar em"
         Height          =   855
         Left            =   2325
         TabIndex        =   30
         Top             =   660
         Visible         =   0   'False
         Width           =   2025
         Begin VB.OptionButton O_Não_Determinado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "&Não determinado"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   270
            TabIndex        =   13
            Top             =   240
            Width           =   1665
         End
         Begin VB.OptionButton O_Conta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "&Conta Corrente"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   270
            TabIndex        =   14
            Top             =   510
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.CommandButton B_Deposita 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Depositar"
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
         TabIndex        =   16
         ToolTipText     =   "Deposita Cartões"
         Top             =   630
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.PictureBox Figura 
         AutoSize        =   -1  'True
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
         Height          =   705
         Left            =   405
         Picture         =   "ManCartoes.frx":4FB32
         ScaleHeight     =   705
         ScaleWidth      =   1770
         TabIndex        =   29
         Top             =   780
         Visible         =   0   'False
         Width           =   1770
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
         Bindings        =   "ManCartoes.frx":5157C
         DataSource      =   "Data2"
         Height          =   315
         Left            =   4455
         TabIndex        =   15
         Top             =   765
         Visible         =   0   'False
         Width           =   1185
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
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   5398
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4366
         Columns(1).Caption=   "Conta"
         Columns(1).Name =   "Conta"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Conta"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1799
         Columns(2).Caption=   "Código"
         Columns(2).Name =   "Código"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Código"
         Columns(2).DataType=   2
         Columns(2).FieldLen=   256
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin VB.Label Label6 
         Caption         =   "Conta"
         Height          =   195
         Left            =   4470
         TabIndex        =   32
         Top             =   540
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Nome_Conta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4470
         TabIndex        =   31
         Top             =   1170
         Visible         =   0   'False
         Width           =   3615
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
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   7260
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   735
      Left            =   8490
      TabIndex        =   21
      Top             =   465
      Width           =   2685
      Begin VB.OptionButton O_Não_Depositados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Não depositado"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   195
         TabIndex        =   4
         Top             =   480
         Width           =   1635
      End
      Begin VB.OptionButton O_Depositados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Depositados"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   195
         TabIndex        =   3
         Top             =   240
         Width           =   1305
      End
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1710
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   825
      End
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
      Left            =   630
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM Cartões ORDER BY Nome"
      Top             =   7260
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   400
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1260
      Width           =   13935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   11220
      TabIndex        =   20
      Top             =   465
      Width           =   2820
      Begin VB.OptionButton O_Valor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Valor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1890
         TabIndex        =   8
         Top             =   210
         Width           =   705
      End
      Begin VB.OptionButton O_Data 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Data Vcto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   450
         Width           =   1080
      End
      Begin VB.OptionButton O_Banco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Administradora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   195
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   1560
      End
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "ManCartoes.frx":51590
      DataSource      =   "datFilial"
      Height          =   330
      Left            =   570
      TabIndex        =   9
      Top             =   82
      Width           =   900
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
      _ExtentX        =   1587
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "ManCartoes.frx":515A8
      DataSource      =   "Data3"
      Height          =   330
      Left            =   7260
      TabIndex        =   41
      Top             =   82
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
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   225
      Left            =   6690
      TabIndex        =   43
      Top             =   135
      Width           =   525
   End
   Begin VB.Label Nome_Fornecedor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8460
      TabIndex        =   42
      Top             =   82
      Width           =   5580
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      Height          =   195
      Left            =   120
      TabIndex        =   38
      Top             =   135
      Width           =   300
   End
   Begin VB.Label lblFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1500
      TabIndex        =   37
      Top             =   82
      Width           =   5085
   End
   Begin VB.Label Tot_Selec 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2370
      TabIndex        =   27
      Top             =   6885
      Width           =   1245
   End
   Begin VB.Label Label5 
      Caption         =   "Não Depositado Selecionado"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6915
      Width           =   2235
   End
   Begin VB.Label Label4 
      Caption         =   "Total Não Depositado"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   6345
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5775
      Width           =   495
   End
   Begin VB.Label Tot_Não 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2370
      TabIndex        =   23
      Top             =   6315
      Width           =   1245
   End
   Begin VB.Label Tot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2370
      TabIndex        =   22
      Top             =   5745
      Width           =   1245
   End
End
Attribute VB_Name = "frmManCartoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCliFor As Recordset
Dim rsContas As Recordset
Dim rsLançamentos As Recordset
Dim gnSomaValSelected As Single
Private gbBaixar As Boolean
Private gnCodAdm As Integer

Sub Soma_Seleção()
  Dim nRowsSelected As Long
  Dim nRow As Long
  Dim bm As Variant

  nRowsSelected = grdCartoes.SelBookmarks.Count

  gnSomaValSelected = 0
  For nRow = 0 To (nRowsSelected - 1)
    bm = grdCartoes.SelBookmarks(nRow)
    grdCartoes.Bookmark = bm
    If grdCartoes.Columns("Depositado").Value = False Then
      gnSomaValSelected = gnSomaValSelected + grdCartoes.Columns("ValDesc").Value
    End If
  Next nRow

  Tot_Selec.Caption = Format(gnSomaValSelected, "###,###,##0.00")

End Sub

Private Sub B_Apaga_Click()
  Dim rsCartoes As Recordset
  Dim nRow As Integer
  Dim bm As Variant
  
  Call StatusMsg("")
  If grdCartoes.SelBookmarks.Count = 0 Then
    DisplayMsg "Realize a seleção de linhas antes de efetuar esta operação."
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente apagar todos os itens selecionados?"
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    Exit Sub
  End If
  
  Set rsCartoes = db.OpenRecordset("SELECT * FROM [Contas a Receber] WHERE Tipo = 'O'", dbOpenDynaset)
  
  Call ws.BeginTrans
  
  For nRow = 0 To (grdCartoes.SelBookmarks.Count - 1)
    bm = grdCartoes.SelBookmarks(nRow)
    grdCartoes.Bookmark = bm
    rsCartoes.FindFirst "Contador = " & grdCartoes.Columns("Contador").Value
    If Not rsCartoes.NoMatch Then
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
        "Cli:" & rsCartoes("Cliente") & "- Seq:" & rsCartoes("Sequência") & "- NF:" & rsCartoes("Nota") & "- Venc:" & rsCartoes("Vencimento") & "- Valor:" & rsCartoes("Valor"), _
        "frmManCartoes_B_Apaga_Click", _
        "Contas a Receber", g_strArquivoSystemLog
      End If
      rsCartoes.Delete
    End If
  Next nRow
  
  Call ws.CommitTrans
  
  rsCartoes.Close
  Set rsCartoes = Nothing
  
  grdCartoes.Update
  
  Call B_Monta_Click

End Sub


Private Sub B_Cancela_Sel_Click()
  grdCartoes.SelBookmarks.RemoveAll
  Call Soma_Seleção
  gbBaixar = True
  Call cmdBaixar_Click
End Sub

Private Sub B_Deposita_Click()
  Dim Saldo_Ant As Currency
  Dim nRow As Long
  Dim bm As Variant
  Dim rsCartoes As Recordset

  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked  As Integer

  If O_Conta.Value = True Then
    If Nome_Conta.Caption = "" Then
      Beep
      DisplayMsg "Informe a conta bancária a depositar."
      Combo_Conta.SetFocus
      Exit Sub
    End If
  End If
  
  If O_Conta.Value = False Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja realizar a operação sem informar a conta bancária?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      DisplayMsg "Operação não realizada."
      Exit Sub
    End If
    If Not frmGerente.gbSenhaGerente Then
      Exit Sub
    End If
  End If
  
  On Error GoTo ErrHandler

  Set rsCartoes = db.OpenRecordset("SELECT * FROM [Contas a Receber] WHERE Tipo = 'O'", dbOpenDynaset)
  
  Call ws.BeginTrans
  blnInTransaction = True
  
  With rsCartoes
    For nRow = 0 To (grdCartoes.SelBookmarks.Count - 1)
      bm = grdCartoes.SelBookmarks(nRow)
      grdCartoes.Bookmark = bm
      .FindFirst "Contador = " & grdCartoes.Columns("Contador").Value
      If Not .NoMatch Then
        If .Fields("Processado").Value = False Then
          '10/09/2007 - Anderson
          'Gera arquivo log do sistema
          If g_bolSystemLog Then
            SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
            "Cli:" & rsCartoes("Cliente") & "- Seq:" & rsCartoes("Sequência") & "- NF:" & rsCartoes("Nota") & "- Venc:" & rsCartoes("Vencimento") & "- Valor:" & rsCartoes("Valor"), _
            "frmManCartoes_B_Deposita_Click", _
            "Contas a Receber", g_strArquivoSystemLog
          End If
          .LockEdits = True
          .Edit
          .Fields("Processado").Value = True
          .Fields("Valor Recebido").Value = .Fields("Valor").Value
          .Fields("Data Recebimento") = Data_Atual
          .Fields("Data Alteração") = Data_Atual
          .Update
        End If
      End If
    Next nRow
  End With
  
  'Atualiza Dinheiro na conta, se for o caso
  rsLançamentos.Index = "Conta"
  rsLançamentos.Seek "<", Val(Combo_Conta.Text), CDate(Data_Atual), 99999999#
  If rsLançamentos.NoMatch Then Saldo_Ant = 0
  If Not rsLançamentos.NoMatch Then
    Saldo_Ant = 0
    If rsLançamentos("Conta") = Val(Combo_Conta.Text) Then
      Saldo_Ant = rsLançamentos("Saldo Atual")
    End If
  End If
  
  rsLançamentos.AddNew
  rsLançamentos("Conta") = Val(Combo_Conta.Text)
  rsLançamentos("Data") = Data_Atual
  rsLançamentos("Descrição") = "Cartões recebidos"
  rsLançamentos("Saldo Anterior") = Saldo_Ant
  rsLançamentos("Crédito") = CDbl(gnSomaValSelected)
  rsLançamentos("Saldo Atual") = Saldo_Ant + CDbl(gnSomaValSelected)
  rsLançamentos.Update
   
  Call ws.CommitTrans
  blnInTransaction = False
  
  rsCartoes.Close
  Set rsCartoes = Nothing
  
  DisplayMsg "Cartões depositados."
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Cnt:" & Combo_Conta.Text & " SldAn:" & Saldo_Ant & " SldAt:" & rsLançamentos("Saldo Atual") & " Nm:" & Nome_Conta.Caption, 80) & "', 'BAIXAR CARTOES')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
  
  grdCartoes.Enabled = True
  
  Combo_Conta.Text = ""
  Nome_Conta = ""
  
  B_Monta.Enabled = True
  B_Cancela_Sel_Click
  B_Monta_Click
  
  Exit Sub
  
ErrHandler:
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

Private Sub B_Monta_Click()
  
  Call StatusMsg("")
  
  If Not IsDate(Vcto_Inicial.Text) Then
    DisplayMsg "Vencimento inicial incorreto."
    Vcto_Inicial.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Vcto_Final.Text) Then
    DisplayMsg "Vencimento final incorreto."
    Vcto_Final.SetFocus
    Exit Sub
  End If
  
  If CDate(Vcto_Final.Text) < CDate(Vcto_Inicial.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual à data final."
    Vcto_Inicial.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(cboFilial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If
  
  B_Monta.Enabled = False
  Call LoadGridCartoes
  B_Monta.Enabled = True
  
End Sub

Private Sub cboAdm_CloseUp()
  Dim bm As Variant
  bm = cboAdm.GetBookmark(0)
  gnCodAdm = cboAdm.Columns("Código").CellValue(bm)
End Sub

Private Sub cboAdm_LostFocus()
  If Len(Trim(cboAdm.Text)) = 0 Then
    gnCodAdm = 0
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

Private Sub cmd_calendarioDtFim_Click()
    Vcto_Final.Text = frmCalendario.gsDateCalender(Vcto_Final.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Vcto_Inicial.Text = frmCalendario.gsDateCalender(Vcto_Inicial.Text)
End Sub

Private Sub cmdBaixar_Click()
  If gbBaixar = False Then
    Call StatusMsg("")
    If gnSomaValSelected = 0 Then
      DisplayMsg "Selecione antes um ou mais cartões não depositados para a operação."
      Exit Sub
    End If
    grdCartoes.Enabled = False
    Seleciona_Tudo.Enabled = False
    B_Monta.Enabled = False
    Figura.Visible = True
    Quadro1.Visible = True
    Label6.Visible = True
    Combo_Conta.Visible = True
    Nome_Conta.Visible = True
    B_Deposita.Visible = True
    Combo_Conta.SetFocus
  Else
    grdCartoes.Enabled = True
    Seleciona_Tudo.Enabled = True
    B_Monta.Enabled = True
    B_Monta.SetFocus
    Figura.Visible = False
    Quadro1.Visible = False
    Label6.Visible = False
    Combo_Conta.Visible = False
    Nome_Conta.Visible = False
    B_Deposita.Visible = False
  End If
  gbBaixar = Not gbBaixar
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
  '28/11/2006 - Anderson
  'Alteração do número de contas bancárias de 99 para 255
  'Solicitado por: 2227883 - SANTA FÉ DO ARAGUAIA PREFEITURA MUNICIPAL
   If Val(Combo_Conta.Text) > 255 Then Exit Sub
   
   rsContas.Index = "Código"
   rsContas.Seek "=", Val(Combo_Conta.Text)
   If rsContas.NoMatch Then Exit Sub
   Nome_Conta.Caption = rsContas("Descrição")
   
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

Private Sub Form_Load()
  
  Call CenterForm(Me)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários")
  gnSomaValSelected = 0
  
  Call GetSettings
  gnCodAdm = 0
  
  cboFilial.Text = gnCodFilial
  cboFilial_LostFocus
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Call SaveSetting("QuickStore", "CRMan", "Data1", Vcto_Inicial.Text)
  Call SaveSetting("QuickStore", "CRMan", "Data2", Vcto_Final.Text)
  
  rsCliFor.Close
  rsContas.Close
  rsLançamentos.Close
  Set rsContas = Nothing
  Set rsLançamentos = Nothing
  Set rsCliFor = Nothing
  
End Sub

Private Sub GetSettings()
  Vcto_Final.Text = GetSetting("QuickStore", "CRMan", "Data2", CDate(Date))
  Vcto_Inicial.Text = GetSetting("QuickStore", "CRMan", "Data1", CDate(Date))
End Sub

Private Sub grdCartoes_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  Call StatusMsg("")
  gsTitle = LoadResString(201)
  gsMsg = "Deseja apagar a seleção atual?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    DisplayMsg "Seleção não apagada."
    Cancel = True
  End If
  DispPromptMsg = False
End Sub

Private Sub grdCartoes_InitColumnProps()
  grdCartoes.Columns("Valor").NumberFormat = "##,###,##0.00"
  grdCartoes.Columns("ValDesc").NumberFormat = "##,###,##0.00"
  grdCartoes.Columns("ValReceb").NumberFormat = "##,###,##0.00"
End Sub

Private Sub grdCartoes_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

  Call Soma_Seleção
  Cancel = False

End Sub

Private Sub Seleciona_Tudo_Click()
  Dim nRow As Integer
  
  grdCartoes.MoveFirst
  For nRow = 0 To (grdCartoes.Rows - 1)
    grdCartoes.SelBookmarks.Add grdCartoes.Bookmark
    grdCartoes.MoveNext
  Next nRow
  grdCartoes.MoveFirst
  
  Call Soma_Seleção
  
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

Private Sub LoadGridCartoes()
  Dim rsCartoes As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sSql As String
  Dim Data_Ini As String
  Dim Data_Fim As String
  Dim Tot_Não_Depositado As Currency
  Dim Total As Currency
  Dim sFilial As String
  
  On Error GoTo ErrHandler
  
  'Verifica a filial
  cboFilial_LostFocus
  If lblFilial.Caption = "" Then
    sFilial = "<> 0"
  Else
    sFilial = "= " & cboFilial
  End If
  
  Data_Ini = gsGetInvDate(Vcto_Inicial.Text)
  Data_Fim = gsGetInvDate(Vcto_Final.Text)
  
  bAllow = grdCartoes.AllowAddNew
  grdCartoes.AllowAddNew = True
  grdCartoes.AllowUpdate = True
  
  sSql = "SELECT Filial, Administradora AS Adm, Cartões.Nome as Nome, Cartão, Vencimento, "
  sSql = sSql & "[Valor Cartão] as Val1, Valor as Val2, Processado, "
  sSql = sSql & "[Valor Recebido] as ValReceb, Contador FROM [Contas a Receber]"
  sSql = sSql + " INNER JOIN Cartões ON [Contas a Receber].Administradora = Cartões.Código"
  sSql = sSql + " WHERE Filial " & sFilial & " AND Vencimento >= " + Data_Ini
  
  If Not IsNull(Combo_Fornecedor.Text) And Len(Combo_Fornecedor.Text) > 0 And Len(Nome_Fornecedor.Caption) > 0 Then
    sSql = sSql + " And Cliente = " + Combo_Fornecedor.Text + " "
  End If
  
  sSql = sSql + " And Vencimento <= " + Data_Fim + " AND Tipo = 'O'"
  
  If gnCodAdm > 0 Then
    sSql = sSql & " And [Contas a Receber].Administradora = " & gnCodAdm
  End If
  
  If O_Depositados.Value = True Then
    sSql = sSql + " AND Processado = True"
  Else
    If O_Não_Depositados.Value = True Then
      sSql = sSql + " AND Processado = False"
    End If
  End If
  
  If O_Banco.Value = True Then sSql = sSql + " ORDER BY Cartões.Nome"
  If O_Data.Value = True Then sSql = sSql + " ORDER BY Vencimento"
  If O_Valor.Value = True Then sSql = sSql + " ORDER BY Valor"
  
  Call StatusMsg("Aguarde, montando tabela...")
  
  Set rsCartoes = db.OpenRecordset(sSql, dbOpenDynaset)

  grdCartoes.RemoveAll
  grdCartoes.Redraw = False
  
  If Not rsCartoes.EOF Then
    With rsCartoes
      .MoveFirst
      Do While Not .EOF
        DoEvents
        sRecord = .Fields("Filial") & vbTab & _
          .Fields("Adm") & vbTab & _
          .Fields("Nome") & vbTab & _
          .Fields("Cartão") & vbTab & _
          .Fields("Vencimento") & vbTab & _
          .Fields("Val1") & vbTab & _
          .Fields("Val2") & vbTab & _
          .Fields("Processado") & vbTab & _
          .Fields("ValReceb") & vbTab & _
          .Fields("Contador")
        grdCartoes.AddItem sRecord
        .MoveNext
      Loop
      .MoveFirst
    End With
    
    'Verifica Totais
    Tot_Não_Depositado = 0
    Total = 0
    rsCartoes.MoveFirst
    Do Until rsCartoes.EOF
      If rsCartoes("Processado") = False Then
        Tot_Não_Depositado = Tot_Não_Depositado + rsCartoes("Val2").Value
      End If
      Total = Total + rsCartoes("Val2").Value
      rsCartoes.MoveNext
    Loop
    
    Tot.Caption = Format(Total, "###,###,##0.00")
    Tot_Não.Caption = Format(Tot_Não_Depositado, "###,###,##0.00")
    Tot_Selec.Caption = "0,00"
    
  Else
  
    DisplayMsg "Nenhuma conta de cartão de crédito encontrada segundo os critérios fornecidos."
  
  End If
  
  gnSomaValSelected = 0

  Call StatusMsg("")
  
  grdCartoes.Scroll -99, -99
  grdCartoes.Redraw = True
  
  grdCartoes.AllowAddNew = bAllow
  grdCartoes.AllowUpdate = bAllow

  rsCartoes.Close
  Set rsCartoes = Nothing
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao ler registros do Contas a Receber."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

