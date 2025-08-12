VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelAcompaEstoq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Movimentação de Estoque Simplificado no Período"
   ClientHeight    =   4005
   ClientLeft      =   1320
   ClientTop       =   1740
   ClientWidth     =   14340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelAcompaEstoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   14340
   Begin VB.TextBox txt_fabricanteMarca 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2415
      TabIndex        =   41
      Top             =   1350
      Width           =   4605
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
      Height          =   315
      Left            =   7830
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sub Classes"
      Top             =   3930
      Visible         =   0   'False
      Width           =   1770
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
      Height          =   315
      Left            =   6030
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Classes"
      Top             =   3930
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame5 
      Caption         =   "Período"
      Height          =   795
      Left            =   120
      TabIndex        =   29
      Top             =   1800
      Width           =   6900
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
         Left            =   2910
         Picture         =   "RelAcompaEstoque.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
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
         Left            =   5940
         Picture         =   "RelAcompaEstoque.frx":4F23C
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   465
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   4650
         TabIndex        =   7
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
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
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3780
         TabIndex        =   31
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   660
         TabIndex        =   30
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Opções"
      Height          =   795
      Left            =   7170
      TabIndex        =   28
      Top             =   1785
      Width           =   7110
      Begin VB.CheckBox chkInativo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Desconsiderar produtos inativos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4380
         TabIndex        =   10
         Top             =   255
         Value           =   1  'Checked
         Width           =   2610
      End
      Begin VB.CheckBox O_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Separar por classe"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   195
         TabIndex        =   8
         Top             =   255
         Width           =   2955
      End
      Begin VB.CheckBox O_Período 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Desconsiderar produtos sem movimento no período"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   195
         TabIndex        =   9
         Top             =   510
         Width           =   4635
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordem"
      Height          =   675
      Left            =   3270
      TabIndex        =   27
      Top             =   2670
      Width           =   3750
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1830
         TabIndex        =   17
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   390
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   690
      Left            =   7140
      TabIndex        =   26
      Top             =   2670
      Width           =   7110
      Begin VB.OptionButton O_Edição 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Edição"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2400
         TabIndex        =   15
         Top             =   285
         Width           =   840
      End
      Begin VB.OptionButton O_Grade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Grade"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4380
         TabIndex        =   14
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton O_Normal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Normal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
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
      Height          =   465
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3420
      Width           =   14130
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   675
      Left            =   120
      TabIndex        =   25
      Top             =   2670
      Width           =   3105
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   390
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Data Data3 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   3720
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   3930
      Visible         =   0   'False
      Width           =   1770
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
      Height          =   315
      Left            =   1920
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   3930
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   210
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3930
      Visible         =   0   'False
      Width           =   1665
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   5550
      Top             =   3870
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Prod 
      Bindings        =   "RelAcompaEstoque.frx":4FB1E
      DataSource      =   "Data3"
      Height          =   345
      Left            =   8220
      TabIndex        =   4
      Top             =   555
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
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "RelAcompaEstoque.frx":4FB32
      DataSource      =   "Data2"
      Height          =   345
      Left            =   8220
      TabIndex        =   1
      Top             =   150
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9948
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
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelAcompaEstoque.frx":4FB46
      DataSource      =   "Data1"
      Height          =   345
      Left            =   990
      TabIndex        =   0
      Top             =   150
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9340
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1614
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Prod_Final 
      Bindings        =   "RelAcompaEstoque.frx":4FB5A
      DataSource      =   "Data3"
      Height          =   345
      Left            =   8220
      TabIndex        =   5
      Top             =   945
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
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "RelAcompaEstoque.frx":4FB6E
      DataSource      =   "Data4"
      Height          =   345
      Left            =   990
      TabIndex        =   2
      Top             =   555
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
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_SubClasse 
      Bindings        =   "RelAcompaEstoque.frx":4FB82
      DataSource      =   "Data5"
      Height          =   345
      Left            =   990
      TabIndex        =   3
      Top             =   945
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
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Fabricante / Marca"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   1410
      Width           =   1875
   End
   Begin VB.Label Nome_SubClasse 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2415
      TabIndex        =   37
      Top             =   945
      Width           =   4605
   End
   Begin VB.Label Nome_Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2415
      TabIndex        =   36
      Top             =   555
      Width           =   4605
   End
   Begin VB.Label Nome_Prod_Final 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9645
      TabIndex        =   35
      Top             =   945
      Width           =   4605
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      Caption         =   "Sub-Classe"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   34
      Top             =   990
      Width           =   825
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      Caption         =   "Classe"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Produto Final"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7170
      TabIndex        =   32
      Top             =   990
      Width           =   1005
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      Caption         =   "Produto Inicial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7170
      TabIndex        =   24
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Nome_Prod 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9645
      TabIndex        =   23
      Top             =   555
      Width           =   4605
   End
   Begin VB.Label Nome_Fornecedor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9645
      TabIndex        =   22
      Top             =   150
      Width           =   4605
   End
   Begin VB.Label Label2 
      Caption         =   "Fornecedor"
      Height          =   255
      Left            =   7170
      TabIndex        =   21
      Top             =   195
      Width           =   825
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2415
      TabIndex        =   20
      Top             =   150
      Width           =   4605
   End
   Begin VB.Label Label1 
      Caption         =   "Filial"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   195
      Width           =   435
   End
End
Attribute VB_Name = "frmRelAcompaEstoq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsProdutos2 As Recordset
Dim rsForn_Prod As Recordset
Dim rsFornecedores As Recordset
Dim rsTempo As Recordset
Dim rsEstoque As Recordset
Dim rsAcompa As Recordset
Dim rsGrade As Recordset
Dim rsEdicoes As Recordset
Dim rsTamanhos As Recordset
Dim rsCores As Recordset
Dim rsClasses As Recordset
Dim rsSubclasses As Recordset
Dim Código As String
Dim Código_Completo As String
Dim Cor As Integer
Dim Tamanho As Integer
Dim Edição As Long
Dim Aux_Cor As Integer
Dim Aux_Tamanho As Integer
Dim Aux_Edição As Long
Dim Mês As Integer


Function Estoque_Anterior(Produto As String, Tamanho As Integer, Cor As Integer, Edição As Long) As Single
 Dim Est As Single
  

  rsEstoque.Index = "Data2"
  
  Est = 0
  
  rsEstoque.Seek ">", Val(Combo.Text), Produto, Tamanho, Cor, Edição, CDate(Data_Ini.Text)
  If Not rsEstoque.NoMatch Then
    If rsEstoque("Filial") = Val(Combo.Text) Then
      If rsEstoque("Produto") = Produto Then
        If rsEstoque("Tamanho") = Tamanho Then
          If rsEstoque("Cor") = Cor Then
            If rsEstoque("Edição") = Edição Then
              Est = rsEstoque("Estoque Final")
            End If
          End If
        End If
      End If
    End If
  End If
  
  
  Estoque_Anterior = Est
  
  
  
      
    


End Function

Sub Grava_Produto()
 Dim Fim As Integer
 Dim Aux_Grade As String
 Dim Aux_Str As String
 Dim Estoque As Single
 Dim T_Entrada As Single
 Dim Compras As Single
 Dim T_Saída As Single
 Dim Vendas As Single
 Dim Aux_Edi As Long
 Dim Aux_Classe As Integer
 Dim Aux_Sub As Integer
  
  
  Fim = False
  rsGrade.Index = "Original"
  rsEdicoes.Index = "Produto"
  
  Aux_Grade = ""
  Aux_Edi = 0
  
  Estoque = 0
  T_Entrada = 0
  Compras = 0
  T_Saída = 0
  
  Do
   If rsProdutos2("Tipo") = "N" Then Fim = True
   
   Rem Acha estoque anterior
   Aux_Tamanho = 0
   Aux_Cor = 0
   Aux_Edição = 0
   
   If rsProdutos2("Tipo") = "G" Then
     
     rsGrade.Seek ">", Código, Aux_Grade
     If Not rsGrade.NoMatch Then
       If rsGrade("Código Original") = Código Then
         Aux_Grade = rsGrade("Código")
         Aux_Str = Trim(Right(rsGrade("Código"), 6))
         Aux_Tamanho = Val(Left(Aux_Str, 3))
         Aux_Cor = Val(Right(Aux_Str, 3))
           
         Estoque = Estoque + Estoque_Anterior(Código, Aux_Tamanho, Aux_Cor, 0)
         T_Entrada = T_Entrada + Trans_Entrada(Código, Aux_Tamanho, Aux_Cor, 0, "TE")
         Compras = Compras + Trans_Entrada(Código, Aux_Tamanho, Aux_Cor, 0, "CO")
         T_Saída = T_Saída + Trans_Entrada(Código, Aux_Tamanho, Aux_Cor, 0, "TS")
         Vendas = Vendas + Trans_Entrada(Código, Aux_Tamanho, Aux_Cor, 0, "VE")
       Else
         Fim = True
       End If
     Else
       Fim = True
     End If
     
   End If  ' If Tipo = "G"
   
    
    
    
   If rsProdutos2("Tipo") = "E" Then
      rsEdicoes.Seek ">", Código, Aux_Edi
      If Not rsEdicoes.NoMatch Then
        If rsEdicoes("Produto") = Código Then
          Aux_Edi = rsEdicoes("Código")
          
          Estoque = Estoque + Estoque_Anterior(Código, 0, 0, Aux_Edi)
          T_Entrada = T_Entrada + Trans_Entrada(Código, 0, 0, Aux_Edi, "TE")
          Compras = Compras + Trans_Entrada(Código, 0, 0, Aux_Edi, "CO")
          T_Saída = T_Saída + Trans_Entrada(Código, 0, 0, Aux_Edi, "TS")
          Vendas = Vendas + Trans_Entrada(Código, 0, 0, Aux_Edi, "VE")
        
        Else
         Fim = True
        End If
      Else
       Fim = True
      End If
      
   End If
      
   
   If rsProdutos2("Tipo") = "N" Then
     '14/01/2005 - Daniel
     'Em algumas bases de dados o campo Produtos.Código está
     'aparecendo com caracteres incorretos tais como ...
     'isto estava gerando o BUG com mensagem 3163
     '
     'Case: São Francisco Móveis e Eletro. de Olinda - PE
     If Len(Código) > 20 Then Código = "0"
     '-------------------------------------------------------
     
     Estoque = Estoque_Anterior(Código, 0, 0, 0)
     T_Entrada = Trans_Entrada(Código, 0, 0, 0, "TE")
     Compras = Trans_Entrada(Código, 0, 0, 0, "CO")
     T_Saída = Trans_Entrada(Código, 0, 0, 0, "TS")
     Vendas = Trans_Entrada(Código, 0, 0, 0, "VE")
     
   End If
   
   
   
  
  Loop While Fim = False
  
  Aux_Classe = 0
  rsClasses.Index = "Código"
  rsClasses.Seek "=", rsProdutos2("Classe")
  If Not rsClasses.NoMatch Then Aux_Classe = rsProdutos2("Classe")
  
  Aux_Sub = 0
  rsSubclasses.Index = "Código"
  rsSubclasses.Seek "=", rsProdutos2("Sub Classe")
  If Not rsSubclasses.NoMatch Then Aux_Sub = rsProdutos2("Sub Classe")
  
  
  

  rsAcompa.Index = "Código"
  rsAcompa.Seek "=", Código, 0, 0, 0
  If Not rsAcompa.NoMatch Then rsAcompa.Delete

  rsAcompa.AddNew
    rsAcompa("Código") = Código
    rsAcompa("Saldo Anterior") = Estoque
    rsAcompa("Transf Entrada") = T_Entrada
    rsAcompa("Compras") = Compras
    rsAcompa("Transf Saída") = T_Saída
    rsAcompa("Vendas") = Vendas
    
    rsAcompa("Saldo") = Estoque + T_Entrada + Compras - T_Saída - Vendas
    
    If Estoque <> 0 Then
       rsAcompa("Giro") = Vendas * 100 / Estoque
    Else
       rsAcompa("Giro") = 0
    End If
    
    rsAcompa("Classe") = Aux_Classe
    If Aux_Classe <> 0 Then rsAcompa("Nome Classe") = rsClasses("Nome")
    rsAcompa("Sub Classe") = Aux_Sub
    If Aux_Sub <> 0 Then rsAcompa("Nome Sub") = rsSubclasses("Nome")
    
    rsAcompa("Nome") = rsProdutos2("Nome")
    rsAcompa("Ordenação") = rsProdutos2("Código Ordenação")
    rsAcompa("Fracionado") = rsProdutos2("Fracionado")
  rsAcompa.Update
  
  
  
    



End Sub

Sub Grava_Produto_Grade()
 Dim Fim As Integer
 Dim Aux_Grade As String
 Dim Aux_Str As String
 Dim Estoque As Single
 Dim T_Entrada As Single
 Dim Compras As Single
 Dim T_Saída As Single
 Dim Vendas As Single
 Dim Aux_Edi As Long
 Dim Aux_Classe As Integer
 Dim Aux_Sub As Integer
  
  
  Fim = False
  rsGrade.Index = "Original"
  rsEdicoes.Index = "Produto"
  
  Aux_Grade = ""
  Aux_Edi = 0
  Aux_Tamanho = 0
  Aux_Cor = 0
  
  Estoque = 0
  T_Entrada = 0
  Compras = 0
  T_Saída = 0
  
  Do
   If rsProdutos2("Tipo") = "N" Then Fim = True
   
   Rem Acha estoque anterior
   Aux_Tamanho = 0
   Aux_Cor = 0
   Aux_Edição = 0
   
   If rsProdutos2("Tipo") = "G" Then
     
     rsGrade.Seek ">", Código, Aux_Grade
     If Not rsGrade.NoMatch Then
       If rsGrade("Código Original") = Código Then
         Aux_Grade = rsGrade("Código")
         Aux_Str = Trim(Right(rsGrade("Código"), 6))
         Aux_Tamanho = Val(Left(Aux_Str, 3))
         Aux_Cor = Val(Right(Aux_Str, 3))
           
         Estoque = Estoque_Anterior(Código, Aux_Tamanho, Aux_Cor, 0)
         T_Entrada = Trans_Entrada(Código, Aux_Tamanho, Aux_Cor, 0, "TE")
         Compras = Trans_Entrada(Código, Aux_Tamanho, Aux_Cor, 0, "CO")
         T_Saída = Trans_Entrada(Código, Aux_Tamanho, Aux_Cor, 0, "TS")
         Vendas = Trans_Entrada(Código, Aux_Tamanho, Aux_Cor, 0, "VE")
         
         GoSub Grava_Registro
       Else
         Fim = True
       End If
     Else
       Fim = True
     End If
     
   End If  ' If Tipo = "G"
   
    
        
   If rsProdutos2("Tipo") = "E" Then
      rsEdicoes.Seek ">", Código, Aux_Edi
      If Not rsEdicoes.NoMatch Then
        If rsEdicoes("Produto") = Código Then
          Aux_Edi = rsEdicoes("Código")
          
          Estoque = Estoque_Anterior(Código, 0, 0, Aux_Edi)
          T_Entrada = Trans_Entrada(Código, 0, 0, Aux_Edi, "TE")
          Compras = Trans_Entrada(Código, 0, 0, Aux_Edi, "CO")
          T_Saída = Trans_Entrada(Código, 0, 0, Aux_Edi, "TS")
          Vendas = Trans_Entrada(Código, 0, 0, Aux_Edi, "VE")
        
          GoSub Grava_Registro
        Else
         Fim = True
        End If
      Else
       Fim = True
      End If
      
   End If
      
   
   
   If rsProdutos2("Tipo") = "N" Then
     Estoque = Estoque_Anterior(Código, 0, 0, 0)
     T_Entrada = Trans_Entrada(Código, 0, 0, 0, "TE")
     Compras = Trans_Entrada(Código, 0, 0, 0, "CO")
     T_Saída = Trans_Entrada(Código, 0, 0, 0, "TS")
     Vendas = Trans_Entrada(Código, 0, 0, 0, "VE")
     
     GoSub Grava_Registro
     
   End If
   
 
   
  
  Loop While Fim = False
  
  
  Exit Sub

Grava_Registro:
  Aux_Classe = 0
  rsClasses.Index = "Código"
  rsClasses.Seek "=", rsProdutos2("Classe")
  If Not rsClasses.NoMatch Then Aux_Classe = rsProdutos2("Classe")
  
  Aux_Sub = 0
  rsSubclasses.Index = "Código"
  rsSubclasses.Seek "=", rsProdutos2("Sub Classe")
  If Not rsSubclasses.NoMatch Then Aux_Sub = rsProdutos2("Sub Classe")
  
  If Aux_Cor <> 0 Then
    rsCores.Index = "Código"
    rsCores.Seek "=", Aux_Cor
    If rsCores.NoMatch Then Aux_Cor = 0
  End If
  
  If Aux_Tamanho <> 0 Then
    rsTamanhos.Index = "Código"
    rsTamanhos.Seek "=", Aux_Tamanho
    If rsTamanhos.NoMatch Then Aux_Tamanho = 0
  End If
  

  rsAcompa.Index = "Código"
  rsAcompa.Seek "=", Código, Aux_Tamanho, Aux_Cor, Aux_Edi
  If Not rsAcompa.NoMatch Then rsAcompa.Delete

  rsAcompa.AddNew
    rsAcompa("Código") = Código
    rsAcompa("Tamanho") = Aux_Tamanho
    If Aux_Tamanho <> 0 Then rsAcompa("Nome Tamanho") = rsTamanhos("Nome")
    rsAcompa("Cor") = Aux_Cor
    If Aux_Cor <> 0 Then rsAcompa("Nome Cor") = rsCores("Nome")
    rsAcompa("Edição") = Aux_Edi
    rsAcompa("Saldo Anterior") = Estoque
    rsAcompa("Transf Entrada") = T_Entrada
    rsAcompa("Compras") = Compras
    rsAcompa("Transf Saída") = T_Saída
    rsAcompa("Vendas") = Vendas
    
    rsAcompa("Saldo") = Estoque + T_Entrada + Compras - T_Saída - Vendas
    
    If Estoque <> 0 Then
       rsAcompa("Giro") = Vendas * 100 / Estoque
    Else
       rsAcompa("Giro") = 0
    End If
    
    rsAcompa("Classe") = Aux_Classe
    If Aux_Classe <> 0 Then rsAcompa("Nome Classe") = rsClasses("Nome")
    rsAcompa("Sub Classe") = Aux_Sub
    If Aux_Sub <> 0 Then rsAcompa("Nome Sub") = rsSubclasses("Nome")
    
    rsAcompa("Nome") = rsProdutos2("Nome")
    rsAcompa("Ordenação") = rsProdutos2("Código Ordenação")
    rsAcompa("Fracionado") = rsProdutos2("Fracionado")
  rsAcompa.Update
  
  Return
  

End Sub

Function Trans_Entrada(Produto As String, Tamanho As Integer, Cor As Integer, Edição As Long, Tipo As String) As Single
  Dim Est As Single
  Dim Aux_Data As Date
  Dim sCriteria As String
  
  Est = 0
  Aux_Data = CDate(Data_Ini.Text)

  rsEstoque.Index = "Data"
  sCriteria = ">="
  
Lp1:
  rsEstoque.Seek sCriteria, Val(Combo.Text), Produto, Tamanho, Cor, Edição, Aux_Data
  If rsEstoque.NoMatch Then GoTo Fim_Func
  If rsEstoque("Filial") <> Val(Combo.Text) Then GoTo Fim_Func
  If rsEstoque("Data") > CDate(Data_Fim.Text) Then GoTo Fim_Func
  If rsEstoque("Produto") <> Produto Then GoTo Fim_Func
  If rsEstoque("Tamanho") <> Tamanho Then GoTo Fim_Func
  If rsEstoque("Cor") <> Cor Then GoTo Fim_Func
  If rsEstoque("Edição") <> Edição Then GoTo Fim_Func
  
  Aux_Data = rsEstoque("Data")
  
  '26/02/2007 - Anderson
  'O relatório não estava considerando a quantidade de devoluções de mercadoria
  'If Tipo = "TE" Then Est = Est + rsEstoque("Transf Entra") + rsEstoque("Ajuste Entra") + rsEstoque("Grátis Entra") + rsEstoque("Empre Entra")
  If Tipo = "TE" Then Est = Est + rsEstoque("Transf Entra") + rsEstoque("Ajuste Entra") + rsEstoque("Grátis Entra") + rsEstoque("Empre Entra") + rsEstoque("Devolução")
  
  If Tipo = "CO" Then Est = Est + rsEstoque("Compras")
  
  If Tipo = "TS" Then Est = Est + rsEstoque("Transf Saída") + rsEstoque("Ajuste Saída") + rsEstoque("Grátis Saída") + rsEstoque("Empre Saída")
  
  If Tipo = "VE" Then Est = Est + rsEstoque("Vendas")
  
  sCriteria = ">"
  
  GoTo Lp1



Fim_Func:
  Trans_Entrada = Est
  


End Function

Private Sub B_Imprime_Click()
  Dim sSql As String
  Dim Str1 As String
  Dim Str_Rel As String
  Dim Fornecedor As Long
  Dim Erro As Integer
  Dim gsCriteria As String
  Dim rsSelection As Recordset
  
  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Data inicial inválida. "
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Data final inválida."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data final deve ser superior a data inicial."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(Combo_Classe.Text) And Trim(Combo_Classe.Text) <> "" Then
    DisplayMsg "Código da classe inválida!"
    Combo_Classe.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(Combo_SubClasse.Text) And Trim(Combo_SubClasse.Text) <> "" Then
    DisplayMsg "Código da sub classe inválida!"
    Combo_SubClasse.SetFocus
    Exit Sub
  End If
  
  If Trim(Combo_Fornecedor.Text) = "" Then
    Combo_Fornecedor.Text = "0"
  End If
  
  If Trim(Combo_Classe.Text) = "" Then
    Combo_Classe.Text = "0"
  End If
  
  If Trim(Combo_SubClasse.Text) = "" Then
    Combo_SubClasse.Text = "0"
  End If
  
  If Trim(Combo_Prod.Text) = "" Then
    Combo_Prod.Text = "0"
  End If
  
  If Trim(Combo_Prod_Final.Text) = "" Then
    Combo_Prod_Final.Text = "0"
  End If
  
  Rem Agora zera o arquivo zzz4

  Call StatusMsg("Aguarde, preparando arquivo temporário ...")
  sSql = "Delete * From [Acompa Estoque]"
  dbTemp.Execute sSql
  Call StatusMsg("")

  gsCriteria = ""
  
  Fornecedor = 0
  If Nome_Fornecedor.Caption <> "" Then
    Fornecedor = Combo_Fornecedor.Text
  End If
  
  If Len(Trim(Combo_Classe.Text)) > 0 Then
    If Combo_Classe.Text <> "0" Then
      If Len(gsCriteria) > 0 Then
          gsCriteria = gsCriteria & " AND "
      End If
      gsCriteria = gsCriteria & "Classe = " & Combo_Classe.Text & ""
    End If
  End If
  
  If Len(Trim(Combo_SubClasse.Text)) > 0 Then
    If Combo_SubClasse.Text <> "0" Then
      If Len(gsCriteria) > 0 Then
          gsCriteria = gsCriteria & " AND "
      End If
      gsCriteria = gsCriteria & "[Sub Classe] = " & Combo_SubClasse.Text
    End If
  End If
  
  If Len(Trim(Combo_Prod.Text)) > 0 Then
    If Combo_Prod.Text <> "0" Then
      If Len(gsCriteria) > 0 Then
        gsCriteria = gsCriteria & " AND "
      End If
      If IsNumeric(Combo_Prod.Text) Then
        gsCriteria = gsCriteria & "[Código Ordenação]>= '" & String(20 - Len(Combo_Prod.Text), "+") & Combo_Prod.Text & "'"
      Else
        gsCriteria = gsCriteria & "[Código Ordenação]>= '" & Combo_Prod.Text & "'"
      End If
    End If
  End If

  If Len(Trim(Combo_Prod_Final.Text)) > 0 Then
      If Combo_Prod_Final.Text <> "0" Then
          If Len(gsCriteria) > 0 Then
              gsCriteria = gsCriteria & " AND "
          End If
          If IsNumeric(Combo_Prod.Text) Then
              gsCriteria = gsCriteria & "[Código Ordenação]<= '" & String(20 - Len(Combo_Prod_Final.Text), "+") & Combo_Prod_Final.Text & "'"
          Else
              gsCriteria = gsCriteria & "[Código Ordenação]<= '" & Combo_Prod_Final.Text & "'"
          End If
      End If
  End If
  
  If Len(txt_fabricanteMarca.Text) > 0 Then
      If Len(gsCriteria) > 0 Then
          gsCriteria = gsCriteria & " AND "
      End If
      
      gsCriteria = gsCriteria & " fabricante = '" & txt_fabricanteMarca.Text & "' "
  End If
  
  If Len(gsCriteria) > 0 Then
      gsCriteria = gsCriteria & " AND "
  End If
  If chkInativo.Value = vbChecked Then
      gsCriteria = gsCriteria & "Desativado = False"
  Else
      gsCriteria = gsCriteria & "Desativado = True OR Desativado = False"
  End If
  
  rsForn_Prod.Index = "Produto"

  '03/07/2006 - Anderson
  'rsProdutos2.FindFirst gsCriteria
  Set rsProdutos2 = db.OpenRecordset("SELECT * FROM Produtos WHERE " & gsCriteria & " ORDER BY [Código Ordenação]", dbOpenDynaset, dbReadOnly)
    
  'Do While Not rsProdutos2.NoMatch
  Do While Not rsProdutos2.EOF
    Código = rsProdutos2("Código")
    
    Call StatusMsg("Verificando produto " & Código)
    
    If Fornecedor <> 0 Then
      rsForn_Prod.Seek "=", Código, Fornecedor
      If Not rsForn_Prod.NoMatch Then
        If O_Normal.Value = True Then Grava_Produto
        If O_Grade.Value = True Then Grava_Produto_Grade
        If O_Edição.Value = True Then Grava_Produto_Grade
      End If
    Else
      If O_Normal.Value = True Then Grava_Produto
      If O_Grade.Value = True Then Grava_Produto_Grade
      If O_Edição.Value = True Then Grava_Produto_Grade
    End If
          
    'rsProdutos2.FindNext gsCriteria
    rsProdutos2.MoveNext
    
  Loop
  
  If O_Período.Value = 1 Then
    sSql = "Delete * From [Acompa Estoque] "
    sSql = sSql & " Where [Transf Entrada] = 0 And [Compras] = 0 "
    sSql = sSql & " And [Transf Saída] = 0 and [Vendas] = 0"
    '26/02/2007 - Anderson
    'Alteração na linha do SQL pois estava excluindo produtos com movimentação
    'sSql = sSql & " And [Saldo] <> 0 "
    'sSql = sSql & " And [Saldo] = 0 "
    dbTemp.Execute sSql
  End If
  
  
  
  
' Rem  Nome do BD
'  With Rel
'    .DataFiles(0) = gsTempDBFileName
'    .DataFiles(1) = gsQuickDBFileName
'  End With

  
  '31/10/2002 - mpdea
  'Corrigido associação com a localização das bases de dados
  With Rel
    .Reset
    .WindowState = crptMaximized
    
    '25/01/2006 - mpdea
    'Exibe botão para configurar impressora
    .WindowShowPrintSetupBtn = True
    
    If O_Normal.Value Then
      .DataFiles(0) = gsTempDBFileName
      .DataFiles(1) = gsTempDBFileName
      .DataFiles(2) = gsQuickDBFileName
    ElseIf O_Grade.Value Then
      .DataFiles(0) = gsTempDBFileName
      .DataFiles(1) = gsQuickDBFileName
    ElseIf O_Edição.Value Then
      .DataFiles(0) = gsTempDBFileName
      .DataFiles(1) = gsQuickDBFileName
    End If
  End With


 Rem Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1
 Rem If B_Arquivo = True Then
 Rem    frmMenu.Relatório.Destination = 2
 Rem    frmMenu.Relatório.PrintFileName = T_Arquivo.Text
 Rem End If

 Rem Nome do arquivo .rpt
 If O_Classe.Value = 0 Then
   If O_Normal.Value = True Then Str1 = gsReportPath & "AcEst1.RPT"
   If O_Grade.Value = True Then Str1 = gsReportPath & "AcEst1G.RPT"
   If O_Edição.Value = True Then Str1 = gsReportPath & "AcEst1E.RPT"
   
   If O_Código.Value = True Then Rel.SortFields(0) = "+{Acompa Estoque.Ordenação}"
   If O_Nome.Value = True Then Rel.SortFields(0) = "+{Acompa Estoque.Nome}"
   
   Rel.SortFields(1) = ""
   Rel.SortFields(2) = ""
 End If
 
 If O_Classe.Value = 1 Then
   If O_Normal.Value = True Then Str1 = gsReportPath & "AcEst2.RPT"
   If O_Grade.Value = True Then Str1 = gsReportPath & "AcEst2G.RPT"
   If O_Edição.Value = True Then Str1 = gsReportPath & "AcEst2E.RPT"
   
   If O_Código.Value = True Then
      Rel.SortFields(0) = "+{Acompa Estoque.Classe}"
      Rel.SortFields(1) = "+{Acompa Estoque.Sub Classe}"
      Rel.SortFields(2) = "+{Acompa Estoque.Ordenação}"
   End If
   If O_Nome.Value = True Then
      Rel.SortFields(0) = "+{Acompa Estoque.Classe}"
      Rel.SortFields(1) = "+{Acompa Estoque.Sub Classe}"
      Rel.SortFields(2) = "+{Acompa Estoque.Nome}"
   End If
   
 End If
 
 Rel.ReportFileName = Str1

 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

  Rel.Formulas(0) = Str_Rel '

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 Rel.Formulas(1) = Str_Rel

 Str_Rel = "data_ini = '"
 Str_Rel = Str_Rel + Data_Ini.Text + "'"
 Rel.Formulas(2) = Str_Rel

 Str_Rel = "data_fim = '" + Data_Fim.Text + "'"
 Rel.Formulas(3) = Str_Rel

 Str_Rel = "fornecedor = '" + Nome_Fornecedor.Caption + "'"
 Rel.Formulas(4) = Str_Rel
  
 '03/07/2006 - Anderson
 'Retirado para que o relatório possa utilizar os filtros de código inicial e código final
' If Len(Trim(Combo_Prod.Text)) > 0 Then
'   If Trim(Combo_Prod.Text) <> "0" Then
'     Rel.SelectionFormula = "{Acompa Estoque.Código} = '" & Trim(Combo_Prod.Text) & "'"
'   End If
' End If

 rsProdutos2.Close
 Set rsProdutos2 = Nothing

 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass

  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
 
 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault
End Sub

Private Sub cmd_calendarioDtFim_Click()
  Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
  Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

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

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(1).Text
  Combo_LostFocus
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
  Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Fornecedor_LostFocus()
  Call StatusMsg("")
 
  Nome_Fornecedor.Caption = ""
  If IsNull(Combo_Fornecedor.Text) Then Exit Sub
  If Combo_Fornecedor.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
  If Val(Combo_Fornecedor.Text) <= 0 Then Exit Sub

  rsFornecedores.Index = "Código"
  rsFornecedores.Seek "=", CLng(Combo_Fornecedor.Text)
  If rsFornecedores.NoMatch Then Exit Sub
  Nome_Fornecedor.Caption = rsFornecedores("Nome")


End Sub

Private Sub Combo_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_LostFocus()
  Call StatusMsg("")
 
  Nome_Empresa.Caption = ""
  If IsNull(Combo.Text) Then Exit Sub
  If Combo.Text = "" Then Exit Sub
  If Not IsNumeric(Combo.Text) Then Exit Sub
  If Val(Combo.Text) < 0 Then Exit Sub
  If Val(Combo.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")

End Sub

Private Sub Combo_Prod_CloseUp()
  Combo_Prod.Text = Combo_Prod.Columns(1).Text
  Combo_Prod_LostFocus
End Sub

Private Sub Combo_Prod_Final_CloseUp()
  Combo_Prod_Final.Text = Combo_Prod_Final.Columns(1).Text
  Combo_Prod_Final_LostFocus
End Sub

Private Sub Combo_Prod_Final_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Prod_Final_LostFocus()
  Call StatusMsg("")
 
  Nome_Prod_Final.Caption = ""
  If IsNull(Combo_Prod_Final.Text) Then Exit Sub
  If Combo_Prod_Final.Text = "" Then Exit Sub
  If Combo_Prod_Final.Text = "0" Then Exit Sub
  
  If rsProdutos.RecordCount > 0 Then
    rsProdutos.MoveFirst
  End If
  rsProdutos.FindFirst "Código = '" & Combo_Prod_Final.Text & "'"
  If Not rsProdutos.NoMatch Then
    Nome_Prod_Final.Caption = rsProdutos("Nome")
  Else
    Nome_Prod_Final.Caption = ""
  End If

End Sub

Private Sub Combo_Prod_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Prod_LostFocus()
  Call StatusMsg("")
 
  Nome_Prod.Caption = ""
  If IsNull(Combo_Prod.Text) Then Exit Sub
  If Combo_Prod.Text = "" Then Exit Sub
  If Combo_Prod.Text = "0" Then Exit Sub
  
  rsProdutos.FindFirst "Código = '" & Combo_Prod.Text & "'"
  If Not rsProdutos.NoMatch Then
    Nome_Prod.Caption = rsProdutos("Nome")
  Else
    Nome_Prod.Caption = ""
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

Private Sub Form_Load()

  Call CenterForm(Me)
  
 
 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName
 Data3.DatabaseName = gsQuickDBFileName
 Data4.DatabaseName = gsQuickDBFileName
 Data5.DatabaseName = gsQuickDBFileName

 Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 Set rsFornecedores = db.OpenRecordset("Cli_For", , dbReadOnly)
 Set rsForn_Prod = db.OpenRecordset("Forn_Prod", , dbReadOnly)
 Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)
 Set rsAcompa = dbTemp.OpenRecordset("Acompa Estoque")
 Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
 Set rsEdicoes = db.OpenRecordset("Edições", , dbReadOnly)
 
 Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
 Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
 Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
 Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
 
 
 Combo.Text = gnCodFilial

 Data_Fim.Text = gsFormatDate(Data_Atual)

 If gbGrade = False Then O_Grade.Enabled = False
 If gbEdicao = False Then O_Edição.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsParametros.Close
    Set rsParametros = Nothing
    rsFornecedores.Close
    Set rsFornecedores = Nothing
    rsForn_Prod.Close
    Set rsForn_Prod = Nothing
    rsEstoque.Close
    Set rsEstoque = Nothing
    rsAcompa.Close
    Set rsAcompa = Nothing
    rsGrade.Close
    Set rsGrade = Nothing
    rsEdicoes.Close
    Set rsEdicoes = Nothing
    rsTamanhos.Close
    Set rsTamanhos = Nothing
    rsCores.Close
    Set rsCores = Nothing
    rsClasses.Close
    Set rsClasses = Nothing
    rsSubclasses.Close
    Set rsSubclasses = Nothing
End Sub

Private Sub txt_fabricanteMarca_LostFocus()
    If Len(Trim(txt_fabricanteMarca.Text)) > 0 Then
        txt_fabricanteMarca.Text = UCase(Trim(txt_fabricanteMarca.Text))
    End If
End Sub
