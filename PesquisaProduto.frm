VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPesquisaProduto 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pesquisar Produtos"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PesquisaProduto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   15150
   Begin VB.CommandButton cmd_pesqTelaDet 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pesquisa &Detalhada"
      Height          =   990
      Left            =   14220
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txt_EntrarQtde 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   3870
      TabIndex        =   11
      Text            =   "1"
      Top             =   4410
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFA324&
      Caption         =   "Incluir na tela anterior"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5355
      Width           =   11310
   End
   Begin VB.Frame fraPesquisa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   -480
      TabIndex        =   21
      Top             =   5820
      Width           =   900
      Begin SSDataWidgets_B.SSDBCombo cboPesquisa1 
         Bindings        =   "PesquisaProduto.frx":4E95A
         Height          =   360
         Left            =   900
         TabIndex        =   14
         Top             =   210
         Width           =   2460
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
         ColumnHeaders   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4419
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
         _ExtentX        =   4339
         _ExtentY        =   635
         _StockProps     =   93
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboPesquisa2 
         Bindings        =   "PesquisaProduto.frx":4E971
         Height          =   360
         Left            =   900
         TabIndex        =   15
         Top             =   630
         Width           =   2460
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
         ColumnHeaders   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4419
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
         _ExtentX        =   4339
         _ExtentY        =   635
         _StockProps     =   93
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboPesquisa3 
         Bindings        =   "PesquisaProduto.frx":4E988
         Height          =   360
         Left            =   900
         TabIndex        =   16
         Top             =   1050
         Width           =   2475
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
         ColumnHeaders   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4419
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
         _ExtentX        =   4366
         _ExtentY        =   635
         _StockProps     =   93
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin VB.Label lblPesquisa 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pesquisa 3"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   1110
         Width           =   885
      End
      Begin VB.Label lblPesquisa 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pesquisa 2"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   705
         Width           =   885
      End
      Begin VB.Label lblPesquisa 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pesquisa 1"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   300
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Data datPesq3 
      Caption         =   "datPesq1"
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
      Height          =   360
      Left            =   -150
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq3"
      Top             =   5940
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Data datPesq2 
      Caption         =   "datPesq1"
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
      Height          =   360
      Left            =   -540
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq2"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Data datPesq1 
      Caption         =   "datPesq1"
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
      Height          =   360
      Left            =   -720
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Frame fraEntrega 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mostrar produtos que"
      Height          =   1080
      Left            =   9210
      TabIndex        =   20
      Top             =   30
      Width           =   3060
      Begin VB.CheckBox chk_considerarInativados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Considerar inativados"
         Height          =   225
         Left            =   210
         TabIndex        =   33
         Top             =   780
         Width           =   2025
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contenham uma OU outra parte"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Top             =   450
         Width           =   2760
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contenham TODAS as partes"
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Top             =   165
         Value           =   -1  'True
         Width           =   2490
      End
   End
   Begin VB.CommandButton cmdAbort 
      Cancel          =   -1  'True
      Caption         =   "Interromper"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   450
      TabIndex        =   17
      Top             =   5940
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   -120
      TabIndex        =   18
      Top             =   5850
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Pesquisar"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   12300
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1860
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pesquisar por                              * Uma ou mais partes separadas por espaço em branco"
      Height          =   1080
      Left            =   45
      TabIndex        =   19
      Top             =   30
      Width           =   9135
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   6090
         TabIndex        =   4
         Top             =   600
         Width           =   2865
      End
      Begin VB.OptionButton optSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Código para &Fornecedor ou parte"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   6060
         TabIndex        =   1
         Top             =   330
         Width           =   2880
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   3525
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   3780
         TabIndex        =   3
         Top             =   600
         Width           =   2235
      End
      Begin VB.OptionButton optSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Nome do produto ou parte"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   2490
      End
      Begin VB.OptionButton optSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Código produto ou parte"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3750
         TabIndex        =   0
         Top             =   315
         Width           =   2250
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridTabelaPrecos 
      Height          =   1470
      Left            =   11400
      TabIndex        =   13
      Top             =   4410
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   2593
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16753444
      BackColorSel    =   16777152
      ForeColorSel    =   -2147483641
      BackColorBkg    =   15066597
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
   Begin SSDataWidgets_B.SSDBGrid grdResultados 
      Height          =   3150
      Left            =   45
      TabIndex        =   10
      Top             =   1170
      Width           =   15075
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   7
      BevelColorHighlight=   16777215
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      MaxSelectedRows =   0
      ForeColorEven   =   4210752
      BackColorEven   =   12648447
      BackColorOdd    =   16777215
      RowHeight       =   423
      ExtraHeight     =   265
      Columns.Count   =   7
      Columns(0).Width=   4075
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   10398
      Columns(1).Caption=   "Descrição"
      Columns(1).Name =   "Descricao"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1799
      Columns(2).Caption=   "Tamanho"
      Columns(2).Name =   "Tamanho"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2328
      Columns(3).Caption=   "Cor"
      Columns(3).Name =   "Cor"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2858
      Columns(4).Caption=   "Classe"
      Columns(4).Name =   "Classe"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2858
      Columns(5).Caption=   "SubClasse"
      Columns(5).Name =   "SubClasse"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1270
      Columns(6).Caption=   "Estoque"
      Columns(6).Name =   "EstoqueAtual"
      Columns(6).Alignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      _ExtentX        =   26591
      _ExtentY        =   5556
      _StockProps     =   79
      BackColor       =   15066597
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_cor 
      BackColor       =   &H00FFA324&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   8580
      TabIndex        =   32
      Top             =   4860
      Width           =   2775
   End
   Begin VB.Label lbl_tamanho 
      BackColor       =   &H00FFA324&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   6750
      TabIndex        =   31
      Top             =   4860
      Width           =   1785
   End
   Begin VB.Label lbl_subClasse 
      BackColor       =   &H00FFA324&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   8580
      TabIndex        =   30
      Top             =   4410
      Width           =   2775
   End
   Begin VB.Label lbl_classe 
      BackColor       =   &H00FFA324&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5910
      TabIndex        =   29
      Top             =   4410
      Width           =   2625
   End
   Begin VB.Label lbl_nomeProduto 
      BackColor       =   &H00FFA324&
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
      Height          =   420
      Left            =   30
      TabIndex        =   28
      Top             =   4860
      Width           =   6660
   End
   Begin VB.Label lbl_codProduto 
      BackColor       =   &H00FFA324&
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
      Height          =   420
      Left            =   30
      TabIndex        =   27
      Top             =   4410
      Width           =   3135
   End
   Begin VB.Label lbl_luminosoGrade 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   26
      Top             =   1110
      Width           =   15135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Qtde"
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
      Height          =   285
      Left            =   3390
      TabIndex        =   25
      Top             =   4485
      Width           =   435
   End
End
Attribute VB_Name = "frmPesquisaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strTipoPesquisa As String '31/08/2006 - Anderson - Indica o tipo de pesquisa que será realizada (Pelo cadastro de produto ou pela consulta de produtos)
Private rsProdPesq As Recordset
Private gsSql As String
Private gsCod() As String
Private gsDesc() As String
Private gsEstoqueAtualProdutoPesq() As String
Private gsProdTamanho() As String
Private gsProdCor() As String
Private gsProdClasse() As String
Private gsProdSubClasse() As String
Private gbToAbort As Boolean

Private rsProdPesqGrade As Recordset
Private rsProdPesqTamanho As Recordset
Private rsProdPesqCor As Recordset
Private rsProdPesqClasse As Recordset
Private rsProdPesqSubClasse As Recordset


Private sCodigoProduto As String

'''Private Sub cboPesquisa1_Change()
'''  If Len(Trim(txtSearch(0).Text)) > 0 Or _
'''    Len(Trim(txtSearch(1).Text)) > 0 Or _
'''    Len(Trim(txtSearch(2).Text)) > 0 Or _
'''    Len(cboPesquisa1.Text) > 0 Or _
'''    Len(cboPesquisa2.Text) > 0 Or _
'''    Len(cboPesquisa3.Text) > 0 Then
'''    cmdClose.Default = False
'''    cmdSearch.Enabled = True
'''    cmdSearch.Default = True
'''  Else
'''    cmdSearch.Enabled = False
'''    cmdSearch.Default = False
'''    cmdClose.Default = True
'''  End If
'''End Sub
'''
'''Private Sub cboPesquisa1_CloseUp()
'''  Dim bm As Variant
'''  bm = cboPesquisa1.GetBookmark(0)
'''  cboPesquisa1.Text = cboPesquisa1.Columns(0).CellText(bm)
'''End Sub
'''
'''Private Sub cboPesquisa2_Change()
'''  If Len(Trim(txtSearch(0).Text)) > 0 Or _
'''    Len(Trim(txtSearch(1).Text)) > 0 Or _
'''    Len(Trim(txtSearch(2).Text)) > 0 Or _
'''    Len(cboPesquisa1.Text) > 0 Or _
'''    Len(cboPesquisa2.Text) > 0 Or _
'''    Len(cboPesquisa3.Text) > 0 Then
'''    cmdClose.Default = False
'''    cmdSearch.Enabled = True
'''    cmdSearch.Default = True
'''  Else
'''    cmdSearch.Enabled = False
'''    cmdSearch.Default = False
'''    cmdClose.Default = True
'''  End If
'''End Sub
'''
'''Private Sub cboPesquisa2_CloseUp()
'''  Dim bm As Variant
'''  bm = cboPesquisa2.GetBookmark(0)
'''  cboPesquisa2.Text = cboPesquisa2.Columns(0).CellText(bm)
'''End Sub
'''
'''Private Sub cboPesquisa3_Change()
'''  If Len(Trim(txtSearch(0).Text)) > 0 Or _
'''    Len(Trim(txtSearch(1).Text)) > 0 Or _
'''    Len(Trim(txtSearch(2).Text)) > 0 Or _
'''    Len(cboPesquisa1.Text) > 0 Or _
'''    Len(cboPesquisa2.Text) > 0 Or _
'''    Len(cboPesquisa3.Text) > 0 Then
'''    cmdClose.Default = False
'''    cmdSearch.Enabled = True
'''    cmdSearch.Default = True
'''  Else
'''    cmdSearch.Enabled = False
'''    cmdSearch.Default = False
'''    cmdClose.Default = True
'''  End If
'''End Sub
'''
'''Private Sub cboPesquisa3_CloseUp()
'''  Dim bm As Variant
'''  bm = cboPesquisa3.GetBookmark(0)
'''  cboPesquisa3.Text = cboPesquisa3.Columns(0).CellText(bm)
'''End Sub

Private Sub cmd_pesqTelaDet_Click()
  frmConsultaProd.Show
End Sub

Private Sub cmdAbort_Click()
  gbToAbort = True
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSearch_Click()
  
  lbl_codProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  lbl_nomeProduto.Caption = ""
  lbl_classe.Caption = ""
  lbl_subClasse.Caption = ""
  lbl_tamanho.Caption = ""
  lbl_cor.Caption = ""
  
  Call SearchString
  'Call CenterForm(Me)
  
  If grdResultados.Rows > 0 Then
      grdResultados.SetFocus
  Else
      If optSearch(0).Value = True Then
          txtSearch(0).SetFocus
      ElseIf optSearch(1).Value = True Then
          txtSearch(1).SetFocus
      Else
          txtSearch(2).SetFocus
      End If
  End If

End Sub

Private Sub Command2_Click()
  Dim intCancel As Integer
  Dim frmX As Form
  
  txt_EntrarQtde.Text = Replace(txt_EntrarQtde.Text, ".", ",")
  
  If Not IsNumeric(txt_EntrarQtde.Text) Then
      MsgBox "Informe uma quantidade válida.", vbInformation, "Atenção"
      txt_EntrarQtde.SetFocus
      Exit Sub
  End If
  
  If Trim(lbl_codProduto.Caption) = "" Then
      MsgBox "Selecione um produto na grade.", vbInformation, "Atenção"
      Exit Sub
  End If
  
  
  gTabelaPrecoAcatadaTelaPesquisaProduto = ""
  
  Select Case nChamaConsulta
    Case 1, 2 'Previne outros códigos de chamada
      
      'Form de origem (chamadas comuns)
      If nChamaConsulta = 1 Then
        '18/01/2006 - mpdea
        'Alterado objeto frmVendaRap2 -> g_frmVendaRapida
        Set frmX = g_frmVendaRapida
      Else
        Set frmX = frmSaidas
      End If
      
      With frmX
        'Insere o item
        .Grade1.Columns(0).Text = lbl_codProduto.Caption
        '''.Grade1.Columns(1).Text = "1"
        
        If txt_EntrarQtde.Text = "" Or txt_EntrarQtde.Text = "0" Then
            .Grade1.Columns(1).Text = "1"
        Else
            .Grade1.Columns(1).Text = txt_EntrarQtde.Text
        End If
        
        If gridTabelaPrecos.RowSel > 0 Then
            gTabelaPrecoAcatadaTelaPesquisaProduto = gridTabelaPrecos.TextMatrix(gridTabelaPrecos.RowSel, 1)
'''            If nChamaConsulta = 1 Then
'''                .Grade1.Columns(3).Text = gridTabelaPrecos.TextMatrix(gridTabelaPrecos.RowSel, 2)
'''            Else
'''                .Grade1.Columns(4).Text = gridTabelaPrecos.TextMatrix(gridTabelaPrecos.RowSel, 2)
'''            End If
        End If
        
        'Atualiza grid
        .Grade1_BeforeColUpdate 0, "", intCancel
        If gridTabelaPrecos.RowSel > 0 Then
            gTabelaPrecoAcatadaTelaPesquisaProduto = gridTabelaPrecos.TextMatrix(gridTabelaPrecos.RowSel, 1)
        End If


        If intCancel = -1 Then Exit Sub
        'Calcula totais
        .Calcula_Linha
        .Recalcula
        'Move para a próxima linha
        .Grade1.MoveNext
        .Grade1.DoClick
      
      End With
      
      Set frmX = Nothing
      
    '04/11/2009 - mpdea
    'Tela de Entradas
    Case 3
      With frmEntrada
        'Remove as linhas em branco
        Dim Str_Aux As String
        Dim bm As Variant
        Dim nRow As Long
        For nRow = .grdItens.Rows - 1 To 0 Step -1
          bm = .grdItens.AddItemBookmark(nRow)
          .grdItens.Bookmark = bm
          Str_Aux = gsHandleNull(.grdItens.Columns("Código").CellText(bm))
          If (Str_Aux = "0" Or Str_Aux = "") And Not IsEmpty(bm) Then
            .grdItens.RemoveItem .grdItens.AddItemRowIndex(bm)
          End If
        Next nRow
        .grdItens.Scroll -99, -99
        .grdItens.Update
        
        'Insere o item
        '''.grdItens.AddItem lbl_codProduto.Caption & vbTab & "1"
        If txt_EntrarQtde.Text = "" Or txt_EntrarQtde.Text = "0" Then
            .grdItens.AddItem lbl_codProduto.Caption & vbTab & "1"
        Else
            .grdItens.AddItem lbl_codProduto.Caption & vbTab & txt_EntrarQtde.Text
        End If
        
        'Atualiza grid
        .grdItens.MoveLast
        .grdItens_BeforeColUpdate 0, "", intCancel
        If intCancel = -1 Then Exit Sub
        'Calcula totais
        .Calcula_Linha
        .Recalcula
        .grdItens.Update
      End With
      
    Case 4
      With frmProdutosCFOP
        .txt_codigoProduto.Text = lbl_codProduto.Caption
        .txt_NomeProduto.Text = lbl_nomeProduto.Caption
      End With
    Case 5
      With frmProdutosCesta
        CodigoProdutoCestaPesq = lbl_codProduto.Caption
        NomeProdutoCestaPesq = lbl_nomeProduto.Caption
      End With
    Case 6
      ' Tela TransferenciaEntreEmpresas
      With frmTransfere
        'Insere o item
        .Grade1.Columns(0).Text = lbl_codProduto.Caption
        .Grade1.Columns(1).Text = lbl_nomeProduto.Caption & " " & lbl_tamanho.Caption & " " & lbl_cor.Caption
        
        If txt_EntrarQtde.Text = "" Or txt_EntrarQtde.Text = "0" Then
            .Grade1.Columns(2).Text = "1"
        Else
            .Grade1.Columns(2).Text = txt_EntrarQtde.Text
        End If

        If intCancel = -1 Then Exit Sub
        'Move para a próxima linha
        .Grade1.MoveNext
        .Grade1.DoClick
      End With
      
    Case 7
      ' Tela frmEtiquetas
      With frmEtiquetas
        'Insere o item
        .txtProduto.Text = lbl_codProduto.Caption
        .lbl_nomeProduto.Caption = lbl_nomeProduto.Caption & " " & lbl_tamanho.Caption & " " & lbl_cor.Caption
        
        If intCancel = -1 Then Exit Sub
      End With
      
      
  End Select
  
  If optSearch(0).Value = True Then
    txtSearch(0).SetFocus
  ElseIf optSearch(1).Value = True Then
    txtSearch(1).SetFocus
  ElseIf optSearch(2).Value = True Then
    txtSearch(2).SetFocus
  End If
     
End Sub

Private Sub Form_Load()

  Set rsProdPesqGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsProdPesqTamanho = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsProdPesqCor = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsProdPesqClasse = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsProdPesqSubClasse = db.OpenRecordset("Sub Classes", , dbReadOnly)

  gsSql = "SELECT *, Produtos.Classe, Produtos.[Sub Classe] FROM Produtos "
  gsSql = gsSql & " LEFT JOIN [Estoque Final] ON "
  gsSql = gsSql & " (Produtos.[Código] = [Estoque Final].produto and [Estoque Final].Filial = " & gnCodFilial
  gsSql = gsSql & ") WHERE [Código] <> '0' "
  'gsSql = gsSql & " and [Estoque Final].Filial = " & gnCodFilial
  gsSql = gsSql & " ORDER BY [Código Ordenação], Tamanho, Cor"
  Set rsProdPesq = db.OpenRecordset(gsSql, dbReadOnly)
  
  Screen.MousePointer = vbHourglass
''  Set rsProdPesq = rsProdutos.Clone
  Screen.MousePointer = vbDefault
  cmdAbort.Enabled = False
  cmdSearch.Enabled = False
  
  datPesq1.DatabaseName = gsQuickDBFileName
  datPesq2.DatabaseName = gsQuickDBFileName
  datPesq3.DatabaseName = gsQuickDBFileName
  
  Me.Left = Screen.Width - Me.Width
  
  If Screen.Height - (Me.Height + 350) > 0 Then
      Me.Top = Screen.Height - (Me.Height + 350)
  Else
      Me.Top = 500
  End If
  
  'Me.Height = 3180
  'Me.Width = 9000
  'Call CenterForm(Me)
  Me.Show
  
'''  If gsPesq1 & gsPesq2 & gsPesq3 = "" Then
'''    'fraPesquisa.Visible = False
'''    cboPesquisa1.Enabled = False
'''    cboPesquisa2.Enabled = False
'''    cboPesquisa3.Enabled = False
'''
'''  Else
'''    fraPesquisa.Visible = True
'''    If gsPesq1 = "" Then
'''      lblPesquisa(0).Visible = False
'''      cboPesquisa1.Visible = False
'''    Else
'''      lblPesquisa(0).Caption = gsPesq1 + ":"
'''    End If
'''
'''    If gsPesq2 = "" Then
'''      lblPesquisa(1).Visible = False
'''      cboPesquisa2.Visible = False
'''    Else
'''      lblPesquisa(1).Caption = gsPesq2 + ":"
'''    End If
'''
'''    If gsPesq3 = "" Then
'''      lblPesquisa(2).Visible = False
'''      cboPesquisa3.Visible = False
'''    Else
'''      lblPesquisa(2).Caption = gsPesq3 + ":"
'''    End If
'''  End If
  
  gridTabelaPrecos.ColWidth(0) = 1
  gridTabelaPrecos.ColWidth(1) = 1700
  gridTabelaPrecos.ColWidth(2) = 1700

  gridTabelaPrecos.Row = 0
  gridTabelaPrecos.TextMatrix(0, 1) = "Tabela de Preço"
  gridTabelaPrecos.TextMatrix(0, 2) = "Valor R$"
  
  txtSearch(1).SetFocus

End Sub

Private Sub SearchString()
  Dim nPos As Integer
  Dim sCod As String
  Dim sText() As String
  Dim nI As Integer
  Dim nK As Integer
  Dim nItem As Integer
  Dim nSum As Integer
  Dim bOk As Boolean
  Dim sError As String
  
  
  'Call CenterForm(Me)
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  ReDim gsCod(0)
  ReDim gsDesc(0)
  ReDim gsEstoqueAtualProdutoPesq(0)
  
  ReDim gsProdTamanho(0)
  ReDim gsProdCor(0)
  ReDim gsProdClasse(0)
  ReDim gsProdSubClasse(0)
  Dim sTamZeros As String
  Dim sCorZeros As String

  If optSearch(0).Value = True Then
    nK = 0
  Else
    If optSearch(1).Value = True Then
      nK = 1
    Else
      nK = 2
    End If
  End If
  nItem = -1
  gbToAbort = False
  cmdAbort.Enabled = True
  cmdSearch.Enabled = False
  sText = Split(UCase(txtSearch(nK).Text), " ", -1, vbTextCompare)
  With rsProdPesq
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        DoEvents
        If gbToAbort Then
          Exit Do
        End If
        nSum = 0
        bOk = True
'''        If Len(Trim(cboPesquisa1.Text)) > 0 Then
'''          bOk = UCase(.Fields("Pesquisa 1").Value) = UCase(cboPesquisa1.Columns(1).Text)
'''        End If
'''        If bOk And Len(Trim(cboPesquisa2.Text)) > 0 Then
'''          bOk = UCase(.Fields("Pesquisa 2").Value) = UCase(cboPesquisa2.Columns(1).Text)
'''        End If
'''        If bOk And Len(Trim(cboPesquisa3.Text)) > 0 Then
'''          bOk = UCase(.Fields("Pesquisa 3").Value) = UCase(cboPesquisa3.Columns(1).Text)
'''        End If
        If bOk = True Then
          For nI = 0 To UBound(sText)
            DoEvents
            If gbToAbort Then
              Exit For
            End If
            If nK = 0 Then
              nPos = InStr(UCase(.Fields("Código").Value), sText(nI))
            Else
              If nK = 1 Then
                nPos = InStr(UCase(.Fields("Nome").Value), sText(nI))
              Else
                nPos = InStr(UCase(.Fields("Código do Fornecedor").Value), sText(nI))
              End If
            End If
            If nPos > 0 Then
                If (chk_considerarInativados.Value = vbUnchecked And .Fields("Desativado").Value = False) Or _
                  chk_considerarInativados.Value = vbChecked Then
                    nSum = nSum + 1
                End If
            End If
          Next nI
          If (optType(1).Value = True And nSum > 0) Or (nSum = UBound(sText) + 1) Then
            nItem = nItem + 1
            ReDim Preserve gsCod(nItem)
            ReDim Preserve gsDesc(nItem)
            ReDim Preserve gsEstoqueAtualProdutoPesq(nItem)
            ReDim Preserve gsProdTamanho(nItem)
            ReDim Preserve gsProdCor(nItem)
            ReDim Preserve gsProdClasse(nItem)
            ReDim Preserve gsProdSubClasse(nItem)

            gsCod(nItem) = .Fields("Código")
            gsDesc(nItem) = .Fields("Nome")
            gsEstoqueAtualProdutoPesq(nItem) = .Fields("Estoque Atual")
            
            If .Fields("Tipo") = "G" Then 'Produto com grade
                sTamZeros = ""
                If Len(.Fields("Tamanho")) = 1 Then
                    sTamZeros = "00"
                ElseIf Len(.Fields("Tamanho")) = 2 Then
                    sTamZeros = "0"
                End If
                
                sCorZeros = ""
                If Len(.Fields("Cor")) = 1 Then
                    sCorZeros = "00"
                ElseIf Len(.Fields("Cor")) = 2 Then
                    sCorZeros = "0"
                End If
                
                gsCod(nItem) = .Fields("Código") & sTamZeros & .Fields("Tamanho") & sCorZeros & .Fields("Cor")
                gsProdTamanho(nItem) = AcharTamanho(.Fields("Tamanho"))
                gsProdCor(nItem) = AcharCor(.Fields("Cor"))
            Else
                gsProdTamanho(nItem) = ""
                gsProdCor(nItem) = ""
            End If
            gsProdClasse(nItem) = AcharClasse(.Fields("Produtos.Classe"))
            gsProdSubClasse(nItem) = AcharSubClasse(.Fields("Produtos.Sub Classe"))
          End If
        End If
        .MoveNext
      Loop
    Else
      gsTitle = "Resultados da Pesquisa"
      gsMsg = "Cadastro de Produtos está vazio."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Exit Sub
    End If
  End With
  If nItem > -1 Then
    Call LoadGrid
  Else
    gsTitle = "Resultados da Pesquisa"
    gsMsg = "Nenhum registro encontrado para as condições fornecidas."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  rsProdPesq.MoveFirst
  cmdAbort.Enabled = False
  cmdSearch.Enabled = True
  Screen.MousePointer = vbDefault

End Sub

Private Sub LoadGrid()
  Dim nI As Integer
  '''grdResultados.Redraw = False
  For nI = 0 To UBound(gsCod)
    'grdResultados.AddItem gsCod(nI) & vbTab & gsDesc(nI) & vbTab & gsEstoqueAtualProdutoPesq(nI)
    grdResultados.AddItem gsCod(nI) & vbTab & gsDesc(nI) & vbTab & gsProdTamanho(nI) & vbTab & _
    gsProdCor(nI) & vbTab & gsProdClasse(nI) & vbTab & gsProdSubClasse(nI) & vbTab & gsEstoqueAtualProdutoPesq(nI)
  Next nI
  grdResultados.Redraw = True
  ReDim gsCod(0)
  ReDim gsDesc(0)
  ReDim gsEstoqueAtualProdutoPesq(0)
  'Me.Height = 6690
  Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsProdPesq.Close
  Set rsProdPesq = Nothing

  rsProdPesqGrade.Close
  Set rsProdPesqGrade = Nothing
  
  rsProdPesqTamanho.Close
  Set rsProdPesqTamanho = Nothing
  
  rsProdPesqCor.Close
  Set rsProdPesqCor = Nothing
  
  rsProdPesqClasse.Close
  Set rsProdPesqClasse = Nothing
  
  rsProdPesqSubClasse.Close
  Set rsProdPesqSubClasse = Nothing
End Sub


Private Function AcharCor(iCor As Integer) As String
On Error GoTo Erro

  AcharCor = ""
  
  rsProdPesqCor.Index = "Código"
  rsProdPesqCor.Seek "=", iCor
  If rsProdPesqCor.NoMatch Then
      AcharCor = ""
      Exit Function
  Else
      AcharCor = LCase(rsProdPesqCor.Fields("Nome").Value)
  End If

  Exit Function
Erro:
  MsgBox "Erro na função AcharCor do produto " & Err.Description, vbInformation, "Atenção"

End Function

Private Function AcharTamanho(iTamanho As Integer) As String
On Error GoTo Erro

  AcharTamanho = ""

  rsProdPesqTamanho.Index = "Código"
  rsProdPesqTamanho.Seek "=", iTamanho
  If rsProdPesqTamanho.NoMatch Then
      AcharTamanho = ""
      Exit Function
  Else
      AcharTamanho = LCase(rsProdPesqTamanho.Fields("Nome").Value)
  End If

  Exit Function
Erro:
  MsgBox "Erro na função AcharTamanho do produto " & Err.Description, vbInformation, "Atenção"

End Function

Private Function AcharClasse(iClasse As Integer) As String
On Error GoTo Erro

  AcharClasse = ""
  
  rsProdPesqClasse.Index = "Código"
  rsProdPesqClasse.Seek "=", iClasse
  If rsProdPesqClasse.NoMatch Then
      AcharClasse = ""
      Exit Function
  Else
      AcharClasse = LCase(rsProdPesqClasse.Fields("Nome").Value)
  End If
  
  Exit Function
Erro:
  MsgBox "Erro na função AcharClasse do produto " & Err.Description, vbInformation, "Atenção"
  
End Function

Private Function AcharSubClasse(iSubClasse As Integer) As String
On Error GoTo Erro

  AcharSubClasse = ""
  
  rsProdPesqSubClasse.Index = "Código"
  rsProdPesqSubClasse.Seek "=", iSubClasse
  If rsProdPesqSubClasse.NoMatch Then
      AcharSubClasse = ""
      Exit Function
  Else
      AcharSubClasse = LCase(rsProdPesqSubClasse.Fields("Nome").Value)
  End If
  
  Exit Function
Erro:
  MsgBox "Erro na função AcharSubClasse do produto " & Err.Description, vbInformation, "Atenção"
  
End Function


Private Sub grdResultados_Click()
On Error GoTo Erro

  Dim CodigoProduto As String
  
  txt_EntrarQtde.Text = "1"
  
  gridTabelaPrecos.Rows = 1
  gridTabelaPrecos.Clear
  gridTabelaPrecos.Row = 0
  gridTabelaPrecos.TextMatrix(0, 1) = "Tabela de Preço"
  gridTabelaPrecos.TextMatrix(0, 2) = "Valor R$"
  
  CodigoProduto = grdResultados.Columns(0).Text
  
  If grdResultados.Columns(2).Text <> "" Then
      ' É produto com Grade (tamanho e cor)
      If Len(CodigoProduto) > 6 Then
          CodigoProduto = Mid(CodigoProduto, 1, Len(CodigoProduto) - 6)
      End If
  End If

  If CodigoProduto <> "0" And CodigoProduto <> "" Then
  
    Dim sSql As String
    Dim rsTabelas As Recordset
    If Len(CodigoProduto) > 0 Then
        sSql = "Select P.Tabela, P.Preço from Preços P, AcessoTabelasDePrecosProdutos A "
        If Funcionario <> "" Then
            sSql = sSql & " where A.Usuario = " & Funcionario
        Else
            sSql = sSql & " where A.Usuario = " & gnUserCode
        End If
        
        sSql = sSql & " And A.Tabela = P.Tabela "
        sSql = sSql & " And P.Produto ='" & CodigoProduto & "'"
    
        Set rsTabelas = db.OpenRecordset(sSql, dbOpenDynaset)
        If rsTabelas.EOF And rsTabelas.BOF Then
            Exit Sub
        End If
        
        rsTabelas.MoveFirst
        While Not rsTabelas.EOF
            gridTabelaPrecos.AddItem vbTab & rsTabelas.Fields("Tabela").Value & vbTab & _
                            FormatNumber(rsTabelas.Fields("Preço").Value, 2) & vbTab
            rsTabelas.MoveNext
        Wend
        rsTabelas.Close
        Set rsTabelas = Nothing
    End If
  End If
  
  gridTabelaPrecos.RowSel = 0
  
  If nChamaConsulta <> 4 And nChamaConsulta <> 5 Then
    'Implementação de pesquisa avançada na tela de consulta do produto
    If strTipoPesquisa = "Pesquisa" Then
      frmConsultaProd.Con_Código.Text = grdResultados.Columns(0).Text
      frmConsultaProd.Con_Código_LostFocus
    Else
      frmProdutos.cboCodigo.Text = grdResultados.Columns(0).Text
      frmProdutos.cboCodigo_LostFocus
    End If
    grdResultados.SetFocus
  End If
  
  sCodigoProduto = grdResultados.Columns(0).Text
  lbl_codProduto.Caption = sCodigoProduto
  lbl_nomeProduto.Caption = grdResultados.Columns(1).Text
  lbl_classe.Caption = grdResultados.Columns(4).Text
  lbl_subClasse.Caption = grdResultados.Columns(5).Text
  lbl_tamanho.Caption = grdResultados.Columns(2).Text
  lbl_cor.Caption = grdResultados.Columns(3).Text
  

  txt_EntrarQtde.SetFocus
 
  Exit Sub
Erro:
  MsgBox "Erro na função Click da grade " & Err.Description, vbInformation, "Atenção"
 
End Sub

Private Sub grdResultados_DblClick()

 
'''  If nChamaConsulta <> 4 And nChamaConsulta <> 5 Then
'''    'Implementação de pesquisa avançada na tela de consulta do produto
'''    If strTipoPesquisa = "Pesquisa" Then
'''      frmConsultaProd.Con_Código.Text = grdResultados.Columns(0).Text
'''      frmConsultaProd.Con_Código_LostFocus
'''    Else
'''      frmProdutos.cboCodigo.Text = grdResultados.Columns(0).Text
'''      frmProdutos.cboCodigo_LostFocus
'''    End If
'''    grdResultados.SetFocus
'''  End If
'''
'''  sCodigoProduto = grdResultados.Columns(0).Text
'''  lbl_codProduto.Caption = sCodigoProduto
'''  lbl_nomeProduto.Caption = grdResultados.Columns(1).Text
  
End Sub

Private Sub grdResultados_GotFocus()
    lbl_luminosoGrade.BackColor = &H800000
End Sub

Private Sub grdResultados_KeyPress(KeyAscii As Integer)
  Dim cCaracter As Variant
   cCaracter = Chr(KeyAscii)
   KeyAscii = Asc(UCase(cCaracter))
   
   If KeyAscii = 13 Then    'Tecla ENTER
         grdResultados_Click
   End If
  
End Sub

Private Sub grdResultados_LostFocus()
    lbl_luminosoGrade.BackColor = &HFFFFFF
End Sub

Private Sub gridTabelaPrecos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridTabelaPrecos.Redraw = False
End Sub

Private Sub gridTabelaPrecos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridTabelaPrecos.RowSel = gridTabelaPrecos.Row
  gridTabelaPrecos.Redraw = True
End Sub

Private Sub optSearch_Click(Index As Integer)
  txtSearch(0).Text = ""
  txtSearch(0).Enabled = False
  txtSearch(1).Text = ""
  txtSearch(1).Enabled = False
  txtSearch(2).Text = ""
  txtSearch(2).Enabled = False
  txtSearch(Index).Text = ""
  txtSearch(Index).Enabled = True
End Sub

Private Sub txt_EntrarQtde_GotFocus()
    With txt_EntrarQtde
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    Command2.Default = True
End Sub

Private Sub txtSearch_Change(Index As Integer)
  If Len(Trim(txtSearch(Index).Text)) > 0 Then
    cmdClose.Default = False
    cmdSearch.Enabled = True
    cmdSearch.Default = True
  Else
    cmdSearch.Enabled = False
    cmdSearch.Default = False
    cmdClose.Default = True
  End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    With txtSearch(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    cmdSearch.Default = True
End Sub
