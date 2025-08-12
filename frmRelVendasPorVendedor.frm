VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelVendasPorVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Vendas por Vendedor"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   7365
   Begin VB.Frame Frame1 
      Caption         =   "Período 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   3720
      Width           =   3615
      Begin MSMask.MaskEdBox mskDataFinal3 
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskDataInicio3 
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label4 
         Caption         =   "até:"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   33
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "De:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   390
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   1
      Left            =   3840
      TabIndex        =   28
      Top             =   2760
      Width           =   3495
      Begin MSMask.MaskEdBox mskDataFinal2 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskDataInicio2 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label4 
         Caption         =   "até:"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   30
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "De:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   390
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   3615
      Begin MSMask.MaskEdBox mskDataFinal1 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskDataInicio1 
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
         Caption         =   "De:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "até:"
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
         Index           =   0
         Left            =   1800
         TabIndex        =   26
         Top             =   390
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   24
      Top             =   3720
      Width           =   1695
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
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
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
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
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1575
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7335
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   240
         Picture         =   "frmRelVendasPorVendedor.frx":0000
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Vendas por Vendedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelVendasPorVendedor.frx":1E68
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   855
         Left            =   2160
         TabIndex        =   22
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Data datFiliais 
      Caption         =   "datFiliais"
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Data datClientes 
      Caption         =   "datClientes"
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For ORDER BY Nome"
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Data datClasse 
      Caption         =   "datClasse"
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
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Classes ORDER BY Nome"
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Data datSubClasse 
      Caption         =   "datSubClasse"
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
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM [Sub Classes] ORDER BY Nome"
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Data datProdutos 
      Caption         =   "datProdutos"
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
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0' ORDER BY Nome"
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   7215
      Begin VB.TextBox txtNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   4575
      End
      Begin VB.TextBox txtNomeVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtNomeOperacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
         Width           =   4575
      End
      Begin SSDataWidgets_B.SSDBCombo cboVendedor 
         Bindings        =   "frmRelVendasPorVendedor.frx":1F58
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         DataFieldList   =   "Nome"
         _Version        =   196617
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4101
         Columns(0).Caption=   "Codigo"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Código"
         Columns(0).FieldLen=   256
         Columns(1).Width=   7355
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).DataField=   "Nome"
         Columns(1).FieldLen=   256
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelVendasPorVendedor.frx":1F74
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   120
         Width           =   1335
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
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Filial"
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperacao 
         Bindings        =   "frmRelVendasPorVendedor.frx":1F8D
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1335
         DataFieldList   =   "Nome"
         _Version        =   196617
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4101
         Columns(0).Caption=   "Codigo"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Código"
         Columns(0).FieldLen=   256
         Columns(1).Width=   7355
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).DataField=   "Nome"
         Columns(1).FieldLen=   256
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label Label1 
         Caption         =   "Filial"
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
         Left            =   120
         TabIndex        =   19
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Vendedor"
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
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Operação"
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
         Left            =   120
         TabIndex        =   17
         Top             =   870
         Width           =   855
      End
   End
   Begin VB.Data datVendedores 
      Caption         =   "datVendedores"
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
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Data datOperacao 
      Caption         =   "datOperacao"
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM [Operações Saída] ORDER BY Nome"
      Top             =   7680
      Width           =   2295
   End
   Begin Crystal.CrystalReport crpView 
      Left            =   360
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   4680
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "frmRelVendasPorVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstFiliais As Recordset
  
  txtNomeFilial.Text = ""
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstFiliais = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & cboFilial.Text, dbOpenSnapshot)
  
  With rstFiliais
    If Not (.BOF And .EOF) Then
      txtNomeFilial.Text = .Fields("Nome") & ""
    End If
    
    If Not rstFiliais Is Nothing Then .Close
    Set rstFiliais = Nothing
  End With
End Sub

Private Sub cboVendedor_CloseUp()
  cboVendedor.Text = cboVendedor.Columns(0).Text
  cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_LostFocus()
  Dim rstFuncionarios As Recordset
  
  txtNomeVendedor.Text = ""
  If Not IsNumeric(cboVendedor.Text) Then Exit Sub
  
  Set rstFuncionarios = db.OpenRecordset(" SELECT Nome FROM Funcionários " & _
                                         " WHERE Código = " & cboVendedor.Text, dbOpenSnapshot)
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      txtNomeVendedor.Text = .Fields("Nome") & ""
    End If
    
    .Close
    Set rstFuncionarios = Nothing
  End With
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  
  Dim intFormulas As Integer
  Dim strNomeArquivo As String
  Dim intContador As Integer
  
  On Error GoTo ErrHandler
  
  Rem Verifica empresa
  If IsNull(txtNomeFilial.Text) Or txtNomeFilial.Text = "" Then
    DisplayMsg "Escolha a empresa."
    cboFilial.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(cboFilial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If
  
  If (Not IsDate(mskDataInicio1(0).Text)) And (Not IsDate(mskDataFinal1(0).Text)) Then
    mskDataInicio1(0).Text = Data_Atual
    mskDataFinal1(0).Text = Data_Atual
  End If
  
  If Not IsDate(mskDataInicio1(0).Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(mskDataFinal1(0).Text) Then
    MsgBox "Data final inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(mskDataInicio1(0).Text) > CDate(mskDataFinal1(0).Text) Then
    MsgBox "A data inicial não pode ser maior que a data final !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If (Not IsDate(mskDataInicio2(1).Text)) And (Not IsDate(mskDataFinal2(1).Text)) Then
    mskDataInicio2(1).Text = Data_Atual
    mskDataFinal2(1).Text = Data_Atual
  End If
  
  If Not IsDate(mskDataInicio2(1).Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(mskDataFinal2(1).Text) Then
    MsgBox "Data final inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(mskDataInicio2(1).Text) > CDate(mskDataFinal2(1).Text) Then
    MsgBox "A data inicial não pode ser maior que a data final !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
    If (Not IsDate(mskDataInicio3(2).Text)) And (Not IsDate(mskDataFinal3(2).Text)) Then
    mskDataInicio3(2).Text = Data_Atual
    mskDataFinal3(2).Text = Data_Atual
  End If
  
  If Not IsDate(mskDataInicio3(2).Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(mskDataFinal3(2).Text) Then
    MsgBox "Data final inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(mskDataInicio3(2).Text) > CDate(mskDataFinal3(2).Text) Then
    MsgBox "A data inicial não pode ser maior que a data final !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If txtNomeVendedor.Text = "" Or IsNull(txtNomeVendedor.Text) Then
    cboVendedor.Text = "0"
  End If
  
  If txtNomeOperacao.Text = "" Or IsNull(txtNomeOperacao.Text) Then
    cboOperacao.Text = "0"
  End If
    
  Call StatusMsg("Exportando...")
  MousePointer = vbHourglass
  
  If g_blnRelVendasPorVendedor(cboFilial.Text, cboVendedor.Text, cboOperacao.Text, mskDataInicio1(0).Text, mskDataFinal1(0).Text, mskDataInicio2(1).Text, mskDataFinal2(1).Text, mskDataInicio3(2).Text, mskDataFinal3(2).Text, pgbProgress) Then
    
    With crpView
    
    'Rem  Nome do BD
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsTempDBFileName
    .DataFiles(3) = gsQuickDBFileName
    .DataFiles(4) = gsQuickDBFileName
    
    Rem Saída
    If optSaidaVideo.Value Then .Destination = crptToWindow
    If optSaidaImpressora.Value Then .Destination = crptToPrinter
      
    strNomeArquivo = gsReportPath & "VendaPorVendedor.rpt"
    .ReportFileName = strNomeArquivo
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crpView
        
    Rem data inicial
    .Formulas(0) = "Periodo1 = '" & "De " & mskDataInicio1(0).Text & " até " & mskDataFinal1(0).Text & "'"
    intFormulas = 1
    .Formulas(intFormulas) = "Periodo2 = '" & "De " & mskDataInicio2(1).Text & " até " & mskDataFinal2(1).Text & "'"
    intFormulas = intFormulas + 1
    .Formulas(intFormulas) = "Periodo3 = '" & "De " & mskDataInicio3(2).Text & " até " & mskDataFinal3(2).Text & "'"
    intFormulas = intFormulas + 1
    If Len(Trim(txtNomeFilial.Text)) > 0 Then
      .Formulas(intFormulas) = "Filtro_Filial = '" & txtNomeFilial.Text & "'"
      intFormulas = intFormulas + 1
    End If
    If Len(Trim(txtNomeVendedor.Text)) > 0 Then
      .Formulas(intFormulas) = "Filtro_Vendedor = '" & txtNomeVendedor.Text & "'"
      intFormulas = intFormulas + 1
    End If
    If Len(Trim(txtNomeOperacao.Text)) > 0 Then
      .Formulas(intFormulas) = "Filtro_Operacao = '" & cboOperacao.Text & " - " & txtNomeOperacao & "'"
      intFormulas = intFormulas + 1
    End If

    'Seta a impressora para relatório
    
    Call SetPrinterName("REL", crpView)

    .Action = 1

    For intContador = 0 To intFormulas - 1
      .Formulas(intContador) = ""
    Next

  End With
  End If

  Call StatusMsg("")
  MousePointer = vbDefault

  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  MsgBox "Erro ao imprimir relatório: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Form_Load()
  
  datFiliais.DatabaseName = gsQuickDBFileName
  datClientes.DatabaseName = gsQuickDBFileName
  datClasse.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
  datProdutos.DatabaseName = gsQuickDBFileName
  datVendedores.DatabaseName = gsQuickDBFileName

  datOperacao.DatabaseName = gsQuickDBFileName

  txtNomeOperacao.Width = 4455
    
  Call CenterForm(Me)
  
End Sub

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinal1(0).Text = frmCalendario.gsDateCalender(mskDataFinal1(0).Text)
  End If
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal1(0).Text = Ajusta_Data(mskDataFinal1(0).Text)
End Sub

Private Sub mskDataInicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicio1(0).Text = frmCalendario.gsDateCalender(mskDataInicio1(0).Text)
  End If
End Sub

Private Sub mskDataInicio_LostFocus()
  mskDataInicio1(0).Text = Ajusta_Data(mskDataInicio1(0).Text)
End Sub

Private Sub cboOperacao_CloseUp()
  cboOperacao.Text = cboOperacao.Columns(0).Text
  cboOperacao_LostFocus
End Sub

Private Sub cboOperacao_LostFocus()
  Dim rstOpSaida As Recordset
  
  txtNomeOperacao.Text = ""
  If Not IsNumeric(cboOperacao.Text) Then Exit Sub
  
  Set rstOpSaida = db.OpenRecordset(" SELECT Nome FROM [Operações Saída] " & _
                                         " WHERE Código = " & cboOperacao.Text, dbOpenSnapshot)
  With rstOpSaida
    If Not (.BOF And .EOF) Then
      txtNomeOperacao.Text = .Fields("Nome") & ""
    End If
    
    .Close
    Set rstOpSaida = Nothing
  End With
End Sub



