VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManConsigEntradas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção de consignação"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManConsigEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   11535
   Begin VB.TextBox txtTotalVendido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0,00"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdGerar 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Gerar Prestação de Contas"
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Geração com todos os ítens 'Selecionados' da Grid para o Acerto com o Fornecedor"
      Top             =   5880
      Width           =   2535
   End
   Begin SSDataWidgets_B.SSDBGrid grdGeral 
      Height          =   3855
      Left            =   120
      TabIndex        =   32
      Top             =   6720
      Width           =   11295
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   16
      RowHeight       =   423
      Columns.Count   =   16
      Columns(0).Width=   503
      Columns(0).Caption=   "Fil"
      Columns(0).Name =   "Filial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1720
      Columns(1).Caption=   "Data"
      Columns(1).Name =   "Data"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3466
      Columns(2).Caption=   "Fornecedor"
      Columns(2).Name =   "Fornecedor"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1164
      Columns(3).Caption=   "Nota"
      Columns(3).Name =   "Nota"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1111
      Columns(4).Caption=   "Seq"
      Columns(4).Name =   "Sequencia"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   900
      Columns(5).Caption=   "Linha"
      Columns(5).Name =   "Linha"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   1984
      Columns(6).Caption=   "Codigo"
      Columns(6).Name =   "Codigo"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   4101
      Columns(7).Caption=   "Nome"
      Columns(7).Name =   "Nome"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   873
      Columns(8).Caption=   "Qtde"
      Columns(8).Name =   "Qtde"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   1773
      Columns(9).Caption=   "Preco Custo"
      Columns(9).Name =   "Preco"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(9).Locked=   -1  'True
      Columns(10).Width=   1270
      Columns(10).Caption=   "Vendido"
      Columns(10).Name=   "Vendido"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(10).Locked=   -1  'True
      Columns(11).Width=   1640
      Columns(11).Caption=   "Prestação"
      Columns(11).Name=   "Prestacao"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(11).Locked=   -1  'True
      Columns(12).Width=   1720
      Columns(12).Caption=   "Selecionado"
      Columns(12).Name=   "Selecionado"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(12).Locked=   -1  'True
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "Acertados"
      Columns(13).Name=   "QtdeAcertada"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(13).Locked=   -1  'True
      Columns(13).HasForeColor=   -1  'True
      Columns(13).ForeColor=   255
      Columns(14).Width=   1244
      Columns(14).Caption=   "Estoque"
      Columns(14).Name=   "Estoque"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(14).Locked=   -1  'True
      Columns(15).Width=   1429
      Columns(15).Caption=   "Acertado"
      Columns(15).Name=   "Acertado"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(15).Locked=   -1  'True
      _ExtentX        =   19923
      _ExtentY        =   6800
      _StockProps     =   79
      Caption         =   "Resultado da pesquisa (Geral)"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab tabTabs 
      Height          =   3855
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sintético"
      TabPicture(0)   =   "frmManConsigEntradas.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdResultado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Analítico"
      TabPicture(1)   =   "frmManConsigEntradas.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdProdutos"
      Tab(1).ControlCount=   1
      Begin SSDataWidgets_B.SSDBGrid grdProdutos 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   11055
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   8
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         RowHeight       =   423
         Columns.Count   =   8
         Columns(0).Width=   2752
         Columns(0).Caption=   "Codigo"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   7779
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   1217
         Columns(2).Caption=   "Qtde"
         Columns(2).Name =   "Qtde"
         Columns(2).Alignment=   1
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   1508
         Columns(3).Caption=   "Vendido"
         Columns(3).Name =   "QtdeConsignada"
         Columns(3).Alignment=   1
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(4).Width=   1852
         Columns(4).Caption=   "Preco Custo"
         Columns(4).Name =   "Preco"
         Columns(4).Alignment=   1
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "LinhaProduto"
         Columns(5).Name =   "LinhaProduto"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   1905
         Columns(6).Caption=   "Selecionado"
         Columns(6).Name =   "Selecionado"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   1508
         Columns(7).Caption=   "Acertado"
         Columns(7).Name =   "Acertado"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         _ExtentX        =   19500
         _ExtentY        =   5741
         _StockProps     =   79
         Caption         =   "Produtos da nota ( Analítico )"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid grdResultado 
         Height          =   3255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   11055
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   8
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         CheckBox3D      =   0   'False
         RowHeight       =   423
         Columns.Count   =   8
         Columns(0).Width=   1058
         Columns(0).Caption=   "Filial"
         Columns(0).Name =   "Filial"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   2117
         Columns(1).Caption=   "Data"
         Columns(1).Name =   "Data"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   5794
         Columns(2).Caption=   "Fornecedor"
         Columns(2).Name =   "Fornecedor"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   1773
         Columns(3).Caption=   "Nota"
         Columns(3).Name =   "NotaFiscal"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(4).Width=   2117
         Columns(4).Caption=   "Sequencia"
         Columns(4).Name =   "Sequencia"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Total Nota"
         Columns(5).Name =   "TotalNota"
         Columns(5).Alignment=   1
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(5).Locked=   -1  'True
         Columns(6).Width=   2117
         Columns(6).Caption=   "Vendido"
         Columns(6).Name =   "TotalConsignada"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(6).Locked=   -1  'True
         Columns(7).Width=   1561
         Columns(7).Caption=   "Visualizar"
         Columns(7).Name =   "Status"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(7).Locked=   -1  'True
         Columns(7).Style=   2
         _ExtentX        =   19500
         _ExtentY        =   5741
         _StockProps     =   79
         Caption         =   "Resultado da pesquisa ( Sintético )"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox txtTotalConsignado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0,00"
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtTotalNota 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0,00"
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data datFornecedores 
      Caption         =   "datFornecedores"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_FOR WHERE Tipo = 'F'"
      Top             =   10680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "Intervalo para Notas Fiscais"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   6100
         TabIndex        =   36
         Top             =   120
         Width           =   2500
         Begin VB.TextBox txtNFFin 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   11
            Top             =   480
            Width           =   1005
         End
         Begin VB.TextBox txtNFIni 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            MaxLength       =   8
            TabIndex        =   10
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   1200
            TabIndex        =   38
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Produtos"
         Height          =   855
         Left            =   2700
         TabIndex        =   33
         Top             =   840
         Width           =   1695
         Begin VB.OptionButton optVendidos 
            Appearance      =   0  'Flat
            Caption         =   "Vendidos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optNaoVendidos 
            Appearance      =   0  'Flat
            Caption         =   "Não Vendidos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   550
            Width           =   1455
         End
      End
      Begin VB.OptionButton optGeral 
         Caption         =   "Visualização Geral"
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   1500
         Width           =   1695
      End
      Begin VB.OptionButton optDetalhada 
         Caption         =   "Detalhada"
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   1500
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame fraPeriodo 
         Appearance      =   0  'Flat
         Caption         =   " Período ( Data de Vendas ) "
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   8640
         TabIndex        =   29
         Top             =   120
         Width           =   2500
         Begin MSMask.MaskEdBox mskDataFinal 
            Height          =   315
            Left            =   1320
            TabIndex        =   13
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDataInicial 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   1320
            TabIndex        =   30
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.ComboBox cboStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmManConsigEntradas.frx":05C2
         Left            =   4560
         List            =   "frmManConsigEntradas.frx":05CF
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ordenação"
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   2535
         Begin VB.OptionButton optFornecedor 
            Appearance      =   0  'Flat
            Caption         =   "Fornecedor"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optNota 
            Appearance      =   0  'Flat
            Caption         =   "Nota"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optSequencia 
            Appearance      =   0  'Flat
            Caption         =   "Sequência"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optData 
            Appearance      =   0  'Flat
            Caption         =   "Data"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H0000C0C0&
         Caption         =   "Pesquisar"
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtNomeFornecedor 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   4755
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmManConsigEntradas.frx":061C
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   975
         DataFieldList   =   "Nome"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   7805
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).DataField=   "Nome"
         Columns(0).FieldLen=   256
         Columns(1).Width=   3731
         Columns(1).Caption=   "Codigo"
         Columns(1).Name =   "Codigo"
         Columns(1).DataField=   "Código"
         Columns(1).FieldLen=   256
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Status da Entrada"
         Height          =   195
         Left            =   4560
         TabIndex        =   28
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Label lblTotalVendido 
      AutoSize        =   -1  'True
      Caption         =   "Total Vendido R$"
      Height          =   195
      Left            =   5760
      TabIndex        =   35
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblMensagem 
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5910
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Consignado (R$)"
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   10710
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Notas (R$)"
      Height          =   255
      Left            =   7560
      TabIndex        =   19
      Top             =   10710
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmManConsigEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_lngSequencia(1 To 3000) As Long
Dim m_intCont      As Integer

Dim m_dblQtde(1 To 100) As Long
Dim m_intContador As Integer

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(1).Text
  txtNomeFornecedor.Text = FindFornecedor
End Sub

Private Sub cboFornecedor_LostFocus()
  txtNomeFornecedor.Text = FindFornecedor
End Sub

Private Sub cmdGerar_Click()
  '17/09/2004 - Daniel
  Dim intAuxi            As Integer
  Dim intI               As Integer
  Dim varBook            As Variant
  Dim strAuxi            As String
  Dim bytAsc             As Byte
  Dim lngFornecedor      As Long
  Dim strSQL             As String
  Dim intLinhadaGrid     As Integer
  
  Dim rstPrestaContas As Recordset
  
  If ExaminaSelecao Then Exit Sub
  
  '17/11/2004 - Daniel
  'If QtdeAcertada = Vendido exit...
  '15/12/2004 - Daniel
  'a coluna acertado mostrará exatamente quantas vezes o
  'produto já foi acertado então cancelamos as próximas linhas
  'If optVendidos.Value Then
  '  For intAuxi = 0 To (grdGeral.SelBookmarks.Count - 1)
  '    varBook = grdGeral.SelBookmarks(intAuxi)
  '    grdGeral.Bookmark = varBook
  '
  '    intLinhadaGrid = intAuxi + 1
  '
  '    If grdGeral.Columns("QtdeAcertada").Text = grdGeral.Columns("Vendido").Text Then
  '      MsgBox "A informação da Linha " & intLinhadaGrid & " já foi acertada, verifique.", vbExclamation, "Acertar 2x"
  '      Exit Sub
  '    End If
  '
  '  Next intAuxi
  'End If
  '---------------------------------------------------
  
  If optGeral.Value Then
    If VerificarColunaSelecionado Then Exit Sub
  End If

  If MsgBox("Deseja gerar a Prestação de Contas? ", vbQuestion + vbYesNo, "Atenção") = vbNo Then Exit Sub

  Set rstPrestaContas = db.OpenRecordset("PrestacaoContas", dbOpenDynaset)

  lblMensagem.Caption = ""

  Call StatusMsg("Aguarde gerando a Prestação de Contas...")
  Screen.MousePointer = vbHourglass

'  grdGeral.MoveFirst
'  For intAuxi = 0 To (grdGeral.Rows - 1)

  For intAuxi = 0 To (grdGeral.SelBookmarks.Count - 1)
    varBook = grdGeral.SelBookmarks(intAuxi)
    grdGeral.Bookmark = varBook
  
    '1º - Atualizar o campo [Entradas - Produtos].Selecionado
    Call AtualizarEntraProd(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text, grdGeral.Columns("Linha").Text)
    
    '2º - Criar registro na table PrestacaoContas
    
    'Pegar apenas o número da string
    'Exemplo: '45 - AAAA devemos guardar na string strAuxi somente o 45
    For intI = 1 To Len(grdGeral.Columns("Fornecedor").Text)
      bytAsc = Asc(Mid(((grdGeral.Columns("Fornecedor").Text)), intI, 1))
      If bytAsc = 32 Then Exit For '32 é o espaço em branco
      strAuxi = strAuxi & Mid(((grdGeral.Columns("Fornecedor").Text)), intI, 1)
    Next intI
    
    lngFornecedor = CLng(strAuxi)
    strAuxi = ""
    
    With rstPrestaContas
      .AddNew
      .Fields("Filial").Value = CByte(grdGeral.Columns("Filial").Text)
      .Fields("Fornecedor").Value = lngFornecedor
      .Fields("Sequencia").Value = CLng(grdGeral.Columns("Sequencia").Text)
      .Fields("Linha").Value = CByte(grdGeral.Columns("Linha").Text)
      .Fields("Produto").Value = CStr(grdGeral.Columns("Codigo").Text)
      .Fields("Custo").Value = CDbl(grdGeral.Columns("Preco").Text)
      .Fields("QtdeOriginal").Value = CDbl(grdGeral.Columns("Qtde").Text)
      .Fields("QtdeDevolvida").Value = 0
      .Fields("QtdeComprada").Value = 0
      .Fields("QtdeVendida").Value = CDbl(grdGeral.Columns("Vendido").Text)
      .Fields("DatadaGeracao").Value = Data_Atual
      .Fields("Resultado").Value = 0
      '15/12/2004 - Daniel
      If optVendidos.Value Then
        .Fields("PeriodoVenda").Value = CDate(mskDataInicial.Text)
      Else
        .Fields("PeriodoVenda").Value = Data_Atual
      End If
      
      If Len(grdGeral.Columns("Vendido").Text) > 0 Then
        .Fields("NotaFiscal").Value = CLng(grdGeral.Columns("Nota").Text)
      Else
        .Fields("NotaFiscal").Value = 0
      End If
      
      'No primeiro momento a QtdeAcertada será zero para este record
      .Fields("QtdeAcertada").Value = 0
      
      .Update
    End With
  
'    grdGeral.MoveNext
  
  Next intAuxi
  
  rstPrestaContas.Close
  Set rstPrestaContas = Nothing
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  MsgBox "Geração finalizada com sucesso.", vbInformation, "Quick Store"
  
  txtTotalVendido.Text = "0,00"
  
  grdGeral.Redraw = False
  grdGeral.RemoveAll
  grdGeral.Refresh
  grdGeral.Redraw = True

End Sub

Private Sub cmdPesquisar_Click()
  '---[ Declaração das variáveis ]---'
    Dim strSQL          As String
    Dim rstEntradas     As Recordset
    Dim dblTotalNotas   As Double
    Dim sngEstoqueAtual As Single
    Dim dblTotVendido   As Double
    Dim dblQtdeAcertada As Double
    Dim lngSeqVenda     As Long
  '---[ Declaração das variáveis ]---'
  
  'Chama função que valida os campos de data
  If optVendidos.Value Then
    If Not ValidaDatas() Then Exit Sub
  End If
  
  '09/12/2004 - Daniel
  'Tratamento para carregar somente o que vendeu para prestar contas
  'buscando as informações da tabela de AcertoConsignacaoEntrada
  'e não mais de Entradas e [Entradas - Produtos] como foi a princípio
  If optVendidos.Value Then
    Dim i As Integer
    'Limpamos as vars modulares
    m_intCont = 0
    m_intContador = 0
    
    For i = 1 To 100
      m_lngSequencia(i) = 0
      m_dblQtde(i) = 0
    Next i
    '---------------------------
    
    Call StatusMsg("Verificando os produtos vendidos, aguarde...")
    Call CarregarGridComBooksVendidos
    Call StatusMsg("")
    Exit Sub
  End If
  
  '16/09/2004 - Daniel
  'Adicionado a string a table [Entradas - Produtos].*
  strSQL = " SELECT Entradas.*, [Operações Entrada].*, [Entradas - Produtos].* "
  strSQL = strSQL & " FROM Entradas, [Operações Entrada], [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Entradas.Filial = " & gnCodFilial
  strSQL = strSQL & " AND Entradas.Operação = [Operações Entrada].Código "
  strSQL = strSQL & " AND [Operações Entrada].Tipo = 'E' "
  
  If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
    strSQL = strSQL & " AND Entradas.Fornecedor = " & CLng(cboFornecedor.Text)
  End If
  
  strSQL = strSQL & " AND Entradas.[Data Acerto Empréstimo] >= #" & Format("01/01/2000", "mm/dd/yyyy") & "# "
  strSQL = strSQL & " AND Entradas.[Data Acerto Empréstimo] <= #" & Format("01/01/2100", "mm/dd/yyyy") & "# "
  
  '08/10/2004 - Daniel
  'Adicionado filtro de notas fiscais
  If Len(txtNFIni.Text) > 0 And Len(txtNFFin.Text) > 0 Then
    If CLng(txtNFIni.Text) <= CLng(txtNFFin.Text) Then
      strSQL = strSQL & " AND Entradas.[Nota Fiscal] >= '" & CStr(txtNFIni.Text) & "'"
      strSQL = strSQL & " AND Entradas.[Nota Fiscal] <= '" & CStr(txtNFFin.Text) & "'"
    End If
  End If
  
  Select Case GetCodigoCombos(cboStatus.Text)
    Case 1
      'strSQL = strSQL & " AND Entradas.ConsignacaoFechada = False "
      strSQL = strSQL & " AND [Entradas - Produtos].Acertado = False "
      strSQL = strSQL & " AND NOT [Entradas - Produtos].Selecionado " 'Critério do campo Selecionado
    Case 2
      'strSQL = strSQL & " AND Entradas.ConsignacaoFechada = True "
      strSQL = strSQL & " AND [Entradas - Produtos].Acertado = True "
      strSQL = strSQL & " AND [Entradas - Produtos].Selecionado " 'Critério do campo Selecionado
  End Select
  
  '16/09/2004 - Daniel
  strSQL = strSQL & " AND [Entradas - Produtos].Sequência = Entradas.Sequência "
  
  '06/10/2004 - Daniel
  If optVendidos.Value Then
    strSQL = strSQL & " AND [Entradas - Produtos].QtdeAtual <> [Entradas - Produtos].Qtde  "
  Else
    strSQL = strSQL & " AND [Entradas - Produtos].QtdeAtual = [Entradas - Produtos].Qtde "
  End If
  
  'Ordena o recordset
  strSQL = strSQL & " ORDER BY "
  
  If optData.Value Then
    strSQL = strSQL & " Entradas.Data "
  End If
  
  If optSequencia.Value Then
    strSQL = strSQL & " Entradas.Sequência "
  End If
  
  If optNota.Value Then
    strSQL = strSQL & " Entradas.[Nota Fiscal] "
  End If
  
  If optFornecedor.Value Then
    strSQL = strSQL & " Entradas.Fornecedor "
  End If
  
  '16/09/2004 - Daniel
  strSQL = strSQL & " , [Entradas - Produtos].Linha "
  
  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  '16/09/2004 - Daniel
  If optDetalhada.Value Then
    grdResultado.Redraw = False
    grdResultado.RemoveAll
  Else
    grdGeral.Redraw = False
    grdGeral.RemoveAll
  End If
  
  dblTotalNotas = 0
    
  Call StatusMsg("Verificando consignações, aguarde...")
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        
        dblTotalNotas = dblTotalNotas + .Fields("Total").Value
        
        '16/09/2004 - Daniel
        If optDetalhada.Value Then
          grdResultado.AddNew
          
          grdResultado.Columns("Filial").Value = .Fields("Entradas.Filial").Value
          grdResultado.Columns("Sequencia").Value = .Fields("Entradas.Sequência").Value
          grdResultado.Columns("Data").Value = .Fields("Data").Value
          grdResultado.Columns("Nota").Value = .Fields("Nota Fiscal").Value
          grdResultado.Columns("TotalNota").Value = Format(.Fields("Total").Value, FORMAT_VALUE)
          
          '14/12/2004 - Daniel
          'Nem entrará neste bloco pois o optVendidos não estará = TRUE
          'If optVendidos.Value Then
          '  Call BuscarSeqVenda(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Código").Value, lngSeqVenda)
          'Else
          '  Call BuscarSeqVenda(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Código").Value, lngSeqVenda)
          'End If
          'grdResultado.Columns("TotalConsignada").Value = Format(getAcertosConsignacao(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Código").Value, lngSeqVenda), FORMAT_VALUE)
          grdResultado.Columns("TotalConsignada").Value = Format(getAcertosConsignacaoNaoVendidos(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Código").Value), FORMAT_VALUE)
          
          If IsNumeric(.Fields("Fornecedor").Value) Then
            grdResultado.Columns("Fornecedor").Value = .Fields("Fornecedor").Value & " - " & FindFornecedor(CLng(.Fields("Fornecedor").Value))
          End If
          grdResultado.Update
        
        Else 'grdGeral
        
          grdGeral.AddNew
          
          grdGeral.Columns("Filial").Value = .Fields("Entradas.Filial").Value
          grdGeral.Columns("Data").Value = .Fields("Data").Value
          
          If IsNumeric(.Fields("Fornecedor").Value) Then
            grdGeral.Columns("Fornecedor").Value = .Fields("Fornecedor").Value & " - " & FindFornecedor(CLng(.Fields("Fornecedor").Value))
          End If
          
          grdGeral.Columns("Sequencia").Value = .Fields("Entradas.Sequência").Value
          grdGeral.Columns("Nota").Value = .Fields("Nota Fiscal").Value
          grdGeral.Columns("Codigo").Value = .Fields("Entradas - Produtos.Código").Value
          grdGeral.Columns("Nome").Value = FindProduto(CStr(.Fields("Entradas - Produtos.Código").Value))
          grdGeral.Columns("Linha").Value = .Fields("Linha").Value
          grdGeral.Columns("Qtde").Value = .Fields("Qtde").Value
          '14/12/2004 - Daniel
          'Nem entrará neste bloco quando optVendido = TRUE
          'If optVendidos.Value Then
          '  Call BuscarSeqVenda(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Código").Value, lngSeqVenda)
          'Else
          '  Call BuscarSeqVenda(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Código").Value, lngSeqVenda)
          'End If
          
          'grdGeral.Columns("Vendido").Value = getAcertosConsignacao(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Código").Value, lngSeqVenda)
          grdGeral.Columns("Vendido").Value = getAcertosConsignacaoNaoVendidos(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Código").Value)
          
          grdGeral.Columns("Preco").Value = Format(.Fields("Preço").Value, FORMAT_VALUE)
          grdGeral.Columns("Prestacao").Value = Format(grdGeral.Columns("Vendido").Value * grdGeral.Columns("Preco").Value, FORMAT_VALUE)
          
          'Adicionado para Visualização as Colunas Selecionado e Acertado
          If .Fields("Selecionado").Value Then
            grdGeral.Columns("Selecionado").Value = "SIM"
          Else
            grdGeral.Columns("Selecionado").Value = "NÃO"
          End If
          
          If .Fields("Acertado").Value Then
            grdGeral.Columns("Acertado").Value = "SIM"
          Else
            grdGeral.Columns("Acertado").Value = "NÃO"
          End If
          
          Call BuscarEstoqueAtual(.Fields("Entradas.Filial").Value, .Fields("Entradas - Produtos.Código").Value, sngEstoqueAtual)
          grdGeral.Columns("Estoque").Value = sngEstoqueAtual
          
          Call BuscarQtdeAcertada(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Linha").Value, dblQtdeAcertada)
          grdGeral.Columns("QtdeAcertada").Value = dblQtdeAcertada
          
          dblTotVendido = dblTotVendido + Format((CDbl(grdGeral.Columns("Preco").Value) * CDbl(grdGeral.Columns("Vendido").Value)), FORMAT_VALUE)
          
          grdGeral.Update
        End If
        
        .MoveNext
      Loop
    End If
    
    txtTotalVendido.Text = Format(dblTotVendido, FORMAT_VALUE)
    'txtTotalNota.Text = Format(dblTotalNotas, FORMAT_VALUE)
    lblMensagem.Caption = "Pesquisa concluída, " & .RecordCount & " registros encontrados."
    .Close
  End With
  
  'Estava sumindo o nome do Fornecedor...
  cboFornecedor_LostFocus
  
  Call StatusMsg("")
  
  '16/09/2004 - Daniel
  If optDetalhada.Value Then
    tabTabs.Tab = 0
    
    grdResultado.MoveFirst
    grdResultado.Redraw = True
  Else
    grdGeral.MoveFirst
    grdGeral.Redraw = True
  End If
  
  Set rstEntradas = Nothing
End Sub

Private Function ValidaDatas() As Boolean
  Dim blnValida As Boolean
  blnValida = True
  
  If Not IsDate(mskDataInicial.Text) Then
    MsgBox "Data inicial inválida, verifique !!", vbCritical, "Quick Store"
    blnValida = False
    mskDataInicial.SetFocus
  End If
  
  If blnValida Then
    If Not IsDate(mskDataFinal.Text) Then
      MsgBox "Data final inválida, verifique !!", vbCritical, "Quick Store"
      blnValida = False
      mskDataFinal.SetFocus
    End If
  End If
  
  If blnValida Then
    If (CDate(mskDataInicial.Text) > CDate(mskDataFinal.Text)) Then
      MsgBox "A data inicial não pode ser maior que a final, verifique !!", vbCritical, "Quick Store"
      blnValida = False
      mskDataFinal.SetFocus
    End If
  End If
  
  If blnValida Then
    If Len(Trim(cboStatus.Text)) <= 0 Then
      MsgBox "Nenhum filtro para status escolhido, verifique !!", vbCritical, "Quick Store"
      blnValida = False
    End If
  End If
  
  ValidaDatas = blnValida
End Function

Private Function FindFornecedor(Optional lngFornecedor As Long)
  Dim strSQL    As String
  Dim rstForn   As Recordset
  Dim lng_Fornecedor As Long
  
  txtNomeFornecedor.Text = ""
  
  If lngFornecedor <= 0 Then
    If Len(Trim(cboFornecedor.Text)) <= 0 Then Exit Function
    If Not IsNumeric(Trim(cboFornecedor.Text)) Then Exit Function
    
    lng_Fornecedor = CLng(cboFornecedor.Text)
  Else
    lng_Fornecedor = lngFornecedor
  End If
  
  strSQL = " SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' AND "
  strSQL = strSQL & " Código = " & lng_Fornecedor
  
  Set rstForn = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstForn
    If Not (.BOF And .EOF) Then
      FindFornecedor = .Fields("Nome").Value & ""
    Else
      FindFornecedor = ""
    End If
    .Close
  End With
  
  Set rstForn = Nothing
End Function

Private Sub Form_Load()
  tabTabs.TabEnabled(1) = False
  Call CenterForm(Me)
  datFornecedores.DatabaseName = gsQuickDBFileName
  
  lblTotalVendido.Visible = False
  txtTotalVendido.Visible = False
  
'  mskDataInicial.Text = ""
'  mskDataFinal.Text = ""
  
  '16/09/2004 - Daniel
  tabTabs.Visible = True
  tabTabs.Top = 1920
  grdGeral.Visible = False
  grdGeral.Top = 6720
  
  cboStatus.Text = "1 - Apenas as em aberto"
  cmdGerar.Visible = False
  
  grdGeral.Columns("Acertado").Visible = False 'Não precisa mostrar...
  
End Sub

Private Sub LoadAnalitico(bytFilial As Byte, lngSequencia As Long)
  Dim strSQL As String
  Dim rstProdutos As Recordset
  
  strSQL = "SELECT * FROM [Entradas - Produtos] WHERE Filial = " & bytFilial & " AND Sequência = " & lngSequencia
  
  Set rstProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  grdProdutos.Redraw = False
  grdProdutos.RemoveAll
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        grdProdutos.AddNew
        
        grdProdutos.Columns("Codigo").Value = .Fields("Código").Value
        grdProdutos.Columns("Nome").Value = FindProduto(.Fields("Código").Value)
        grdProdutos.Columns("Qtde").Value = .Fields("Qtde").Value
        grdProdutos.Columns("Preco").Value = Format(.Fields("Preço").Value, FORMAT_VALUE)
        grdProdutos.Columns("QtdeConsignada").Value = getProdutoAcertoConsignacao(bytFilial, lngSequencia, .Fields("Linha").Value, .Fields("Código").Value)
        
        If .Fields("Selecionado").Value Then
          grdProdutos.Columns("Selecionado").Value = "SIM"
        Else
          grdProdutos.Columns("Selecionado").Value = "NÃO"
        End If
        
        If .Fields("Acertado").Value Then
          grdProdutos.Columns("Acertado").Value = "SIM"
        Else
          grdProdutos.Columns("Acertado").Value = "NÃO"
        End If
        
        grdProdutos.Update
        
        .MoveNext
      Loop
    End If
    .Close
  End With
  
  grdProdutos.MoveFirst
  grdProdutos.Redraw = True
  
  Set rstProdutos = Nothing
End Sub

Private Sub grdGeral_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
'  DispPromptMsg = False
'  If gbPodeApagar = False Then
'    Beep
'    Cancel = True
'    Exit Sub
'  End If
'
'  If bGridBeforeDelete Then
'    Call StatusMsg("Seleção de itens apagada.")
'    Cancel = False
'  Else
'    Cancel = True
'  End If
End Sub

Public Function bGridBeforeDelete() As Boolean
'  gsTitle = LoadResString(201)
'  gsMsg = "Apagar seleção atual?"
'  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
'  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'
'  If gnResponse = vbNo Then
'    bGridBeforeDelete = False
'  Else
'    bGridBeforeDelete = True
'  End If
End Function

Private Sub grdResultado_DblClick()
  LoadAnalitico grdResultado.Columns("Filial").Value, grdResultado.Columns("Sequencia").Value
  tabTabs.Tab = 1
  tabTabs.TabEnabled(1) = True
End Sub

Private Function FindProduto(strCodigo As String)
  Dim strSQL    As String
  Dim rstProd   As Recordset
  
  strSQL = " SELECT Nome FROM Produtos WHERE Código = '" & strCodigo & "'"
  
  Set rstProd = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstProd
    If Not (.BOF And .EOF) Then
      FindProduto = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  
  Set rstProd = Nothing
End Function

Private Function getAcertosConsignacao(bytFilial As Byte, lngSequencia As Long, ByVal strProduto As String, ByVal SeqVenda As Long) As Double
  Dim strSQL        As String
  Dim rstAcerConsig As Recordset
  Dim dblVend       As Double
  Dim intAuxi       As Integer
  Dim i             As Integer
  
  '---[Limpando as modulares]---
    m_intCont = 0
    m_intContador = 0

    For i = 1 To 100
      m_lngSequencia(i) = 0
      m_dblQtde(i) = 0
    Next i
  '-----------------------------
  
  strSQL = ""
  strSQL = " SELECT * FROM AcertoConsignacaoEntrada WHERE "
  strSQL = strSQL & " Filial = " & bytFilial & " AND Sequencia = " & lngSequencia
  '22/09/2004 - Daniel - Adicionado AND abaixo
  strSQL = strSQL & " AND CodigoProduto = '" & strProduto & "'"
  
  '06/10/2004 - Daniel - Adicionado AND abaixo
  If optVendidos.Value Then
    strSQL = strSQL & " AND DataAcerto >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
    strSQL = strSQL & " AND DataAcerto <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  End If
  
  '14/12/2004 - Daniel
  'Otimização para pegar a Quantidade de Venda da Sequência de Venda Exata
  strSQL = strSQL & " AND SequenciaVenda = " & SeqVenda
  
  
  dblVend = 0
  
  Set rstAcerConsig = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstAcerConsig
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .MoveLast
      .MoveFirst
      'Do Until .EOF
      '  dblVend = dblVend + .Fields("QtdeVendida").Value
      '
      ' .MoveNext
      'Loop
      
      If .RecordCount = 1 Then
        dblVend = dblVend + .Fields("QtdeVendida").Value
      Else
      
        m_intContador = m_intContador + 1
        
        If m_intContador = 1 Then 'Na Primeira vez criamos o Array com as Sequências
          Do Until .EOF
            intAuxi = intAuxi + 1
          
            m_dblQtde(intAuxi) = .Fields("QtdeVendida").Value
          
           .MoveNext
          Loop
        End If
        
        dblVend = m_dblQtde(m_intContador)
      
      End If
      
      .Close
    End If
  End With
  
  getAcertosConsignacao = dblVend
  
  Set rstAcerConsig = Nothing
  
End Function

Private Function getProdutoAcertoConsignacao(bytFilial As Byte, lngSequencia As Long, intLinha As Integer, strCodigoProduto As String) As Double
  Dim strSQL        As String
  Dim rstProdAcerConsig As Recordset
  Dim dblVend     As Double
  
  strSQL = " SELECT * FROM AcertoConsignacaoEntrada WHERE "
  strSQL = strSQL & " Filial = " & bytFilial & " AND Sequencia = " & lngSequencia
  strSQL = strSQL & " AND LinhaProduto = " & intLinha
  strSQL = strSQL & " AND CodigoProduto = '" & strCodigoProduto & "'"
  
  dblVend = 0
  
  Set rstProdAcerConsig = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstProdAcerConsig
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do Until .EOF
      dblVend = dblVend + .Fields("QtdeVendida").Value
      .MoveNext
    Loop
    .Close
  End With
  
  getProdutoAcertoConsignacao = dblVend
  
  Set rstProdAcerConsig = Nothing
End Function

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinal.Text = frmCalendario.gsDateCalender(mskDataFinal.Text)
  End If
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal.Text = Ajusta_Data(mskDataFinal.Text)
End Sub

Private Sub mskDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicial.Text = frmCalendario.gsDateCalender(mskDataInicial.Text)
  End If
End Sub

Private Sub mskDataInicial_LostFocus()
  mskDataInicial.Text = Ajusta_Data(mskDataInicial.Text)
End Sub

Private Sub optDetalhada_Click()
  tabTabs.Visible = True
  tabTabs.Top = 1920
  grdGeral.Visible = False
  grdGeral.Top = 6720
  cmdGerar.Visible = False
  
  lblTotalVendido.Visible = False
  txtTotalVendido.Visible = False
End Sub

Private Sub optGeral_Click()
  tabTabs.Visible = False
  grdGeral.Visible = True
  grdGeral.Top = 1920
  cmdGerar.Visible = True
  
  lblTotalVendido.Visible = True
  txtTotalVendido.Visible = True
End Sub

Private Sub AtualizarEntraProd(ByVal bytFilial As Byte, ByVal lngSequencia As Long, ByVal bytLinha As Byte)
  Dim rstEntraProd As Recordset
  Dim strSQL       As String
  
  strSQL = "SELECT Selecionado FROM [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & bytFilial
  strSQL = strSQL & " AND Sequência = " & lngSequencia
  strSQL = strSQL & " AND Linha = " & bytLinha
  
  Set rstEntraProd = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntraProd
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      .Edit
      .Fields("Selecionado").Value = True
      .Update
    End If
    .Close
  End With
  
  Set rstEntraProd = Nothing

End Sub

Private Function VerificarColunaSelecionado() As Boolean
  Dim intAuxi  As Integer
  'Dim intLinha As Integer
  Dim varBook  As Variant

  For intAuxi = 0 To (grdGeral.SelBookmarks.Count - 1)
    varBook = grdGeral.SelBookmarks(intAuxi)
    grdGeral.Bookmark = varBook
    
    If grdGeral.Columns("Selecionado").Value = "SIM" Then
      VerificarColunaSelecionado = True
      MsgBox "Alguma linha já foi Selecionada para Prestação de Contas, verifique.", vbExclamation, "Coluna Selecionado"
      Exit Function
    End If

  Next intAuxi

'  grdGeral.MoveFirst
'
'  For intAuxi = 0 To (grdGeral.Rows - 1)
'    intLinha = intAuxi + 1
'
'    If grdGeral.Columns("Selecionado").Value = "SIM" Then
'      VerificarColunaSelecionado = True
'      MsgBox "A informação da Linha " & intLinha & " já foi selecionada para Prestação de Contas.", vbExclamation, "Atenção"
'      Exit Function
'    End If
'
'  grdGeral.MoveNext
'  Next intAuxi

End Function

Private Sub BuscarEstoqueAtual(ByVal Filial As Byte, ByVal CodProduto As String, ByRef EstoqueAtual As Single)
  Dim rstEstoqueFinal As Recordset
  Dim strSQL          As String

  strSQL = "SELECT [Estoque Atual] FROM [Estoque Final] "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Produto = '" & CodProduto & "'"
  
  Set rstEstoqueFinal = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEstoqueFinal
    If Not (.BOF And .EOF) Then
      .MoveFirst
      EstoqueAtual = .Fields("Estoque Atual").Value
    End If
    .Close
  End With
  
  Set rstEstoqueFinal = Nothing

End Sub

Private Function ExaminaSelecao() As Boolean
  If grdGeral.SelBookmarks.Count < 1 Then
    MsgBox "Não existem informações selecionadas para Geração da Prestação de Contas, verifique.", vbExclamation, "Atenção"
    ExaminaSelecao = True
  End If
End Function

Private Sub BuscarQtdeAcertada(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte, ByRef QtdeAcertada As Double)
  Dim rstPrestacaoContas As Recordset
  Dim strQuery           As String

  QtdeAcertada = 0

  strQuery = "SELECT QtdeAcertada FROM PrestacaoContas "
  strQuery = strQuery & " WHERE Filial = " & Filial
  strQuery = strQuery & " AND Sequencia = " & Sequencia
  strQuery = strQuery & " AND Linha = " & Linha

  Set rstPrestacaoContas = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstPrestacaoContas.RecordCount = 0 Then
    rstPrestacaoContas.Close
    Set rstPrestacaoContas = Nothing
    Exit Sub
  End If

  With rstPrestacaoContas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        QtdeAcertada = QtdeAcertada + .Fields("QtdeAcertada").Value
        
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstPrestacaoContas = Nothing

End Sub

Private Function ValidarObjetos() As Boolean
  
  If Len(txtNomeFornecedor.Text) <= 0 Then
    ValidarObjetos = False
    MsgBox "Fornecedor inválido, verifique.", vbExclamation, "Quick Store"
    cboFornecedor.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataInicial.Text) Then
    ValidarObjetos = False
    MsgBox "Data Inicial inválida, verifique.", vbExclamation, "Quick Store"
    mskDataInicial.SetFocus
    Exit Function
  End If

  If Not IsDate(mskDataFinal.Text) Then
    ValidarObjetos = False
    MsgBox "Data Final inválida, verifique.", vbExclamation, "Quick Store"
    mskDataFinal.SetFocus
    Exit Function
  End If
  
  If CDate(mskDataFinal.Text) < CDate(mskDataInicial.Text) Then
    ValidarObjetos = False
    MsgBox "Data Final menor que a Inicial, verifique.", vbExclamation, "Quick Store"
    mskDataFinal.SetFocus
    Exit Function
  End If

End Function

Private Sub CarregarGridComBooksVendidos()
  'Private desenvolvida em 13/12/2004
  Dim rstAcerto       As Recordset
  Dim rstEntradas     As Recordset
  Dim rstBooks        As Recordset
  Dim strSQL          As String
  Dim sngEstoqueAtual As Single
  Dim dblTotalNotas   As Double
  Dim dblTotVendido   As Double
  Dim dblQtdeAcertada As Double

  Dim lngSeqVenda     As Long

  'Esvaziar a tabela temporária
  db.Execute "DELETE * FROM BooksVendidos"
  'Abrimos a BooksVendidos para o addnew
  Set rstBooks = db.OpenRecordset("BooksVendidos", dbOpenDynaset)

  'Dentro deste SELECT tenho X Sequências de Entradas
  'é exatamente destas entradas que precisamos
  strSQL = "SELECT Filial, Sequencia, LinhaProduto FROM AcertoConsignacaoEntrada "
  strSQL = strSQL & " WHERE  Filial = " & gnCodFilial
  strSQL = strSQL & " AND DataAcerto >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND DataAcerto <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " ORDER BY DataAcerto "

  Set rstAcerto = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstAcerto
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
          rstBooks.AddNew
            rstBooks.Fields("Filial").Value = .Fields("Filial").Value
            rstBooks.Fields("Sequencia").Value = .Fields("Sequencia").Value
            rstBooks.Fields("Linha").Value = .Fields("LinhaProduto").Value
          rstBooks.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstAcerto = Nothing
  
  'Agora dentro da tabela temporária BooksVendidos
  'eliminamos as sequências que não são do fornecedor
  'escolhido
  
  If rstBooks.RecordCount = 0 Then Exit Sub
  
  With rstBooks
    .MoveFirst
    
    Do Until .EOF
      If FornecedorDiff(.Fields("Filial").Value, .Fields("Sequencia").Value) Then .Delete
      
     .MoveNext
    Loop
    .Close
  End With
  
  Set rstBooks = Nothing
  
  'A partir deste momento temos apenas as Entradas que nos
  'interessam e pertencem ao Fornecedor escolhido
  strSQL = ""
  strSQL = "SELECT BooksVendidos.* , Entradas.*, [Entradas - Produtos].* "
  strSQL = strSQL & " FROM BooksVendidos, Entradas, [Entradas - Produtos] "
  strSQL = strSQL & " WHERE BooksVendidos.Filial = Entradas.Filial "
  strSQL = strSQL & " AND Entradas.Sequência = BooksVendidos.Sequencia "
  strSQL = strSQL & " AND [Entradas - Produtos].Filial = BooksVendidos.Filial "
  strSQL = strSQL & " AND [Entradas - Produtos].Sequência = BooksVendidos.Sequencia "
  strSQL = strSQL & " AND [Entradas - Produtos].Linha = BooksVendidos.Linha "

  'Adicionado filtro de notas fiscais
  If Len(txtNFIni.Text) > 0 And Len(txtNFFin.Text) > 0 Then
    If CLng(txtNFIni.Text) <= CLng(txtNFFin.Text) Then
      strSQL = strSQL & " AND Entradas.[Nota Fiscal] >= '" & CStr(txtNFIni.Text) & "'"
      strSQL = strSQL & " AND Entradas.[Nota Fiscal] <= '" & CStr(txtNFFin.Text) & "'"
    End If
  End If

  Select Case GetCodigoCombos(cboStatus.Text)
    Case 1
      'strSQL = strSQL & " AND Entradas.ConsignacaoFechada = False "
      strSQL = strSQL & " AND [Entradas - Produtos].Acertado = False "
      strSQL = strSQL & " AND NOT [Entradas - Produtos].Selecionado " 'Critério do campo Selecionado
    Case 2
      'strSQL = strSQL & " AND Entradas.ConsignacaoFechada = True "
      strSQL = strSQL & " AND [Entradas - Produtos].Acertado = True "
      strSQL = strSQL & " AND [Entradas - Produtos].Selecionado " 'Critério do campo Selecionado
  End Select

  'Ordena o recordset
  strSQL = strSQL & " ORDER BY "

  If optData.Value Then
    strSQL = strSQL & " Entradas.Data "
  End If

  If optSequencia.Value Then
    strSQL = strSQL & " Entradas.Sequência "
  End If

  If optNota.Value Then
    strSQL = strSQL & " Entradas.[Nota Fiscal] "
  End If

  If optFornecedor.Value Then
    strSQL = strSQL & " Entradas.Fornecedor "
  End If

  '16/09/2004 - Daniel
  strSQL = strSQL & " , [Entradas - Produtos].Linha "

  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  '16/09/2004 - Daniel
  If optDetalhada.Value Then
    grdResultado.Redraw = False
    grdResultado.RemoveAll
  Else
    grdGeral.Redraw = False
    grdGeral.RemoveAll
  End If
  
  dblTotalNotas = 0
    
  Call StatusMsg("Verificando consignações, aguarde...")
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        
        dblTotalNotas = dblTotalNotas + .Fields("Total").Value
        
        '16/09/2004 - Daniel
        If optDetalhada.Value Then
          grdResultado.AddNew
          
          grdResultado.Columns("Filial").Value = .Fields("Entradas.Filial").Value
          grdResultado.Columns("Sequencia").Value = .Fields("Entradas.Sequência").Value
          grdResultado.Columns("Data").Value = .Fields("Data").Value
          grdResultado.Columns("Nota").Value = .Fields("Nota Fiscal").Value
          grdResultado.Columns("TotalNota").Value = Format(.Fields("Total").Value, FORMAT_VALUE)
          '14/12/2004 - Daniel
          Call BuscarSeqVenda(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Código").Value, lngSeqVenda)
          'grdResultado.Columns("TotalConsignada").Value = Format(getAcertosConsignacao(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Código").Value, lngSeqVenda), FORMAT_VALUE)
          grdResultado.Columns("TotalConsignada").Value = Format(getAcertosConsignacao(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Código").Value, lngSeqVenda), FORMAT_VALUE)
          
          If IsNumeric(.Fields("Fornecedor").Value) Then
            grdResultado.Columns("Fornecedor").Value = .Fields("Fornecedor").Value & " - " & FindFornecedor(CLng(.Fields("Fornecedor").Value))
          End If
          grdResultado.Update
        
        Else 'grdGeral
        
          grdGeral.AddNew
          
          grdGeral.Columns("Filial").Value = .Fields("Entradas.Filial").Value
          grdGeral.Columns("Data").Value = .Fields("Data").Value
          
          If IsNumeric(.Fields("Fornecedor").Value) Then
            grdGeral.Columns("Fornecedor").Value = .Fields("Fornecedor").Value & " - " & FindFornecedor(CLng(.Fields("Fornecedor").Value))
          End If
          
          grdGeral.Columns("Sequencia").Value = .Fields("Entradas.Sequência").Value
          grdGeral.Columns("Nota").Value = .Fields("Nota Fiscal").Value
          'grdGeral.Columns("Codigo").Value = .Fields("Entradas - Produtos.Código").Value
          grdGeral.Columns("Codigo").Value = .Fields("Código").Value
          'grdGeral.Columns("Nome").Value = FindProduto(CStr(.Fields("Entradas - Produtos.Código").Value))
          grdGeral.Columns("Nome").Value = FindProduto(CStr(.Fields("Código").Value))
          'grdGeral.Columns("Linha").Value = .Fields("Linha").Value
          grdGeral.Columns("Linha").Value = .Fields("Entradas - Produtos.Linha").Value
          grdGeral.Columns("Qtde").Value = .Fields("Qtde").Value
          'grdGeral.Columns("Vendido").Value = getAcertosConsignacao(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Código").Value)
          Call BuscarSeqVenda(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Código").Value, lngSeqVenda)
          grdGeral.Columns("Vendido").Value = getAcertosConsignacao(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Código").Value, lngSeqVenda)
          grdGeral.Columns("Preco").Value = Format(.Fields("Preço").Value, FORMAT_VALUE)
          grdGeral.Columns("Prestacao").Value = Format(grdGeral.Columns("Vendido").Value * grdGeral.Columns("Preco").Value, FORMAT_VALUE)
          
          'Adicionado para Visualização as Colunas Selecionado e Acertado
          If .Fields("Selecionado").Value Then
            grdGeral.Columns("Selecionado").Value = "SIM"
          Else
            grdGeral.Columns("Selecionado").Value = "NÃO"
          End If
          
          If .Fields("Acertado").Value Then
            grdGeral.Columns("Acertado").Value = "SIM"
          Else
            grdGeral.Columns("Acertado").Value = "NÃO"
          End If
          
          'Call BuscarEstoqueAtual(.Fields("Entradas.Filial").Value, .Fields("Entradas - Produtos.Código").Value, sngEstoqueAtual)
          Call BuscarEstoqueAtual(.Fields("Entradas.Filial").Value, .Fields("Código").Value, sngEstoqueAtual)
          grdGeral.Columns("Estoque").Value = sngEstoqueAtual
          
          'Call BuscarQtdeAcertada(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Linha").Value, dblQtdeAcertada)
          Call BuscarQtdeAcertada(.Fields("Entradas.Filial").Value, .Fields("Entradas.Sequência").Value, .Fields("Entradas - Produtos.Linha").Value, dblQtdeAcertada)
          grdGeral.Columns("QtdeAcertada").Value = dblQtdeAcertada
          
          dblTotVendido = dblTotVendido + Format((CDbl(grdGeral.Columns("Preco").Value) * CDbl(grdGeral.Columns("Vendido").Value)), FORMAT_VALUE)
          
          grdGeral.Update
        End If
        
        .MoveNext
      Loop
    End If
    
    txtTotalVendido.Text = Format(dblTotVendido, FORMAT_VALUE)
    'txtTotalNota.Text = Format(dblTotalNotas, FORMAT_VALUE)
    lblMensagem.Caption = "Pesquisa concluída, " & .RecordCount & " registros encontrados."
    .Close
  End With
  
  'Estava sumindo o nome do Fornecedor...
  cboFornecedor_LostFocus
  
  Call StatusMsg("")
  
  '16/09/2004 - Daniel
  If optDetalhada.Value Then
    tabTabs.Tab = 0
    
    grdResultado.MoveFirst
    grdResultado.Redraw = True
  Else
    grdGeral.MoveFirst
    grdGeral.Redraw = True
  End If
  
  Set rstEntradas = Nothing

End Sub

Private Function FornecedorDiff(ByVal Filial As Byte, ByVal Seq As Long) As Boolean
  'Function desenvolvida em 13/12/2004
  Dim rstEntradas As Recordset
  Dim strSQL      As String
  
  strSQL = "SELECT Fornecedor FROM Entradas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequência = " & Seq
  
  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .Fields("Fornecedor").Value <> CLng(cboFornecedor.Text) Then FornecedorDiff = True
      
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing

End Function

Private Sub BuscarSeqVenda(ByVal Filial As Byte, ByVal Seq As Long, ByVal CodProd As String, ByRef SeqVenda As Long)
  Dim strSQL        As String
  Dim rstAcerConsig As Recordset
  Dim intAuxi       As Integer
  Dim i             As Integer
  
  strSQL = "SELECT * FROM AcertoConsignacaoEntrada WHERE "
  strSQL = strSQL & " Filial = " & Filial & " AND Sequencia = " & Seq
  strSQL = strSQL & " AND CodigoProduto = '" & CodProd & "'"
  
  If optVendidos.Value Then
    strSQL = strSQL & " AND DataAcerto >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
    strSQL = strSQL & " AND DataAcerto <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  End If

  Set rstAcerConsig = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  '---[Limpando as modulares]---
    m_intCont = 0
    m_intContador = 0

    For i = 1 To 100
      m_lngSequencia(i) = 0
      m_dblQtde(i) = 0
    Next i
  '-----------------------------
  
  With rstAcerConsig
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .MoveLast
      .MoveFirst
      
      If .RecordCount = 1 Then
        SeqVenda = .Fields("SequenciaVenda").Value
      Else
        m_intCont = m_intCont + 1
        
        If m_intCont = 1 Then 'Na Primeira vez criamos o Array com as Sequências
          Do Until .EOF
            intAuxi = intAuxi + 1
          
            m_lngSequencia(intAuxi) = .Fields("SequenciaVenda").Value
          
           .MoveNext
          Loop
        End If
        
        SeqVenda = m_lngSequencia(m_intCont)
        
      End If
      
    End If
    .Close
  End With
  
  Set rstAcerConsig = Nothing

End Sub

Private Function getAcertosConsignacaoNaoVendidos(bytFilial As Byte, lngSequencia As Long, ByVal strProduto As String) As Double
  Dim strSQL        As String
  Dim rstAcerConsig As Recordset
  Dim dblVend       As Double
  
  strSQL = ""
  strSQL = " SELECT * FROM AcertoConsignacaoEntrada WHERE "
  strSQL = strSQL & " Filial = " & bytFilial & " AND Sequencia = " & lngSequencia
  '22/09/2004 - Daniel - Adicionado AND abaixo
  strSQL = strSQL & " AND CodigoProduto = '" & strProduto & "'"
  
  '06/10/2004 - Daniel - Adicionado AND abaixo
  If optVendidos.Value Then
    strSQL = strSQL & " AND DataAcerto >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
    strSQL = strSQL & " AND DataAcerto <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  End If
  
  dblVend = 0
  
  Set rstAcerConsig = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstAcerConsig
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        dblVend = dblVend + .Fields("QtdeVendida").Value
      
       .MoveNext
      Loop
      .Close
    End If
  End With
  
  getAcertosConsignacaoNaoVendidos = dblVend
  
  Set rstAcerConsig = Nothing
  
End Function
