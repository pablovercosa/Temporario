VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"Produtos.frx":0000
   ClientHeight    =   7965
   ClientLeft      =   300
   ClientTop       =   480
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   30
   Icon            =   "Produtos.frx":00B0
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7965
   ScaleWidth      =   11340
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultar (F3)"
      Height          =   375
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Consultar Detalhes do Produto"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Data Data7 
      Caption         =   "TabPrecos"
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
      Left            =   6690
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Tabela From [Tabela de Preços] ORDER BY Tabela"
      Top             =   7380
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data datProdNome 
      Caption         =   "datProdNome"
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
      Height          =   360
      Left            =   5280
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código  FROM Produtos WHERE Código <> '0' ORDER BY [Código Ordenação]"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Data datProdCod 
      Caption         =   "datProdCod"
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
      Height          =   360
      Left            =   2850
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0' ORDER BY [Código Ordenação]"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Frame Frame7 
      Height          =   7230
      Left            =   30
      TabIndex        =   128
      Top             =   -30
      Width           =   11235
      Begin VB.ComboBox cboMotivoDesoneracao 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         ItemData        =   "Produtos.frx":4EA0A
         Left            =   2205
         List            =   "Produtos.frx":4EA2F
         TabIndex        =   234
         Top             =   6810
         Width           =   8940
      End
      Begin SSDataWidgets_B.SSDBCombo cboCodigo 
         Bindings        =   "Produtos.frx":4EBC4
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   3405
         DataFieldList   =   "Código"
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
         UseExactRowCount=   0   'False
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3175
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Código"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Código"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   9525
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6006
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
         DataFieldToDisplay=   "Código"
      End
      Begin TabDlg.SSTab tabItens 
         Height          =   5205
         Left            =   75
         TabIndex        =   13
         Top             =   1560
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   9181
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         Tab             =   2
         TabsPerRow      =   7
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
         TabCaption(0)   =   "&Características"
         TabPicture(0)   =   "Produtos.frx":4EBDD
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame8"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Fornecedores"
         TabPicture(1)   =   "Produtos.frx":4EBF9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame11"
         Tab(1).Control(1)=   "Data2_2"
         Tab(1).Control(2)=   "Data2_1"
         Tab(1).Control(3)=   "Frame3"
         Tab(1).Control(4)=   "DropDown2_1"
         Tab(1).Control(5)=   "Grade2_1"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "&Impostos/Comissão"
         TabPicture(2)   =   "Produtos.frx":4EC15
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame10"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "tabComissaoIndices"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Cá&lculos"
         TabPicture(3)   =   "Produtos.frx":4EC31
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame12"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "I&magem"
         TabPicture(4)   =   "Produtos.frx":4EC4D
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame4"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "&Web"
         TabPicture(5)   =   "Produtos.frx":4EC69
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame9"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "&Outros"
         TabPicture(6)   =   "Produtos.frx":4EC85
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "fraConfiguraEtiqueta"
         Tab(6).Control(1)=   "fraIdentificacaoCusto"
         Tab(6).Control(2)=   "fraEtiq"
         Tab(6).Control(3)=   "mskDataValidade"
         Tab(6).Control(4)=   "txtLote"
         Tab(6).Control(5)=   "txtVolumagem"
         Tab(6).Control(6)=   "lblTitle(5)"
         Tab(6).Control(7)=   "lblTitle(4)"
         Tab(6).Control(8)=   "lblTitle(1)"
         Tab(6).ControlCount=   9
         Begin VB.Frame Frame11 
            Caption         =   "Pareamento de Códigos de produto do Fornecedor"
            Height          =   1095
            Left            =   -74880
            TabIndex        =   238
            Top             =   3060
            Width           =   10890
            Begin VB.CommandButton cmd_consultarPareamentoFornecedor 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Todos os produtos do Fornecedor"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5550
               Style           =   1  'Graphical
               TabIndex        =   242
               Top             =   660
               Width           =   5145
            End
            Begin VB.CommandButton cmd_salvarPareamentoCodigoProdutoFornecedor 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Realizar pareamento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   241
               Top             =   660
               Width           =   5145
            End
            Begin VB.Label lbl_fornCodigo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               Height          =   315
               Left            =   210
               TabIndex        =   240
               Top             =   270
               Width           =   2115
            End
            Begin VB.Label lbl_fornNome 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               Height          =   315
               Left            =   2370
               TabIndex        =   239
               Top             =   270
               Width           =   8325
            End
         End
         Begin VB.Frame fraConfiguraEtiqueta 
            Caption         =   "Configuração para impressão de etiquetas"
            Height          =   1155
            Left            =   -71895
            TabIndex        =   230
            Top             =   390
            Width           =   4095
            Begin VB.TextBox txtDividirPrecoEtiqueta 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2160
               MaxLength       =   2
               TabIndex        =   231
               Text            =   "1"
               Top             =   435
               Width           =   375
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Dividir preço do produto em:          vez(es) sem juro(s)."
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   232
               Top             =   480
               Width           =   3945
            End
         End
         Begin TabDlg.SSTab tabComissaoIndices 
            Height          =   3915
            Left            =   7830
            TabIndex        =   40
            Top             =   420
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   6906
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
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
            TabCaption(0)   =   "Comi&ssão"
            TabPicture(0)   =   "Produtos.frx":4ECA1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "L_Comiss1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "L_Comiss2"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Comissão_Sobre"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Comissão"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Ín&dices"
            TabPicture(1)   =   "Produtos.frx":4ECBD
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lblTitle(10)"
            Tab(1).Control(1)=   "lblTitle(11)"
            Tab(1).Control(2)=   "lblTitle(13)"
            Tab(1).Control(3)=   "mskIndiceIcmsRetido"
            Tab(1).Control(4)=   "mskIndice"
            Tab(1).Control(5)=   "fraIndicePrecoEntrada"
            Tab(1).ControlCount=   6
            Begin VB.Frame fraIndicePrecoEntrada 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   810
               Left            =   -74910
               TabIndex        =   227
               Top             =   1590
               Width           =   2955
               Begin MSMask.MaskEdBox mskIndicePrecoEntrada 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   44
                  Top             =   270
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  Appearance      =   0
                  MaxLength       =   7
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   "0.00000"
                  PromptChar      =   " "
               End
               Begin VB.Label lblTitle 
                  Caption         =   "Não incide IPI no valor calculado pelo índice"
                  Height          =   435
                  Index           =   14
                  Left            =   1230
                  TabIndex        =   229
                  Top             =   240
                  Width           =   1650
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblTitle 
                  AutoSize        =   -1  'True
                  Caption         =   "Cálculo do Preço de Entrada"
                  Height          =   195
                  Index           =   12
                  Left            =   30
                  TabIndex        =   228
                  Top             =   60
                  Width           =   2025
               End
            End
            Begin VB.TextBox Comissão 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   330
               Left            =   1890
               MaxLength       =   6
               TabIndex        =   42
               Top             =   1080
               Width           =   735
            End
            Begin VB.CheckBox Comissão_Sobre 
               Appearance      =   0  'Flat
               Caption         =   "Comissão do produto sobrepõe comissão do vendedor"
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   120
               TabIndex        =   41
               Top             =   600
               Width           =   2625
            End
            Begin MSMask.MaskEdBox mskIndice 
               Height          =   315
               Left            =   -74850
               TabIndex        =   43
               Top             =   1080
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.000"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskIndiceIcmsRetido 
               Height          =   315
               Left            =   -74880
               TabIndex        =   45
               Top             =   2760
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               MaxLength       =   7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "0.00000"
               PromptChar      =   " "
            End
            Begin VB.Label L_Comiss2 
               Caption         =   "%"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2730
               TabIndex        =   226
               Top             =   1110
               Width           =   255
            End
            Begin VB.Label L_Comiss1 
               Caption         =   "Comissão do produto"
               Enabled         =   0   'False
               Height          =   225
               Left            =   120
               TabIndex        =   225
               Top             =   1140
               Width           =   1710
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "ICMS Retido - Base complementar"
               Height          =   195
               Index           =   13
               Left            =   -74880
               TabIndex        =   224
               Top             =   2490
               Visible         =   0   'False
               Width           =   2430
            End
            Begin VB.Label lblTitle 
               Caption         =   "Para aumentar os preços 10% use 1,10 e para diminuir 10% use 0,9"
               Height          =   465
               Index           =   11
               Left            =   -74850
               TabIndex        =   223
               Top             =   390
               Width           =   2955
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Financeiro"
               Height          =   195
               Index           =   10
               Left            =   -74850
               TabIndex        =   222
               Top             =   870
               Width           =   735
            End
         End
         Begin VB.Frame fraIdentificacaoCusto 
            Caption         =   "Identificação de Custo do Produto"
            Height          =   2295
            Left            =   -74895
            TabIndex        =   216
            Top             =   1590
            Width           =   10875
            Begin VB.CommandButton cmdCustoTecido 
               Caption         =   "..."
               Default         =   -1  'True
               Height          =   315
               Left            =   3285
               TabIndex        =   119
               ToolTipText     =   "Ao teclar {ENTER} o Custo do Tecido será calculado através da fórmula (Consumo de Tecido * Preço do metro linear)"
               Top             =   720
               Width           =   375
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   0
               Left            =   2040
               TabIndex        =   117
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin VB.CommandButton cmdCalcularCusto 
               BackColor       =   &H00C0FFFF&
               Caption         =   "&Calcular"
               Height          =   375
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   127
               ToolTipText     =   "Click aqui para atualizar o 'CUSTO' do Produto"
               Top             =   1800
               Width           =   1215
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   1
               Left            =   2040
               TabIndex        =   118
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   2
               Left            =   2040
               TabIndex        =   120
               Top             =   1080
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   3
               Left            =   2040
               TabIndex        =   121
               Top             =   1440
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   4
               Left            =   2040
               TabIndex        =   122
               Top             =   1800
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   5
               Left            =   5040
               TabIndex        =   123
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   6
               Left            =   5040
               TabIndex        =   124
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   7
               Left            =   5040
               TabIndex        =   125
               Top             =   1080
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCustoProduto 
               Height          =   315
               Index           =   8
               Left            =   5040
               TabIndex        =   126
               Top             =   1440
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label lblCustoProduto 
               AutoSize        =   -1  'True
               Caption         =   "Lavanderia"
               Height          =   195
               Index           =   4
               Left            =   360
               TabIndex        =   221
               Top             =   1860
               Width           =   795
            End
            Begin VB.Label lblCustoProduto 
               AutoSize        =   -1  'True
               Caption         =   "Mão de obra facção                                             Outros custos"
               Height          =   195
               Index           =   3
               Left            =   360
               TabIndex        =   220
               Top             =   1500
               Width           =   4455
            End
            Begin VB.Label lblCustoProduto 
               AutoSize        =   -1  'True
               Caption         =   "Custo do Tecido                                                    Aviamentos"
               Height          =   195
               Index           =   2
               Left            =   360
               TabIndex        =   219
               Top             =   1140
               Width           =   4335
            End
            Begin VB.Label lblCustoProduto 
               AutoSize        =   -1  'True
               Caption         =   "Preço do metro linear                                           Estamparia"
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   218
               Top             =   780
               Width           =   4260
            End
            Begin VB.Label lblCustoProduto 
               AutoSize        =   -1  'True
               Caption         =   "Consumo de Tecido                                              Bordado"
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   217
               Top             =   420
               Width           =   4065
            End
         End
         Begin VB.Frame fraEtiq 
            Caption         =   "Impressão de etiquetas na impressora térmica"
            Height          =   1155
            Left            =   -67785
            TabIndex        =   215
            Top             =   390
            Visible         =   0   'False
            Width           =   3765
            Begin VB.CheckBox chkImprimirPrecoEtiq 
               Appearance      =   0  'Flat
               Caption         =   "Imprimir o preço do produto na etiqueta"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   240
               TabIndex        =   116
               Top             =   720
               Width           =   3450
            End
            Begin VB.CheckBox chkImprimirUmaEtiq 
               Appearance      =   0  'Flat
               Caption         =   "Imprimir etiquetas simples"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   360
               Width           =   3135
            End
         End
         Begin MSMask.MaskEdBox mskDataValidade 
            Height          =   285
            Left            =   -73530
            TabIndex        =   114
            ToolTipText     =   "Pressione F2 para obter o calendário"
            Top             =   1185
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.TextBox txtLote 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   -73530
            MaxLength       =   15
            TabIndex        =   113
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtVolumagem 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   -73530
            MaxLength       =   4
            TabIndex        =   112
            Top             =   480
            Width           =   1200
         End
         Begin VB.Frame Frame9 
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
            Height          =   3615
            Left            =   -74880
            TabIndex        =   200
            Top             =   360
            Width           =   10815
            Begin VB.Frame fraPromocao 
               Caption         =   "Promoção"
               Height          =   1215
               Left            =   3960
               TabIndex        =   203
               Top             =   720
               Width           =   6615
               Begin MSComCtl2.DTPicker dtpDataInicial 
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   108
                  Top             =   720
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DateIsNull      =   -1  'True
                  Format          =   65404929
                  CurrentDate     =   37333
               End
               Begin MSComCtl2.DTPicker dtpDataFinal 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   109
                  Top             =   720
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DateIsNull      =   -1  'True
                  Format          =   65404929
                  CurrentDate     =   37333
               End
               Begin SSDataWidgets_B.SSDBCombo cboTablePromocional 
                  Bindings        =   "Produtos.frx":4ECD9
                  DataSource      =   "Data7"
                  Height          =   315
                  Left            =   240
                  TabIndex        =   107
                  Top             =   720
                  Width           =   1815
                  DataFieldList   =   "Tabela"
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
                  BackColorOdd    =   14737632
                  Columns(0).Width=   3200
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   93
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
               End
               Begin VB.Label lblTitle 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tabela de Preços"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   205
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "até"
                  Height          =   195
                  Index           =   1
                  Left            =   4440
                  TabIndex        =   204
                  Top             =   780
                  Width           =   240
               End
            End
            Begin VB.Frame fraAtributos 
               Caption         =   "Atributos"
               Height          =   1095
               Left            =   360
               TabIndex        =   104
               Top             =   2160
               Width           =   2655
               Begin VB.CheckBox chkPesquisa123 
                  Appearance      =   0  'Flat
                  Caption         =   "Pesquisa 123"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   240
                  TabIndex        =   106
                  Top             =   720
                  Width           =   1695
               End
               Begin VB.CheckBox chkFabricante 
                  Appearance      =   0  'Flat
                  Caption         =   "Fabricante"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   240
                  TabIndex        =   105
                  Top             =   360
                  Width           =   1695
               End
            End
            Begin VB.Frame fraTabPrecos 
               Caption         =   "Preço de Venda"
               Height          =   1215
               Left            =   360
               TabIndex        =   202
               Top             =   720
               Width           =   2655
               Begin SSDataWidgets_B.SSDBCombo cboTableVenda 
                  Bindings        =   "Produtos.frx":4ECED
                  DataSource      =   "Data7"
                  Height          =   315
                  Left            =   240
                  TabIndex        =   103
                  Top             =   600
                  Width           =   2175
                  DataFieldList   =   "Tabela"
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
                  BackColorOdd    =   14737632
                  Columns(0).Width=   3200
                  _ExtentX        =   3836
                  _ExtentY        =   556
                  _StockProps     =   93
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
               End
            End
            Begin VB.Frame fraBonus 
               Caption         =   "Bônus"
               Height          =   1095
               Left            =   3960
               TabIndex        =   201
               Top             =   2160
               Visible         =   0   'False
               Width           =   2055
               Begin VB.TextBox txtBonus 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   315
                  Left            =   120
                  MaxLength       =   6
                  TabIndex        =   110
                  Text            =   "0"
                  Top             =   480
                  Width           =   1815
               End
            End
            Begin VB.CommandButton cmdUpdateBatch 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Atualização em grupo"
               Height          =   420
               Left            =   8160
               Style           =   1  'Graphical
               TabIndex        =   111
               Top             =   2880
               Width           =   2415
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estas configurações serão utilizadas para o produto na Loja Virtual"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   206
               Top             =   240
               Width           =   4785
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Caminho e Nome do Arquivo Imagem  "
            Height          =   3645
            Left            =   -74880
            TabIndex        =   197
            Top             =   390
            Width           =   10785
            Begin VB.TextBox Nome_Foto 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   105
               TabIndex        =   101
               Top             =   285
               Width           =   5775
            End
            Begin VB.CommandButton B_Acha_Foto 
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
               Left            =   6030
               Picture         =   "Produtos.frx":4ED01
               Style           =   1  'Graphical
               TabIndex        =   102
               Top             =   240
               Width           =   435
            End
            Begin MSComDlg.CommonDialog cdgOpen 
               Left            =   6000
               Top             =   720
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
               Filter          =   "Imagem (*.bmp;*.wmf,*.jpg, *.gif)|*.bmp;*.wmf;*.jpg;*.gif"
            End
            Begin VB.Label Label22 
               Caption         =   "Formatos preferíveis: JPG, BMP ou WMF   Tamanho preferível: 270 X 220 pixels (Larg X Alt)"
               Height          =   480
               Left            =   2520
               TabIndex        =   199
               Top             =   2550
               Width           =   3675
            End
            Begin VB.Label Label50 
               Caption         =   $"Produtos.frx":4EE03
               Height          =   915
               Left            =   360
               TabIndex        =   198
               Top             =   1200
               Width           =   5535
            End
            Begin VB.Image imgFoto 
               BorderStyle     =   1  'Fixed Single
               Height          =   3360
               Left            =   6600
               Stretch         =   -1  'True
               Top             =   165
               Width           =   4110
            End
         End
         Begin VB.Frame Frame12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3915
            Left            =   -74880
            TabIndex        =   171
            Top             =   390
            Width           =   10830
            Begin MSMask.MaskEdBox C_Compra_Outros_V 
               Height          =   285
               Left            =   3480
               TabIndex        =   70
               Top             =   2670
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
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
            Begin MSMask.MaskEdBox C_Compra_Finan_V 
               Height          =   285
               Left            =   3480
               TabIndex        =   66
               Top             =   2340
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
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
            Begin MSMask.MaskEdBox C_Compra_Frete_V 
               Height          =   285
               Left            =   3480
               TabIndex        =   54
               Top             =   1410
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
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
            Begin VB.Frame Frame6 
               Caption         =   "Opção para o Cálculo"
               Height          =   1665
               Left            =   8640
               TabIndex        =   92
               Top             =   180
               Width           =   2055
               Begin VB.OptionButton O_C_Venda 
                  Appearance      =   0  'Flat
                  Caption         =   "Valor de V&enda"
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   75
                  TabIndex        =   93
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   1770
               End
               Begin VB.OptionButton O_C_Lucro_Val 
                  Appearance      =   0  'Flat
                  Caption         =   "Lucro em &Valores"
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   75
                  TabIndex        =   94
                  Top             =   615
                  Width           =   1695
               End
               Begin VB.OptionButton O_C_Lucro_Perc 
                  Appearance      =   0  'Flat
                  Caption         =   "Lucro em &Percentual"
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   75
                  TabIndex        =   95
                  Top             =   855
                  Width           =   1815
               End
               Begin VB.CommandButton cmdCalcular 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Calcular"
                  Height          =   420
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   96
                  Top             =   1155
                  Width           =   1815
               End
               Begin VB.Label Label21 
                  Caption         =   "Manter:"
                  Height          =   240
                  Left            =   555
                  TabIndex        =   173
                  Top             =   210
                  Width           =   975
               End
            End
            Begin VB.Frame fraDemais 
               Caption         =   "Para Demais Rotinas"
               Height          =   1110
               Left            =   8640
               TabIndex        =   97
               Top             =   1860
               Width           =   2085
               Begin VB.OptionButton O_Manter_Lucro_Perc 
                  Appearance      =   0  'Flat
                  Caption         =   "Lucro em &Percentual"
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   75
                  TabIndex        =   100
                  Top             =   825
                  Width           =   1815
               End
               Begin VB.OptionButton O_Manter_Lucro_Val 
                  Appearance      =   0  'Flat
                  Caption         =   "Lucro em &Valores"
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   75
                  TabIndex        =   99
                  Top             =   600
                  Width           =   1800
               End
               Begin VB.OptionButton O_Manter_Venda 
                  Appearance      =   0  'Flat
                  Caption         =   "Valor de V&enda"
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   75
                  TabIndex        =   98
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   1665
               End
               Begin VB.Label Label3 
                  Caption         =   "Manter:"
                  Height          =   240
                  Left            =   600
                  TabIndex        =   172
                  Top             =   195
                  Width           =   975
               End
            End
            Begin VB.CheckBox Fixa_ICM_V_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   7275
               TabIndex        =   75
               Top             =   705
               Width           =   210
            End
            Begin VB.CheckBox Fixa_ICM_V_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6960
               TabIndex        =   74
               Top             =   705
               Value           =   1  'Checked
               Width           =   225
            End
            Begin VB.CheckBox Fixa_IPI_V_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   7275
               TabIndex        =   79
               Top             =   1035
               Width           =   210
            End
            Begin VB.CheckBox Fixa_IPI_V_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6960
               TabIndex        =   78
               Top             =   1035
               Value           =   1  'Checked
               Width           =   225
            End
            Begin VB.CheckBox Fixa_Imp_V_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   7275
               TabIndex        =   83
               Top             =   1365
               Width           =   210
            End
            Begin VB.CheckBox Fixa_Imp_V_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6960
               TabIndex        =   82
               Top             =   1365
               Value           =   1  'Checked
               Width           =   225
            End
            Begin VB.CheckBox Fixa_Outros_V_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   7275
               TabIndex        =   87
               Top             =   1695
               Width           =   210
            End
            Begin VB.CheckBox Fixa_Outros_V_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6960
               TabIndex        =   86
               Top             =   1695
               Value           =   1  'Checked
               Width           =   225
            End
            Begin VB.CheckBox Fixa_Desc_C_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3225
               TabIndex        =   49
               Top             =   1095
               Width           =   210
            End
            Begin VB.CheckBox Fixa_Desc_C_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2940
               TabIndex        =   48
               Top             =   1095
               Value           =   1  'Checked
               Width           =   225
            End
            Begin VB.CheckBox Fixa_Frete_C_Valor 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3225
               TabIndex        =   53
               Top             =   1425
               Value           =   1  'Checked
               Width           =   210
            End
            Begin VB.CheckBox Fixa_Frete_C_Perc 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2940
               TabIndex        =   52
               Top             =   1425
               Width           =   225
            End
            Begin VB.CheckBox Fixa_ICM_C_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3225
               TabIndex        =   57
               Top             =   1755
               Width           =   210
            End
            Begin VB.CheckBox Fixa_ICM_C_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2940
               TabIndex        =   56
               Top             =   1755
               Value           =   1  'Checked
               Width           =   225
            End
            Begin VB.CheckBox Fixa_IPI_C_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3225
               TabIndex        =   61
               Top             =   2055
               Width           =   210
            End
            Begin VB.CheckBox Fixa_IPI_C_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2940
               TabIndex        =   60
               Top             =   2055
               Value           =   1  'Checked
               Width           =   225
            End
            Begin VB.CheckBox Fixa_Custo_C_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3225
               TabIndex        =   65
               Top             =   2385
               Width           =   210
            End
            Begin VB.CheckBox Fixa_Custo_C_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2940
               TabIndex        =   64
               Top             =   2385
               Value           =   1  'Checked
               Width           =   225
            End
            Begin VB.CheckBox Fixa_Outros_C_Valor 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3225
               TabIndex        =   69
               Top             =   2715
               Width           =   210
            End
            Begin VB.CheckBox Fixa_Outros_C_Perc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2940
               TabIndex        =   68
               Top             =   2715
               Value           =   1  'Checked
               Width           =   225
            End
            Begin MSMask.MaskEdBox C_Compra_Sem_Nota 
               Height          =   285
               Left            =   2040
               TabIndex        =   71
               Top             =   3015
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Compra_Outros_P 
               Height          =   285
               Left            =   2040
               TabIndex        =   67
               Top             =   2670
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Compra_Finan_P 
               Height          =   285
               Left            =   2040
               TabIndex        =   63
               Top             =   2340
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Compra_IPI_P 
               Height          =   285
               Left            =   2040
               TabIndex        =   59
               Top             =   2010
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Compra_ICM_P 
               Height          =   285
               Left            =   2040
               TabIndex        =   55
               Top             =   1710
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Compra_Frete_P 
               Height          =   285
               Left            =   2040
               TabIndex        =   51
               Top             =   1380
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Compra_Desc_P 
               Height          =   285
               Left            =   2040
               TabIndex        =   47
               Top             =   1050
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Compra_Desc_V 
               Height          =   285
               Left            =   3480
               TabIndex        =   50
               Top             =   1050
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   12632319
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
            Begin MSMask.MaskEdBox C_Compra_Valor 
               Height          =   285
               Left            =   3480
               TabIndex        =   46
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   12648384
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
            Begin MSMask.MaskEdBox C_Venda_Outros_V 
               Height          =   285
               Left            =   7560
               TabIndex        =   88
               Top             =   1650
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_Outros_P 
               Height          =   285
               Left            =   6000
               TabIndex        =   85
               Top             =   1650
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_Imp_V 
               Height          =   285
               Left            =   7560
               TabIndex        =   84
               Top             =   1320
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_Imp_P 
               Height          =   285
               Left            =   6000
               TabIndex        =   81
               Top             =   1320
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_IPI_P 
               Height          =   285
               Left            =   6000
               TabIndex        =   77
               Top             =   990
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_ICM_P 
               Height          =   285
               Left            =   6000
               TabIndex        =   73
               Top             =   660
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_Valor 
               Height          =   330
               Left            =   6960
               TabIndex        =   72
               Top             =   3450
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   12648384
               Format          =   "###,###,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_Margem_V 
               Height          =   285
               Left            =   7575
               TabIndex        =   91
               Top             =   2340
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   12648384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_Margem_P 
               Height          =   285
               Left            =   6000
               TabIndex        =   90
               Top             =   2340
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   12648384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox C_Venda_Sem_Nota 
               Height          =   285
               Left            =   7575
               TabIndex        =   89
               Top             =   3015
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   16753444
               PromptChar      =   "_"
            End
            Begin VB.Label C_Compra_ICM_V 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   3480
               TabIndex        =   58
               Top             =   1710
               Width           =   975
            End
            Begin VB.Label C_Compra_IPI_V 
               Appearance      =   0  'Flat
               BackColor       =   &H00E5E5E5&
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   3480
               TabIndex        =   62
               Top             =   2010
               Width           =   975
            End
            Begin VB.Label lblCalculo 
               Caption         =   "Lucro"
               Height          =   225
               Index           =   15
               Left            =   4665
               TabIndex        =   196
               Top             =   2370
               Width           =   450
            End
            Begin VB.Label C_Compra_Custo 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   2865
               TabIndex        =   195
               Top             =   3450
               Width           =   1590
            End
            Begin VB.Label Label52 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "Percentual"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5970
               TabIndex        =   194
               Top             =   330
               Width           =   870
            End
            Begin VB.Label Label53 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "Valor"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7560
               TabIndex        =   193
               Top             =   330
               Width           =   975
            End
            Begin VB.Label C_Venda_IPI_V 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFA324&
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   7560
               TabIndex        =   80
               Top             =   990
               Width           =   975
            End
            Begin VB.Label C_Venda_ICM_V 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFA324&
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   7560
               TabIndex        =   76
               Top             =   660
               Width           =   975
            End
            Begin VB.Label Label42 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "Valor"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3480
               TabIndex        =   192
               Top             =   330
               Width           =   975
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Preço de Venda com IPI"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   9
               Left            =   4710
               TabIndex        =   191
               Top             =   3495
               Width           =   1755
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " ICM de Venda"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   10
               Left            =   4680
               TabIndex        =   190
               Top             =   705
               Width           =   1050
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " IPI de Venda"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   11
               Left            =   4680
               TabIndex        =   189
               Top             =   1035
               Width           =   975
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Outros Impostos"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   12
               Left            =   4680
               TabIndex        =   188
               Top             =   1365
               Width           =   1245
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Outros"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   13
               Left            =   4680
               TabIndex        =   187
               Top             =   1695
               Width           =   540
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Vendas SEM Impostos"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   14
               Left            =   4680
               TabIndex        =   186
               Top             =   3045
               Width           =   1620
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Preço do produto"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   185
               Top             =   765
               Width           =   1290
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Desconto Fornecedor"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   184
               Top             =   1095
               Width           =   1590
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Frete"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   183
               Top             =   1425
               Width           =   435
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " ICMS de Compra"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   182
               Top             =   1755
               Width           =   1245
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " IPI de Compra"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   181
               Top             =   2055
               Width           =   1080
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Custo Financeiro"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   180
               Top             =   2385
               Width           =   1245
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Compra SEM Impostos"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   179
               Top             =   3045
               Width           =   1650
            End
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "Percentual"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2025
               TabIndex        =   178
               Top             =   330
               Width           =   870
            End
            Begin VB.Label lblCalculo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   " Outros"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   177
               Top             =   2715
               Width           =   540
            End
            Begin VB.Label lblCalculo 
               Caption         =   "Preço de Custo Calculado"
               Height          =   255
               Index           =   8
               Left            =   105
               TabIndex        =   176
               Top             =   3510
               Width           =   1965
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Compra"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   345
               Left            =   165
               TabIndex        =   175
               Top             =   180
               Width           =   975
            End
            Begin VB.Line Line1 
               X1              =   60
               X2              =   10665
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Venda"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFA324&
               Height          =   345
               Left            =   4770
               TabIndex        =   174
               Top             =   180
               Width           =   795
            End
            Begin VB.Line Line2 
               BorderWidth     =   3
               X1              =   4560
               X2              =   4560
               Y1              =   195
               Y2              =   3720
            End
            Begin VB.Line Line3 
               X1              =   105
               X2              =   10710
               Y1              =   3345
               Y2              =   3345
            End
         End
         Begin VB.Data Data2_2 
            Caption         =   "Data4"
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
            Left            =   -66960
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT Nome, Código, [Fone 1] FROM Cli_For WHERE (Tipo = 'F' OR Tipo = 'R' OR Tipo = 'O') And Inativo = False ORDER BY Nome;"
            Top             =   2220
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.Data Data2_1 
            Caption         =   "Data4"
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
            Left            =   -68880
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Forn_Prod"
            Top             =   2220
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   -74880
            TabIndex        =   139
            Top             =   4140
            Width           =   10890
            Begin VB.Label Última_Compra 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   9120
               TabIndex        =   145
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Último_Fornecedor 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2160
               TabIndex        =   143
               Top             =   600
               Width           =   8535
            End
            Begin VB.Label Label28 
               Caption         =   "Último Fornecedor"
               Height          =   210
               Left            =   180
               TabIndex        =   142
               Top             =   600
               Width           =   1545
            End
            Begin VB.Label Último_Custo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2160
               TabIndex        =   141
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label26 
               Caption         =   $"Produtos.frx":4EF06
               Height          =   255
               Left            =   180
               TabIndex        =   140
               Top             =   285
               Width           =   8880
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Impostos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4035
            Left            =   90
            TabIndex        =   138
            Top             =   330
            Width           =   7665
            Begin VB.CommandButton cmd_tributosProduto 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Tributos do produto"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   243
               Top             =   450
               Width           =   7245
            End
            Begin VB.CommandButton cmd_vincularCFOPs_produto 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Vincular CFOPs ao Produto"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   237
               Top             =   1650
               Width           =   7245
            End
            Begin VB.CommandButton cmd_clonarCaract_produto 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Clonar Características de um produto"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   236
               Top             =   2910
               Width           =   7245
            End
            Begin VB.Data datGrupoFiscal 
               Caption         =   "datGrupoFiscal"
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
               Left            =   6450
               Options         =   0
               ReadOnly        =   -1  'True
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "SELECT Código, Nome FROM GrupoFiscal ORDER BY Nome"
               Top             =   810
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Data datCodigoNBM 
               Caption         =   "datCodigoNBM"
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
               Left            =   6480
               Options         =   0
               ReadOnly        =   -1  'True
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "SELECT Codigo, Nome FROM AliquotasNCM ORDER BY Codigo"
               Top             =   270
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Data Data4_1 
               Caption         =   "Data4"
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
               Left            =   6480
               Options         =   0
               ReadOnly        =   -1  'True
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "Con_C_Fiscal"
               Top             =   510
               Visible         =   0   'False
               Width           =   1140
            End
         End
         Begin VB.Frame Frame8 
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
            Height          =   4320
            Left            =   -74925
            TabIndex        =   137
            Top             =   360
            Width           =   10935
            Begin VB.CommandButton cmd_cesta_de_produtos 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Cesta de Produtos"
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
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   235
               Top             =   3810
               Width           =   10845
            End
            Begin VB.TextBox txtCubagem 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   4845
               MaxLength       =   6
               TabIndex        =   21
               Top             =   2280
               Width           =   855
            End
            Begin VB.TextBox txtPesoBruto 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   4845
               MaxLength       =   12
               TabIndex        =   19
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox txtPesoLiquido 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   4845
               MaxLength       =   12
               TabIndex        =   20
               Top             =   1920
               Width           =   855
            End
            Begin VB.CheckBox chkUsarDescrAdic 
               Appearance      =   0  'Flat
               Caption         =   "Usar Descr. Adicional"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3690
               TabIndex        =   28
               Top             =   3450
               Width           =   2010
            End
            Begin VB.CheckBox chkDontAllowDesc 
               Appearance      =   0  'Flat
               Caption         =   "Não permitir desconto neste produto (VR)"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   27
               Top             =   3450
               Width           =   3630
            End
            Begin VB.TextBox Cod_Fornecedor 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   5370
               MaxLength       =   15
               TabIndex        =   25
               ToolTipText     =   "Caso o fornecedor use um código diferente, digite aqui este código."
               Top             =   2655
               Width           =   2430
            End
            Begin VB.CommandButton cmdImprimirEtiqueta 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Etiqueta"
               Height          =   375
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   207
               Top             =   90
               Width           =   1380
            End
            Begin VB.TextBox Fabricante 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   330
               Left            =   5370
               MaxLength       =   15
               TabIndex        =   38
               Top             =   3060
               Width           =   2445
            End
            Begin VB.TextBox txtCasasDecimais 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   315
               Left            =   6960
               MaxLength       =   8
               TabIndex        =   23
               ToolTipText     =   "Número de Casas Decimais para Quantidades"
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox Unid 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   6960
               MaxLength       =   5
               TabIndex        =   22
               ToolTipText     =   "Máximo 6 caracteres"
               Top             =   1560
               Width           =   855
            End
            Begin VB.CheckBox Não_Lista 
               Appearance      =   0  'Flat
               Caption         =   "Não incluir nas listas de preços"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   26
               Top             =   3105
               Width           =   3255
            End
            Begin VB.Frame Frame14 
               Caption         =   "Estoque"
               BeginProperty Font 
                  Name            =   "WeblySleek UI Semibold"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1335
               Left            =   7890
               TabIndex        =   161
               Top             =   2040
               Width           =   3045
               Begin VB.TextBox Locali 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   315
                  Left            =   1035
                  MaxLength       =   15
                  TabIndex        =   37
                  Top             =   960
                  Width           =   1935
               End
               Begin VB.TextBox Estoque_Ideal 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   315
                  Left            =   1560
                  MaxLength       =   8
                  TabIndex        =   36
                  Top             =   600
                  Width           =   1410
               End
               Begin VB.TextBox Estoque_Min 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   315
                  Left            =   1560
                  MaxLength       =   8
                  TabIndex        =   35
                  Top             =   240
                  Width           =   1410
               End
               Begin VB.CheckBox Controla_Estoque 
                  Appearance      =   0  'Flat
                  Caption         =   "Controlar estoque"
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
                  Height          =   255
                  Left            =   120
                  TabIndex        =   34
                  Top             =   0
                  Value           =   1  'Checked
                  Width           =   1905
               End
               Begin VB.Label Label19 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Local"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   164
                  Top             =   990
                  Width           =   570
               End
               Begin VB.Label Label16 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Estoque Mínimo"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   163
                  Top             =   270
                  Width           =   1245
               End
               Begin VB.Label Label17 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Estoque Ideal"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   162
                  Top             =   630
                  Width           =   1155
               End
            End
            Begin VB.TextBox Nome_Nota 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   1440
               MaxLength       =   80
               TabIndex        =   15
               Top             =   1080
               Width           =   5360
            End
            Begin VB.Frame Frame13 
               Caption         =   "Pesquisa"
               BeginProperty Font 
                  Name            =   "WeblySleek UI Semibold"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1560
               Left            =   120
               TabIndex        =   152
               Top             =   1440
               Width           =   3255
               Begin SSDataWidgets_B.SSDBCombo Combo_Pesq1 
                  Bindings        =   "Produtos.frx":4EFB4
                  DataSource      =   "Data4"
                  Height          =   315
                  Left            =   960
                  TabIndex        =   16
                  Top             =   240
                  Width           =   720
                  DataFieldList   =   "Nome"
                  _Version        =   196617
                  RowHeight       =   423
                  Columns.Count   =   2
                  Columns(0).Width=   7699
                  Columns(0).Caption=   "Nome"
                  Columns(0).Name =   "Nome"
                  Columns(0).CaptionAlignment=   0
                  Columns(0).DataField=   "Nome"
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   256
                  Columns(1).Width=   1535
                  Columns(1).Caption=   "Código"
                  Columns(1).Name =   "Código"
                  Columns(1).Alignment=   1
                  Columns(1).CaptionAlignment=   1
                  Columns(1).DataField=   "Código"
                  Columns(1).DataType=   3
                  Columns(1).FieldLen=   256
                  _ExtentX        =   1270
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   12648447
               End
               Begin SSDataWidgets_B.SSDBCombo Combo_Pesq2 
                  Bindings        =   "Produtos.frx":4EFC8
                  DataSource      =   "Data5"
                  Height          =   315
                  Left            =   960
                  TabIndex        =   17
                  Top             =   660
                  Width           =   720
                  DataFieldList   =   "Nome"
                  _Version        =   196617
                  RowHeight       =   423
                  Columns.Count   =   2
                  Columns(0).Width=   8811
                  Columns(0).Caption=   "Nome"
                  Columns(0).Name =   "Nome"
                  Columns(0).CaptionAlignment=   0
                  Columns(0).DataField=   "Nome"
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   256
                  Columns(1).Width=   1429
                  Columns(1).Caption=   "Código"
                  Columns(1).Name =   "Código"
                  Columns(1).Alignment=   1
                  Columns(1).CaptionAlignment=   1
                  Columns(1).DataField=   "Código"
                  Columns(1).DataType=   3
                  Columns(1).FieldLen=   256
                  _ExtentX        =   1270
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   12648447
               End
               Begin SSDataWidgets_B.SSDBCombo Combo_Pesq3 
                  Bindings        =   "Produtos.frx":4EFDC
                  DataSource      =   "Data6"
                  Height          =   315
                  Left            =   960
                  TabIndex        =   18
                  Top             =   1080
                  Width           =   720
                  DataFieldList   =   "Nome"
                  _Version        =   196617
                  RowHeight       =   423
                  Columns.Count   =   2
                  Columns(0).Width=   8334
                  Columns(0).Caption=   "Nome"
                  Columns(0).Name =   "Nome"
                  Columns(0).CaptionAlignment=   0
                  Columns(0).DataField=   "Nome"
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   256
                  Columns(1).Width=   2064
                  Columns(1).Caption=   "Código"
                  Columns(1).Name =   "Código"
                  Columns(1).Alignment=   1
                  Columns(1).CaptionAlignment=   1
                  Columns(1).DataField=   "Código"
                  Columns(1).DataType=   3
                  Columns(1).FieldLen=   256
                  _ExtentX        =   1270
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   12648447
               End
               Begin VB.Label Label_Pesq3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "3"
                  Height          =   210
                  Left            =   120
                  TabIndex        =   158
                  Top             =   1132
                  Width           =   840
               End
               Begin VB.Label Nome_Pesq3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   157
                  Top             =   1080
                  Width           =   1455
               End
               Begin VB.Label Label_Pesq2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "2"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   156
                  Top             =   720
                  Width           =   870
               End
               Begin VB.Label Nome_Pesq2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   155
                  Top             =   660
                  Width           =   1455
               End
               Begin VB.Label Label_Pesq1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "1"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   154
                  Top             =   285
                  Width           =   870
               End
               Begin VB.Label Nome_Pesq1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   153
                  Top             =   240
                  Width           =   1455
               End
            End
            Begin VB.Frame Preços 
               Caption         =   "Preços"
               Height          =   1900
               Left            =   7890
               TabIndex        =   147
               Top             =   0
               Width           =   3045
               Begin VB.TextBox Custo_Médio 
                  Appearance      =   0  'Flat
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
                  Left            =   1650
                  TabIndex        =   33
                  Top             =   1540
                  Width           =   1335
               End
               Begin MSMask.MaskEdBox medPreco 
                  Height          =   285
                  Index           =   1
                  Left            =   1650
                  TabIndex        =   30
                  Top             =   555
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  BorderStyle     =   0
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox medPreco 
                  Height          =   285
                  Index           =   0
                  Left            =   1650
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  BorderStyle     =   0
                  Appearance      =   0
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox medPreco 
                  Height          =   285
                  Index           =   2
                  Left            =   1650
                  TabIndex        =   31
                  Top             =   885
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  BorderStyle     =   0
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox medPreco 
                  Height          =   285
                  Index           =   3
                  Left            =   1650
                  TabIndex        =   32
                  Top             =   1215
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  BorderStyle     =   0
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PromptChar      =   "_"
               End
               Begin VB.Label Label27 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Custo Médio"
                  Height          =   255
                  Left            =   90
                  TabIndex        =   170
                  Top             =   1560
                  Width           =   1545
               End
               Begin VB.Label lblNomePreco 
                  BackStyle       =   0  'Transparent
                  Height          =   255
                  Index           =   3
                  Left            =   90
                  TabIndex        =   151
                  Top             =   1230
                  Width           =   1545
               End
               Begin VB.Label lblNomePreco 
                  BackStyle       =   0  'Transparent
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  TabIndex        =   150
                  Top             =   900
                  Width           =   1545
               End
               Begin VB.Label lblNomePreco 
                  BackStyle       =   0  'Transparent
                  Height          =   255
                  Index           =   1
                  Left            =   90
                  TabIndex        =   149
                  Top             =   570
                  Width           =   1545
               End
               Begin VB.Label lblNomePreco 
                  BackStyle       =   0  'Transparent
                  Height          =   255
                  Index           =   0
                  Left            =   90
                  TabIndex        =   148
                  Top             =   240
                  Width           =   1545
               End
            End
            Begin VB.TextBox Obs 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Top             =   90
               Width           =   6375
            End
            Begin MSMask.MaskEdBox Desconto 
               Height          =   315
               Left            =   6960
               TabIndex        =   24
               Top             =   2280
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cubagem"
               Height          =   195
               Left            =   3480
               TabIndex        =   212
               Top             =   2340
               Width           =   675
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso Bruto (Kg)"
               Height          =   195
               Left            =   3480
               TabIndex        =   211
               Top             =   1620
               Width           =   1125
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso Líquido (Kg)"
               Height          =   195
               Left            =   3480
               TabIndex        =   210
               Top             =   1980
               Width           =   1230
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Código p/ Fornecedor"
               Height          =   195
               Left            =   3600
               TabIndex        =   208
               Top             =   2715
               Width           =   1560
            End
            Begin VB.Shape Shape1 
               Height          =   360
               Left            =   3480
               Top             =   2640
               Width           =   4335
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Fabricante/Marca"
               Height          =   255
               Left            =   3480
               TabIndex        =   168
               Top             =   3090
               Width           =   1485
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto (%)"
               Height          =   195
               Left            =   5760
               TabIndex        =   167
               Top             =   2340
               Width           =   1005
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Casas Decimais"
               Height          =   195
               Left            =   5760
               TabIndex        =   166
               Top             =   1980
               Width           =   1095
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unid. de Venda"
               Height          =   195
               Left            =   5760
               TabIndex        =   165
               Top             =   1620
               Width           =   1095
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Nome para Nota"
               Height          =   255
               Left            =   120
               TabIndex        =   160
               Top             =   1110
               Width           =   1335
            End
            Begin VB.Label Qtde_Carac 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   6840
               TabIndex        =   159
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição       Inf. Adicional"
               Height          =   390
               Left            =   120
               TabIndex        =   146
               Top             =   570
               Width           =   1260
            End
         End
         Begin SSDataWidgets_B.SSDBDropDown DropDown2_1 
            Bindings        =   "Produtos.frx":4EFF0
            Height          =   1065
            Left            =   -74070
            TabIndex        =   169
            Top             =   930
            Width           =   7770
            DataFieldList   =   "Nome"
            ListAutoValidate=   0   'False
            _Version        =   196617
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorOdd    =   12648447
            RowHeight       =   423
            ExtraHeight     =   53
            Columns.Count   =   3
            Columns(0).Width=   9260
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Nome"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   2090
            Columns(1).Caption=   "Código"
            Columns(1).Name =   "Código"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Código"
            Columns(1).DataType=   3
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Caption=   "Fone 1"
            Columns(2).Name =   "Fone 1"
            Columns(2).CaptionAlignment=   0
            Columns(2).DataField=   "Fone 1"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            _ExtentX        =   13705
            _ExtentY        =   1879
            _StockProps     =   77
         End
         Begin SSDataWidgets_B.SSDBGrid Grade2_1 
            Bindings        =   "Produtos.frx":4F006
            Height          =   2610
            Left            =   -74880
            TabIndex        =   39
            Top             =   420
            Width           =   10875
            _Version        =   196617
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            SelectTypeRow   =   1
            BackColorOdd    =   12648447
            RowHeight       =   423
            ExtraHeight     =   53
            Columns(0).Width=   3200
            _ExtentX        =   19182
            _ExtentY        =   4604
            _StockProps     =   79
            Caption         =   "Fornecedores"
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
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Data de Validade"
            Height          =   195
            Index           =   5
            Left            =   -74895
            TabIndex        =   214
            Top             =   1230
            Width           =   1215
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   195
            Index           =   4
            Left            =   -74895
            TabIndex        =   213
            Top             =   885
            Width           =   315
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Volumagem"
            Height          =   195
            Index           =   1
            Left            =   -74895
            TabIndex        =   209
            Top             =   525
            Width           =   810
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Produto"
         Height          =   1335
         Left            =   8160
         TabIndex        =   5
         Top             =   120
         Width           =   3015
         Begin VB.CheckBox Desativado 
            Appearance      =   0  'Flat
            Caption         =   "Produto Inativo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   960
            Width           =   1545
         End
         Begin VB.CheckBox chkWebIncluded 
            Appearance      =   0  'Flat
            Caption         =   "Produto Web"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   75
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox Fracionado 
            Appearance      =   0  'Flat
            Caption         =   "Qtde Fracionada"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1395
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton O_Produto_Normal 
            Appearance      =   0  'Flat
            Caption         =   "&Normal"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   60
            TabIndex        =   6
            Top             =   225
            Value           =   -1  'True
            Width           =   990
         End
         Begin VB.OptionButton O_Produto_Grade 
            Appearance      =   0  'Flat
            Caption         =   "&Grade (roupas, etc..)"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   60
            TabIndex        =   8
            Top             =   465
            Width           =   1995
         End
         Begin VB.OptionButton O_Produto_Edição 
            Appearance      =   0  'Flat
            Caption         =   "&Edição (revistas, etc..)"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   60
            TabIndex        =   9
            Top             =   705
            Width           =   2085
         End
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Moeda 
         Bindings        =   "Produtos.frx":4F01C
         DataSource      =   "Data3"
         Height          =   315
         Left            =   4875
         TabIndex        =   1
         Top             =   240
         Width           =   870
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
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3810
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
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
         Bindings        =   "Produtos.frx":4F030
         DataSource      =   "Data1"
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   1125
         Width           =   945
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
         GroupHeaders    =   0   'False
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         RowHeight       =   476
         Columns.Count   =   2
         Columns(0).Width=   5450
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3545
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1667
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Sub_Classe 
         Bindings        =   "Produtos.frx":4F044
         DataSource      =   "Data2"
         Height          =   315
         Left            =   4875
         TabIndex        =   4
         Top             =   1125
         Width           =   870
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
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         Columns(0).Width=   3200
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBCombo cboNome 
         Bindings        =   "Produtos.frx":4F058
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   690
         Width           =   7335
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
         UseExactRowCount=   0   'False
         BackColorEven   =   15066597
         BackColorOdd    =   12648447
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   9525
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3175
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   12938
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label lblMotivoDesoneracao 
         Alignment       =   2  'Center
         Caption         =   "Motivo Desoneração"
         Height          =   255
         Left            =   450
         TabIndex        =   233
         Top             =   6870
         Width           =   1695
      End
      Begin VB.Label Ordenação 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   780
         TabIndex        =   144
         Top             =   315
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Nome_Sub_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5760
         TabIndex        =   136
         Top             =   1125
         Width           =   2295
      End
      Begin VB.Label Nome_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1680
         TabIndex        =   135
         Top             =   1125
         Width           =   2415
      End
      Begin VB.Label Nome_Moeda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5760
         TabIndex        =   134
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   1150
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Subclasse"
         Height          =   255
         Left            =   4110
         TabIndex        =   130
         Top             =   1155
         Width           =   825
      End
      Begin VB.Label Label14 
         Caption         =   "Moeda"
         Height          =   255
         Left            =   4260
         TabIndex        =   129
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Left            =   -90
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   7290
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   2130
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Sub_Classe"
      Top             =   7140
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq3"
      Top             =   7380
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   7530
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq2"
      Top             =   7260
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   3990
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Moeda"
      Top             =   7260
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Left            =   5700
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq1"
      Top             =   7260
      Visible         =   0   'False
      Width           =   1935
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   9855
      Top             =   7200
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
      Bands           =   "Produtos.frx":4F072
   End
End
Attribute VB_Name = "frmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'17/09/2002 - mpdea
'Comentado referências ao TextBox txtBonus da pasta WEB
'(aguardando definições para implementação futura)

'17/05/2005 - Daniel
'Na guia "Outros", adicionamos uma nova frame que estará visível
'apenas para a Miss Nuvem. Os campos desta frame gerenciaram a
'impressão de uma ou duas etiquetas na térmica e a impressão ou
'não do preço do produto

'04/07/2005 - Daniel
'Na guia "Outros", adicionamos uma nova frame que estará visível
'apenas para a Zue.
'Finalidade: Informações do Custo de Produto

Option Explicit

Dim Num_Registro As Variant
Dim rsProdutos2 As Recordset
Dim rsClasses As Recordset
Dim rsSub_Classes As Recordset
Dim rsTamanhos As Recordset
Dim rsMoedas As Recordset
Dim rsCliFor As Recordset
Dim rsForn_Prod As Recordset
Dim rsC_Fiscal As Recordset
Dim rsEstoque As Recordset
Dim rsResumo As Recordset
Dim rsParametros As Recordset
Dim rsTabelas As Recordset
Dim rsPrecos As Recordset
Dim rsFuncionarios As Recordset
Dim rsPesquisa1 As Recordset
Dim rsPesquisa2 As Recordset
Dim rsPesquisa3 As Recordset
Dim Rec_Fornec As Recordset
Dim Rec_Preços As Recordset
'Dim rsProdutos22 As Recordset
Dim bGravando As Boolean
Dim rsProdutosFiltrados As Recordset
Dim rsProdutosFiltradosNome As Recordset


Private gbComboDropped As Boolean

Private Type Grade_1
  Código As Integer
  Nome As String
End Type

Dim Grade_Tamanhos(49) As Grade_1
Dim Grade_Cores(49) As Grade_1
'04/05/2004 - Daniel
'Var de controle para manutenções feitas para a Embalavi
Dim m_blnEmbalavi As Boolean

'15/10/2007 - Anderson
'Implementação de CFOP por Produto
'15/12/2006 - Anderson - Alteração para o cadastro de CFOP por serviço
Dim Rec_CFOP As Recordset
Dim rsOperacaoSaida As Recordset
Dim rsProdutoCFOP As Recordset
Dim Sincroniza As Boolean

' Variaveis para tratamento do Componente WebApi
Dim rsCNPJ As Recordset
Dim http As MSXML2.XMLHTTP
'Dim sRETORNO_ERRO_WEBAPI As String
Dim alvoEnderecoWEBAPI As String
'Dim tentativasWEBAPI As Integer
'Dim sCodigoProdutoDELETE As String

'Private Sub Acha_CódigoVELHO(Optional Grade As Integer)
'
'  Dim Maior As Double
'  Dim Conta As Integer
'  Dim Prod As String
'
'  On Error GoTo Deu_Erro
'
'  Call StatusMsg("")
'  If Not IsNull(Num_Registro) Then
'    DisplayMsg "Não é possível mostrar o próximo código disponível com um produto ativo. Pressione o botão Limpa antes."
'    Exit Sub
'  End If
'
'  Call StatusMsg("")
'
'  Maior = 0
'  Prod = ""
'  Conta = 0
'  rsProdutos22.Index = "Código"
'
'Lp1:
'  rsProdutos22.Seek ">", Prod
'  If rsProdutos22.NoMatch Then GoTo Fim_Lp
'  Prod = rsProdutos22("Código")
'
'  Conta = Conta + 1
'  If Conta = 10 Then
'    Call StatusMsg("Aguarde, verificando " & Prod)
'    DoEvents
'    Conta = 0
'  End If
'
'  If Not IsNumeric(Prod) Then GoTo Lp1
'  If CDbl(Prod) > Maior Then Maior = CDbl(Prod)
'
'  GoTo Lp1
'
'Fim_Lp:
'  Maior = Maior + 1
'  cboCodigo.Text = Maior
'  Call StatusMsg("")
'  Exit Sub
'
'Deu_Erro:
'  DisplayMsg "Não foi possível achar o próximo código."
'
'End Sub

Private Sub FindNextProdutoCode()
'  Dim rsCode As Recordset
'  Dim sSQL As String
'
'  sSQL = "SELECT Last(Código) AS LastCode FROM Produtos WHERE Código LIKE ""#*"" ORDER BY Código"
'  Set rsCode = db.OpenRecordset(sSQL, dbOpenSnapshot)
'  Debug.Print sSQL
'  Debug.Print rsCode!LastCode
'  rsCode.Close
'  Set rsCode = Nothing
  

  Dim sSql As String
  Dim rsProd As Recordset
  Dim nCod As Double

  Call StatusMsg("")
  If Not IsNull(Num_Registro) Then
    DisplayMsg "Não é possível mostrar o próximo código disponível com um produto ativo. Pressione o botão Limpa antes."
    Exit Sub
  End If

  Call StatusMsg("Aguarde...")

  Set rsProd = rsProdutos.Clone
  If Not rsProd.EOF Then
    rsProd.MoveLast
    If IsNumeric(rsProd("Código").Value) Then
      nCod = CDbl(gsHandleNull(rsProd("Código").Value & ""))
    Else
      nCod = 0
    End If
  Else
    If rsProd.EOF And rsProd.BOF Then
      nCod = 0
    Else
      nCod = CDbl(gsHandleNull(rsProd("Código").Value & ""))
    End If
  End If

  cboCodigo.Text = nCod + 1
  rsProd.Close
  Set rsProd = Nothing

  Call StatusMsg("")

End Sub

Private Sub Calcula_Custo1()
 Dim Custo As Double
 Dim Fixa_Desc As String
 Dim Fixa_Frete As String
 Dim Fixa_ICM As String
 Dim Fixa_IPI As String
 Dim Fixa_Custo As String
 Dim Fixa_Outros As String
 
 Dim Compra_Desc_V As Double
 Dim Compra_Desc_P As Double
 Dim Compra_Valor As Double
 Dim Compra_Frete_V As Double
 Dim Compra_Frete_P As Double
 Dim Compra_ICM_V As Double
 Dim Compra_ICM_P As Double
 Dim Compra_IPI_V As Double
 Dim Compra_IPI_P As Double
 Dim Compra_Finan_V As Double
 Dim Compra_Finan_P As Double
 Dim Compra_Outros_V As Double
 Dim Compra_Outros_P As Double
 
 If Not IsNumeric(C_Compra_Valor.Text) Then C_Compra_Valor.Text = 0
 
 If CDbl(C_Compra_Valor.Text) = 0 Then Exit Sub
 
 If Not IsNumeric(C_Compra_Finan_V.Text) Then C_Compra_Finan_V.Text = 0
 If Not IsNumeric(C_Compra_Finan_P.Text) Then C_Compra_Finan_P.Text = 0
 If Not IsNumeric(C_Compra_Outros_V.Text) Then C_Compra_Outros_V.Text = 0
 If Not IsNumeric(C_Compra_Outros_P.Text) Then C_Compra_Outros_P.Text = 0
 If Not IsNumeric(C_Compra_Desc_P.Text) Then C_Compra_Desc_P.Text = 0
 If Not IsNumeric(C_Compra_Desc_V.Text) Then C_Compra_Desc_V.Text = 0
 If Not IsNumeric(C_Compra_Frete_P.Text) Then C_Compra_Frete_P.Text = 0
 If Not IsNumeric(C_Compra_Frete_V.Text) Then C_Compra_Frete_V.Text = 0
 If Not IsNumeric(C_Compra_ICM_P.Text) Then C_Compra_ICM_P.Text = 0
 If Not IsNumeric(C_Compra_ICM_V.Caption) Then C_Compra_ICM_V.Caption = 0
 If Not IsNumeric(C_Compra_IPI_P.Text) Then C_Compra_IPI_P.Text = 0
 If Not IsNumeric(C_Compra_IPI_V.Caption) Then C_Compra_IPI_V.Caption = 0
 
 Fixa_Desc = "P"
 If Fixa_Desc_C_Valor.Value = 1 Then Fixa_Desc = "V"
 Fixa_Frete = "P"
 If Fixa_Frete_C_Valor.Value = 1 Then Fixa_Frete = "V"
 Fixa_ICM = "P"
 If Fixa_ICM_C_Valor.Value = 1 Then Fixa_ICM = "V"
 Fixa_IPI = "P"
 If Fixa_IPI_C_Valor.Value = 1 Then Fixa_IPI = "V"
 Fixa_Custo = "P"
 If Fixa_Custo_C_Valor.Value = 1 Then Fixa_Custo = "V"
 Fixa_Outros = "P"
 If Fixa_Outros_C_Valor.Value = 1 Then Fixa_Outros = "V"
 
 Compra_Desc_V = C_Compra_Desc_V.Text
 Compra_Desc_P = C_Compra_Desc_P.Text
 Compra_Valor = C_Compra_Valor.Text
 Compra_Frete_V = C_Compra_Frete_V.Text
 Compra_Frete_P = C_Compra_Frete_P.Text
 Compra_ICM_V = C_Compra_ICM_V.Caption
 Compra_ICM_P = C_Compra_ICM_P.Text
 Compra_IPI_V = C_Compra_IPI_V.Caption
 Compra_IPI_P = C_Compra_IPI_P.Text
 Compra_Finan_V = C_Compra_Finan_V.Text
 Compra_Finan_P = C_Compra_Finan_P.Text
 Compra_Outros_V = C_Compra_Outros_V.Text
 Compra_Outros_P = C_Compra_Outros_P.Text
 
 Calcula_Custo Custo, Fixa_Desc, Compra_Desc_V, Compra_Desc_P, _
 Compra_Valor, Fixa_Frete, Compra_Frete_V, Compra_Frete_P, Fixa_ICM, _
 Compra_ICM_V, Compra_ICM_P, Fixa_IPI, Compra_IPI_V, Compra_IPI_P, _
 Fixa_Custo, Compra_Finan_V, Compra_Finan_P, Fixa_Outros, _
 Compra_Outros_V, Compra_Outros_P
 
 C_Compra_Desc_V.Text = Compra_Desc_V
 C_Compra_Desc_P.Text = Compra_Desc_P
 C_Compra_Frete_V.Text = Compra_Frete_V
 C_Compra_Frete_P.Text = Compra_Frete_P
 C_Compra_ICM_V.Caption = Compra_ICM_V
 C_Compra_ICM_P.Text = Compra_ICM_P
 C_Compra_IPI_V.Caption = Compra_IPI_V
 C_Compra_IPI_P.Text = Compra_IPI_P
 C_Compra_Finan_V.Text = Compra_Finan_V
 C_Compra_Finan_P.Text = Compra_Finan_P
 C_Compra_Outros_V.Text = Compra_Outros_V
 C_Compra_Outros_P.Text = Compra_Outros_P
 
 C_Compra_Custo.Caption = Format(Custo, "###,###,###,##0.00")
 
End Sub

Private Sub Calcula_Venda()
 Dim Aux As Double

 If Not IsNumeric(C_Venda_ICM_P.Text) Then C_Venda_ICM_P.Text = 0
 If Not IsNumeric(C_Venda_ICM_V.Caption) Then C_Venda_ICM_V.Caption = 0
 If Not IsNumeric(C_Venda_IPI_P.Text) Then C_Venda_IPI_P.Text = 0
 If Not IsNumeric(C_Venda_IPI_V.Caption) Then C_Venda_IPI_V.Caption = 0
 If Not IsNumeric(C_Venda_Imp_P.Text) Then C_Venda_Imp_P.Text = 0
 If Not IsNumeric(C_Venda_Imp_V.Text) Then C_Venda_Imp_V.Text = 0
 If Not IsNumeric(C_Venda_Outros_P.Text) Then C_Venda_Outros_P.Text = 0
 If Not IsNumeric(C_Venda_Outros_V.Text) Then C_Venda_Outros_V.Text = 0
 
 If Not IsNumeric(C_Venda_Valor.Text) Then Exit Sub
 If CDbl(C_Venda_Valor.Text) = 0 Then Exit Sub
 
 If Fixa_ICM_V_Perc.Value = 1 Then
   Aux = CDbl(C_Venda_Valor.Text) * CDbl(C_Venda_ICM_P.Text) / 100
   C_Venda_ICM_V.Caption = Format(Aux, "###,##0.00")
 Else
   Aux = CDbl(C_Venda_ICM_V.Caption) / CDbl(C_Venda_Valor.Text) * 100
   C_Venda_ICM_P.Text = Format(Aux, "##0.00")
 End If
 
 If Fixa_IPI_V_Perc.Value = 1 Then
   Aux = CDbl(C_Venda_Valor.Text) * CDbl(C_Venda_IPI_P.Text) / 100
   C_Venda_IPI_V.Caption = Format(Aux, "###,##0.00")
 Else
   Aux = CDbl(C_Venda_IPI_V.Caption) / CDbl(C_Venda_Valor.Text) * 100
   C_Venda_IPI_P.Text = Format(Aux, "##0.00")
 End If
 
 If Fixa_Imp_V_Perc.Value = 1 Then
   Aux = CDbl(C_Venda_Valor.Text) * CDbl(C_Venda_Imp_P.Text) / 100
   C_Venda_Imp_V.Text = Format(Aux, "###,##0.00")
 Else
   Aux = CDbl(C_Venda_Imp_V.Text) / CDbl(C_Venda_Valor.Text) * 100
   C_Venda_Imp_P.Text = Format(Aux, "##0.00")
 End If
 
 If Fixa_Outros_V_Perc.Value = 1 Then
   Aux = CDbl(C_Venda_Valor.Text) * CDbl(C_Venda_Outros_P.Text) / 100
   C_Venda_Outros_V.Text = Format(Aux, "###,##0.00")
 Else
   Aux = CDbl(C_Venda_Outros_V.Text) / CDbl(C_Venda_Valor.Text) * 100
   C_Venda_Outros_P.Text = Format(Aux, "##0.00")
 End If
 
End Sub

Public Sub ClearScreen(Optional Encontra As Integer)
  Dim nI As Integer
  
  If bGravando = True Then Exit Sub
  
  cboMotivoDesoneracao.Text = ""
  
  cboCodigo.Text = ""
  
  Ordenação.Caption = ""
  
  O_Produto_Normal.Value = True
  
  ActiveBar1.Tools("miOpShowEdicoes").Enabled = False
  ActiveBar1.Tools("miOpConsultaEstoque").Enabled = False
  ActiveBar1.Tools("miOpInicializaEstoque").Enabled = False
  
  'WEB
  chkWebIncluded.Value = vbUnchecked
  
  Fracionado.Value = 0
  
  If ActiveBar1.Tools("miOpFreezeNome").Checked = False Then
    cboNome.Text = ""
  End If
  If ActiveBar1.Tools("miOpFreezeClasse").Checked = False Then
    Combo_Classe.Text = ""
    Nome_Classe.Caption = ""
  End If
  If ActiveBar1.Tools("miOpFreezeSubClasse").Checked = False Then
    Combo_Sub_Classe.Text = ""
    Nome_Sub_Classe.Caption = ""
  End If
  If ActiveBar1.Tools("miOpFreezeMoeda").Checked = False Then
    Combo_Moeda.Text = ""
    Nome_Moeda.Caption = ""
  End If
  If ActiveBar1.Tools("miOpFreezeUnidadeVenda").Checked = False Then
    Unid.Text = ""
  End If
  If ActiveBar1.Tools("miOpFreezeFabricante").Checked = False Then
    Fabricante.Text = ""
  End If
  If ActiveBar1.Tools("miOpFreezeImpostos").Checked = False Then
    
    '11/03/2011 - Andrea
    'CSO.Text = ""
  End If
  
  '11/02/2010 - mpdea
  txtCasasDecimais.Text = ""
  
  Controla_Estoque.Value = 1
  Desativado.Value = 0
  Não_Lista.Value = 0
  Cod_Fornecedor.Text = ""
  Estoque_Min.Text = ""
  Estoque_Ideal.Text = ""
  '30/10/2007 - Anderson
  'Campo utilizado para gerar o relatório de produtos a comprar
  'Solicitante: Kings Cross
  'txtEspacoFisicoTotal.Text = ""
  Locali.Text = ""
  Desconto.Text = ""
  'Desconto_Max.Text = ""
    
  '29/12/2003 - mpdea
  'Não permitir desconto neste item
  chkDontAllowDesc.Value = vbUnchecked
  
  '20/08/2004 - Daniel
  'Usar Descrição Adicional
  chkUsarDescrAdic.Value = vbUnchecked
  
  Combo_Pesq1.Text = ""
  Combo_Pesq1_LostFocus
  Combo_Pesq2.Text = ""
  Combo_Pesq2_LostFocus
  Combo_Pesq3.Text = ""
  Combo_Pesq3_LostFocus
  
  '23/06/2004 - mpdea
  'Limpa os campos peso bruto e líquido
  txtPesoBruto.Text = "0"
  txtPesoLiquido.Text = "0"
  txtCubagem.Text = "0"
  
  Nome_Nota.Text = ""
  Obs.Text = ""
  
  Última_Compra.Caption = ""
  Último_Custo.Caption = ""
  Custo_Médio.Text = ""
  Último_Fornecedor.Caption = ""
  
  
   For nI = 0 To 3
     medPreco(nI).Text = "0"
   Next nI
   
  
  Nome_Foto.Text = ""
   
  'Tela de Comissão
  Comissão_Sobre.Value = 0
  Comissão.Text = ""
  
  '23/08/2004 - Daniel
  'Campo IndiceFinanceiro
  mskIndice.Mask = ""
  mskIndice.Text = ""
  
  
  '19/09/2005 - mpdea
  'Índice para cálculo do Preço de Entrada
  If g_blnIndicePrecoEntrada Then
    mskIndicePrecoEntrada.Text = 1
  End If
  
  
  '----------------------------------------------------------------------------
  'Data.............: 17/05/2006
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descrição........: Índice ICMS Retido - Base complementar
  mskIndiceIcmsRetido.Text = 1
  '----------------------------------------------------------------------------
  
  
  Comissão_Sobre_Click
  

  'parte de custo
  C_Compra_Valor.Text = 0
  C_Compra_Desc_P.Text = 0
  C_Compra_Desc_V.Text = 0
  Fixa_Desc_C_Perc.Value = 1
  Fixa_Desc_C_Perc.Enabled = False
  Fixa_Desc_C_Valor.Value = 0
  Fixa_Desc_C_Valor.Enabled = True
  
  C_Compra_Frete_P.Text = 0
  C_Compra_Frete_V.Text = 0
  Fixa_Frete_C_Perc.Value = 0
  Fixa_Frete_C_Perc.Enabled = True
  Fixa_Frete_C_Valor.Value = 1
  Fixa_Frete_C_Valor.Enabled = False
  
  C_Compra_ICM_P.Text = 0
  C_Compra_ICM_V.Caption = 0
  Fixa_ICM_C_Perc.Value = 1
  Fixa_ICM_C_Perc.Enabled = False
  Fixa_ICM_C_Valor.Value = 0
  Fixa_ICM_C_Valor.Enabled = True
  
  C_Compra_IPI_P.Text = 0
  C_Compra_IPI_V.Caption = 0
  Fixa_IPI_C_Perc.Value = 1
  Fixa_IPI_C_Perc.Enabled = False
  Fixa_IPI_C_Valor.Value = 0
  Fixa_IPI_C_Valor.Enabled = True
  
  C_Compra_Finan_P.Text = 0
  C_Compra_Finan_V.Text = 0
  Fixa_Custo_C_Perc.Value = 1
  Fixa_Custo_C_Perc.Enabled = False
  Fixa_Custo_C_Valor.Value = 0
  Fixa_Custo_C_Valor.Enabled = True
   
  C_Compra_Outros_P.Text = 0
  C_Compra_Outros_V.Text = 0
  Fixa_Outros_C_Perc.Value = 1
  Fixa_Outros_C_Perc.Enabled = False
  Fixa_Outros_C_Valor.Value = 0
  Fixa_Outros_C_Valor.Enabled = True
   
  C_Compra_Sem_Nota.Text = 0
  
  C_Venda_Valor.Text = 0
  
  C_Venda_ICM_P.Text = 0
  C_Venda_ICM_V.Caption = 0
  Fixa_ICM_V_Perc.Value = 1
  Fixa_ICM_V_Perc.Enabled = False
  Fixa_ICM_V_Valor.Value = 0
  Fixa_ICM_V_Valor.Enabled = True
  
  C_Venda_IPI_P.Text = 0
  C_Venda_IPI_V.Caption = 0
  Fixa_IPI_V_Perc.Value = 1
  Fixa_IPI_V_Perc.Enabled = False
  Fixa_IPI_V_Valor.Value = 0
  Fixa_IPI_V_Valor.Enabled = True
  
  C_Venda_Imp_P.Text = 0
  C_Venda_Imp_V.Text = 0
  Fixa_Imp_V_Perc.Value = 1
  Fixa_Imp_V_Perc.Enabled = False
  Fixa_Imp_V_Valor.Value = 0
  Fixa_Imp_V_Valor.Enabled = True
  
  C_Venda_Outros_P.Text = 0
  C_Venda_Outros_V.Text = 0
  Fixa_Outros_V_Perc.Value = 1
  Fixa_Outros_V_Perc.Enabled = False
  Fixa_Outros_V_Valor.Value = 0
  Fixa_Outros_V_Valor.Enabled = True
  
  C_Venda_Sem_Nota.Text = 0
  C_Venda_Margem_P.Text = 0
  C_Venda_Margem_V.Text = 0
  
  '------------------------------------------
  'Tab Web
  
  chkFabricante.Value = vbUnchecked
  chkPesquisa123.Value = vbUnchecked
  
  cboTableVenda.Text = ""
  cboTablePromocional.Text = ""

  dtpDataInicial.Value = Now
  dtpDataFinal.Value = Now

  'txtBonus.Text = "0"
  
  '10/11/2004 - Daniel
  'Limpando os conteúdos da tab Outros
  txtVolumagem.Text = ""
  txtLote.Text = ""
  mskDataValidade.Mask = ""
  mskDataValidade.Text = ""
  mskDataValidade.Mask = "##/##/####"
  '------------------------------------------
  
  '17/05/2005 - Daniel
  'Case: Miss Nuvem
  'Tratamento para os campos: ImprimirUmaEtiq e ImprimirPrecoEtiq
  chkImprimirUmaEtiq.Value = vbUnchecked
  chkImprimirPrecoEtiq.Value = vbUnchecked
  '------------------------------------------
  
  '04/07/2005 - Daniel
  'Case.....: Zue
  'Projeto..: Informações do Custo de Produto
  Dim intClear As Integer
  
  For intClear = 0 To 8
    txtCustoProduto(intClear).Mask = ""
    txtCustoProduto(intClear).Text = ""
    txtCustoProduto(intClear).Format = "##0.00"
  Next intClear
  '------------------------------------------
  
  '17/04/2007 - Anderson
  'Inclusão do campo para dividir preco da etiqueta - Exigencia do procon
  txtDividirPrecoEtiqueta.Text = 1
  '------------------------------------------
  
  Call StatusMsg("")
  
  If Not rsProdutos2.EOF Then
    On Error Resume Next
    rsProdutos2.MoveFirst
    rsProdutos2.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  
  tabItens.Tab = 0
  tabItens.TabEnabled(1) = False
  tabItens.TabEnabled(2) = False
  tabItens.TabEnabled(3) = False
  tabItens.TabEnabled(4) = False
  tabItens.TabEnabled(5) = False
  tabItens.TabEnabled(6) = False
  
  On Error Resume Next
  If Encontra <> 1 Then cboCodigo.SetFocus
  On Error GoTo 0
  
End Sub

Sub ShowRecord()
  Dim Aux_Str As String
  Dim nI As Integer
  Dim sTableName As String
  Dim nPos As Integer
  Dim MotivoDeseneracao As String
  
  On Error GoTo Processa_Erro
  
  lbl_fornCodigo.Caption = ""
  lbl_fornNome.Caption = ""
  
  Select Case rsProdutos2("MotivoDesoneracaoICMS").Value
    Case "1 "
      cboMotivoDesoneracao.Text = "1 =Táxi"
    Case "2 "
      cboMotivoDesoneracao.Text = "2 =Deficiente Físico"
    Case "3 "
      cboMotivoDesoneracao.Text = "3 =Produtor Agropecuário"
    Case "4 "
      cboMotivoDesoneracao.Text = "4 =Frotista/Locadora"
    Case "5 "
      cboMotivoDesoneracao.Text = "5 =Diplomático/Consular"
    Case "6 "
      cboMotivoDesoneracao.Text = "6 =Utilitários e Motocicletas da Amazônia Ocidental e Áreas de Livre Comércio (Resolução 714/88 e 790/94 – CONTRAN e suas alterações)"
    Case "7 "
      cboMotivoDesoneracao.Text = "7 =SUFRAMA"
    Case "8 "
      cboMotivoDesoneracao.Text = "8 =Venda a Órgão Público"
    Case "9 "
      cboMotivoDesoneracao.Text = "9 =Outros. (NT 2011/004)"
    Case "10"
      cboMotivoDesoneracao.Text = "10 =Deficiente Condutor (Convênio ICMS 38/12)"
    Case "11"
      cboMotivoDesoneracao.Text = "11 =Deficiente Não Condutor (Convênio ICMS 38/12)"
    Case Else
      cboMotivoDesoneracao.Text = ""
    End Select
  
  cboCodigo.Text = rsProdutos2("Código")

  O_Produto_Normal.Value = (rsProdutos2("Tipo") = "N")
  O_Produto_Grade.Value = (rsProdutos2("Tipo") = "G")
  O_Produto_Edição.Value = (rsProdutos2("Tipo") = "E")
  
  'WEB
  chkWebIncluded.Value = IIf(rsProdutos2("WebIncluded").Value, vbChecked, vbUnchecked)
    
  gsTipoProduto = rsProdutos2("Tipo").Value
  gsCodProduto = rsProdutos2("Código").Value
  
  ActiveBar1.Tools("miOpConsultaEstoque").Enabled = True
  ActiveBar1.Tools("miOpInicializaEstoque").Enabled = True
  
  ActiveBar1.Tools("miOpCheckGrid").Enabled = gbGrade
  O_Produto_Grade.Enabled = gbGrade

  O_Produto_Edição.Enabled = gbEdicao
  ActiveBar1.Tools("miOpShowEdicoes").Enabled = gbEdicao
  
  txtCasasDecimais.Text = gsHandleNull(rsProdutos2("QtdeCasasDecimais"))
  Fracionado.Value = 0
  If rsProdutos2("Fracionado") = True Then Fracionado.Value = 1
  cboNome.Text = rsProdutos2("Nome") & ""
  
  Ordenação.Caption = rsProdutos2("Código Ordenação") & ""
  
  Combo_Classe.Text = rsProdutos2("Classe")
  Combo_Classe_LostFocus
  Combo_Sub_Classe.Text = rsProdutos2("Sub Classe")
  Combo_Sub_Classe_LostFocus
  Combo_Moeda.Text = rsProdutos2("Moeda")
  Combo_Moeda_LostFocus
  
  'Select Case tabItens.Tab
  
    'Case 0
    
      Controla_Estoque.Value = 0
      If rsProdutos2("Estoque") = True Then Controla_Estoque.Value = 1
      
      Desativado.Value = 0
      If rsProdutos2("Desativado") = True Then Desativado.Value = 1
      
      Não_Lista.Value = 0
      If rsProdutos2("Não Incluir na Tabela") = True Then Não_Lista.Value = 1
      
      Unid.Text = rsProdutos2("Unidade Venda") & ""
      Cod_Fornecedor.Text = rsProdutos2("Código do Fornecedor") & ""
      Estoque_Min.Text = rsProdutos2("Estoque Mínimo")
      Estoque_Ideal.Text = rsProdutos2("Estoque Ideal")
      '30/10/2007 - Anderson
      'Campo utilizado para gerar o relatório de produtos a comprar
      'Solicitante: Kings Cross
      'txtEspacoFisicoTotal.Text = "" & rsProdutos2("EspacoFisicoTotal")
      Locali.Text = rsProdutos2("Localização") & ""
      Desconto.Text = rsProdutos2("Desconto") & ""
      
  
      '29/12/2003 - mpdea
      'Não permitir desconto neste item
      chkDontAllowDesc.Value = IIf(rsProdutos2.Fields("DontAllowDesc").Value, vbChecked, vbUnchecked)
      
      '20/08/2004 - Daniel
      If rsProdutos2.Fields("UsaDescrAdic").Value Then
        chkUsarDescrAdic.Value = vbChecked
      Else
        chkUsarDescrAdic.Value = vbUnchecked
      End If
      
      Fabricante.Text = rsProdutos2("Fabricante") & ""
      
      Combo_Pesq1.Text = rsProdutos2("Pesquisa 1")
      Combo_Pesq1_LostFocus
      Combo_Pesq2.Text = rsProdutos2("Pesquisa 2")
      Combo_Pesq2_LostFocus
      Combo_Pesq3.Text = rsProdutos2("Pesquisa 3")
      Combo_Pesq3_LostFocus
      
      'Acha preços
      rsPrecos.Index = "Tabela"
      
      For nI = 0 To 3
        medPreco(nI).Text = "0"
        nPos = InStr(lblNomePreco(nI).Caption, ":")
        If nPos > 0 Then
          sTableName = Mid(lblNomePreco(nI).Caption, 1, nPos - 1)
        Else
          sTableName = lblNomePreco(nI).Caption
        End If
        rsPrecos.Seek "=", sTableName, rsProdutos2("Código")
        If Not rsPrecos.NoMatch Then
          If g_bln5CasasDecimais Then '04/05/2004 - Daniel
            medPreco(nI).Text = Format(rsPrecos("Preço"), "##,###,##0.00000")
          '30/04/2007 - Anderson - Implementação de 3 casas decimais
          ElseIf g_bln3CasasDecimais Then
            medPreco(nI).Text = Format(rsPrecos("Preço"), "##,###,##0.000")
          Else
            medPreco(nI).Text = Format(rsPrecos("Preço"), "##,###,##0.00")
          End If
        End If
      Next nI
      
      Nome_Nota.Text = rsProdutos2("Nome Nota") & ""
      Obs.Text = rsProdutos2("Obs") & ""
      
      txtPesoBruto.Text = rsProdutos2("PesoBruto") & ""
      txtPesoLiquido.Text = rsProdutos2("PesoLiquido") & ""
      txtCubagem.Text = rsProdutos2("Cubagem") & ""
  
    'Case 1
      Última_Compra.Caption = rsProdutos2("Última Compra") & ""
      Último_Custo.Caption = Format(rsProdutos2("Último Custo"), "###,##0.00")
      '04/05/2004 - Daniel
      'Case: Embalavi
      If g_bln5CasasDecimais Then
        Custo_Médio.Text = Format(rsProdutos2("Custo Médio"), "###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Custo_Médio.Text = Format(rsProdutos2("Custo Médio"), "###,##0.000")
      Else
        Custo_Médio.Text = Format(rsProdutos2("Custo Médio"), "###,##0.00")
      End If
      
      Último_Fornecedor.Caption = ""
      If rsProdutos2("Último Fornecedor") <> 0 Then
        Aux_Str = str(rsProdutos2("Último Fornecedor")) + " - "
        rsCliFor.Index = "Código"
        rsCliFor.Seek "=", rsProdutos2("Último Fornecedor")
        If Not rsCliFor.NoMatch Then
          Aux_Str = Aux_Str + rsCliFor("Nome")
        End If
        Último_Fornecedor.Caption = Aux_Str
       End If
      
       Call LoadFornec
      
    'Case 2
      'Impostos
      
      '--------------------------------------------------------------------
      '23/08/2004 - Daniel
      'Campo IndiceFinanceiro
      mskIndice.Text = rsProdutos2("IndiceFinanceiro").Value
      
      
      '19/09/2005 - mpdea
      'Índice para cálculo do Preço de Entrada
      If g_blnIndicePrecoEntrada Then
        mskIndicePrecoEntrada.Text = rsProdutos2("IndicePrecoEntrada").Value
      End If
      
      
      '----------------------------------------------------------------------------
      'Data.............: 17/05/2006
      'Desenvolvedor....: mpdea
      'Solicitante......: Pavinato
      'Descrição........: Índice ICMS Retido - Base complementar
      mskIndiceIcmsRetido.Text = rsProdutos2("IndiceIcmsRetido").Value
      '----------------------------------------------------------------------------
      
      
    'Case 3
      If rsProdutos2("Comissão Sobrepõe") = True Then Comissão_Sobre.Value = 1
      If rsProdutos2("Comissão Sobrepõe") = False Then Comissão_Sobre.Value = 0
      Comissão.Text = rsProdutos2("Comissão") & ""
 
    'Case 4
     'Parte de custo
     C_Compra_Valor.Text = rsProdutos2("Custo Preço Valor")
     C_Compra_Desc_P.Text = rsProdutos2("Custo Desconto Perc")
     C_Compra_Desc_V.Text = rsProdutos2("Custo Desconto Valor")
     If rsProdutos2("Custo Desconto Fixo") = "P" Then
       Fixa_Desc_C_Perc.Value = 1
       Fixa_Desc_C_Perc.Enabled = False
       Fixa_Desc_C_Valor.Value = 0
       Fixa_Desc_C_Valor.Enabled = True
     Else
       Fixa_Desc_C_Perc.Value = 0
       Fixa_Desc_C_Perc.Enabled = True
       Fixa_Desc_C_Valor.Value = 1
       Fixa_Desc_C_Valor.Enabled = False
     End If
     
     C_Compra_Frete_P.Text = rsProdutos2("Custo Frete Perc")
     C_Compra_Frete_V.Text = rsProdutos2("Custo Frete Valor")
     If rsProdutos2("Custo Frete Fixo") = "P" Then
       Fixa_Frete_C_Perc.Value = 1
       Fixa_Frete_C_Perc.Enabled = False
       Fixa_Frete_C_Valor.Value = 0
       Fixa_Frete_C_Valor.Enabled = True
     Else
       Fixa_Frete_C_Perc.Value = 0
       Fixa_Frete_C_Perc.Enabled = True
       Fixa_Frete_C_Valor.Value = 1
       Fixa_Frete_C_Valor.Enabled = False
     End If
     
     
     C_Compra_ICM_P.Text = rsProdutos2("Custo ICM Compra Perc")
     C_Compra_ICM_V.Caption = rsProdutos2("Custo ICM Compra Valor")
     If rsProdutos2("Custo ICM Compra Fixo") = "P" Then
       Fixa_ICM_C_Perc.Value = 1
       Fixa_ICM_C_Perc.Enabled = False
       Fixa_ICM_C_Valor.Value = 0
       Fixa_ICM_C_Valor.Enabled = True
     Else
       Fixa_ICM_C_Perc.Value = 0
       Fixa_ICM_C_Perc.Enabled = True
       Fixa_ICM_C_Valor.Value = 1
       Fixa_ICM_C_Valor.Enabled = False
     End If
     
     C_Compra_IPI_P.Text = rsProdutos2("Custo IPI Compra Perc")
     C_Compra_IPI_V.Caption = rsProdutos2("Custo IPI Compra Valor")
     If rsProdutos2("Custo IPI Compra Fixo") = "P" Then
       Fixa_IPI_C_Perc.Value = 1
       Fixa_IPI_C_Perc.Enabled = False
       Fixa_IPI_C_Valor.Value = 0
       Fixa_IPI_C_Valor.Enabled = True
     Else
       Fixa_IPI_C_Perc.Value = 0
       Fixa_IPI_C_Perc.Enabled = True
       Fixa_IPI_C_Valor.Value = 1
       Fixa_IPI_C_Valor.Enabled = False
     End If
       
     C_Compra_Finan_P.Text = rsProdutos2("Custo Custo Finan Perc")
     C_Compra_Finan_V.Text = rsProdutos2("Custo Custo Finan Valor")
     If rsProdutos2("Custo Custo Finan Fixo") = "P" Then
       Fixa_Custo_C_Perc.Value = 1
       Fixa_Custo_C_Perc.Enabled = False
       Fixa_Custo_C_Valor.Value = 0
       Fixa_Custo_C_Valor.Enabled = True
     Else
       Fixa_Custo_C_Perc.Value = 0
       Fixa_Custo_C_Perc.Enabled = True
       Fixa_Custo_C_Valor.Value = 1
       Fixa_Custo_C_Valor.Enabled = False
     End If
     
     C_Compra_Outros_P.Text = rsProdutos2("Custo Outros Compra Perc")
     C_Compra_Outros_V.Text = rsProdutos2("Custo Outros Compra Valor")
     If rsProdutos2("Custo Outros Compra Fixo") = "P" Then
       Fixa_Outros_C_Perc.Value = 1
       Fixa_Outros_C_Perc.Enabled = False
       Fixa_Outros_C_Valor.Value = 0
       Fixa_Outros_C_Valor.Enabled = True
     Else
       Fixa_Outros_C_Perc.Value = 0
       Fixa_Outros_C_Perc.Enabled = True
       Fixa_Outros_C_Valor.Value = 1
       Fixa_Outros_C_Valor.Enabled = False
     End If
     
     
     C_Compra_Sem_Nota.Text = rsProdutos2("Custo Perc Compra Sem")
     C_Compra_Custo.Caption = rsProdutos2("Custo Custo Calculado")
    
     C_Venda_Valor.Text = rsProdutos2("Custo Preço Venda")
     
     C_Venda_ICM_P.Text = rsProdutos2("Custo ICM Venda Perc")
     C_Venda_ICM_V.Caption = rsProdutos2("Custo ICM Venda Valor")
     If rsProdutos2("Custo ICM Venda Fixo") = "P" Then
       Fixa_ICM_V_Perc.Value = 1
       Fixa_ICM_V_Perc.Enabled = False
       Fixa_ICM_V_Valor.Value = 0
       Fixa_ICM_V_Valor.Enabled = True
     Else
       Fixa_ICM_V_Perc.Value = 0
       Fixa_ICM_V_Perc.Enabled = True
       Fixa_ICM_V_Valor.Value = 1
       Fixa_ICM_V_Valor.Enabled = False
     End If
     
     C_Venda_IPI_P.Text = rsProdutos2("Custo IPI Venda Perc")
     C_Venda_IPI_V.Caption = rsProdutos2("Custo IPI Venda Valor")
     If rsProdutos2("Custo IPI Venda Fixo") = "P" Then
       Fixa_IPI_V_Perc.Value = 1
       Fixa_IPI_V_Perc.Enabled = False
       Fixa_IPI_V_Valor.Value = 0
       Fixa_IPI_V_Valor.Enabled = True
     Else
       Fixa_IPI_V_Perc.Value = 0
       Fixa_IPI_V_Perc.Enabled = True
       Fixa_IPI_V_Valor.Value = 1
       Fixa_IPI_V_Valor.Enabled = False
     End If
     
     C_Venda_Imp_P.Text = rsProdutos2("Custo Impostos Perc")
     C_Venda_Imp_V.Text = rsProdutos2("Custo Impostos Valor")
     If rsProdutos2("Custo Impostos Fixo") = "P" Then
       Fixa_Imp_V_Perc.Value = 1
       Fixa_Imp_V_Perc.Enabled = False
       Fixa_Imp_V_Valor.Value = 0
       Fixa_Imp_V_Valor.Enabled = True
     Else
       Fixa_Imp_V_Perc.Value = 0
       Fixa_Imp_V_Perc.Enabled = True
       Fixa_Imp_V_Valor.Value = 1
       Fixa_Imp_V_Valor.Enabled = False
     End If
     
     C_Venda_Outros_P.Text = rsProdutos2("Custo Outros Venda Perc")
     C_Venda_Outros_V.Text = rsProdutos2("Custo Outros Venda Valor")
     If rsProdutos2("Custo Outros Venda Fixo") = "P" Then
       Fixa_Outros_V_Perc.Value = 1
       Fixa_Outros_V_Perc.Enabled = False
       Fixa_Outros_V_Valor.Value = 0
       Fixa_Outros_V_Valor.Enabled = True
     Else
       Fixa_Outros_V_Perc.Value = 0
       Fixa_Outros_V_Perc.Enabled = True
       Fixa_Outros_V_Valor.Value = 1
       Fixa_Outros_V_Valor.Enabled = False
     End If
     
     C_Venda_Sem_Nota.Text = rsProdutos2("Custo Perc Venda Sem") & ""
     C_Venda_Margem_P.Text = rsProdutos2("Custo Lucro Perc") & ""
     C_Venda_Margem_V.Text = rsProdutos2("Custo Lucro Valor") & ""
     
     O_Manter_Venda.Value = (rsProdutos2("Custo Manter") & "" = "P")
     O_Manter_Lucro_Val.Value = (rsProdutos2("Custo Manter") & "" = "L")
     O_Manter_Lucro_Perc.Value = (rsProdutos2("Custo Manter") & "" = "U")
    
      Calcula_Custo1
      Calcula_Venda
    
    'Case 5
      Nome_Foto.Text = rsProdutos2("Foto") & ""
      Call Nome_Foto_LostFocus
      
      
      'WEB
      
      '------------------------------------------------------------------------------
      '08/05/2002 - mpdea
      'Alterado exibição da tab WEB
      '
      '13/08/2002 - mpdea
      'Alterado exibição do atributo Pesquisa 123
      '------------------------------------------------------------------------------
      With rsProdutos2
        chkFabricante.Value = IIf(.Fields("WebAttribFabricante").Value, vbChecked, vbUnchecked)
        
        If gsPesq1 <> "" Or gsPesq2 <> "" Or gsPesq3 <> "" Then
          chkPesquisa123.Value = IIf(.Fields("WebAttribPesquisa123").Value, vbChecked, vbUnchecked)
        Else
          chkPesquisa123.Value = vbUnchecked
        End If
        
'        txtBonus.Text = CLng("0" & .Fields("WebBonus").Value)
        If Not IsNull(.Fields("WebSaleTablePrice").Value) Then
        cboTableVenda.Text = .Fields("WebSaleTablePrice").Value & ""
        Else
        cboTableVenda.Text = Mid(lblNomePreco(1).Caption, 1, Len(lblNomePreco(1).Caption) - 1)
        End If
        cboTablePromocional.Text = .Fields("WebOfferTablePrice").Value & ""
        If .Fields("WebSaleTablePrice").Value & "" <> "" Then
          If Not IsNull(.Fields("WebOfferDateStart").Value) Then
            dtpDataInicial.Value = .Fields("WebOfferDateStart").Value
          End If
          If Not IsNull(.Fields("WebOfferDateEnd").Value) Then
            dtpDataFinal.Value = .Fields("WebOfferDateEnd").Value
          End If
        Else
          dtpDataInicial.Value = Date
          dtpDataFinal.Value = Date
        End If
      End With
      
      

  'End Select
  
  
  'Tab Outros
  
  '27/09/2004 - mpdea
  'Volumagem
  txtVolumagem.Text = rsProdutos2.Fields("Volumagem").Value & ""
  '10/11/2004 - Daniel
  'Campos: Lote e DataValidade
  txtLote.Text = rsProdutos2.Fields("Lote").Value & ""
  If IsDate(rsProdutos2.Fields("DataValidade").Value) Then
    mskDataValidade.Text = rsProdutos2.Fields("DataValidade").Value
  Else
    mskDataValidade.Mask = ""
    mskDataValidade.Text = ""
    mskDataValidade.Mask = "##/##/####"
  End If
  
  '17/05/2005 - Daniel
  'Case: Miss Nuvem
  'Tratamento para os campos ImprimirUmaEtiq e ImprimirPrecoEtiq
  If rsProdutos2.Fields("ImprimirUmaEtiq").Value Then
    chkImprimirUmaEtiq.Value = vbChecked
  Else
    chkImprimirUmaEtiq.Value = vbUnchecked
  End If
  
  If rsProdutos2.Fields("ImprimirPrecoEtiq").Value Then
    chkImprimirPrecoEtiq.Value = vbChecked
  Else
    chkImprimirPrecoEtiq.Value = vbUnchecked
  End If
  
  '04/07/2005 - Daniel
  'Case.....: Zue
  'Projeto..: Informações do Custo de Produto
  '
  'ConsumoDeTecido
  If IsNumeric(rsProdutos2.Fields("ConsumoDeTecido").Value) Then
    txtCustoProduto(0).Text = Format(rsProdutos2.Fields("ConsumoDeTecido").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(0).Text = "0"
  End If
  'PrecoDoMetroLinear
  If IsNumeric(rsProdutos2.Fields("PrecoDoMetroLinear").Value) Then
    txtCustoProduto(1).Text = Format(rsProdutos2.Fields("PrecoDoMetroLinear").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(1).Text = "0"
  End If
  'CustoDoTecido
  If IsNumeric(rsProdutos2.Fields("CustoDoTecido").Value) Then
    txtCustoProduto(2).Text = Format(rsProdutos2.Fields("CustoDoTecido").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(2).Text = "0"
  End If
  'VlMaoDeObraFaccao
  If IsNumeric(rsProdutos2.Fields("VlMaoDeObraFaccao").Value) Then
    txtCustoProduto(3).Text = Format(rsProdutos2.Fields("VlMaoDeObraFaccao").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(3).Text = "0"
  End If
  'VlLavanderia
  If IsNumeric(rsProdutos2.Fields("VlLavanderia").Value) Then
    txtCustoProduto(4).Text = Format(rsProdutos2.Fields("VlLavanderia").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(4).Text = "0"
  End If
  'VlBordado
  If IsNumeric(rsProdutos2.Fields("VlBordado").Value) Then
    txtCustoProduto(5).Text = Format(rsProdutos2.Fields("VlBordado").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(5).Text = "0"
  End If
  'VlEstamparia
  If IsNumeric(rsProdutos2.Fields("VlEstamparia").Value) Then
    txtCustoProduto(6).Text = Format(rsProdutos2.Fields("VlEstamparia").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(6).Text = "0"
  End If
  'VlAviamentos
  If IsNumeric(rsProdutos2.Fields("VlAviamentos").Value) Then
    txtCustoProduto(7).Text = Format(rsProdutos2.Fields("VlAviamentos").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(7).Text = "0"
  End If
  'OutrosCustos
  If IsNumeric(rsProdutos2.Fields("OutrosCustos").Value) Then
    txtCustoProduto(8).Text = Format(rsProdutos2.Fields("OutrosCustos").Value, FORMAT_VALUE)
  Else
    txtCustoProduto(8).Text = "0"
  End If
  '----------------------------------------------------------------------------------------
  
  tabItens.TabEnabled(1) = True
  tabItens.TabEnabled(2) = True
  '20/04/2006 - mpdea
  'Corrigido aplicação da permissão de acesso a pasta de cálculos (custo)
  tabItens.TabEnabled(3) = (rsFuncionarios("Custo Produtos") = True)
  tabItens.TabEnabled(4) = True
  tabItens.TabEnabled(5) = chkWebIncluded.Value = vbChecked
  tabItens.TabEnabled(6) = True
  
  If tabItens.Tab = 5 And Not tabItens.TabEnabled(5) Then
    tabItens.Tab = 0
  End If
  
  '17/04/2007 - Anderson
  'Inclusão do campo para dividir preco da etiqueta - Exigencia do procon
  txtDividirPrecoEtiqueta.Text = rsProdutos2("DividirPrecoEtiqueta")
  
  '15/10/2007 - Anderson
  'Implementação de CFOP
  'Call LoadOperacao
  
  Num_Registro = rsProdutos2.Bookmark
  
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
      Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset)
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
    Case "miOpSearch"
      Call SearchRecord
    Case "miOpPassword"
    Case "miOpCheckGrid"
      Call CheckGrid
    Case "miOpShowEdicoes"
      Call ConsultaEdicoes
    Case "miOpConsultaEstoque"
      Call ConsultaEstoque
    Case "miOpInicializaEstoque"
      Call InicializaEstoque
    Case "miComplRefresh"
      Screen.MousePointer = vbHourglass
      datProdCod.Refresh
      datProdNome.Refresh
      Data1.Refresh
      Data2.Refresh
      Data2_2.Refresh
      Data3.Refresh
      Data4.Refresh
      Data5.Refresh
      Data6.Refresh
      Data7.Refresh
      'Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset)
      FiltraProdutos
'
      'Call InitFormProdutos(True)
      Call InitFormProdutos(Not IsNull(Num_Registro))
      Screen.MousePointer = vbDefault
      
      '26/07/2002 - mpdea
      'Atualiza exibição do registro na janela
      If Not IsNull(Num_Registro) Then ShowRecord
      
    '
    'Opções
    Case "miOpFreezeNome"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Produtos", "Mantem Nome", Tool.Checked)
    
    Case "miOpFreezeClasse"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Produtos", "Mantem Classe", Tool.Checked)
    
    Case "miOpFreezeSubClasse"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Produtos", "Mantem Subclasse", Tool.Checked)
    
    Case "miOpFreezeMoeda"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Produtos", "Mantem Moeda", Tool.Checked)
    
    Case "miOpFreezeUnidadeVenda"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Produtos", "Mantem Unidadevenda", Tool.Checked)
    
    Case "miOpFreezeFabricante"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Produtos", "Mantem Fabricante", Tool.Checked)
    
    Case "miOpFreezeImpostos"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Produtos", "Mantem Impostos", Tool.Checked)
      
    '13/03/2013-Alexandre Afornali
    'Adicionando botao mi0pFilter
    Case "miOpFilter"
      rsParametros.Edit
      If rsParametros("FiltrarProdutosInativos").Value = 0 Then
        rsParametros("FiltrarProdutosInativos").Value = 1
      Else
        rsParametros("FiltrarProdutosInativos").Value = 0
      End If
      rsParametros.Update
      FiltraProdutos
  
  End Select
  '07-10-2015 Jean e Mauro Comentado a chamada da função abaixo para melhorar performance
  'FiltraProdutos
  
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpOrdem"
      Screen.MousePointer = vbHourglass
      Call StatusMsg("Aguarde, reordenando a base...")
      Select Case Tool.CBListIndex
        Case 0
          Set rsProdutos2 = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset, dbSeeChanges)
        Case 1
          Set rsProdutos2 = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Nome]", dbOpenDynaset, dbSeeChanges)
        Case 2
          Set rsProdutos2 = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0'ORDER BY [Código do Fornecedor]", dbOpenDynaset, dbSeeChanges)
        Case 3
          Set rsProdutos2 = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Pesquisa 1]", dbOpenDynaset, dbSeeChanges)
        Case 4
          Set rsProdutos2 = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Pesquisa 2]", dbOpenDynaset, dbSeeChanges)
        Case 5
          Set rsProdutos2 = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Pesquisa 3]", dbOpenDynaset, dbSeeChanges)
      End Select
      Call StatusMsg("")
      Screen.MousePointer = vbDefault
  End Select
End Sub

Private Sub B_Acha_Foto_Click()
  On Error GoTo ErrHandler

  With cdgOpen
    .CancelError = True
    .DialogTitle = "Escolha o arquivo que contém a imagem do produto"
    .Filter = "Todas as Figuras|*.bmp; *.dib; *.gif; *.jpg; *.wmf; *.emf; *.ico; *.cur" & _
      "|Imagens Bitmaps (*.bmp; *.dib)|*.bmp; *.dib" & _
      "|Imagens GIF (*.gif)|*.gif" & _
      "|Imagens JPEG (*.jpg)|*.jpg" & _
      "|Metafiles (*.wmf; *.emf)|*.wmf; *.emf" & _
      "|Ícones (*.ico; *.cur)|*.ico; *.cur" & _
      "|Todos os Arquivos (*.*)|*.*"
    .FilterIndex = 1
    .Flags = cdlOFNPathMustExist & cdlOFNHideReadOnly
    .ShowOpen
    Nome_Foto.Text = .Filename
    Nome_Foto_LostFocus
  End With
  Exit Sub

ErrHandler:
  If Err.Number <> cdlCancel Then
    MsgBox "Ocorreu o erro nº " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  End If
End Sub

Private Sub DeleteRecord()

  Call StatusMsg("")

  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre um produto antes."
    Exit Sub
  End If

  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  
  If Not IsNull(Num_Registro) Then
    frmApagaProdutos.Combo_Produto1.Text = rsProdutos2("Código")
    frmApagaProdutos.Combo_Produto2.Text = rsProdutos2("Código")
    frmApagaProdutos.Nome_Produto1.Caption = rsProdutos2("Nome")
    frmApagaProdutos.Nome_Produto2.Caption = rsProdutos2("Nome")
    'sCodigoProdutoDELETE = rsProdutos2("Código")
    
  End If
  frmApagaProdutos.Show vbModal
  
  datProdCod.Refresh
  datProdNome.Refresh
  
  Call ClearScreen

End Sub


Private Sub Calc_Perc_Lucro()

  Dim min As Double
  Dim Max As Double
  Dim Atual As Double
  Dim Anterior As Double
  Dim Desejado As Double
  Dim Diferença As Double
  Dim Lucro As Double
  Dim Lucro_Perc As Double
  Dim Fim As Integer
  Dim i As Integer
  Dim Venda As Double
 
  Dim Venda_ICM_V As Double
  Dim Venda_ICM_P As Double
  Dim Venda_IPI_V As Double
  Dim Venda_IPI_P As Double
  Dim Venda_Imp_V As Double
  Dim Venda_Imp_P As Double
  Dim Venda_Outros_V As Double
  Dim Venda_Outros_P As Double
  Dim Compra_Valor As Double
  Dim Compra_Desc_V As Double
  Dim Compra_Frete_V As Double
  Dim Compra_Finan_V As Double
  Dim Compra_Outros_V As Double
  Dim Compra_ICM_V As Double
  Dim Compra_IPI_V As Double
 
  Call StatusMsg("")
  If CDbl(C_Compra_Valor.Text) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Impossível calcular, preço de compra igual a 0."
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  If Not IsNumeric(C_Venda_ICM_P.Text) Then C_Venda_ICM_P.Text = 0
  If Not IsNumeric(C_Venda_ICM_V.Caption) Then C_Venda_ICM_V.Caption = 0
  If Not IsNumeric(C_Venda_IPI_P.Text) Then C_Venda_IPI_P.Text = 0
  If Not IsNumeric(C_Venda_IPI_V.Caption) Then C_Venda_IPI_V.Caption = 0
  If Not IsNumeric(C_Venda_Imp_P.Text) Then C_Venda_Imp_P.Text = 0
  If Not IsNumeric(C_Venda_Imp_V.Text) Then C_Venda_Imp_V.Text = 0
  If Not IsNumeric(C_Venda_Outros_P.Text) Then C_Venda_Outros_P.Text = 0
  If Not IsNumeric(C_Venda_Outros_V.Text) Then C_Venda_Outros_V.Text = 0
  If Not IsNumeric(C_Venda_Sem_Nota.Text) Then C_Venda_Sem_Nota.Text = 0
  If Not IsNumeric(C_Compra_Sem_Nota.Text) Then C_Compra_Sem_Nota.Text = 0
  If Not IsNumeric(C_Compra_Valor.Text) Then C_Compra_Valor.Text = 0
  If Not IsNumeric(C_Compra_Desc_V.Text) Then C_Compra_Desc_V.Text = 0
  If Not IsNumeric(C_Compra_Frete_V.Text) Then C_Compra_Frete_V.Text = 0
  If Not IsNumeric(C_Compra_Finan_V.Text) Then C_Compra_Finan_V.Text = 0
  If Not IsNumeric(C_Compra_Outros_V.Text) Then C_Compra_Outros_V.Text = 0
  If Not IsNumeric(C_Compra_ICM_V.Caption) Then C_Compra_ICM_V.Caption = 0
  If Not IsNumeric(C_Compra_IPI_V.Caption) Then C_Compra_IPI_V.Caption = 0

 If Not IsNumeric(C_Venda_Margem_P.Text) Then Exit Sub
 
 Max = CDbl(C_Compra_Valor.Text)
 Desejado = CDbl(C_Venda_Margem_P.Text)
 Max = Max * 500
 If Max < 1000 Then Max = 10000
 
 Call StatusMsg("")
 
 Fim = False
 
 Atual = (min + Max) / 2
 
 Diferença = 0.0001
 
 i = 0
 
 Do
  Rem Seta variáveis
  Venda_ICM_V = C_Venda_ICM_V.Caption
  Venda_ICM_P = C_Venda_ICM_P.Text
  Venda_IPI_V = C_Venda_IPI_V.Caption
  Venda_IPI_P = C_Venda_IPI_P.Text
  Venda_Imp_V = C_Venda_Imp_V.Text
  Venda_Imp_P = C_Venda_Imp_P.Text
  Venda_Outros_V = C_Venda_Outros_V.Text
  Venda_Outros_P = C_Venda_Outros_P.Text
  Compra_Valor = C_Compra_Valor.Text
  Compra_Desc_V = C_Compra_Desc_V.Text
  Compra_Frete_V = C_Compra_Frete_V.Text
  Compra_Finan_V = C_Compra_Finan_V.Text
  Compra_Outros_V = C_Compra_Outros_V.Text
  Compra_ICM_V = C_Compra_ICM_V.Caption
  Compra_IPI_V = C_Compra_IPI_V.Caption
 
  Venda = Atual

  Call StatusMsg("Aguarde, calculando " & str(i))
  DoEvents

  Calcula_Lucro Venda, Venda_ICM_P, Venda_ICM_V, _
    Venda_IPI_P, Venda_IPI_V, Venda_Imp_P, Venda_Imp_V, _
    Venda_Outros_V, Venda_Outros_P, C_Venda_Sem_Nota.Text, C_Compra_Sem_Nota.Text, _
    -Fixa_ICM_V_Perc.Value, -Fixa_IPI_V_Perc.Value, -Fixa_Imp_V_Perc.Value, _
    -Fixa_Outros_V_Perc.Value, Lucro, Compra_Valor, Compra_Desc_V, _
    Compra_Frete_V, Compra_Finan_V, Compra_Outros_V, Compra_ICM_V, Compra_IPI_V
   
  C_Venda_ICM_V.Caption = Venda_ICM_V
  C_Venda_ICM_P.Text = Venda_ICM_P
  C_Venda_IPI_V.Caption = Venda_IPI_V
  C_Venda_IPI_P.Text = Venda_IPI_P
  C_Venda_Imp_V.Text = Venda_Imp_V
  C_Venda_Imp_P.Text = Venda_Imp_P
  C_Venda_Outros_V.Text = Venda_Outros_V
  C_Venda_Outros_P.Text = Venda_Outros_P
   
  Lucro_Perc = (Lucro / (CDbl(C_Compra_Valor.Text) - CDbl(C_Compra_Desc_V.Text))) * CDbl(100)
  If Abs(Lucro_Perc - Desejado) < Diferença Then
    Fim = True
  End If
  If Fim = False Then
    If Lucro_Perc > Desejado Then
       Max = Atual
       Atual = (min + Atual) / 2
    Else
       min = Atual
       Atual = (Atual + Max) / 2
    End If
    i = i + 1
    If i > 150 Then Diferença = 0.001
    If i > 250 Then Diferença = 0.1
    If i > 1000 Then
      DisplayMsg "Não foi possível calcular o preço."
      Exit Sub
    End If
  End If
 Loop Until Fim = True
 
 C_Venda_Valor.Text = Format(Atual, "#########0.00")
 
 'DisplayMsg "Preço calculado em " + str(i) + " tentativas."
 
 Call Calc_Valor
 
End Sub

Private Sub Calc_Valor()
  Dim Lucro As Double
  Dim Venda As Double
  Dim Margem As Double
  
  Dim Venda_ICM_V As Double
  Dim Venda_ICM_P As Double
  Dim Venda_IPI_V As Double
  Dim Venda_IPI_P As Double
  Dim Venda_Imp_V As Double
  Dim Venda_Imp_P As Double
  Dim Venda_Outros_V As Double
  Dim Venda_Outros_P As Double
  Dim Compra_Valor As Double
  Dim Compra_Desc_V As Double
  Dim Compra_Frete_V As Double
  Dim Compra_Finan_V As Double
  Dim Compra_Outros_V As Double
  Dim Compra_ICM_V As Double
  Dim Compra_IPI_V As Double
  
  Call StatusMsg("")
  
  If CDbl(C_Compra_Valor.Text) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Impossível calcular, preço de compra igual a 0."
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If Not IsNumeric(C_Venda_ICM_P.Text) Then C_Venda_ICM_P.Text = 0
  If Not IsNumeric(C_Venda_ICM_V.Caption) Then C_Venda_ICM_V.Caption = 0
  If Not IsNumeric(C_Venda_IPI_P.Text) Then C_Venda_IPI_P.Text = 0
  If Not IsNumeric(C_Venda_IPI_V.Caption) Then C_Venda_IPI_V.Caption = 0
  If Not IsNumeric(C_Venda_Imp_P.Text) Then C_Venda_Imp_P.Text = 0
  If Not IsNumeric(C_Venda_Imp_V.Text) Then C_Venda_Imp_V.Text = 0
  If Not IsNumeric(C_Venda_Outros_P.Text) Then C_Venda_Outros_P.Text = 0
  If Not IsNumeric(C_Venda_Outros_V.Text) Then C_Venda_Outros_V.Text = 0
  If Not IsNumeric(C_Venda_Sem_Nota.Text) Then C_Venda_Sem_Nota.Text = 0
  If Not IsNumeric(C_Compra_Sem_Nota.Text) Then C_Compra_Sem_Nota.Text = 0
  If Not IsNumeric(C_Compra_Valor.Text) Then C_Compra_Valor.Text = 0
  If Not IsNumeric(C_Compra_Desc_V.Text) Then C_Compra_Desc_V.Text = 0
  If Not IsNumeric(C_Compra_Frete_V.Text) Then C_Compra_Frete_V.Text = 0
  If Not IsNumeric(C_Compra_Finan_V.Text) Then C_Compra_Finan_V.Text = 0
  If Not IsNumeric(C_Compra_Outros_V.Text) Then C_Compra_Outros_V.Text = 0
  If Not IsNumeric(C_Compra_ICM_V.Caption) Then C_Compra_ICM_V.Caption = 0
  If Not IsNumeric(C_Compra_IPI_V.Caption) Then C_Compra_IPI_V.Caption = 0
  
  Venda_ICM_V = C_Venda_ICM_V.Caption
  Venda_ICM_P = C_Venda_ICM_P.Text
  Venda_IPI_V = C_Venda_IPI_V.Caption
  Venda_IPI_P = C_Venda_IPI_P.Text
  Venda_Imp_V = C_Venda_Imp_V.Text
  Venda_Imp_P = C_Venda_Imp_P.Text
  Venda_Outros_V = C_Venda_Outros_V.Text
  Venda_Outros_P = C_Venda_Outros_P.Text
  Compra_Valor = C_Compra_Valor.Text
  Compra_Desc_V = C_Compra_Desc_V.Text
  Compra_Frete_V = C_Compra_Frete_V.Text
  Compra_Finan_V = C_Compra_Finan_V.Text
  Compra_Outros_V = C_Compra_Outros_V.Text
  Compra_ICM_V = C_Compra_ICM_V.Caption
  Compra_IPI_V = C_Compra_IPI_V.Caption
    
  Venda = CDbl(C_Venda_Valor.Text)

  Call Calcula_Lucro(Venda, _
    Venda_ICM_P, Venda_ICM_V, _
    Venda_IPI_P, Venda_IPI_V, _
    Venda_Imp_P, Venda_Imp_V, _
    Venda_Outros_V, Venda_Outros_P, _
    C_Venda_Sem_Nota.Text, C_Compra_Sem_Nota.Text, _
    -Fixa_ICM_V_Perc.Value, -Fixa_IPI_V_Perc.Value, _
    -Fixa_Imp_V_Perc.Value, -Fixa_Outros_V_Perc.Value, _
    Lucro, _
    Compra_Valor, Compra_Desc_V, _
    Compra_Frete_V, Compra_Finan_V, _
    Compra_Outros_V, Compra_ICM_V, Compra_IPI_V)

  C_Venda_ICM_V.Caption = Venda_ICM_V
  C_Venda_ICM_P.Text = Venda_ICM_P
  C_Venda_IPI_V.Caption = Venda_IPI_V
  C_Venda_IPI_P.Text = Venda_IPI_P
  C_Venda_Imp_V.Text = Venda_Imp_V
  C_Venda_Imp_P.Text = Venda_Imp_P
  C_Venda_Outros_V.Text = Venda_Outros_V
  C_Venda_Outros_P.Text = Venda_Outros_P

  C_Venda_Margem_V.Text = Lucro

  If CDbl(C_Compra_Valor.Text) - CDbl(C_Compra_Desc_V.Text) = 0 Then
    Exit Sub
  End If

  Margem = Lucro / (CDbl(C_Compra_Valor.Text) - CDbl(C_Compra_Desc_V.Text))
  Margem = Margem * 100
  C_Venda_Margem_P.Text = Format(Margem, "##0.00")

End Sub

Private Sub Calc_Valor_Lucro()

  Dim min As Double
  Dim Max As Double
  Dim Atual As Double
  Dim Anterior As Double
  Dim Desejado As Double
  Dim Diferença As Double
  Dim Lucro As Double
  Dim Fim As Integer
  Dim i As Integer
  Dim Venda As Double
 
  Dim Venda_ICM_V As Double
  Dim Venda_ICM_P As Double
  Dim Venda_IPI_V As Double
  Dim Venda_IPI_P As Double
  Dim Venda_Imp_V As Double
  Dim Venda_Imp_P As Double
  Dim Venda_Outros_V As Double
  Dim Venda_Outros_P As Double
  Dim Compra_Valor As Double
  Dim Compra_Desc_V As Double
  Dim Compra_Frete_V As Double
  Dim Compra_Finan_V As Double
  Dim Compra_Outros_V As Double
  Dim Compra_ICM_V As Double
  Dim Compra_IPI_V As Double
 
  Call StatusMsg("")
  If CDbl(C_Compra_Valor.Text) = 0 Then
     gsTitle = LoadResString(201)
     gsMsg = "Impossível calcular, preço de compra igual a 0."
     gnStyle = vbOKOnly + vbCritical
     gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
     Exit Sub
  End If
 
  If Not IsNumeric(C_Venda_ICM_P.Text) Then C_Venda_ICM_P.Text = 0
  If Not IsNumeric(C_Venda_ICM_V.Caption) Then C_Venda_ICM_V.Caption = 0
  If Not IsNumeric(C_Venda_IPI_P.Text) Then C_Venda_IPI_P.Text = 0
  If Not IsNumeric(C_Venda_IPI_V.Caption) Then C_Venda_IPI_V.Caption = 0
  If Not IsNumeric(C_Venda_Imp_P.Text) Then C_Venda_Imp_P.Text = 0
  If Not IsNumeric(C_Venda_Imp_V.Text) Then C_Venda_Imp_V.Text = 0
  If Not IsNumeric(C_Venda_Outros_P.Text) Then C_Venda_Outros_P.Text = 0
  If Not IsNumeric(C_Venda_Outros_V.Text) Then C_Venda_Outros_V.Text = 0
  If Not IsNumeric(C_Venda_Sem_Nota.Text) Then C_Venda_Sem_Nota.Text = 0
  If Not IsNumeric(C_Compra_Sem_Nota.Text) Then C_Compra_Sem_Nota.Text = 0
  If Not IsNumeric(C_Compra_Valor.Text) Then C_Compra_Valor.Text = 0
  If Not IsNumeric(C_Compra_Desc_V.Text) Then C_Compra_Desc_V.Text = 0
  If Not IsNumeric(C_Compra_Frete_V.Text) Then C_Compra_Frete_V.Text = 0
  If Not IsNumeric(C_Compra_Finan_V.Text) Then C_Compra_Finan_V.Text = 0
  If Not IsNumeric(C_Compra_Outros_V.Text) Then C_Compra_Outros_V.Text = 0
  If Not IsNumeric(C_Compra_ICM_V.Caption) Then C_Compra_ICM_V.Caption = 0
  If Not IsNumeric(C_Compra_IPI_V.Caption) Then C_Compra_IPI_V.Caption = 0
 
  Max = CDbl(C_Venda_Margem_V.Text)
  Desejado = Max
  Max = Max * 200
  
  Call StatusMsg("")
  
  Fim = False
  
  Atual = (min + Max) / 2
  
  Diferença = 0.0000001
  
  i = 0
  
  Do
    Venda_ICM_V = C_Venda_ICM_V.Caption
    Venda_ICM_P = C_Venda_ICM_P.Text
    Venda_IPI_V = C_Venda_IPI_V.Caption
    Venda_IPI_P = C_Venda_IPI_P.Text
    Venda_Imp_V = C_Venda_Imp_V.Text
    Venda_Imp_P = C_Venda_Imp_P.Text
    Venda_Outros_V = C_Venda_Outros_V.Text
    Venda_Outros_P = C_Venda_Outros_P.Text
    Compra_Valor = C_Compra_Valor.Text
    Compra_Desc_V = C_Compra_Desc_V.Text
    Compra_Frete_V = C_Compra_Frete_V.Text
    Compra_Finan_V = C_Compra_Finan_V.Text
    Compra_Outros_V = C_Compra_Outros_V.Text
    Compra_ICM_V = C_Compra_ICM_V.Caption
    Compra_IPI_V = C_Compra_IPI_V.Caption
      
    Venda = Atual
    
    
    Calcula_Lucro Venda, Venda_ICM_P, Venda_ICM_V, _
        Venda_IPI_P, Venda_IPI_V, Venda_Imp_P, Venda_Imp_V, _
        Venda_Outros_V, Venda_Outros_P, C_Venda_Sem_Nota.Text, C_Compra_Sem_Nota.Text, _
        -Fixa_ICM_V_Perc.Value, -Fixa_IPI_V_Perc.Value, -Fixa_Imp_V_Perc.Value, _
        -Fixa_Outros_V_Perc.Value, Lucro, Compra_Valor, Compra_Desc_V, _
        Compra_Frete_V, Compra_Finan_V, Compra_Outros_V, Compra_ICM_V, Compra_IPI_V
    
    C_Venda_ICM_V.Caption = Venda_ICM_V
    C_Venda_ICM_P.Text = Venda_ICM_P
    C_Venda_IPI_V.Caption = Venda_IPI_V
    C_Venda_IPI_P.Text = Venda_IPI_P
    C_Venda_Imp_V.Text = Venda_Imp_V
    C_Venda_Imp_P.Text = Venda_Imp_P
    C_Venda_Outros_V.Text = Venda_Outros_V
    C_Venda_Outros_P.Text = Venda_Outros_P
    
    If Abs(Lucro - Desejado) < Diferença Then
      Fim = True
    End If
    If Fim = False Then
      If Lucro > Desejado Then
         Max = Atual
         Atual = (min + Atual) / 2
      Else
         min = Atual
         Atual = (Atual + Max) / 2
      End If
      i = i + 1
      If i > 150 Then Diferença = 0.001
      If i > 250 Then Diferença = 0.1
      If i > 1000 Then
        gsTitle = LoadResString(201)
        gsMsg = "Não foi possível calcular o preço."
        gnStyle = vbOKOnly + vbCritical
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Exit Sub
      End If
    End If
  Loop Until Fim = True
  
  C_Venda_Valor.Text = Format(Atual, "#########0.00")
  
  'DisplayMsg "Preço calculado em " + str(i) + " tentativas."
  
  Call Calc_Valor
  
End Sub

Private Sub ConsultaEstoque()
  Call StatusMsg("")
  frmConsultaEstoque.Show vbModal
End Sub


Private Sub ConsultaEdicoes()
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre ou grave um produto antes."
    Exit Sub
  End If
  
  If O_Produto_Edição.Value = False Then
    DisplayMsg "Este produto não usa edições."
    Exit Sub
  End If
   
  frmEdicoes.Código.Caption = rsProdutos2("Código")
  frmEdicoes.Show vbModal
 
End Sub

Private Sub SearchRecord()
  Dim F As Form
  Set F = New frmPesquisaProduto
  F.Show
End Sub

Private Sub CheckGrid()
 Call StatusMsg("")

 If IsNull(Num_Registro) Then
   DisplayMsg "Encontre um produto antes."
   Exit Sub
 End If

 If rsProdutos2("Tipo") <> "G" Then
   DisplayMsg "Este produto não usa grade."
   Exit Sub
 End If
 
 ' Tratamento Integração com Wise (inicio)
 Dim nI As Integer
 frmGrade.txt_unidadeProduto = Unid
 
 If Preços.Enabled = True Then
      For nI = 0 To 3
        If lblNomePreco(nI).Caption = "A VISTA:" Then
          With rsPrecos
            .Index = "Tabela"
            '.Seek "=", sTableName, cboCodigo.Text
            If .NoMatch Then
          
              If IsNull(medPreco(nI).Text) Or medPreco(nI).Text = "0" Then
                 frmGrade.txt_valorProduto = "0"
              ElseIf Not IsNumeric(medPreco(nI).Text) Then
                 frmGrade.txt_valorProduto = "0"
              Else
                 frmGrade.txt_valorProduto = CSng(medPreco(nI).Text)
              End If
            Else
              If IsNull(medPreco(nI).Text) Or medPreco(nI).Text = "0" Then
                 frmGrade.txt_valorProduto = "0"
              ElseIf Not IsNumeric(medPreco(nI).Text) Then
                 frmGrade.txt_valorProduto = "0"
              Else
                 frmGrade.txt_valorProduto = CSng(medPreco(nI).Text)
              End If
            End If
            
            frmGrade.txt_valorProduto = Replace(frmGrade.txt_valorProduto, ".", "")
            frmGrade.txt_valorProduto = Replace(frmGrade.txt_valorProduto, ",", ".")
          End With
        End If
      Next nI
  End If
  
 ' Fim tratamento

 frmGrade.Cod_Prod.Caption = cboCodigo.Text
 frmGrade.Nome_Prod.Caption = cboNome.Text
 frmGrade.Classe.Caption = rsProdutos2("Classe")
 frmGrade.Sub_Classe.Caption = rsProdutos2("Sub Classe")
 
 frmGrade.Show vbModal
End Sub


'
'Public Function WebRequest(url As String) As Boolean
'On Error GoTo trata_WebApiErro
'    ' Função que chama o componente WebApi
'    Dim retWebApiAux As String
'
'    http.Open "GET", url, False
'    http.Send
'
'    retWebApiAux = http.statusText
'
'    If retWebApiAux = "OK" Then
'        sRETORNO_ERRO_WEBAPI = "Cadastrado com sucesso!"
'        WebRequest = True
'    Else
'        sRETORNO_ERRO_WEBAPI = http.responseText
'        WebRequest = False
'    End If
'
'    Exit Function
'
'trata_WebApiErro:
'    Dim iInStrRet As Integer
'    iInStrRet = InStr(1, Err.Description, "timed out")
'    If iInStrRet > 0 And tentativasWEBAPI < 4 Then
'        tentativasWEBAPI = tentativasWEBAPI + 1
'        WebRequest url
'    End If
'    sRETORNO_ERRO_WEBAPI = Err.Description
'End Function


Private Sub UpdateRecord()
  Dim Erro As Integer
  Dim nI As Integer
  Dim sTableName As String
  Dim nPos As Integer
  Dim strAux As String
  Dim sCh As String
  Dim lngRet As Long
  Dim intRet As Integer
  
  Dim sngPreco As Single
  
  '13/05/2004 - mpdea
  'Implementado controle da transação
  'Corrigindo RT-3034
  Dim blnInTransaction As Boolean
  
  
  '15/09/2005 - mpdea
  'Retorno de eventos Validate
  Dim blnCancel As Boolean
  
  Dim blnNaoExisteTabelaPrecoCusto As Boolean
  
  On Error GoTo ErrHandler
    
  
  
  If IsNull(Num_Registro) And gbDemoVersion Then
    rsProdutos2.MoveLast
    rsProdutos2.MoveFirst
    If rsProdutos2.RecordCount >= NMAXREGDEMO Then
      gsTitle = LoadResString(201)
      gsMsg = LoadResString(13)
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Exit Sub
    End If
  End If
  
  cboCodigo.Text = gsHandleNull(cboCodigo.Text)
  
  If cboCodigo.Text = "0" Then
    DisplayMsg "Código incorreto, verifique."
    cboCodigo.SetFocus
    Exit Sub
  End If
  
  '----------------------------------------------------------------------------
  '02/10/2008 - mpdea
  'Permitir ponto como parte do código do produto
  For nI = 1 To Len(cboCodigo.Text)
    sCh = UCase(Mid(cboCodigo.Text, nI, 1))
    If Not ((sCh >= "A" And sCh <= "Z") Or (sCh >= "0" And sCh <= "9") Or sCh = "-" Or sCh = "." Or sCh = "/") Then
      DisplayMsg "Caracter inválido em código de produto. Verifique."
      cboCodigo.SetFocus
      Exit Sub
    End If
  Next nI
'  For nI = 1 To Len(cboCodigo.Text)
'    sCh = UCase(Mid(cboCodigo.Text, nI, 1))
'    If Not ((sCh >= "A" And sCh <= "Z") Or (sCh >= "0" And sCh <= "9") Or sCh = "-") Then
'      DisplayMsg "Caracter inválido em código de produto. Verifique."
'      cboCodigo.SetFocus
'      Exit Sub
'    End If
'  Next nI
  '----------------------------------------------------------------------------
  
  '26/05/2004 - Daniel
  'Tratamento para 0 'zero' a esquerda
  If Not gbZeroEsquerda Then
    'Verifica se o primeiro caracter é zero
    If Left(cboCodigo.Text, 1) = "0" Then
      DisplayMsg "Código do produto não pode iniciar com zero."
      cboCodigo.SetFocus
      Exit Sub
    End If
  End If
  
  '----------------------------------------------------------------------------
  '02/10/2008 - mpdea
  'Permitir até 14 caracteres no código do produto
  If O_Produto_Grade.Value = True Then
    If Len(cboCodigo.Text) > 14 Or Len(cboCodigo.Text) < 4 Then
      DisplayMsg "Para produtos com grade, o tamanho do 'código do produto' deve ser entre 4 e 14 dígitos."
      cboCodigo.SetFocus
      Exit Sub
    End If
  End If
'  'Se tiver grade o tamanho do código deve ser 8 carac.
'  If O_Produto_Grade.Value = True Then
'    If Len(cboCodigo.Text) > 8 Or Len(cboCodigo.Text) < 4 Then
'      DisplayMsg "Produto que usa grade deve ter entre 4 e 8 dígitos."
'      cboCodigo.SetFocus
'      Exit Sub
'    End If
'  End If
  '----------------------------------------------------------------------------
  
  If O_Produto_Edição.Value = True Then
    If Len(Trim(cboCodigo.Text)) <> 13 Then
      DisplayMsg "Produto com edição deve ter 13 dígitos."
      cboCodigo.SetFocus
      Exit Sub
    End If
    If Not IsNumeric(cboCodigo.Text) Then
      DisplayMsg "Produto com edição deve ser numérico."
      Exit Sub
    End If
  End If
  
  If cboNome.Text & "" = "" Then
    DisplayMsg "Nome incorreto, verifique."
    cboNome.SetFocus
    Exit Sub
  End If
  
  
  '10/07/2006 - mpdea
  'Verifica alteração do campo nome
  If Not IsNull(Num_Registro) Then
    If rsProdutos2.Fields("Nome").Value <> cboNome.Text Then
      strAux = "Confirma alteração do nome do produto?"
      strAux = strAux & vbCrLf & vbCrLf & "De: " & rsProdutos2.Fields("Nome").Value
      strAux = strAux & vbCrLf & "Para: " & cboNome.Text & "?"
      Select Case MsgBox(strAux, vbExclamation + vbYesNoCancel, "Atenção")
        Case vbNo
          cboNome.Text = rsProdutos2.Fields("Nome").Value
        Case vbCancel
          Exit Sub
      End Select
    End If
  End If
  
  
  If Nome_Classe.Caption = "" Then
    DisplayMsg "Classe incorreta, por favor verifique."
    Combo_Classe.SetFocus
    Exit Sub
  End If
  
  If Nome_Sub_Classe.Caption = "" Then
    DisplayMsg "Sub Classe incorreta, por favor verifique."
    Combo_Sub_Classe.SetFocus
    Exit Sub
  End If
  
    
  '23/01/2003 - mpdea
  'Moeda padrão para o modo limitado
  If Not gblnQuickFull Then
    Combo_Moeda.Text = "1"
  Else
    If Nome_Moeda.Caption = "" Then
      DisplayMsg "Moeda incorreta, por favor verifique."
      Combo_Moeda.SetFocus
      Exit Sub
    End If
  End If
  
  txtCasasDecimais.Text = gsHandleNull(txtCasasDecimais.Text)
  If Not IsNumeric(txtCasasDecimais.Text) Then
    DisplayMsg "Quantidade de casas decimais inválida."
    txtCasasDecimais.SetFocus
    Exit Sub
  End If
  If CInt(txtCasasDecimais.Text) < 0 Or CInt(txtCasasDecimais.Text) > 4 Then
    DisplayMsg "Quantidade de casas decimais deve estar entre 0 e 4."
    txtCasasDecimais.SetFocus
    Exit Sub
  End If

  If Not IsNumeric(Estoque_Min.Text) Then Estoque_Min.Text = "0"
  If CSng(Estoque_Min.Text) < 0 Then Estoque_Min.Text = "0"

  If Not IsNumeric(Estoque_Ideal.Text) Then Estoque_Ideal.Text = "0"
  If CSng(Estoque_Ideal.Text) < 0 Then Estoque_Ideal.Text = "0"

  If CSng(Estoque_Min.Text) > 0 Then
    If CSng(Estoque_Ideal.Text) < CSng(Estoque_Min.Text) Then
      gsTitle = LoadResString(201)
      gsMsg = "Estoque ideal não pode ser inferior ao estoque mínimo."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Estoque_Ideal.SetFocus
      Exit Sub
    End If
  End If

  '30/10/2007 - Anderson
  'Campo utilizado para gerar o relatório de produtos a comprar
  'Solicitante: Kings Cross
  'If Not IsNumeric(txtEspacoFisicoTotal.Text) Then txtEspacoFisicoTotal.Text = "0"
  'If CDbl(txtEspacoFisicoTotal.Text) < 0 Then txtEspacoFisicoTotal.Text = "0"
  
  Desconto.Text = gsHandleNull(Desconto.Text)
  If CSng(Desconto.Text) < 0# Then Desconto.Text = "0"
  If CSng(Desconto.Text) > 99.99 Then
    gsTitle = LoadResString(201)
    gsMsg = "Desconto não pode ser superior a 99.99%."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If gsPesq1 <> "" Then
    Call Combo_Pesq1_LostFocus
    If Nome_Pesq1.Caption = "" And Combo_Pesq1.Text <> "" And Combo_Pesq1.Text <> "0" Then
      DisplayMsg "Pesquisa inválida."
      Combo_Pesq1.SetFocus
      Exit Sub
    End If
  Else
    Combo_Pesq1.Text = "0"
  End If
  
  If gsPesq2 <> "" Then
    Call Combo_Pesq2_LostFocus
    If Nome_Pesq2.Caption = "" And Combo_Pesq2.Text <> "" And Combo_Pesq2.Text <> "0" Then
      DisplayMsg "Pesquisa inválida."
      Combo_Pesq2.SetFocus
      Exit Sub
    End If
  Else
    Combo_Pesq2.Text = "0"
  End If
  
  If gsPesq3 <> "" Then
    Call Combo_Pesq3_LostFocus
    If Nome_Pesq3.Caption = "" And Combo_Pesq3.Text <> "" And Combo_Pesq3.Text <> "0" Then
      DisplayMsg "Pesquisa inválida."
      Combo_Pesq3.SetFocus
      Exit Sub
    End If
  Else
    Combo_Pesq3.Text = "0"
  End If
  
  ' If IsNull(Desconto_Max.Text) Then Desconto_Max.Text = 0
  ' If Not IsNumeric(Desconto_Max.Text) Then Desconto_Max.Text = 0
  ' If CDbl(Desconto_Max.Text) < 0 Then Desconto_Max.Text = 0
  ' If CDbl(Desconto_Max.Text) > 100 Then Desconto_Max.Text = 0
  ' Desconto_Max.Text = CDbl(Desconto_Max.Text)
  
  Comissão.Text = gsHandleNull(Comissão.Text)
  If Not IsNumeric(Comissão.Text) Then Comissão.Text = "0"
  
  If Len(Trim(txtPesoBruto.Text)) > 0 Then
    If Not IsNumeric(txtPesoBruto.Text) Then
      gsTitle = LoadResString(201)
      gsMsg = "Peso Bruto inválido."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      txtPesoBruto.SetFocus
      Exit Sub
    End If
  End If

  If Len(Trim(txtPesoLiquido.Text)) > 0 Then
    If Not IsNumeric(txtPesoLiquido.Text) Then
      gsTitle = LoadResString(201)
      gsMsg = "Peso Líquido inválido."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      txtPesoLiquido.SetFocus
      Exit Sub
    End If
  End If
  
  
  '07/01/2004 - mpdea
  'Validação para o AM
  '
  'Verifica o cadastramento de preços > 0
  If UCase(gstrGetEstadoFilial(gnCodFilial)) = "AM" Then
    If Preços.Enabled Then
      For nI = 0 To 3
        If lblNomePreco(nI).Caption <> "" Or nI = 0 Then
          Call IsDataType(dtSingle, CSng(medPreco(nI).Text), sngPreco)
          
          If sngPreco <= 0 Then
            MsgBox "Preço deve ser maior do que zero.", vbExclamation, "Atenção"
            SelectAllText medPreco(nI), True
            Exit Sub
          End If
          
        End If
      Next nI
    End If
  End If
  
  
  '19/09/2005 - mpdea
  'Validação do Índice para cálculo do Preço de Entrada
  If g_blnIndicePrecoEntrada Then
    blnCancel = False
    Call mskIndicePrecoEntrada_Validate(blnCancel)
    If blnCancel Then Exit Sub
  End If
  
  
  '----------------------------------------------------------------------------
  'Data.............: 17/05/2006
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descrição........: Índice ICMS Retido - Base complementar
  blnCancel = False
  Call mskIndiceIcmsRetido_Validate(blnCancel)
  If blnCancel Then Exit Sub
  '----------------------------------------------------------------------------
  
  
  '---------------------------------------------------------------------------------
  'WEB - Loja Virtual
  '---------------------------------------------------------------------------------
  If chkWebIncluded.Value = vbChecked Then
    
    Sincroniza = True
    If medPreco(1).Text = 0# Then
     DisplayMsg "WEB - Produto não pode ter preço zero para a Loja Virtual. Não será sincronizado"
     Sincroniza = False
     medPreco(1).SetFocus
    End If
    
    'Verificações para o produto ser do tipo Web
    If Fracionado.Value = vbChecked Then
      DisplayMsg "WEB - Produto não pode ser 'Fracionado' para a Loja Virtual."
      Fracionado.SetFocus
      Exit Sub
    End If
    
    If Controla_Estoque.Value = vbUnchecked Then
      DisplayMsg "WEB - Produto deve ter 'Estoque controlado' para a Loja Virtual."
      tabItens.Tab = 0
      Controla_Estoque.SetFocus
      Exit Sub
    End If
    
    '18/06/2007 - Anderson
    'Retirada a validação da unidade de venda para que o cliente possa emitir notas fiscais com diversos produtos no Quick Web
    'Unidade de venda deve ser: g, K ou Kg
    'Select Case LCase(Unid.Text)
    '  Case "g", "k", "kg"
        'OK
    '  Case Else
        
        '03/06/2004 - mpdea
        'Verifica se possui unidade de venda padrão
    '    Select Case GetWebUnitSaleDefault
    '      Case "k", "g"
            'OK
    '      Case Else
    '        DisplayMsg "WEB - Produto deve ter unidade do tipo 'kg', 'k' ou 'g' para a Loja Virtual."
    '        tabItens.Tab = 0
    '        SelectAllText Unid, True
    '        Exit Sub
    '    End Select
        
    'End Select
    
    '18/06/2007 - Anderson
    'validação do peso do produto
    If txtPesoBruto.Text <= 0 Then
      If MsgBox("O peso bruto é essencial para o cálculo do frete do produto na loja virtual e neste momento não está informado. Deseja continuar assim mesmo?", vbQuestion + vbYesNo, "Quick Store") = vbNo Then
        Exit Sub
      End If
    End If
    
    If txtPesoLiquido.Text <= 0 Then
      If MsgBox("O peso líquido é essencial para o cálculo do frete do produto na loja virtual e neste momento não está informado. Deseja continuar assim mesmo?", vbQuestion + vbYesNo, "Quick Store") = vbNo Then
        Exit Sub
      End If
    End If
    
    If chkFabricante.Value = vbChecked And Fabricante.Text = "" Then
      DisplayMsg "WEB - Opção 'Fabricante' selecionada e não preenchida."
      tabItens.Tab = 0
      Fabricante.SetFocus
      Exit Sub
    End If
    
    If chkPesquisa123.Value = vbChecked And Nome_Pesq1.Caption = "" _
       And Nome_Pesq2.Caption = "" And Nome_Pesq3.Caption = "" Then
      
      DisplayMsg "WEB - Opção 'Pesquisa123' selecionada e não preenchida."
      tabItens.Tab = 0
      If Combo_Pesq1.Visible Then
        Combo_Pesq1.SetFocus
      ElseIf Combo_Pesq2.Visible Then
        Combo_Pesq2.SetFocus
      ElseIf Combo_Pesq3.Visible Then
        Combo_Pesq3.SetFocus
      End If
      Exit Sub
    End If
    
'    'Bônus
'    If Not IsDataType(dtLong, txtBonus.Text, lngRet) Then
'      DisplayMsg "WEB - Quantidade de Bônus incorreta."
'      tabItens.Tab = 6
'      SelectAllText txtBonus, True
'      Exit Sub
'    End If
'    txtBonus.Text = lngRet
    
    'Verifica se as tabela de venda está correta
    If Not gbCheckTabPreco(cboTableVenda.Text) Then
      DisplayMsg "WEB - Tabela de venda incorreta."
      tabItens.Tab = 5
      SelectAllText cboTableVenda, True
      Exit Sub
    End If
    
    'Promoção
    If gbCheckTabPreco(cboTablePromocional.Text) Then
    
      If Not IsDate(dtpDataInicial.Value) Then
        DisplayMsg "Web - Data início da Promoção incorreta."
        tabItens.Tab = 5
        SelectAllText dtpDataInicial, True
        Exit Sub
      End If
  
      If Not IsDate(dtpDataFinal.Value) Then
        DisplayMsg "Web - Data Final da Promoção incorreta."
        tabItens.Tab = 5
        SelectAllText dtpDataFinal, True
        Exit Sub
      End If
    
      If CDate(dtpDataFinal.Value) < CDate(dtpDataInicial.Value) Then
        DisplayMsg "Web - Data incorreta."
        tabItens.Tab = 5
        SelectAllText dtpDataFinal, True
        Exit Sub
      End If
    
    ElseIf cboTablePromocional.Text <> "" Then
      DisplayMsg "Web - Tabela de promoção incorreta."
      tabItens.Tab = 5
      SelectAllText cboTablePromocional, True
      Exit Sub
    End If

  End If
  '---------------------------------------------------------------------------------
  
  
  '---------------------------------------------------------------------------------
  '27/09/2004 - mpdea
  'Volumagem
  If txtVolumagem.Text = "" Then txtVolumagem.Text = "0"
  If Not IsDataType(dtInteger, txtVolumagem.Text, intRet) Then
    DisplayMsg "Outros - Valor para volumagem deve estar entre 0 e 9999."
    tabItens.Tab = 6
    SelectAllText txtVolumagem, True
    Exit Sub
  ElseIf intRet < 0 Then
    DisplayMsg "Outros - Valor para volumagem deve estar entre 0 e 9999."
    tabItens.Tab = 6
    SelectAllText txtVolumagem, True
    Exit Sub
  End If
  '---------------------------------------------------------------------------------
  
  
  Screen.MousePointer = vbHourglass
  
  Call StatusMsg("Gravando...")
  
  bGravando = True
  
  Call ws.BeginTrans: blnInTransaction = True
  
  With rsProdutos2
  
    If IsNull(Num_Registro) Then
      .AddNew
      .Fields("Código").Value = cboCodigo.Text
      .Fields("Código Ordenação").Value = Gera_Ordenação(cboCodigo.Text)
      strAux = "Registro inserido"
    Else
      .Edit
      strAux = "Registro alterado"
    End If
    
    Call MoveData
    
    If cboMotivoDesoneracao.Text <> "" Then
      .Fields("MotivoDesoneracaoICMS").Value = Left(cboMotivoDesoneracao.Text, 2)
    Else
      .Fields("MotivoDesoneracaoICMS").Value = "0"
    End If
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
  
  End With   'rsProdutos2
  
  If Preços.Enabled = True Then
    For nI = 0 To 3
      If lblNomePreco(nI).Caption <> "" Or nI = 0 Then
        If IsNull(medPreco(nI).Text) Then medPreco(nI).Text = "0"
        If medPreco(nI).Text = "" Then medPreco(nI).Text = "0"
        If Not IsNumeric(medPreco(nI).Text) Then medPreco(nI).Text = "0"
        
        nPos = InStr(lblNomePreco(nI).Caption, ":")
        If nPos > 0 Then
          sTableName = Mid(lblNomePreco(nI).Caption, 1, nPos - 1)
        Else
          sTableName = lblNomePreco(nI).Caption
        End If
        With rsPrecos
          .Index = "Tabela"
          .Seek "=", sTableName, cboCodigo.Text
          If .NoMatch Then
            .AddNew
            .Fields("Tabela").Value = sTableName
            .Fields("Produto").Value = cboCodigo.Text
            .Fields("Preço").Value = CSng(medPreco(nI).Text)
            .Fields("Data Alteração").Value = Format(Date, "dd/mm/yyyy")
            .Update
          Else
            .Edit
            .Fields("Preço").Value = CSng(medPreco(nI).Text)
            .Fields("Data Alteração").Value = Format(Date, "dd/mm/yyyy")
            .Update
          End If
        End With
'        Call Grava_Tabela(sTableName)
        Call CheckConfigTablePrice(sTableName)
      End If
    Next nI
  End If
  
  
  Call ws.CommitTrans: blnInTransaction = False
  
'----------------------------------------------------------------------------------------
' Joga dados do produto para o banco do GestoPDV por causa do PAF
'----------------------------------------------------------------------------------------
   Dim GestoBD As Database
   Dim ItemEstoque As Recordset
   Dim SitTrib As String
   Dim Preco As Double
   Dim ItemEstoqueAlmox As Recordset
   Dim Origem As Integer
   Dim VerificaPAF As Boolean
   Call frmParametros.VerificaPAF
   If frmParametros.VerificaPAF = True Then
 
     Preco = medPreco(1).Text
'''     Select Case Sit_Trib.Text
'''     Case "000", "100", "200"
'''       SitTrib = "T"
'''       Origem = Left(Sit_Trib.Text, 1)
'''     Case "010", "110", "210"
'''       SitTrib = "F"
'''       Origem = Left(Sit_Trib.Text, 1)
'''     Case "040", "140", "240"
'''       SitTrib = "I"
'''       Origem = Left(Sit_Trib.Text, 1)
'''     Case "041", "141", "241"
'''       SitTrib = "N"
'''       Origem = Left(Sit_Trib.Text, 1)
'''     Case Else
'''       SitTrib = "I"
'''       Origem = 0
'''     End Select
     Set rsParametros = db.OpenRecordset("Select * from [Parâmetros Filial] Where Filial = " & gnCodFilial & "")
     
     Dim fso As New FileSystemObject
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists(rsParametros("BancoPDV").Value & "\Gesto.mde") Then

     Set GestoBD = OpenDatabase(rsParametros("BancoPDV").Value & "\Gesto.mde", False, False)
     Set ItemEstoque = GestoBD.OpenRecordset("Select * from ItemEstoque Where CODIGO_FORNECEDOR =  '" & cboCodigo.Text & "';")
     Set ItemEstoqueAlmox = GestoBD.OpenRecordset("Select * from ItemEstoqueAlmox Where Codigo_Item =  '" & cboCodigo.Text & "';")
     If ItemEstoque.EOF Then
       ItemEstoque.AddNew
       ItemEstoque!CODIGO_FORNECEDOR = Left(cboCodigo.Text, 14)
       ItemEstoque!COD_FABRICANTE = 0
       ItemEstoque!UNIDADE_MEDIDA = Left(Unid.Text, 2)
'       ItemEstoque!icmsProduto = Perc_ICM.Text
       ItemEstoque!SituacaoTributaria = SitTrib
       ItemEstoque!QUANTIDADE_DISPONIVEL = 0
       ItemEstoque!DESCRICAO = Left(cboNome.Text, 50)
       ItemEstoque!PRECO_CONSUMIDOR = Preco
       ItemEstoque!PRECO_ATACADO = Preco
       ItemEstoque!PRECO_TECNICO = Preco
       ItemEstoque!PRECO_CONSUM_FABRICA = Preco
       ItemEstoque!preco_pdv = Preco
       ItemEstoque!GRU_GRUPO = 2
       ItemEstoque!CODIGO_LINHA = 99
'       ItemEstoque!CodSitTribA = Origem
       ItemEstoque.Update
       ItemEstoqueAlmox.AddNew
       ItemEstoqueAlmox!Codigo_Item = Left(cboCodigo.Text, 14)
       ItemEstoqueAlmox!Codigo_Almox = 1
       ItemEstoqueAlmox!Qtde_Disponivel = 0
       ItemEstoqueAlmox.Update
     Else
       ItemEstoque.Edit
       ItemEstoque!UNIDADE_MEDIDA = Left(Unid.Text, 2)
'       ItemEstoque!icmsProduto = Perc_ICM.Text
       ItemEstoque!SituacaoTributaria = SitTrib
       ItemEstoque!DESCRICAO = Left(cboNome.Text, 50)
       ItemEstoque!PRECO_CONSUMIDOR = Preco
       ItemEstoque!PRECO_ATACADO = Preco
       ItemEstoque!PRECO_TECNICO = Preco
       ItemEstoque!PRECO_CONSUM_FABRICA = Preco
       ItemEstoque!preco_pdv = Preco
'       ItemEstoque!CodSitTribA = Origem
       ItemEstoque.Update
        If ItemEstoqueAlmox.EOF Then
          ItemEstoqueAlmox.AddNew
          ItemEstoqueAlmox!Codigo_Item = Left(cboCodigo.Text, 14)
          ItemEstoqueAlmox!Codigo_Almox = 1
          ItemEstoqueAlmox!Qtde_Disponivel = 0
          ItemEstoqueAlmox.Update
        End If
      End If
   End If
   End If
'-----------------------------------------------------------------------------------------------

  If tabItens.TabEnabled(1) = False Then
    tabItens.TabEnabled(1) = True
    tabItens.TabEnabled(2) = True
    '20/04/2006 - mpdea
    'Corrigido aplicação da permissão de acesso a pasta de cálculos (custo)
    tabItens.TabEnabled(3) = (rsFuncionarios("Custo Produtos").Value = True)
    tabItens.TabEnabled(4) = True
    'tabItens.TabEnabled(5) = True
    tabItens.TabEnabled(6) = True
  End If
  
  ActiveBar1.Tools("miOpCheckGrid").Enabled = gbGrade
  O_Produto_Grade.Enabled = gbGrade

  O_Produto_Edição.Enabled = gbEdicao
  ActiveBar1.Tools("miOpShowEdicoes").Enabled = gbEdicao
  
  cboCodigo.Enabled = True
  
  ActiveBar1.Tools("miOpConsultaEstoque").Enabled = True
  ActiveBar1.Tools("miOpInicializaEstoque").Enabled = True
  
  DBEngine.Idle dbRefreshCache
  
  Call StatusMsg(strAux & " com sucesso.")
  
  tabItens.Tab = 0
  
'  Call StatusMsg("")
  Screen.MousePointer = vbDefault

  bGravando = False
  
 
  Exit Sub
  
ErrHandler:
  If blnInTransaction Then ws.Rollback
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(201)
  gsMsg = "Erro na Atualização de Produto."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  bGravando = False
  Exit Sub

End Sub

Public Sub MoveData()
    
  With rsProdutos2
  
    If O_Produto_Normal.Value = True Then
      .Fields("Tipo").Value = "N"
    Else
      If O_Produto_Grade.Value = True Then
        .Fields("Tipo").Value = "G"
      Else
        If O_Produto_Edição.Value = True Then
          .Fields("Tipo").Value = "E"
        End If
      End If
    End If
     
'    DoEvents
    
    .Fields("Nome").Value = cboNome.Text
    .Fields("Classe").Value = Val(Combo_Classe.Text)
    .Fields("Sub Classe").Value = Val(Combo_Sub_Classe.Text)
    .Fields("Moeda").Value = Val(Combo_Moeda.Text)
    
'    DoEvents
    
    .Fields("Fracionado").Value = (Fracionado.Value = 1)
    .Fields("QtdeCasasDecimais").Value = txtCasasDecimais.Text
    .Fields("Estoque").Value = (Controla_Estoque.Value = 1)
    .Fields("Desativado").Value = (Desativado.Value = 1)
    .Fields("Não Incluir na Tabela").Value = (Não_Lista.Value = 1)
    
'    DoEvents
    
    .Fields("Unidade Venda").Value = Unid.Text
    .Fields("Código do Fornecedor").Value = LTrim(RTrim(Cod_Fornecedor.Text))
    .Fields("Estoque Ideal").Value = CSng(gsHandleNull(Estoque_Ideal.Text))
    '30/10/2007 - Anderson
    'Campo utilizado para gerar o relatório de produtos a comprar
    'Solicitante: Kings Cross
    '.Fields("EspacoFisicoTotal").Value = CDbl(gsHandleNull(txtEspacoFisicoTotal.Text))
    .Fields("Estoque Mínimo").Value = CSng(gsHandleNull(Estoque_Min.Text))
    .Fields("Localização").Value = Locali.Text
    .Fields("Desconto").Value = CSng(Desconto.Text)
    .Fields("Custo Médio").Value = CSng(gsHandleNull(Custo_Médio.Text))
    ' .Fields("Desconto Máximo").Value = CDbl(Desconto_Max.Text)
    
    
    '29/12/2003 - mpdea
    'Não permitir desconto neste item
    .Fields("DontAllowDesc").Value = chkDontAllowDesc.Value = vbChecked
    
    '20/08/2004 - Daniel
    'Usar Descrição Adicional
    If chkUsarDescrAdic.Value Then
      .Fields("UsaDescrAdic").Value = True
    Else
      .Fields("UsaDescrAdic").Value = False
    End If
    
'    DoEvents
    
    Call HandleNullsInTextBoxes
     
'    DoEvents
    
    .Fields("Pesquisa 1").Value = Combo_Pesq1.Text
    .Fields("Pesquisa 2").Value = Combo_Pesq2.Text
    .Fields("Pesquisa 3").Value = Combo_Pesq3.Text
    
    .Fields("Nome Nota").Value = LTrim(RTrim(Nome_Nota.Text))
    .Fields("OBS").Value = Obs.Text
    .Fields("Fabricante").Value = Fabricante.Text
    
    .Fields("PesoLiquido").Value = CSng(gsHandleNull(txtPesoLiquido.Text))
    .Fields("PesoBruto").Value = CSng(gsHandleNull(txtPesoBruto.Text))
    '01/11/2004 - Daniel
    If IsNumeric(txtCubagem.Text) Then
      .Fields("Cubagem").Value = Format(txtCubagem.Text, FORMAT_VALUE)
    Else
      .Fields("Cubagem").Value = 0
    End If
    
'    DoEvents
    
    
    'IMPOSTOS
    .Fields("Tipo ICM").Value = "N"

    .Fields("Foto").Value = Nome_Foto.Text & ""
     
    '--------------------------------------------------------------------
    
'    DoEvents
    
    'Comissão
    .Fields("Comissão Sobrepõe").Value = (Comissão_Sobre.Value = 1)
    .Fields("Comissão").Value = CSng(Comissão.Text)
    .Fields("Data Alteração").Value = Format(Date, "dd/mm/yyyy")
    
    '23/08/2004 - Daniel
    'Campo IndiceFinanceiro
    If IsNumeric(mskIndice.Text) Then
      .Fields("IndiceFinanceiro").Value = Format((CDbl(mskIndice.Text)), "##,###,###,##0.000")
    End If
    
    
    '19/09/2005 - mpdea
    'Índice para cálculo do Preço de Entrada
    If g_blnIndicePrecoEntrada Then
      If IsNumeric(mskIndicePrecoEntrada.Text) Then
        .Fields("IndicePrecoEntrada").Value = Format((CDbl(mskIndicePrecoEntrada.Text)), "#0.00000")
      Else
        .Fields("IndicePrecoEntrada").Value = 1
      End If
    End If
    
    
    '----------------------------------------------------------------------------
    'Data.............: 17/05/2006
    'Desenvolvedor....: mpdea
    'Solicitante......: Pavinato
    'Descrição........: Índice ICMS Retido - Base complementar
    If IsNumeric(mskIndiceIcmsRetido.Text) Then
      .Fields("IndiceIcmsRetido").Value = Format((CDbl(mskIndiceIcmsRetido.Text)), "#0.00000")
    Else
      .Fields("IndiceIcmsRetido").Value = 1
    End If
    '----------------------------------------------------------------------------
    
    
    'Parte de custo
    .Fields("Custo Preço Valor").Value = CDbl(C_Compra_Valor.Text)
    .Fields("Custo Desconto Perc").Value = CDbl(C_Compra_Desc_P.Text)
    .Fields("Custo Desconto Valor").Value = CDbl(C_Compra_Desc_V.Text)
    If Fixa_Desc_C_Perc.Value = 1 Then
      .Fields("Custo Desconto Fixo").Value = "P"
    Else
      .Fields("Custo Desconto Fixo").Value = "V"
    End If
    
'    DoEvents
    
    .Fields("Custo Frete Perc").Value = CDbl(C_Compra_Frete_P.Text)
    .Fields("Custo Frete Valor").Value = CDbl(C_Compra_Frete_V.Text)
    If Fixa_Frete_C_Perc.Value = 1 Then
      .Fields("Custo Frete Fixo").Value = "P"
    Else
      .Fields("Custo Frete Fixo").Value = "V"
    End If
    
'    DoEvents
    
    .Fields("Custo ICM Compra Perc").Value = CDbl(C_Compra_ICM_P.Text)
    .Fields("Custo ICM Compra Valor").Value = CDbl(C_Compra_ICM_V.Caption)
    If Fixa_ICM_C_Perc.Value = 1 Then
      .Fields("Custo ICM Compra Fixo").Value = "P"
    Else
      .Fields("Custo ICM Compra Fixo").Value = "V"
    End If
    
    .Fields("Custo IPI Compra Perc").Value = CDbl(C_Compra_IPI_P.Text)
    .Fields("Custo IPI Compra Valor").Value = CDbl(C_Compra_IPI_V.Caption)
    If Fixa_IPI_C_Perc.Value = 1 Then
      .Fields("Custo IPI Compra Fixo").Value = "P"
    Else
      .Fields("Custo IPI Compra Fixo").Value = "V"
    End If
    
'    DoEvents
    
    .Fields("Custo Custo Finan Perc").Value = CDbl(C_Compra_Finan_P.Text)
    .Fields("Custo Custo Finan Valor").Value = CDbl(C_Compra_Finan_V.Text)
    If Fixa_Custo_C_Perc.Value = 1 Then
      .Fields("Custo Custo Finan Fixo").Value = "P"
    Else
      .Fields("Custo Custo Finan Fixo").Value = "V"
    End If
    
    .Fields("Custo Outros Compra Perc").Value = CDbl(C_Compra_Outros_P.Text)
    .Fields("Custo Outros Compra Valor").Value = CDbl(C_Compra_Outros_V.Text)
    If Fixa_Outros_C_Perc.Value = 1 Then
      .Fields("Custo Outros Compra Fixo").Value = "P"
    Else
      .Fields("Custo Outros Compra Fixo").Value = "V"
    End If
    
'    DoEvents
    
    .Fields("Custo Perc Compra Sem").Value = CDbl(C_Compra_Sem_Nota.Text)
    .Fields("Custo Custo Calculado").Value = CDbl(C_Compra_Custo.Caption)
    
    .Fields("Custo Preço Venda").Value = CDbl(C_Venda_Valor.Text)
    .Fields("Custo ICM Venda Perc").Value = CDbl(C_Venda_ICM_P.Text)
    .Fields("Custo ICM Venda Valor").Value = CDbl(C_Venda_ICM_V.Caption)
    If Fixa_ICM_V_Perc.Value = 1 Then
      .Fields("Custo ICM Venda Fixo").Value = "P"
    Else
      .Fields("Custo ICM Venda Fixo").Value = "V"
    End If
    
'    DoEvents
    
    .Fields("Custo IPI Venda Perc").Value = CDbl(C_Venda_IPI_P.Text)
    .Fields("Custo IPI Venda Valor").Value = CDbl(C_Venda_IPI_V.Caption)
    If Fixa_IPI_V_Perc.Value = 1 Then
      .Fields("Custo IPI Venda Fixo").Value = "P"
    Else
      .Fields("Custo IPI Venda Fixo").Value = "V"
    End If
    
    .Fields("Custo Impostos Perc").Value = CDbl(C_Venda_Imp_P.Text)
    .Fields("Custo Impostos Valor").Value = CDbl(C_Venda_Imp_V.Text)
    If Fixa_Imp_V_Perc.Value = 1 Then
      .Fields("Custo Impostos Fixo").Value = "P"
    Else
      .Fields("Custo Impostos Fixo").Value = "V"
    End If
    
'    DoEvents
    
    .Fields("Custo Outros Venda Perc").Value = CDbl(C_Venda_Outros_P.Text)
    .Fields("Custo Outros Venda Valor").Value = CDbl(C_Venda_Outros_V.Text)
    If Fixa_Outros_V_Perc.Value = 1 Then
      .Fields("Custo Outros Venda Fixo").Value = "P"
    Else
      .Fields("Custo Outros Venda Fixo").Value = "V"
    End If
    
    .Fields("Custo Perc Venda Sem").Value = CDbl(C_Venda_Sem_Nota.Text)
    .Fields("Custo Lucro Perc").Value = CDbl(C_Venda_Margem_P.Text)
    .Fields("Custo Lucro Valor").Value = CDbl(C_Venda_Margem_V.Text)
    
'    DoEvents
    
    If O_Manter_Venda.Value = True Then
      .Fields("Custo Manter").Value = "P"
    Else
      If O_Manter_Lucro_Val.Value = True Then
        .Fields("Custo Manter").Value = "L"
      Else
        .Fields("Custo Manter").Value = "U"
      End If
    End If
     
    gsCodProduto = .Fields("Código").Value
    gsTipoProduto = .Fields("Tipo").Value
    
    
    '---------------------------------------------------------------------------------
    '08/05/2002 - mpdea
    '             Incluído gravação dos dados WEB
    '
    '09/05/2002 - mpdea
    '             Incluído verificação para exclusão do produto na Loja Virtual
    '<<-------------------------------------------------------------------------------
    .Fields("WebIncluded").Value = chkWebIncluded.Value = vbChecked
    
    'Produto tipo WEB
    If chkWebIncluded.Value = vbChecked Then
      If Sincroniza = True Then
        .Fields("WebSynchronize").Value = True
      Else
        .Fields("WebSynchronize").Value = False
      End If
      .Fields("WebAttribFabricante").Value = chkFabricante.Value = vbChecked
      .Fields("WebAttribPesquisa123").Value = chkPesquisa123.Value = vbChecked
      .Fields("WebSaleTablePrice").Value = cboTableVenda.Text
      If gbCheckTabPreco(cboTablePromocional.Text) Then
        .Fields("WebOfferTablePrice").Value = cboTablePromocional.Text
        .Fields("WebOfferDateStart").Value = dtpDataInicial.Value
        .Fields("WebOfferDateEnd").Value = dtpDataFinal.Value
      Else
        .Fields("WebOfferTablePrice").Value = ""
        .Fields("WebOfferDateStart").Value = Null
        .Fields("WebOfferDateEnd").Value = Null
      End If
'      .Fields("WebBonus").Value = CLng(txtBonus.Text)
    'Produto tipo não WEB
    Else
      'Verifica se o produto já foi do tipo WEB
      Select Case UCase(.Fields("WebLastOp").Value & "")
        Case "U", "I" 'U = update, I = Insert
          'Caso tenha sido, inclui o mesmo na lista de produtos a excluir
          'da Loja Virtual
          Call db.Execute("INSERT INTO WEB_ProdutosExcluir (Codigo) VALUES ('" & _
                          .Fields("Código").Value & "')", dbFailOnError)
          'e altera última operação WEB para vazio
          .Fields("WebLastOp").Value = ""
      End Select
    End If
    
    '27/09/2004 - mpdea
    'Volumagem
    rsProdutos2.Fields("Volumagem").Value = CInt(txtVolumagem.Text)
    '10/11/2004 - Daniel
    'Campos: Lote e DataValidade
    If Len(txtLote.Text) > 0 Then rsProdutos2.Fields("Lote").Value = Trim(txtLote.Text)
    If IsDate(mskDataValidade.Text) Then rsProdutos2.Fields("DataValidade").Value = CDate(mskDataValidade.Text)
    
    '17/05/2005 - Daniel
    'Case: Miss Nuvem
    'Tratamento para os campos: ImprimirUmaEtiq e ImprimirPrecoEtiq
    If chkImprimirUmaEtiq.Value = vbChecked Then
      rsProdutos2.Fields("ImprimirUmaEtiq").Value = True
    Else
      rsProdutos2.Fields("ImprimirUmaEtiq").Value = False
    End If
    '
    If chkImprimirPrecoEtiq.Value = vbChecked Then
      rsProdutos2.Fields("ImprimirPrecoEtiq").Value = True
    Else
      rsProdutos2.Fields("ImprimirPrecoEtiq").Value = False
    End If
    '
    '04/07/2005 - Daniel
    'Case.....: Zue
    'Projeto..: Informações do Custo de Produto
    '
    'ConsumoDeTecido
    If IsNumeric(txtCustoProduto(0).Text) Then
       rsProdutos2.Fields("ConsumoDeTecido").Value = Format(txtCustoProduto(0).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("ConsumoDeTecido").Value = 0
    End If
    'PrecoDoMetroLinear
    If IsNumeric(txtCustoProduto(1).Text) Then
       rsProdutos2.Fields("PrecoDoMetroLinear").Value = Format(txtCustoProduto(1).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("PrecoDoMetroLinear").Value = 0
    End If
    'CustoDoTecido
    If IsNumeric(txtCustoProduto(2).Text) Then
       rsProdutos2.Fields("CustoDoTecido").Value = Format(txtCustoProduto(2).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("CustoDoTecido").Value = 0
    End If
    'VlMaoDeObraFaccao
    If IsNumeric(txtCustoProduto(3).Text) Then
       rsProdutos2.Fields("VlMaoDeObraFaccao").Value = Format(txtCustoProduto(3).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("VlMaoDeObraFaccao").Value = 0
    End If
    'VlLavanderia
    If IsNumeric(txtCustoProduto(4).Text) Then
       rsProdutos2.Fields("VlLavanderia").Value = Format(txtCustoProduto(4).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("VlLavanderia").Value = 0
    End If
    'VlBordado
    If IsNumeric(txtCustoProduto(5).Text) Then
       rsProdutos2.Fields("VlBordado").Value = Format(txtCustoProduto(5).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("VlBordado").Value = 0
    End If
    'VlEstamparia
    If IsNumeric(txtCustoProduto(6).Text) Then
       rsProdutos2.Fields("VlEstamparia").Value = Format(txtCustoProduto(6).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("VlEstamparia").Value = 0
    End If
    'VlAviamentos
    If IsNumeric(txtCustoProduto(7).Text) Then
       rsProdutos2.Fields("VlAviamentos").Value = Format(txtCustoProduto(7).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("VlAviamentos").Value = 0
    End If
    'OutrosCustos
    If IsNumeric(txtCustoProduto(8).Text) Then
       rsProdutos2.Fields("OutrosCustos").Value = Format(txtCustoProduto(8).Text, FORMAT_VALUE)
    Else
      rsProdutos2.Fields("OutrosCustos").Value = 0
    End If
    
    '26/06/2013-Alexandre Afornali
    'Lei de olho nos impostos
'    Dim rsNCM As Recordset
'    Set rsNCM = db.OpenRecordset("AliquotasNCM")
'    If (Left$(Sit_Trib, 1) = "0") Then 'Nacional
'      While Not rsNCM.EOF
'        If (rsNCM("Codigo").Value = cboCodigoNBM.Text) Then
'          rsProdutos2("AliqNCM") = rsNCM("AliqNacional")
'          rsNCM.MoveLast
'          rsNCM.MoveNext
'        Else
'          rsNCM.MoveNext
'        End If
'      Wend
'    Else 'Importado
'      While Not rsNCM.EOF
'        If (rsNCM("Codigo").Value = cboCodigoNBM.Text) Then
'          rsProdutos2("AliqNCM") = rsNCM("AliqImportacao")
'          rsNCM.MoveLast
'          rsNCM.MoveNext
'        Else
'          rsNCM.MoveNext
'        End If
'      Wend
'    End If
    
    '17/04/2007 - Anderson
    'Inclusão do campo para dividir preco da etiqueta - Exigencia do procon
    rsProdutos2.Fields("DividirPrecoEtiqueta").Value = txtDividirPrecoEtiqueta.Text

  End With
  
End Sub

Private Sub HandleNullsInTextBoxes()
  
  Combo_Pesq1.Text = gsHandleNull(Combo_Pesq1.Text)
  Combo_Pesq2.Text = gsHandleNull(Combo_Pesq2.Text)
  Combo_Pesq3.Text = gsHandleNull(Combo_Pesq3.Text)
 
  C_Compra_Valor.Text = gsHandleNull(C_Compra_Valor.Text)
  C_Compra_Desc_P.Text = gsHandleNull(C_Compra_Desc_P.Text)
  C_Compra_Desc_V.Text = gsHandleNull(C_Compra_Desc_V.Text)
  C_Compra_Frete_P.Text = gsHandleNull(C_Compra_Frete_P.Text)
  C_Compra_Frete_V.Text = gsHandleNull(C_Compra_Frete_V.Text)
  C_Compra_IPI_P.Text = gsHandleNull(C_Compra_IPI_P.Text)
  C_Compra_IPI_V.Caption = gsHandleNull(C_Compra_IPI_V.Caption)
  C_Compra_ICM_P.Text = gsHandleNull(C_Compra_ICM_P.Text)
  C_Compra_ICM_V.Caption = gsHandleNull(C_Compra_ICM_V.Caption)
  C_Compra_Finan_P.Text = gsHandleNull(C_Compra_Finan_P.Text)
  C_Compra_Finan_V.Text = gsHandleNull(C_Compra_Finan_V.Text)
  C_Compra_Outros_P.Text = gsHandleNull(C_Compra_Outros_P.Text)
  C_Compra_Outros_V.Text = gsHandleNull(C_Compra_Outros_V.Text)
  C_Compra_Sem_Nota.Text = gsHandleNull(C_Compra_Sem_Nota.Text)
  C_Compra_Custo.Caption = gsHandleNull(C_Compra_Custo.Caption)
  
  C_Venda_Valor.Text = gsHandleNull(C_Venda_Valor.Text)
  C_Venda_ICM_P.Text = gsHandleNull(C_Venda_ICM_P.Text)
  C_Venda_ICM_V.Caption = gsHandleNull(C_Venda_ICM_V.Caption)
  C_Venda_IPI_P.Text = gsHandleNull(C_Venda_IPI_P.Text)
  C_Venda_IPI_V.Caption = gsHandleNull(C_Venda_IPI_V.Caption)
  C_Venda_Imp_P.Text = gsHandleNull(C_Venda_Imp_P.Text)
  C_Venda_Imp_V.Text = gsHandleNull(C_Venda_Imp_V.Text)
  C_Venda_Outros_P.Text = gsHandleNull(C_Venda_Outros_P.Text)
  C_Venda_Outros_V.Text = gsHandleNull(C_Venda_Outros_V.Text)
  C_Venda_Sem_Nota.Text = gsHandleNull(C_Venda_Sem_Nota.Text)
  C_Venda_Margem_P.Text = gsHandleNull(C_Venda_Margem_P.Text)
  C_Venda_Margem_V.Text = gsHandleNull(C_Venda_Margem_V.Text)

End Sub

Private Sub InicializaEstoque()

  Dim Mes_Atual As Integer
  Dim Ano_Atual As Integer
  Dim Aux_Str As String
  Dim nValEstoque As Single

  Call StatusMsg("")

  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(210)
    gsMsg = "Encontre ou grave um produto antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  Aux_Str = Format$(Data_Atual, "dd/mm/yyyy")
  Ano_Atual = Val(Right(Aux_Str, 4))
  Mes_Atual = Val(Mid(Aux_Str, 4, 2))

  Call StatusMsg("")

  If rsProdutos2("Tipo") = "G" Then
    gsTitle = LoadResString(201)
    gsMsg = "Produtos que usam Grade devem ter o estoque atualizado através de ajuste de entrada."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If rsProdutos2("Tipo") = "E" Then
    gsTitle = LoadResString(201)
    gsMsg = "Produtos que usam Edição devem ter o estoque atualizado através de ajuste de entrada."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  rsEstoque.Index = "Data"
  rsEstoque.Seek ">", gnCodFilial, cboCodigo.Text, 0, 0
  If Not rsEstoque.NoMatch Then
    If rsEstoque("Produto") = cboCodigo.Text Then
      Beep
      gsTitle = LoadResString(201)
      gsMsg = "Estoque deste produto já foi inicializado."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Exit Sub
    End If
  End If
  
  nValEstoque = frmInicializaEstoque.gnInitializeEstoque(rsProdutos2("Código"))
  
  If nValEstoque > 0 Then
    With rsEstoque
      .AddNew
      .Fields("Filial") = gnCodFilial
      .Fields("Data") = Data_Atual
      .Fields("Produto") = rsProdutos2("Código")
      .Fields("Tamanho") = 0
      .Fields("Cor") = 0
      .Fields("Classe") = rsProdutos2("Classe")
      .Fields("Sub Classe") = rsProdutos2("Sub Classe")
      .Fields("Ajuste Entra") = nValEstoque
      .Fields("Estoque Final") = nValEstoque
      .Update
    End With
    Rem Acerta Estoque Final
    Call Grava_Estoque_Final(gnCodFilial, rsProdutos2("Código"), 0, 0, 0, nValEstoque, CDate(Data_Atual))
    DisplayMsg "Estoque inicializado."
  Else
    DisplayMsg "Estoque não alterado."
  End If
  
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsProdutos2
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  With rsProdutos2
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  With rsProdutos2
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With
End Sub

Private Sub MoveNext()
  On Error Resume Next
  With rsProdutos2
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Private Sub C_Compra_Desc_P_GotFocus()
  C_Compra_Desc_P.SelStart = 0
  C_Compra_Desc_P.SelLength = Len(C_Compra_Desc_P)
End Sub

Private Sub C_Compra_Desc_P_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Desc_P_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_Desc_V_GotFocus()
  C_Compra_Desc_V.SelStart = 0
  C_Compra_Desc_V.SelLength = Len(C_Compra_Desc_V)
End Sub

Private Sub C_Compra_Desc_V_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Desc_V_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_Finan_P_GotFocus()
  C_Compra_Finan_P.SelStart = 0
  C_Compra_Finan_P.SelLength = Len(C_Compra_Finan_P)
End Sub

Private Sub C_Compra_Finan_P_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Finan_P_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_Finan_V_GotFocus()
  C_Compra_Finan_V.SelStart = 0
  C_Compra_Finan_V.SelLength = Len(C_Compra_Finan_V)
End Sub

Private Sub C_Compra_Finan_V_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Finan_V_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_Frete_P_GotFocus()
  C_Compra_Frete_P.SelStart = 0
  C_Compra_Frete_P.SelLength = Len(C_Compra_Frete_P)
End Sub

Private Sub C_Compra_Frete_P_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Frete_P_LostFocus()
Calcula_Custo1
End Sub

Private Sub C_Compra_Frete_V_GotFocus()
  C_Compra_Frete_V.SelStart = 0
  C_Compra_Frete_V.SelLength = Len(C_Compra_Frete_V)
End Sub

Private Sub C_Compra_Frete_V_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Frete_V_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_ICM_P_GotFocus()
  C_Compra_ICM_P.SelStart = 0
  C_Compra_ICM_P.SelLength = Len(C_Compra_ICM_P)
End Sub

Private Sub C_Compra_ICM_P_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_ICM_P_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_IPI_P_GotFocus()
  C_Compra_IPI_P.SelStart = 0
  C_Compra_IPI_P.SelLength = Len(C_Compra_IPI_P)
End Sub

Private Sub C_Compra_IPI_P_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_IPI_P_LostFocus()
Calcula_Custo1
End Sub

Private Sub C_Compra_Outros_P_GotFocus()
  C_Compra_Outros_P.SelStart = 0
  C_Compra_Outros_P.SelLength = Len(C_Compra_Outros_P)
End Sub

Private Sub C_Compra_Outros_P_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Outros_P_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_Outros_V_GotFocus()
  C_Compra_Outros_V.SelStart = 0
  C_Compra_Outros_V.SelLength = Len(C_Compra_Outros_V)
End Sub

Private Sub C_Compra_Outros_V_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Outros_V_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_Sem_Nota_GotFocus()
  C_Compra_Sem_Nota.SelStart = 0
  C_Compra_Sem_Nota.SelLength = Len(C_Compra_Sem_Nota)
End Sub

Private Sub C_Compra_Sem_Nota_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Sem_Nota_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Compra_Valor_GotFocus()
  C_Compra_Valor.SelStart = 0
  C_Compra_Valor.SelLength = Len(C_Compra_Valor)
End Sub

Private Sub C_Compra_Valor_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub C_Compra_Valor_LostFocus()
  Calcula_Custo1
End Sub

Private Sub C_Venda_ICM_P_GotFocus()
  C_Venda_ICM_P.SelStart = 0
  C_Venda_ICM_P.SelLength = Len(C_Venda_ICM_P)
End Sub

Private Sub C_Venda_Imp_P_GotFocus()
  C_Venda_Imp_P.SelStart = 0
  C_Venda_Imp_P.SelLength = Len(C_Venda_Imp_P)
End Sub

Private Sub C_Venda_Imp_V_GotFocus()
  C_Venda_Imp_V.SelStart = 0
  C_Venda_Imp_V.SelLength = Len(C_Venda_Imp_V)
End Sub

Private Sub C_Venda_IPI_P_GotFocus()
  C_Venda_IPI_P.SelStart = 0
  C_Venda_IPI_P.SelLength = Len(C_Venda_IPI_P)
End Sub

Private Sub C_Venda_Margem_P_GotFocus()
  C_Venda_Margem_P.SelStart = 0
  C_Venda_Margem_P.SelLength = Len(C_Venda_Margem_P)
End Sub

Private Sub C_Venda_Margem_V_GotFocus()
  C_Venda_Margem_V.SelStart = 0
  C_Venda_Margem_V.SelLength = Len(C_Venda_Margem_V)
End Sub

Private Sub C_Venda_Outros_P_GotFocus()
  C_Venda_Outros_P.SelStart = 0
  C_Venda_Outros_P.SelLength = Len(C_Venda_Outros_P)
End Sub

Private Sub C_Venda_Outros_V_GotFocus()
  C_Venda_Outros_V.SelStart = 0
  C_Venda_Outros_V.SelLength = Len(C_Venda_Outros_V)
End Sub

Private Sub C_Venda_Sem_Nota_GotFocus()
  C_Venda_Sem_Nota.SelStart = 0
  C_Venda_Sem_Nota.SelLength = Len(C_Venda_Sem_Nota)
End Sub

Private Sub C_Venda_Valor_GotFocus()
  C_Venda_Valor.SelStart = 0
  C_Venda_Valor.SelLength = Len(C_Venda_Valor)
End Sub

Private Sub C_Venda_Valor_LostFocus()
  Call Calcula_Venda
End Sub

Private Sub cboCodigo_Click()
'  If Not rsProdutos2.EOF Then
'    If Len(Trim(cboCodigo.Text)) > 0 Then
'      rsProdutos2.FindFirst "Código = '" & cboCodigo.Text & "'"
'      If Not rsProdutos2.NoMatch Then
'        Call ShowRecord
'      End If
'    End If
'  End If
End Sub

Private Sub cboCodigo_CloseUp()
  If Not rsProdutos2.EOF Then
    If Len(Trim(cboCodigo.Text)) > 0 Then
      rsProdutos2.FindFirst "Código = '" & cboCodigo.Text & "'"
      If Not rsProdutos2.NoMatch Then
        Call ShowRecord
      End If
    End If
  End If
  gbComboDropped = False
End Sub

Private Sub cboCodigo_DropDown()
  gbComboDropped = True
End Sub

Private Sub cboCodigo_GotFocus()
  cboCodigo.SelStart = 0
  cboCodigo.SelLength = Len(cboCodigo.Text)
End Sub

Private Sub cboCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
'  If (Shift And vbAltMask) > 0 Then
'    If KeyCode = 40 Or KeyCode = 38 Then
'      Call cboCodigo_DropDown
'    End If
'    Exit Sub
'  Else
'    If Not gbComboDropped Then
'      Call Form_KeyDown(KeyCode, Shift)
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyF5 Then
'        Call FindNextProdutoCode
'      Else
'        If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then
'          Call HandleKeyDown(KeyCode, Shift)
'        End If
'      End If
'    End If
'  End If

 

End Sub

Private Sub cboCodigo_KeyPress(KeyAscii As Integer)
  If Len(cboCodigo.Text) >= rsProdutos2("Código").Size Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If

'' colocada a clausula - And Keyascii <> 13
  If Not IsNull(Num_Registro) And KeyAscii <> 13 Then
    Beep
    KeyAscii = 0
    Exit Sub
  End If
  
  '26/05/2004 - Daniel
  'Tratamento para 0 'zero' a esquerda
  If Not gbZeroEsquerda Then
    'Verifica se o primeiro caracter é zero
    If cboCodigo.Text = "" And Chr(KeyAscii) = "0" Then
    ''só tinha a atribuição - keyascii = 0
      If KeyAscii <> vbKeyReturn Or _
         KeyAscii <> vbKeyBack Or _
         KeyAscii <> vbKeyTab Then KeyAscii = 0
    End If
  End If
  
Exit Sub


  If Glob_Cod_Alfa = False Then
      KeyAscii = Verifica_Tecla_Integer(KeyAscii)
  Else
      ''Ao passar pela transformação do ALFA, anula a tecla ENTER
      KeyAscii = Verifica_Tecla_Código(KeyAscii)
  End If
  
End Sub

Public Sub cboCodigo_LostFocus()
  Dim nI As Integer
  Dim sCh As String
    
  gbComboDropped = False
  cboCodigo.Text = UCase(Trim(cboCodigo.Text & ""))
  
  If Len(cboCodigo.Text) = 0 Then
    Exit Sub
  End If
  
 
  
  
  
'  For nI = 1 To Len(cboCodigo.Text)
'    sCh = Mid(cboCodigo.Text, nI, 1)
'    If Not ((sCh >= "A" And sCh <= "Z") Or (sCh >= "0" And sCh <= "9") Or sCh = "-") Then
'      DisplayMsg "Caracter inválido em código de produto. Verifique."
'      'cboCodigo.SetFocus
'      Exit Sub
'    End If
'  Next nI

 
  If IsNumeric(cboCodigo.Text) Then
'    cboCodigo.Text = Format(cboCodigo.Text, String(Len(cboCodigo.Text) - 1, "#") & "0")
  End If
  If Not rsProdutos2.EOF Then
    If Len(Trim(cboCodigo.Text)) > 0 Then
      If Mid(cboCodigo.Text, Len(cboCodigo.Text), 1) = "*" Then
        rsProdutos2.FindFirst "Código Like '*" & cboCodigo.Text & "'"
        If Not rsProdutos2.NoMatch Then
          Call ShowRecord
          Exit Sub
        End If
      End If
      rsProdutos2.FindFirst "Código = '" & cboCodigo.Text & "'"
      If Not rsProdutos2.NoMatch Then
        Call ShowRecord
      End If
    End If
  End If
End Sub

Private Sub cboCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'  gbComboDropped = True
End Sub


Private Sub cboNome_Change()
  '''Me.Caption = " Produtos - " & cboNome.Text
End Sub

Private Sub cboNome_CloseUp()
  If Not rsProdutos2.EOF Then
    rsProdutos2.FindFirst "Código = '" & cboNome.Columns(1).Text & "'"
    If Not rsProdutos2.NoMatch Then
      Call ShowRecord
    End If
  End If
End Sub

Private Sub cboNome_DropDown()
  gbComboDropped = True
End Sub

Private Sub cboNome_KeyPress(KeyAscii As Integer)
  If Len(cboNome.Text) >= rsProdutos2("Nome").Size Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub cboNome_LostFocus()
  gbComboDropped = False
  If Not rsProdutos2.EOF Then
    If Len(Trim(cboNome.Text)) > 0 Then
      If Mid(cboNome.Text, Len(cboNome.Text), 1) = "*" Then
        rsProdutos2.FindFirst "Nome Like '*" & cboNome.Text & "'"
        If Not rsProdutos2.NoMatch Then
          Call ShowRecord
        End If
      End If
    End If
  End If
  If IsNull(Nome_Nota.Text) Or (Nome_Nota.Text) = "" Then
    Nome_Nota.Text = RTrim(LTrim(cboNome.Text))
  End If
End Sub

Private Sub chkWebIncluded_Click()
  tabItens.TabEnabled(5) = chkWebIncluded.Value = vbChecked
  '18/06/2007 - Anderson
  'Alteração os labels de pesagem para serem informados em Kilogramas.
  If chkWebIncluded Then
    Label31.Caption = "Peso Bruto (Kg)"
    Label44.Caption = "Peso Líquido (Kg)"
    txtPesoBruto.ToolTipText = "Use a unidade de medida em kilograma para informar o peso bruto."
    txtPesoLiquido.ToolTipText = "Use a unidade de medida em kilograma para informar o peso líquido."
  Else
    Label31.Caption = "Peso Bruto"
    Label44.Caption = "Peso Líquido"
    txtPesoBruto.ToolTipText = ""
    txtPesoLiquido.ToolTipText = ""
  End If
End Sub

Private Sub cmd_cesta_de_produtos_Click()
  Dim formulario As Form
  
  Set formulario = New frmProdutosCesta
  
  formulario.CodigoProduto = cboCodigo.Text
  formulario.nomeProduto = cboNome
  
  formulario.Show
End Sub

Private Sub cmd_clonarCaract_produto_Click()
  Dim formulario As Form

  Set formulario = New frmProdutosCFOP_Clone

  formulario.CodigoProduto = cboCodigo.Text
  formulario.nomeProduto = cboNome

  If Not IsNull(rsProdutos2("CodigoNBM").Value) Then
      formulario.CodigoNCM = rsProdutos2("CodigoNBM").Value
  Else
      formulario.CodigoNCM = ""
  End If

  If Not IsNull(rsProdutos2("Situação Tributária").Value) Then
      formulario.SituacaoTributaria = rsProdutos2("Situação Tributária").Value
  Else
      formulario.SituacaoTributaria = ""
  End If
  
  If Not IsNull(rsProdutos2("CodigoBeneficio").Value) Then
      formulario.CodigoBeneficio = rsProdutos2("CodigoBeneficio").Value
  Else
      formulario.CodigoBeneficio = ""
  End If
  
  If Not IsNull(rsProdutos2("Percentual IPI").Value) Then
      formulario.PercentualIPI_Saida = rsProdutos2("Percentual IPI").Value
  Else
      formulario.PercentualIPI_Saida = ""
  End If
  
  If Not IsNull(rsProdutos2("Percentual_IPI_Entrada").Value) Then
      formulario.PercentualIPI_Entrada = rsProdutos2("Percentual_IPI_Entrada").Value
  Else
      formulario.PercentualIPI_Entrada = ""
  End If
  
  If Not IsNull(rsProdutos2("Percentual ICM Saida").Value) Then
      formulario.PercentualICMS_Saida = rsProdutos2("Percentual ICM Saida").Value
  Else
      formulario.PercentualICMS_Saida = ""
  End If
  
  If Not IsNull(rsProdutos2("Percentual ICM Entrada").Value) Then
      formulario.PercentualICMS_Entrada = rsProdutos2("Percentual ICM Entrada").Value
  Else
      formulario.PercentualICMS_Entrada = ""
  End If

  If Not IsNull(rsProdutos2("Percentual_ICMSST_Saida").Value) Then
      formulario.PercentualICMSST_Saida = rsProdutos2("Percentual_ICMSST_Saida").Value
  Else
      formulario.PercentualICMSST_Saida = ""
  End If

  If Not IsNull(rsProdutos2("Percentual_ICMSST_Entrada").Value) Then
      formulario.PercentualICMSST_Entrada = rsProdutos2("Percentual_ICMSST_Entrada").Value
  Else
      formulario.PercentualICMSST_Entrada = ""
  End If

  If Not IsNull(rsProdutos2("BaseCalculoICMSST_Saida").Value) Then
      formulario.ValorBC_ICMSST_Saida = rsProdutos2("BaseCalculoICMSST_Saida").Value
  Else
      formulario.ValorBC_ICMSST_Saida = ""
  End If

  If Not IsNull(rsProdutos2("BaseCalculoICMSST_Entrada").Value) Then
      formulario.ValorBC_ICMSST_Entrada = rsProdutos2("BaseCalculoICMSST_Entrada").Value
  Else
      formulario.ValorBC_ICMSST_Entrada = ""
  End If

  If Not IsNull(rsProdutos2("TipoSituacaoTributariaPIS").Value) Then
      formulario.TipoSituacaoTributariaPIS = rsProdutos2("TipoSituacaoTributariaPIS").Value
  Else
      formulario.TipoSituacaoTributariaPIS = 0
  End If
  
  If Not IsNull(rsProdutos2("Classe")) Then
      formulario.ClasseProduto = rsProdutos2("Classe")
  Else
      formulario.ClasseProduto = ""
  End If
  
  If Not IsNull(rsProdutos2("Sub Classe")) Then
      formulario.SubClasseProduto = rsProdutos2("Sub Classe")
  Else
      formulario.SubClasseProduto = ""
  End If

  formulario.Show
End Sub

Private Sub cmd_consultarPareamentoFornecedor_Click()
On Error GoTo Erro

  If lbl_fornCodigo.Caption = "" Then
    MsgBox "Selecione um fornecedor na grade acima para realizar o pareamento de código de produto", vbInformation, "Atenção"
    Exit Sub
  End If
  
  If cboNome.Text = "" Then
    MsgBox "Selecione um produto válido", vbInformation, "Atenção"
    Exit Sub
  End If

  Dim objFornPar As frmProdutosPareamentoProdutoForn

  Set objFornPar = New frmProdutosPareamentoProdutoForn
  
  objFornPar.pCodigoFornecedor = lbl_fornCodigo.Caption
  objFornPar.pNomeFornecedor = lbl_fornNome.Caption
  
  objFornPar.Show
  
  Exit Sub
Erro:
  MsgBox "Erro na chamada da tela de pareamento " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub cmd_salvarPareamentoCodigoProdutoFornecedor_Click()
On Error GoTo Erro

  If lbl_fornCodigo.Caption = "" Then
    MsgBox "Selecione um fornecedor na grade acima para realizar o pareamento de código de produto", vbInformation, "Atenção"
    Exit Sub
  End If
  
  If cboNome.Text = "" Then
    MsgBox "Selecione um produto válido", vbInformation, "Atenção"
    Exit Sub
  End If

  Dim objFornPar As frmProdutosPareamentoProdutoForn

  Set objFornPar = New frmProdutosPareamentoProdutoForn

  objFornPar.pCodigoProduto = cboCodigo.Text
  objFornPar.pNomeProduto = cboNome.Text
  
  If O_Produto_Normal.Value = True Then
      objFornPar.pTipoProduto = "NORMAL"
  ElseIf O_Produto_Grade.Value = True Then
      objFornPar.pTipoProduto = "GRADE"
  End If
  
  objFornPar.pCodigoFornecedor = lbl_fornCodigo.Caption
  objFornPar.pNomeFornecedor = lbl_fornNome.Caption
  
  objFornPar.Show
  
  Exit Sub
Erro:
  MsgBox "Erro na chamada da tela de pareamento " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub cmd_tributosProduto_Click()
  Dim formulario As Form
  
  Set formulario = New frmProdutosTributos
  
  formulario.CodigoProduto = cboCodigo.Text
  formulario.nomeProduto = cboNome
  
  If Not IsNull(rsProdutos2("CodigoNBM")) Then
      formulario.NCM = rsProdutos2("CodigoNBM")
  Else
      formulario.NCM = ""
  End If
  
  If Not IsNull(rsProdutos2("Situação Tributária")) Then
      formulario.SituacaoTributaria = rsProdutos2("Situação Tributária")
  Else
      formulario.SituacaoTributaria = ""
  End If
  
  If Not IsNull(rsProdutos2("SituacaoTributariaEntrada")) Then
      formulario.SituacaoTributariaEntrada = rsProdutos2("SituacaoTributariaEntrada")
  Else
      formulario.SituacaoTributariaEntrada = ""
  End If
  
  If Not IsNull(rsProdutos2("Percentual ICM Saida")) Then
      formulario.ICMS_Saida = rsProdutos2("Percentual ICM Saida")
  Else
      formulario.ICMS_Saida = 0
  End If
  
  If Not IsNull(rsProdutos2("Percentual ICM Entrada")) Then
      formulario.ICMS_EntradaDev = rsProdutos2("Percentual ICM Entrada")
  Else
      formulario.ICMS_EntradaDev = 0
  End If
  
  If Not IsNull(rsProdutos2("Percentual_ICMSST_Saida")) Then
      formulario.ICMSST_Saida = rsProdutos2("Percentual_ICMSST_Saida")
  Else
      formulario.ICMSST_Saida = 0
  End If
  
  If Not IsNull(rsProdutos2("BaseCalculoICMSST_Saida")) Then
      formulario.ICMSST_SaidaBaseCalculo = rsProdutos2("BaseCalculoICMSST_Saida")
  Else
      formulario.ICMSST_SaidaBaseCalculo = 0
  End If
  
  If Not IsNull(rsProdutos2("Percentual_ICMSST_Entrada")) Then
      formulario.ICMSST_Entrada = rsProdutos2("Percentual_ICMSST_Entrada")
  Else
      formulario.ICMSST_Entrada = 0
  End If
  
  If Not IsNull(rsProdutos2("BaseCalculoICMSST_Entrada")) Then
      formulario.ICMSST_EntradaBaseCalculo = rsProdutos2("BaseCalculoICMSST_Entrada")
  Else
      formulario.ICMSST_EntradaBaseCalculo = 0
  End If
  
  If Not IsNull(rsProdutos2("Percentual IPI")) Then
      formulario.Perc_IPISaida = rsProdutos2("Percentual IPI")
  Else
      formulario.Perc_IPISaida = 0
  End If
  
  If Not IsNull(rsProdutos2("Percentual_IPI_Entrada")) Then
      formulario.Perc_IPIEntrada = rsProdutos2("Percentual_IPI_Entrada")
  Else
      formulario.Perc_IPIEntrada = 0
  End If
  
  If Not IsNull(rsProdutos2("TipoSituacaoTributariaPIS")) Then
      formulario.SituacaoTributariaPIS = rsProdutos2("TipoSituacaoTributariaPIS")
  Else
      formulario.SituacaoTributariaPIS = -1
  End If
  
  If Not IsNull(rsProdutos2("CodigoBeneficio")) Then
      formulario.CodigoBeneficio = rsProdutos2("CodigoBeneficio")
  Else
      formulario.CodigoBeneficio = ""
  End If
  
  If Not IsNull(rsProdutos2("Redução ICM")) Then
      formulario.Perc_RedBC_ICM = rsProdutos2("Redução ICM")
  Else
      formulario.Perc_RedBC_ICM = 0
  End If
  
  formulario.Show
End Sub

Private Sub cmd_vincularCFOPs_produto_Click()
  Dim formulario As Form
  
  Set formulario = New frmProdutosCFOP
  
  formulario.CodigoProduto = cboCodigo.Text
  formulario.nomeProduto = cboNome
  
  formulario.Show

End Sub

Private Sub cmdCalcular_Click()
  If O_C_Venda.Value = True Then
    Call Calc_Valor
  Else
    If O_C_Lucro_Val.Value = True Then
      Call Calc_Valor_Lucro
    Else
      If O_C_Lucro_Perc.Value = True Then
        Call Calc_Perc_Lucro
      End If
    End If
  End If
End Sub

Private Sub cmdCalcularCusto_Click()
  '04/07/2005 - Daniel
  'Case.....: Zue
  'Projeto..: Informações do Custo de Produto
  Dim rstPrecos          As Recordset
  Dim strSQL             As String
  Dim dblCustoDoProduto  As Double
  Dim dblMaoDeObraFaccao As Double
  Dim dblLavanderia      As Double
  Dim dblBordado         As Double
  Dim dblEstamparia      As Double
  Dim dblAviamentos      As Double
  Dim dblOutrosCustos    As Double
  Dim dblCUSTO           As Double
  
  dblCustoDoProduto = 0
  dblMaoDeObraFaccao = 0
  dblLavanderia = 0
  dblBordado = 0
  dblEstamparia = 0
  dblAviamentos = 0
  dblOutrosCustos = 0
  dblCUSTO = 0
  
  On Error GoTo TratarErro
  
  If IsNumeric(txtCustoProduto(2).Text) Then dblCustoDoProduto = CDbl(txtCustoProduto(2).Text)
  If IsNumeric(txtCustoProduto(3).Text) Then dblMaoDeObraFaccao = CDbl(txtCustoProduto(3).Text)
  If IsNumeric(txtCustoProduto(4).Text) Then dblLavanderia = CDbl(txtCustoProduto(4).Text)
  If IsNumeric(txtCustoProduto(5).Text) Then dblBordado = CDbl(txtCustoProduto(5).Text)
  If IsNumeric(txtCustoProduto(6).Text) Then dblEstamparia = CDbl(txtCustoProduto(6).Text)
  If IsNumeric(txtCustoProduto(7).Text) Then dblAviamentos = CDbl(txtCustoProduto(7).Text)
  If IsNumeric(txtCustoProduto(8).Text) Then dblOutrosCustos = CDbl(txtCustoProduto(8).Text)
  
  dblCUSTO = Format(dblCustoDoProduto + _
             dblMaoDeObraFaccao + _
             dblLavanderia + _
             dblBordado + _
             dblEstamparia + _
             dblAviamentos + _
             dblOutrosCustos, FORMAT_VALUE)
             
  'Atualizar ou Criar a tabela CUSTO para o produto
  strSQL = ""
  strSQL = "SELECT * FROM Preços WHERE Tabela = 'CUSTO' AND Produto = '" & Trim(cboCodigo.Text) & "'"
  
  Set rstPrecos = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstPrecos.RecordCount = 0 Then 'Criar
    With rstPrecos
      .AddNew
      .Fields("Tabela").Value = "CUSTO"
      .Fields("Produto").Value = Trim(cboCodigo.Text) & ""
      .Fields("Preço").Value = dblCUSTO
      .Fields("Data Alteração").Value = CStr(Data_Atual)
      .Update
      .Close
    End With
  Else 'Editar
    With rstPrecos
      .Edit
      .Fields("Preço").Value = dblCUSTO
      .Fields("Data Alteração").Value = CStr(Data_Atual)
      .Update
      .Close
    End With
  End If
  
  Set rstPrecos = Nothing
  
  'Atualizar o objeto medPreco
  If lblNomePreco(0).Caption = "CUSTO: " Then medPreco(0).Text = dblCUSTO
  If lblNomePreco(1).Caption = "CUSTO: " Then medPreco(1).Text = dblCUSTO
  If lblNomePreco(2).Caption = "CUSTO: " Then medPreco(2).Text = dblCUSTO
  If lblNomePreco(3).Caption = "CUSTO: " Then medPreco(3).Text = dblCUSTO
  
  'Invocamos a Update Record
  Screen.MousePointer = vbHourglass
  Call UpdateRecord
  Screen.MousePointer = vbDefault
  
  MsgBox "O CUSTO do Produto " & Trim(cboCodigo.Text) & " foi atualizado com sucesso.", vbInformation, "Quick Store"
  
  Exit Sub

TratarErro:
  Screen.MousePointer = vbHourglass
  MsgBox "Erro ao calcular o custo." & vbCrLf & "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Sub

Private Sub cmdCustoTecido_Click()
  '13/07/2005 - Daniel
  'Case: Zue
  If IsNumeric(txtCustoProduto(0).Text) And IsNumeric(txtCustoProduto(1).Text) Then CalcularCustoDoTecido
End Sub

Private Sub cmdImprimirEtiqueta_Click()
  Dim strNomeProduto    As String
  Dim strCodigoProduto  As String
  Dim intFile           As Integer
  Dim strPath           As String
  Dim strTemp           As String
  
  Dim rstPrecos         As Recordset
  Dim dblPreco          As Double
  
  strNomeProduto = Trim(cboNome.Text)
  strCodigoProduto = Trim(cboCodigo.Text)
  
  intFile = FreeFile
  
  strPath = InputBox("Insira a porta de impressão", "Etiqueta de Código de Barras", "LPT1")
  If Len(Trim(strPath)) <= 0 Then Exit Sub
  
  strTemp = InputBox("Insira a temperatura", "Etiqueta de Código de Barras", "12")
  If Len(Trim(strTemp)) <= 0 Then Exit Sub
  
  '---[ Pega o preço de custo do produto ]---'
    dblPreco = 0
    Set rstPrecos = db.OpenRecordset("SELECT * FROM Preços WHERE Tabela = 'CUSTO' AND Produto = '" & cboCodigo.Text & "'")
    
    With rstPrecos
      If Not (.BOF And .EOF) Then
        dblPreco = .Fields("Preço").Value
      End If
      
      .Close
      Set rstPrecos = Nothing
    End With
  '---[ Pega o preço de custo do produto ]---'
  
  Open strPath For Output As #intFile
    Print #intFile, Chr(2) & "L" & Chr(13)
    Print #intFile, "D11"
    Print #intFile, "H" & strTemp
    Print #intFile, "121102000200105" & Mid(sTranslateInvalidChar(strNomeProduto), 1, 25) & Chr(13)
    Print #intFile, "121102000100105" & Mid(sTranslateInvalidChar(strNomeProduto), 26, 25) & Chr(13)
    Print #intFile, "121102000000105" & Mid(sTranslateInvalidChar(strNomeProduto), 51, 25) & Chr(13)
    Print #intFile, "1E2201000000265A" & strCodigoProduto & Chr(13)             '275
    Print #intFile, "121102000200265" & Format(CStr(dblPreco), FORMAT_VALUE) & Chr(13)              '275
    Print #intFile, "E"
  Close #intFile
End Sub

Private Sub cmdQuery_Click()
  Dim F As Form
  Call StatusMsg("Aguarde....")
  gsCodProduto = cboCodigo.Text
  'Set F = New frmConsultaProd
  'F.Show vbModal
  frmConsultaProd.Show 'vbModal
  'Set F = Nothing
  Call StatusMsg("")
End Sub

Private Sub cmdUpdateBatch_Click()
  frmProdutosBatchUpdate.Show vbModal
End Sub

'Private Sub Código_DblClick()
'  Call FindNextProdutoCode
'End Sub
'
'Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyF5 Then
'    If Not IsNull(Num_Registro) Then
'      DisplayMsg "Pressione o botão Novo antes."
'      Exit Sub
'    End If
'    Call FindNextProdutoCode
'  End If
'End Sub
'
'Private Sub Código_KeyPress(KeyAscii As Integer)
' If Glob_Cod_Alfa = False Then
'     KeyAscii = Verifica_Tecla_Integer(KeyAscii)
' Else
'     KeyAscii = Verifica_Tecla_Código(KeyAscii)
' End If
'End Sub
'
'Public Sub Código_LostFocus()
'  Dim Aux As Variant
'
'  Call StatusMsg("")
'
'  'Set rsProdutos2 = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY Código", dbOpenDynaset)
'
'  rsProdutos2.FindFirst "Código = '" & gsHandleNull(cboCodigo.Text) & "'"
'  If Not rsProdutos2.NoMatch Then
'    Call ShowRecord
'  Else
' '   gsTitle = LoadResString(201)
' '   gsMsg = "Registro não localizado."
' '   gnStyle = vbOKOnly + vbExclamation
' '   gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'    If Not IsNull(Num_Registro) Then
'      rsProdutos2.Bookmark = Num_Registro
'    End If
'  End If
'
'End Sub


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

Private Sub Combo_Moeda_CloseUp()
  Combo_Moeda.Text = Combo_Moeda.Columns(1).Text
  Combo_Moeda_LostFocus
End Sub

Private Sub Combo_Moeda_LostFocus()
  Dim Aux As Variant
  
  Nome_Moeda.Caption = ""
  Aux = Combo_Moeda.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 99 Then Exit Sub
  
  rsMoedas.Index = "Código"
  rsMoedas.Seek "=", Val(Aux)
  If rsMoedas.NoMatch Then Exit Sub
  
  Nome_Moeda.Caption = rsMoedas("Nome")
  

End Sub

Private Sub Combo_Pesq1_CloseUp()
  Combo_Pesq1.Text = Combo_Pesq1.Columns(1).Text
  Combo_Pesq1_LostFocus
End Sub

Private Sub Combo_Pesq1_InitColumnProps()
  With Combo_Pesq1
    .Columns(0).Width = 4500
    .Columns(1).Width = 1000
  End With
End Sub

Private Sub Combo_Pesq1_KeyPress(KeyAscii As Integer)
  If Not Combo_Pesq1.DroppedDown Then
    KeyAscii = gnLimitKeyPress(Combo_Pesq1, 5, KeyAscii, True)
  End If
End Sub

Private Sub Combo_Pesq1_LostFocus()
  Dim Aux As Variant
  
  Nome_Pesq1.Caption = ""
  Aux = Combo_Pesq1.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 99999 Then Exit Sub
  
  rsPesquisa1.Index = "Código"
  rsPesquisa1.Seek "=", Val(Aux)
  If rsPesquisa1.NoMatch Then Exit Sub
  
  Nome_Pesq1.Caption = rsPesquisa1("Nome") & ""

End Sub

Private Sub Combo_Pesq2_CloseUp()
  Combo_Pesq2.Text = Combo_Pesq2.Columns(1).Text
  Combo_Pesq2_LostFocus
End Sub

Private Sub Combo_Pesq2_InitColumnProps()
  With Combo_Pesq2
    .Columns(0).Width = 4500
    .Columns(1).Width = 1000
  End With
End Sub

Private Sub Combo_Pesq2_KeyPress(KeyAscii As Integer)
  If Not Combo_Pesq2.DroppedDown Then
    KeyAscii = gnLimitKeyPress(Combo_Pesq2, 5, KeyAscii, True)
  End If
End Sub

Private Sub Combo_Pesq2_LostFocus()
  Dim Aux As Variant
  
  Nome_Pesq2.Caption = ""
  Aux = Combo_Pesq2.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 99999 Then Exit Sub
  
  rsPesquisa2.Index = "Código"
  rsPesquisa2.Seek "=", Val(Aux)
  If rsPesquisa2.NoMatch Then Exit Sub
  
  Nome_Pesq2.Caption = rsPesquisa2("Nome") & ""
  
End Sub

Private Sub Combo_Pesq3_CloseUp()
  Combo_Pesq3.Text = Combo_Pesq3.Columns(1).Text
  Combo_Pesq3_LostFocus
End Sub

Private Sub Combo_Pesq3_InitColumnProps()
  With Combo_Pesq3
    .Columns(0).Width = 4500
    .Columns(1).Width = 1000
  End With
End Sub

Private Sub Combo_Pesq3_KeyPress(KeyAscii As Integer)
  If Not Combo_Pesq3.DroppedDown Then
    KeyAscii = gnLimitKeyPress(Combo_Pesq3, 5, KeyAscii, True)
  End If
End Sub

Private Sub Combo_Pesq3_LostFocus()
  Dim Aux As Variant
  
  Nome_Pesq3.Caption = ""
  Aux = Combo_Pesq3.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 99999 Then Exit Sub
  
  rsPesquisa3.Index = "Código"
  rsPesquisa3.Seek "=", Val(Aux)
  If rsPesquisa3.NoMatch Then Exit Sub
  
  Nome_Pesq3.Caption = rsPesquisa3("Nome") & ""

End Sub

Private Sub Combo_Sub_Classe_CloseUp()
 Combo_Sub_Classe.Text = Combo_Sub_Classe.Columns(1).Text
 Combo_Sub_Classe_LostFocus
End Sub

Private Sub Combo_Sub_Classe_LostFocus()
  Dim Aux As Variant
  
  Nome_Sub_Classe.Caption = ""
  Aux = Combo_Sub_Classe.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsSub_Classes.Index = "Código"
  rsSub_Classes.Seek "=", Val(Aux)
  If rsSub_Classes.NoMatch Then Exit Sub
  
  If Not IsNull(rsSub_Classes("Nome")) Then
      Nome_Sub_Classe.Caption = rsSub_Classes("Nome")
  End If

End Sub

Private Sub Comissão_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Comissão_LostFocus()
 If IsNull(Comissão.Text) Then Exit Sub
 If Not IsNumeric(Comissão.Text) Then Exit Sub
 Comissão = CDbl(Comissão.Text)
End Sub

Private Sub Comissão_Sobre_Click()
 If Comissão_Sobre.Value = 0 Then
   L_Comiss1.Enabled = False
   Comissão.Enabled = False
   L_Comiss2.Enabled = False
 End If
 If Comissão_Sobre.Value = 1 Then
   L_Comiss1.Enabled = True
   Comissão.Enabled = True
   L_Comiss2.Enabled = True
 End If
 
End Sub


Private Sub Controla_Estoque_Click()
  Estoque_Min.Enabled = True
  Estoque_Min.Text = ""
  Estoque_Min.Enabled = (Controla_Estoque.Value = 1)
  Estoque_Ideal.Enabled = True
  Estoque_Ideal.Text = ""
  Estoque_Ideal.Enabled = (Controla_Estoque.Value = 1)
  '30/10/2007 - Anderson
  'Campo utilizado para gerar o relatório de produtos a comprar
  'Solicitante: Kings Cross
  If g_bolRelatorioCompra Then
    'txtEspacoFisicoTotal.Enabled = True
    'txtEspacoFisicoTotal.Text = ""
    'txtEspacoFisicoTotal.Enabled = (Controla_Estoque.Value = 1)
  Else
  
  End If
End Sub

Private Sub Custo_Médio_LostFocus()
  '04/05/2004 - Daniel
  'Case: Embalavi
  If g_bln5CasasDecimais Then
    Custo_Médio.Text = Format(Custo_Médio.Text, "###,###,##0.00000")
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    Custo_Médio.Text = Format(Custo_Médio.Text, "###,###,##0.000")
  Else
    Custo_Médio.Text = Format(Custo_Médio.Text, "###,###,##0.00")
  End If
End Sub

Private Sub DropDown2_1_Click()
  Grade2_1.Columns(1).Text = DropDown2_1.Columns(1).Text
  Grade2_1.Columns(2).Text = DropDown2_1.Columns(0).Text
  Grade2_1.Columns(3).Text = DropDown2_1.Columns(2).Text
End Sub

Private Sub Estoque_Ideal_KeyPress(KeyAscii As Integer)
  If Fracionado.Value = 0 Then
    KeyAscii = Verifica_Tecla_Integer(KeyAscii)
  End If
End Sub

Private Sub Estoque_Ideal_LostFocus()
  If Fracionado.Value = 1 And txtCasasDecimais.Text <> "" Then
    If CInt(txtCasasDecimais.Text) > 0 Then
      Estoque_Ideal.Text = Format(Estoque_Ideal.Text, String(8, "#") & "0." & String(CInt(txtCasasDecimais.Text), "0"))
    Else
      Estoque_Ideal.Text = Format(Estoque_Ideal.Text, String(8, "#") & "0")
    End If
  End If
End Sub

Private Sub txtEspacoFisicoTotal_KeyPress(KeyAscii As Integer)
  If Fracionado.Value = 0 Then
    KeyAscii = Verifica_Tecla_Integer(KeyAscii)
  End If
End Sub

Private Sub txtEspacoFisicoTotal_LostFocus()
  If Fracionado.Value = 1 And txtCasasDecimais.Text <> "" Then
    If CInt(txtCasasDecimais.Text) > 0 Then
      'txtEspacoFisicoTotal.Text = Format(txtEspacoFisicoTotal.Text, String(8, "#") & "0." & String(CInt(txtCasasDecimais.Text), "0"))
    Else
      'txtEspacoFisicoTotal.Text = Format(txtEspacoFisicoTotal.Text, String(8, "#") & "0")
    End If
  End If
End Sub

Private Sub Estoque_Min_KeyPress(KeyAscii As Integer)
  If Fracionado.Value = 0 Then
    KeyAscii = Verifica_Tecla_Integer(KeyAscii)
  End If
End Sub

Private Sub Estoque_Min_LostFocus()
  If Fracionado.Value = 1 And txtCasasDecimais.Text <> "" Then
    If CInt(txtCasasDecimais.Text) > 0 Then
      Estoque_Min.Text = Format(Estoque_Min.Text, String(8, "#") & "0." & String(CInt(txtCasasDecimais.Text), String(8, "#") & "0"))
    Else
      Estoque_Min.Text = Format(Estoque_Min.Text, String(8, "#") & "0")
    End If
  End If
End Sub


Private Sub Fabricante_LostFocus()
    If Len(Trim(Fabricante.Text)) > 0 Then
        Fabricante.Text = UCase(Trim(Fabricante.Text))
    End If
End Sub

Private Sub Fixa_Custo_C_Perc_Click()
 
  If Fixa_Custo_C_Perc.Value = 1 Then
    Fixa_Custo_C_Perc.Enabled = False
    Fixa_Custo_C_Valor.Enabled = True
    Fixa_Custo_C_Valor.Value = 0
  End If
 
End Sub

Private Sub Fixa_Custo_C_Valor_Click()

  If Fixa_Custo_C_Valor.Value = 1 Then
    Fixa_Custo_C_Valor.Enabled = False
    Fixa_Custo_C_Perc.Enabled = True
    Fixa_Custo_C_Perc.Value = 0
  End If
 
End Sub

Private Sub Fixa_Desc_C_Perc_Click()

  If Fixa_Desc_C_Perc.Value = 1 Then
    Fixa_Desc_C_Perc.Enabled = False
    Fixa_Desc_C_Valor.Enabled = True
    Fixa_Desc_C_Valor.Value = 0
  End If

End Sub

Private Sub Fixa_Desc_C_Valor_Click()

  If Fixa_Desc_C_Valor.Value = 1 Then
    Fixa_Desc_C_Valor.Enabled = False
    Fixa_Desc_C_Perc.Enabled = True
    Fixa_Desc_C_Perc.Value = 0
  End If
 
End Sub

Private Sub Fixa_Frete_C_Perc_Click()
 
  If Fixa_Frete_C_Perc.Value = 1 Then
    Fixa_Frete_C_Perc.Enabled = False
    Fixa_Frete_C_Valor.Enabled = True
    Fixa_Frete_C_Valor.Value = 0
  End If

End Sub

Private Sub Fixa_Frete_C_Valor_Click()

  If Fixa_Frete_C_Valor.Value = 1 Then
    Fixa_Frete_C_Valor.Enabled = False
    Fixa_Frete_C_Perc.Enabled = True
    Fixa_Frete_C_Perc.Value = 0
  End If

End Sub

Private Sub Fixa_ICM_C_Perc_Click()
 
  If Fixa_ICM_C_Perc.Value = 1 Then
    Fixa_ICM_C_Perc.Enabled = False
    Fixa_ICM_C_Valor.Enabled = True
    Fixa_ICM_C_Valor.Value = 0
  End If

End Sub


Private Sub Fixa_ICM_C_Valor_Click()

  If Fixa_ICM_C_Valor.Value = 1 Then
    Fixa_ICM_C_Valor.Enabled = False
    Fixa_ICM_C_Perc.Enabled = True
    Fixa_ICM_C_Perc.Value = 0
  End If

End Sub

Private Sub Fixa_ICM_V_Perc_Click()
 
  If Fixa_ICM_V_Perc.Value = 1 Then
    Fixa_ICM_V_Perc.Enabled = False
    Fixa_ICM_V_Valor.Enabled = True
    Fixa_ICM_V_Valor.Value = 0
  End If

End Sub

Private Sub Fixa_ICM_V_Valor_Click()
 
  If Fixa_ICM_V_Valor.Value = 1 Then
    Fixa_ICM_V_Valor.Enabled = False
    Fixa_ICM_V_Perc.Enabled = True
    Fixa_ICM_V_Perc.Value = 0
  End If
 
End Sub

Private Sub Fixa_Imp_V_Perc_Click()
 
  If Fixa_Imp_V_Perc.Value = 1 Then
    Fixa_Imp_V_Perc.Enabled = False
    Fixa_Imp_V_Valor.Enabled = True
    Fixa_Imp_V_Valor.Value = 0
  End If

End Sub

Private Sub Fixa_Imp_V_Valor_Click()
 
  If Fixa_Imp_V_Valor.Value = 1 Then
    Fixa_Imp_V_Valor.Enabled = False
    Fixa_Imp_V_Perc.Enabled = True
    Fixa_Imp_V_Perc.Value = 0
  End If
 
End Sub

Private Sub Fixa_IPI_C_Perc_Click()
 
  If Fixa_IPI_C_Perc.Value = 1 Then
    Fixa_IPI_C_Perc.Enabled = False
    Fixa_IPI_C_Valor.Enabled = True
    Fixa_IPI_C_Valor.Value = 0
  End If

End Sub

Private Sub Fixa_IPI_C_Valor_Click()

  If Fixa_IPI_C_Valor.Value = 1 Then
    Fixa_IPI_C_Valor.Enabled = False
    Fixa_IPI_C_Perc.Enabled = True
    Fixa_IPI_C_Perc.Value = 0
  End If

End Sub

Private Sub Fixa_IPI_V_Perc_Click()
 
  If Fixa_IPI_V_Perc.Value = 1 Then
    Fixa_IPI_V_Perc.Enabled = False
    Fixa_IPI_V_Valor.Enabled = True
    Fixa_IPI_V_Valor.Value = 0
  End If

End Sub

Private Sub Fixa_IPI_V_Valor_Click()
 
  If Fixa_IPI_V_Valor.Value = 1 Then
    Fixa_IPI_V_Valor.Enabled = False
    Fixa_IPI_V_Perc.Enabled = True
    Fixa_IPI_V_Perc.Value = 0
  End If
 
End Sub

Private Sub Fixa_Outros_C_Perc_Click()
 
  If Fixa_Outros_C_Perc.Value = 1 Then
    Fixa_Outros_C_Perc.Enabled = False
    Fixa_Outros_C_Valor.Enabled = True
    Fixa_Outros_C_Valor.Value = 0
  End If
 
End Sub

Private Sub Fixa_Outros_C_Valor_Click()

  If Fixa_Outros_C_Valor.Value = 1 Then
    Fixa_Outros_C_Valor.Enabled = False
    Fixa_Outros_C_Perc.Enabled = True
    Fixa_Outros_C_Perc.Value = 0
  End If

End Sub

Private Sub Fixa_Outros_V_Perc_Click()
 
  If Fixa_Outros_V_Perc.Value = 1 Then
    Fixa_Outros_V_Perc.Enabled = False
    Fixa_Outros_V_Valor.Enabled = True
    Fixa_Outros_V_Valor.Value = 0
  End If

End Sub

Private Sub Fixa_Outros_V_Valor_Click()
 
  If Fixa_Outros_V_Valor.Value = 1 Then
    Fixa_Outros_V_Valor.Enabled = False
    Fixa_Outros_V_Perc.Enabled = True
    Fixa_Outros_V_Perc.Value = 0
  End If

End Sub

Private Sub Form_Activate()
  Me.KeyPreview = True
End Sub

Private Sub Form_Deactivate()
  Me.KeyPreview = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyF1 Then
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
      'Call objHelp.Show(strfile, "QuickStore10Help")
      Call objHelp.Show(strfile, "QuickStore10Help", 10007)
      Set objHelp = Nothing
  
      Exit Sub
  End If
  
  Dim Tool As ActiveBarLibraryCtl.Tool
  Dim nX As Byte
  
  Set Tool = New ActiveBarLibraryCtl.Tool
  
  
  '24/09/2002 - mpdea
  'Incluído Exit Sub e atribuição de retorno aos argumentos (byref)
  If KeyCode = vbKeyF3 Then
    cmdQuery_Click
    KeyCode = 0: Shift = 0
    Exit Sub
  End If
  
  
  If gbComboDropped Then
    Call cboCodigo_KeyDown(KeyCode, Shift)
  Else
    If KeyCode = vbKeyP And Shift = vbCtrlMask + vbAltMask Then
      'Atalho oculto para Preços
      For nX = 0 To 3
        If medPreco(nX).Visible Then
          medPreco(nX).SetFocus
          KeyCode = 0: Shift = 0
          Exit Sub
        End If
      Next nX
    ElseIf KeyCode = vbKeyF5 Then
      Call FindNextProdutoCode
    '-------------------------------------------------------
    '22/01/2003 - mpdea
    'Verifica Quick em modo limitado
    ElseIf KeyCode = vbKeyF7 And gblnQuickFull Then
      Tool.Name = "miOpCheckGrid"
      Call ActiveBar1_Click(Tool)
      KeyCode = 0: Shift = 0
    ElseIf KeyCode = vbKeyF8 And gblnQuickFull Then
      Tool.Name = "miOpShowEdicoes"
      Call ActiveBar1_Click(Tool)
      KeyCode = 0: Shift = 0
    '-------------------------------------------------------
    Else
      Call HandleKeyDown(KeyCode, Shift)
    End If
  End If
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim blnMultiline As Boolean
  
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
  
  
  '26/09/2002 - mpdea
  'Incluído verificação para controles Multiline
  If KeyAscii = vbKeyReturn Then
  
    On Error Resume Next
    blnMultiline = Me.ActiveControl.MultiLine
    On Error GoTo 0
    
    If Not blnMultiline Then
      SendKeys "{Tab}"
      KeyAscii = 0
    End If
  End If
  
End Sub

'06/01/2006 - mpdea
'Incluído tratamento de erros
Public Sub Form_Load()
 
  Dim sSql As String
  
  On Error GoTo ErrHandler
  
  Screen.MousePointer = vbHourglass
  Call CenterForm(Me)
  
  gbComboDropped = False
  KeyPreview = True
  
  '06/02/2006 - mpdea
  'Seta Banco de dados para controle Data
  datGrupoFiscal.DatabaseName = gsQuickDBFileName
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data2_1.DatabaseName = gsQuickDBFileName
  Data2_2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data4_1.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  Data6.DatabaseName = gsQuickDBFileName
  Data7.DatabaseName = gsQuickDBFileName
  '20/07/2005 - Daniel
  'Acesso a tabela de Código NBM
  datCodigoNBM.DatabaseName = gsQuickDBFileName
  
  Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset)
  Set rsProdutos2 = rsProdutos.Clone
  '----------------------------------------------------------------------------------------------------------------------
  '13/03/2013-Alexandre Afornali
  'Case Infosocial para nao filtrar produtos
  FiltraProdutos
  
  'datProdNome.Refresh
  '----------------------------------------------------------------------------------------------------------------------
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsMoedas = db.OpenRecordset("Moedas", , dbReadOnly)
  Set rsC_Fiscal = db.OpenRecordset("Classificação Fiscal", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque")
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsForn_Prod = db.OpenRecordset("Forn_Prod", , dbReadOnly)
'tirei daqui
  Set rsTabelas = db.OpenRecordset("Tabela de Preços")
  Set rsPrecos = db.OpenRecordset("Preços")
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsPesquisa1 = db.OpenRecordset("Pesquisa 1", , dbReadOnly)
  Set rsPesquisa2 = db.OpenRecordset("Pesquisa 2", , dbReadOnly)
  Set rsPesquisa3 = db.OpenRecordset("Pesquisa 3", , dbReadOnly)
'  Set rsProdutos22 = db.OpenRecordset("Produtos", , dbReadOnly)
  '15/10/2007 - Anderson
  'Alteração para o cadastro de CFOP por produto
  Set rsOperacaoSaida = db.OpenRecordset("Operações Saída", , dbReadOnly)
  Set rsProdutoCFOP = db.OpenRecordset("ProdutoCFOP", , dbReadOnly)

  
  Call InitFormProdutos
  
  Call ActiveBarLoadToolTips(frmProdutos)
  
  Call ClearScreen
  
  
  '19/09/2005 - mpdea
  'Índice para cálculo do Preço de Entrada
  fraIndicePrecoEntrada.Visible = g_blnIndicePrecoEntrada
  
  
  '26/07/2002 - mpdea
  'Somente ativa o botão se houver permissão para gravar
  cmdUpdateBatch.Enabled = gbPodeGravar 'ou ActiveBar1.Tools("miOpUpdate").Enabled
  
  '09-07-2015 Jean Ricardo Zanella - Função para verificar se usuario tem permissão para alterar preços
  If blnPermissaoAlterarPrecos(gnUserCode) = False Then
    medPreco(0).Enabled = False
    medPreco(1).Enabled = False
    medPreco(2).Enabled = False
    medPreco(3).Enabled = False
    Custo_Médio.Enabled = False
  End If
  
  '22/01/2003 - mpdea
  'Modo limitado
  If Not gblnQuickFull Then
    
    O_Produto_Grade.Visible = False
    O_Produto_Edição.Visible = False
    
    tabItens.TabVisible(3) = False
    
    With ActiveBar1
      With .Bands("tbrComplem")
        .Tools("miOpCheckGrid").Visible = False
        .Tools("miOpShowEdicoes").Visible = False
      End With
      .RecalcLayout
      .Refresh
    End With
    
    'Moeda
    Label14.Visible = False
    Combo_Moeda.Visible = False
    Nome_Moeda.Visible = False
  End If
  
  '31/01/2008 - Anderson
  'Inclusão do serial da Mello e Maruko QS71248-282
  '-----------------------------------------------------------------------------
  ' Data......: 18/09/2003
  ' Dev.......: Maikel
  ' Case......: TV Shopping
  cmdImprimirEtiqueta.Visible = CheckSerialCaseMod("QS39945-043", "QS39944-959", "QS71248-282")
  '-----------------------------------------------------------------------------
  
  '-----------------------------------------------------------------------------
  '17/05/2005 - Daniel
  'Personalizações para a empresa Miss Nuvem
  fraEtiq.Visible = CheckSerialCaseMod("QS71134-847")
  '-----------------------------------------------------------------------------
  
  '-----------------------------------------------------------------------------
  '04/07/2005 - Daniel
  'Personalizações para a empresa Zue (Identificações do Custo do Produto)
  fraIdentificacaoCusto.Visible = CheckSerialCaseMod("QS71258-374")
  '-----------------------------------------------------------------------------
  
  '20/08/2004 - Daniel
  'Usar Descrição Adicional
  'Case: Stralango (Na emissão de NF ao invés de sair a Descrição do Produto
  'sairá a Descrição Adicional
  If CheckSerialCaseMod("QS40937-051") Then
    chkUsarDescrAdic.Visible = True
  Else
    chkUsarDescrAdic.Visible = False
  End If
  
  '-----------------------------------------------------------------------------
  ' 04/05/2004 - Daniel
  ' Case: Embalavi
  m_blnEmbalavi = CheckSerialCaseMod("QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162")
  
  With medPreco(0)
    If g_bln5CasasDecimais Then
      .Format = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Format = "##,###,##0.000"
    Else
      .Format = "##,###,##0.00"
    End If
  End With
  
  With medPreco(1)
    If g_bln5CasasDecimais Then
      .Format = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Format = "##,###,##0.000"
    Else
      .Format = "##,###,##0.00"
    End If
  End With
  
  With medPreco(2)
    If g_bln5CasasDecimais Then
      .Format = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Format = "##,###,##0.000"
    Else
      .Format = "##,###,##0.00"
    End If
  End With
  
  With medPreco(3)
    If g_bln5CasasDecimais Then
      .Format = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementação de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Format = "##,###,##0.000"
    Else
      .Format = "##,###,##0.00"
    End If
  End With
  '-----------------------------------------------------------------------------
  
  '30/10/2007 - Anderson
  'Campo utilizado para gerar o relatório de produtos a comprar
  'Solicitante: Kings Cross
  'txtEspacoFisicoTotal.Visible = g_bolRelatorioCompra
  'lblEspacoFisicoTotal.Visible = g_bolRelatorioCompra
  If Not g_bolRelatorioCompra Then
    Frame14.Height = 1335
  End If
  
  Screen.MousePointer = vbDefault
  
  cboTableVenda.Text = rsParametros("Tabela 1")
  
    '******************************* INICIO:     Tratamento para o WEBAPI
  'Recuperar todos os CNPJs
  Set rsCNPJ = db.OpenRecordset("Select CGC From [Parâmetros Filial]")
  Set http = CreateObject("MSXML2.ServerXMLHTTP")
  '******************************* FIM:        Tratamento para o WEBAPI

  Exit Sub
    
ErrHandler:
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  gbComboDropped = (cboCodigo.DroppedDown = True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  '29/08/2003 - mpdea
  'Comentado desvio no descarregamento do form
  'que estava ocasionando diversos erros

'  If UnloadMode <> 4 Then
'    Cancel = True
'    Me.Hide
'    Me.ClearScreen
'  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsProdutos2.Close
  rsClasses.Close
  rsSub_Classes.Close
  rsMoedas.Close
  rsC_Fiscal.Close
  rsEstoque.Close
  rsCliFor.Close
  rsForn_Prod.Close
  rsParametros.Close
  rsTabelas.Close
  rsPrecos.Close
  rsFuncionarios.Close
  rsPesquisa1.Close
  rsPesquisa2.Close
  rsPesquisa3.Close
'  rsProdutos22.Close
  '15/10/2007 - Anderson - Alteração para o cadastro de CFOP por produto
  rsOperacaoSaida.Close
  rsProdutoCFOP.Close
  '
  Set rsProdutos2 = Nothing
  Set rsClasses = Nothing
  Set rsSub_Classes = Nothing
  Set rsMoedas = Nothing
  Set rsC_Fiscal = Nothing
  Set rsEstoque = Nothing
  Set rsCliFor = Nothing
  Set rsForn_Prod = Nothing
  Set rsParametros = Nothing
  Set rsTabelas = Nothing
  Set rsPrecos = Nothing
  Set rsFuncionarios = Nothing
  Set rsPesquisa1 = Nothing
  Set rsPesquisa2 = Nothing
  Set rsPesquisa3 = Nothing
'  Set rsProdutos22 = Nothing

  '15/10/2007 - Anderson - Alteração para o cadastro de CFOP por produto
  Set rsOperacaoSaida = Nothing
  Set rsProdutoCFOP = Nothing

  rsCNPJ.Close
  Set rsCNPJ = Nothing
  Set http = Nothing
  
End Sub

Public Sub InitFormProdutos(Optional ByVal sTabEnabled As Boolean = False)
  Dim sResult As String
  Dim bVisible As Boolean
  Dim nOrdem As Byte
  
  ActiveBar1.Tools("miOpOrdem").CBList.Clear
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 0, "Por Código Produto"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 1, "Por Nome do Produto"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 2, "Por Código para o Fornecedor"
  
  nOrdem = 3
  If gsPesq1 <> "" Then
    ActiveBar1.Tools("miOpOrdem").CBList.InsertItem nOrdem, "Por " & gsPesq1
    nOrdem = nOrdem + 1
  End If
  If gsPesq2 <> "" Then
    ActiveBar1.Tools("miOpOrdem").CBList.InsertItem nOrdem, "Por " & gsPesq2
    nOrdem = nOrdem + 1
  End If
  If gsPesq3 <> "" Then
    ActiveBar1.Tools("miOpOrdem").CBList.InsertItem nOrdem, "Por " & gsPesq3
  End If
  
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  
  sResult = GetSetting("QuickStore", "Produtos", "Mantem Nome", False)
  ActiveBar1.Tools("miOpFreezeNome").Checked = CBool(sResult)
  
  sResult = GetSetting("QuickStore", "Produtos", "Mantem Classe", False)
  ActiveBar1.Tools("miOpFreezeClasse").Checked = CBool(sResult)
  
  sResult = GetSetting("QuickStore", "Produtos", "Mantem Subclasse", False)
  ActiveBar1.Tools("miOpFreezeSubClasse").Checked = CBool(sResult)
  
  sResult = GetSetting("QuickStore", "Produtos", "Mantem Moeda", False)
  ActiveBar1.Tools("miOpFreezeMoeda").Checked = CBool(sResult)
  
  sResult = GetSetting("QuickStore", "Produtos", "Mantem Unidadevenda", False)
  ActiveBar1.Tools("miOpFreezeUnidadeVenda").Checked = CBool(sResult)
  
  sResult = GetSetting("QuickStore", "Produtos", "Mantem Fabricante", False)
  ActiveBar1.Tools("miOpFreezeFabricante").Checked = CBool(sResult)
  
  sResult = GetSetting("QuickStore", "Produtos", "Mantem Impostos", False)
  ActiveBar1.Tools("miOpFreezeImpostos").Checked = CBool(sResult)
  
'''  tabItens.TabEnabled(1) = sTabEnabled
'''  tabItens.TabEnabled(2) = sTabEnabled
'''  tabItens.TabEnabled(3) = sTabEnabled
'''  tabItens.TabEnabled(4) = sTabEnabled
'''  tabItens.TabEnabled(5) = chkWebIncluded.Value = vbChecked
  
  '31/07/2002 - mpdea
  'Loja Virtual (Web)
  chkWebIncluded.Visible = gblnWorkWeb
  tabItens.TabVisible(5) = gblnWorkWeb
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  
  lblNomePreco(0).Caption = "CUSTO: "
  lblNomePreco(1).Caption = rsParametros("Tabela 1") & ""
  If Len(lblNomePreco(1).Caption) > 0 Then
    lblNomePreco(1).Caption = lblNomePreco(1).Caption & ":"
  End If
  lblNomePreco(2).Caption = rsParametros("Tabela 2") & ""
  If Len(lblNomePreco(2).Caption) > 0 Then
    lblNomePreco(2).Caption = lblNomePreco(2).Caption & ":"
  End If
  lblNomePreco(3).Caption = rsParametros("Tabela 3") & ""
  If Len(lblNomePreco(3).Caption) > 0 Then
    lblNomePreco(3).Caption = lblNomePreco(3).Caption & ":"
  End If
  
  gsPesq1 = rsParametros("Nome Pesquisa 1") & ""
  gsPesq2 = rsParametros("Nome Pesquisa 2") & ""
  gsPesq3 = rsParametros("Nome Pesquisa 3") & ""
  
  Glob_Cod_Alfa = rsParametros("Usar Códigos Alfa")
    
  gbGrade = rsParametros("Usar Grade") = True
  gbEdicao = rsParametros("Usar Edições") = True
  gbServico = rsParametros("Usar Serviços") = True
  gbCaixas = rsParametros("Usa Vários Caixas") = True
  '20/11/2006 - Anderson
  'Considerar saldo anterior ao movimentar o caixa
  gbSaldoAnterior = rsParametros("ConsiderarSaldoAnterior") = True

  
  ActiveBar1.Tools("miOpCheckGrid").Enabled = (gbGrade = True)
  O_Produto_Grade.Enabled = (gbGrade = True)

  O_Produto_Edição.Enabled = (gbEdicao = True)
  ActiveBar1.Tools("miOpShowEdicoes").Enabled = (gbEdicao = True)
  
  
  If gsPesq1 = "" Then
    bVisible = False
  Else
    bVisible = True
    Label_Pesq1.Caption = gsPesq1 + ":"
  End If
  Label_Pesq1.Visible = bVisible
  Combo_Pesq1.Visible = bVisible
  Nome_Pesq1.Visible = bVisible
  
  If gsPesq2 = "" Then
    bVisible = False
  Else
    bVisible = True
    Label_Pesq2.Caption = gsPesq2 + ":"
  End If
  Label_Pesq2.Visible = bVisible
  Combo_Pesq2.Visible = bVisible
  Nome_Pesq2.Visible = bVisible
  
  If gsPesq3 = "" Then
    bVisible = False
  Else
    bVisible = True
    Label_Pesq3.Caption = gsPesq3 + ":"
  End If
  Label_Pesq3.Visible = bVisible
  Combo_Pesq3.Visible = bVisible
  Nome_Pesq3.Visible = bVisible
  
  
  '13/08/2002 - mpdea
  'Atributo da pasta WEB para envio de informações da Pesquisa 123
  chkPesquisa123.Visible = gsPesq1 <> "" Or gsPesq2 <> "" Or gsPesq3 <> ""
  
  
  Preços.Visible = rsParametros("Três Tabelas")
  
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", gnUserCode
  
  ActiveBar1.RecalcLayout
  ActiveBar1.Refresh
  
  Frame13.Visible = gsPesq1 <> "" Or gsPesq2 <> "" Or gsPesq3 <> ""
End Sub

'Private Sub LoadCombosProdutos()
'  cboCodigo.RemoveAll
'  With rsProdCod
'    .MoveFirst
'    Do While Not .EOF
'      cboCodigo.AddItem .Fields(0) & vbTab & .Fields(1)
'      .MoveNext
'    Loop
'  End With
'  cboNome.RemoveAll
'  With rsProdNome
'    .MoveFirst
'    Do While Not .EOF
'      cboNome.AddItem .Fields(0) & vbTab & .Fields(1)
'      .MoveNext
'    Loop
'  End With
'End Sub
'
Private Sub Fracionado_Click()
  If Fracionado.Value = 1 Then
    txtCasasDecimais.Enabled = True
  Else
    txtCasasDecimais.Enabled = True
    txtCasasDecimais.Text = gsHandleNull(txtCasasDecimais.Text & "")
    txtCasasDecimais.Enabled = False
  End If
End Sub

Private Sub Grade2_1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
 Dim Aux As Variant
 Dim Resp As Integer
 
  Call StatusMsg("")
  Aux = Grade2_1.Columns(ColIndex).Text
  If ColIndex = 1 Then  'Código
    If IsNull(Aux) Then
      Cancel = True
      Exit Sub
    End If
    If Not IsNumeric(Aux) Then
      Cancel = True
      Exit Sub
    End If
    If Val(Aux) > 99999999 Or Val(Aux) < 1 Then
      Cancel = True
      Exit Sub
    End If
    
    rsCliFor.Index = "Código"
    rsCliFor.Seek "=", Val(Aux)
    If rsCliFor.NoMatch Then
       MsgBox ("Código de fornecedor inválido.")
       Cancel = True
       Exit Sub
    End If
    
    Rem Verifica se este forn. já está neste produto
    rsForn_Prod.Index = "Produto"
    rsForn_Prod.Seek "=", rsProdutos2("Código"), Val(Aux)
    If Not rsForn_Prod.NoMatch Then
      MsgBox "Este fornecedor já está cadastrado neste produto.", vbExclamation
      Cancel = True
      Exit Sub
    End If
    
  End If
  
  Grade2_1.Columns(0).Text = cboCodigo.Text

End Sub

Private Sub Grade2_1_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  Dim Resposta As Integer
  
  Call StatusMsg("")
  Resposta = MsgBox("Deseja apagar o(s) fornecedor(es) selecionados ? ", 260, "Atenção")
  If Resposta = 7 Then
    DisplayMsg "Fornecedores não apagados."
    Cancel = True
  End If
  
  DispPromptMsg = False

End Sub

Private Sub Grade2_1_Click()
On Error GoTo Erro

  Dim bm As Variant
  Dim nRow As Long
  
  bm = Grade2_1.GetBookmark(0)
  lbl_fornCodigo.Caption = Grade2_1.Columns(1).CellValue(bm)
  lbl_fornNome.Caption = Grade2_1.Columns(2).CellValue(bm)
  
  If Grade2_1.Columns(1).CellValue(bm) = "" Or Grade2_1.Columns(1).CellValue(bm) = "0" Then
    ' Buscar código do produto do fornecedor na tabela de pareamento
    ' Se existir, então mostra na tela
    
    
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro no tratamento de pesquisa de Código do Produto do Fornecedor " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub Grade2_1_LostFocus()

 If Grade2_1.RowChanged = True Then
   Grade2_1.Update
 End If

End Sub

Private Sub medPreco_GotFocus(Index As Integer)
  medPreco(Index).SelStart = 0
  medPreco(Index).SelLength = medPreco(Index).MaxLength
End Sub

Private Sub medPreco_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub mskDataValidade_KeyDown(KeyCode As Integer, Shift As Integer)
'A tecla está pressionada para baixo
  If KeyCode = vbKeyF2 Then
    mskDataValidade.Text = frmCalendario.gsDateCalender(mskDataValidade.Text)
  End If
End Sub

Private Sub mskDataValidade_LostFocus()
  mskDataValidade.Text = Ajusta_Data(mskDataValidade.Text)
End Sub


'19/09/2005 - mpdea
'Validação do Índice para cálculo do Preço de Entrada
Private Sub mskIndicePrecoEntrada_Validate(Cancel As Boolean)
  Dim dblValue As Double
  
  If g_blnIndicePrecoEntrada Then
    If IsDataType(dtDouble, mskIndicePrecoEntrada.Text, dblValue) Then
      mskIndicePrecoEntrada.Text = dblValue
    Else
      DisplayMsg "Índice para cálculo do Preço de Entrada inválido."
      tabItens.Tab = 2
      tabComissaoIndices.Tab = 1
      mskIndicePrecoEntrada.SetFocus
      Cancel = True
    End If
  End If
End Sub

'----------------------------------------------------------------------------
'Data.............: 17/05/2006
'Desenvolvedor....: mpdea
'Solicitante......: Pavinato
'Descrição........: Validação do índice ICMS Retido - Base complementar
Private Sub mskIndiceIcmsRetido_Validate(Cancel As Boolean)
  Dim dblValue As Double
  
  If IsDataType(dtDouble, mskIndiceIcmsRetido.Text, dblValue) Then
    mskIndiceIcmsRetido.Text = dblValue
  Else
    DisplayMsg "Índice para cálculo do ICMS Retido - Base complementar inválido."
    tabItens.Tab = 2
    tabComissaoIndices.Tab = 1
    mskIndiceIcmsRetido.SetFocus
    Cancel = True
  End If
End Sub
'----------------------------------------------------------------------------

Private Sub Nome_Foto_Change()
  If Len(Trim(Nome_Foto.Text)) = 0 Then
    On Error Resume Next
    imgFoto.Picture = LoadPicture()
  End If
End Sub

Private Sub Nome_Foto_LostFocus()

  Call StatusMsg("")
  imgFoto.Picture = LoadPicture("")
  If IsNull(Nome_Foto.Text) Then
    Exit Sub
  ElseIf Nome_Foto.Text = "" Then
    Exit Sub
  End If
  
  On Error GoTo ErrFoto
  imgFoto.Picture = LoadPicture(gsConvertImagePath(Nome_Foto.Text))
  Exit Sub

ErrFoto:
  '07/12/2005 - mpdea
  'Modificado para que apresente a mensagem na barra de status
  StatusMsg "Não foi possível abrir a imagem. Verifique o nome e o formato do arquivo indicado."

End Sub

Private Sub Nome_Nota_Change()
  Qtde_Carac.Caption = Len(Nome_Nota.Text)
End Sub

Private Sub O_Produto_Edição_Click()
  Fracionado.Value = 0
  Fracionado.Enabled = False
  txtCasasDecimais.Text = "0"
  txtCasasDecimais.Enabled = False
End Sub

Private Sub O_Produto_Grade_Click()
  Fracionado.Value = 0
  Fracionado.Enabled = False
  txtCasasDecimais.Text = "0"
  txtCasasDecimais.Enabled = False
End Sub

Private Sub O_Produto_Normal_Click()
  Fracionado.Enabled = True
  Fracionado.Value = 0
  txtCasasDecimais.Enabled = True
End Sub

Private Sub LoadFornec()
  Dim Cód As String
  Dim sSql As String

  'Arruma Grade
  Cód = cboCodigo.Text
 
  sSql = "SELECT Produto, Fornecedor, Cli_For.Nome, Cli_For.[Fone 1] FROM [Forn_Prod]"
  sSql = sSql + " INNER JOIN Cli_For ON Forn_Prod.Fornecedor = Cli_For.Código  WHERE Produto ='" + Cód + "'"
 
  Call StatusMsg("Aguarde, montando tabela.")
  
  Set Rec_Fornec = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade2_1.Visible = False
  Grade2_1.DataMode = 1

  Set Data2_1.Recordset = Rec_Fornec

  Grade2_1.DataMode = 0

  Grade2_1.ReBind
 
  Grade2_1.Columns(0).Visible = False 'Cod Prod
  Grade2_1.Columns(1).Width = 1200 'Cod Forn
  Grade2_1.Columns(2).Locked = True 'Nome Cli
  Grade2_1.Columns(2).Width = 5800
  Grade2_1.Columns(3).Locked = True 'Tel Cli
  Grade2_1.Columns(3).Width = 2000
  
  Grade2_1.Visible = True
  
  Grade2_1.Columns(1).DropDownHwnd = DropDown2_1.hwnd
  
'  Nome_Foto_LostFocus
  
  Call StatusMsg("")

End Sub


'11/03/2011 - Andrea
'CSO SN - Código da Situacao Tributaria - Simples Nacional
Private Sub CSO_GotFocus()
  'CSO.SelStart = 0
  'CSO.SelLength = CSO.MaxLength
End Sub


Private Sub tabItens_Click(PreviousTab As Integer)

  If Not IsNull(Num_Registro) Then
    If PreviousTab = 1 Then
      Grade2_1.MoveNext
      Grade2_1.MovePrevious
    End If
    Select Case tabItens.Tab
      Case 1
        Call LoadFornec
      Case 4
        Calcula_Custo1
        Calcula_Venda
      Case 5
'        Call Nome_Foto_LostFocus
    End Select
'    Call ShowRecord
  End If
  
End Sub

'Private Sub txtBonus_KeyPress(KeyAscii As Integer)
'  KeyAscii = gnSomenteNumero(KeyAscii)
'End Sub

Private Sub txtCasasDecimais_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub txtDividirPrecoEtiqueta_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub txtPesoBruto_GotFocus()
  txtPesoBruto.SelStart = 0
  txtPesoBruto.SelLength = txtPesoBruto.MaxLength
End Sub

Private Sub txtPesoLiquido_GotFocus()
  txtPesoLiquido.SelStart = 0
  txtPesoLiquido.SelLength = txtPesoLiquido.MaxLength
End Sub

'03/06/2004 - mpdea
'Obtém o unidade de venda padrão para os produtos do tipo web (Loja virtual)
Private Function GetWebUnitSaleDefault() As String
  Dim rstX As Recordset
  Dim strSQL As String
  Dim strRet As String
  
  strSQL = "SELECT UnitSaleDefault FROM WEB_Config WHERE ID = 1;"
  Set rstX = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rstX
    If Not (.BOF And .EOF) Then
      strRet = .Fields("UnitSaleDefault").Value & ""
    End If
    .Close
  End With
  Set rstX = Nothing
  
  GetWebUnitSaleDefault = strRet
  
End Function

Private Sub txtVolumagem_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub CalcularCustoDoTecido()
  '04/07/2005 - Daniel
  'Case.....: Zue
  'Projeto..: Informações do Custo de Produto
  Dim dblConsumoDoTecido  As Double
  Dim dblPrecoMetroLinear As Double
  Dim dblCustoDoTecido    As Double
  
  dblConsumoDoTecido = 0
  dblPrecoMetroLinear = 0
  dblCustoDoTecido = 0
  
  On Error GoTo TratarErro
  
  If IsNumeric(txtCustoProduto(0).Text) Then dblConsumoDoTecido = Format(txtCustoProduto(0).Text, FORMAT_VALUE)
  If IsNumeric(txtCustoProduto(1).Text) Then dblPrecoMetroLinear = Format(txtCustoProduto(1).Text, FORMAT_VALUE)
  
  dblCustoDoTecido = Format(dblConsumoDoTecido * dblPrecoMetroLinear, FORMAT_VALUE)
  
  txtCustoProduto(2).Text = dblCustoDoTecido
  
  Exit Sub

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Sub


'''''-----------------------------------------------------
'''''15/12/2006 - Anderson
'''''Alteração para o cadastro de CFOP por produto
''''Private Sub LoadOperacao()
''''  Dim Cód As String
''''  Dim sSql As String
''''
''''  'Arruma Grade
''''  Cód = cboCodigo.Text
''''
''''    '30/03/2011 - Andrea
''''    'Incluído campo CSO
''''    sSql = "SELECT CodProduto, CodOperacao, [Operações Saída].Nome, CFOP, ProdutoCFOP.CSO FROM ProdutoCFOP "
''''    sSql = sSql + " INNER JOIN [Operações Saída] ON ProdutoCFOP.CodOperacao = [Operações Saída].Código  WHERE CodProduto ='" + Cód + "' "
''''
''''    GradeCFOP.Visible = False
''''    GradeCFOP.DataMode = 1
''''
''''    Set Rec_CFOP = db.OpenRecordset(sSql, dbOpenDynaset)
''''
''''    Set DataCFOP1.Recordset = Rec_CFOP
''''
''''  If Cód = "0" Then
''''
''''    GradeCFOP.Enabled = False
''''    'Exit Sub
''''  Else
''''    GradeCFOP.Enabled = True
''''
''''  End If
''''
''''  'GradeCFOP.Enabled = True
''''
''''  GradeCFOP.DataMode = 0
''''
''''  GradeCFOP.ReBind
''''
''''  GradeCFOP.Columns(0).Visible = False 'Cod Prod
''''  GradeCFOP.Columns(1).Width = 1000 'Cod Operação
''''  GradeCFOP.Columns(1).Caption = "Cód. Operação"
''''  GradeCFOP.Columns(2).Locked = True 'Nome da Operação
''''  GradeCFOP.Columns(2).Width = 4000
''''  GradeCFOP.Columns(3).Width = 1000 'CFOP
''''  '30/03/2011 - Andrea
''''  'GradeCFOP.Columns(4).Width = 1000 'CSO
''''
''''  GradeCFOP.Visible = True
''''  GradeCFOP.Columns(1).DropDownHwnd = DropDownCFOP.hwnd
''''  '-----------------------------------------------------
''''
'''''  Nome_Foto_LostFocus
''''
''''  Call StatusMsg("")
''''
''''End Sub



Private Sub FiltraProdutos()
  datProdCod.DatabaseName = gsQuickDBFileName

  '13/03/2013-Alexandre Afornali
  'Adicinar novo recordset para codigo com apenas produtos ativos
  Set rsProdutosFiltrados = db.OpenRecordset("SELECT * FROM [Produtos] WHERE Desativado = 0 ORDER BY [Código]")

  'Alterada a linha da abertura da tabela para case Infosocial de filtrar produtos inativos
  'Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial")
  'datProdCod.Refresh

  datProdNome.DatabaseName = gsQuickDBFileName
  '13/03/2013-Alexandre Afornali
  'Adicinar novo recordset para nome com apenas produtos ativos
  Set rsProdutosFiltradosNome = db.OpenRecordset("SELECT Nome, Código FROM Produtos WHERE Desativado = 0 ORDER BY Nome")
  'If para carregar ou nao produtos inativos
  If rsParametros("FiltrarProdutosInativos").Value = 0 Then
    Set datProdCod.Recordset = rsProdutos.Clone
  Else
    Set datProdCod.Recordset = rsProdutosFiltrados.Clone
  End If
    'Adicionar if para filtrar nome apenas de produtos
    If rsParametros("FiltrarProdutosInativos").Value = 0 Then
    Set datProdNome.Recordset = rsProdutosNome.Clone
  Else
    Set datProdNome.Recordset = rsProdutosFiltradosNome.Clone
  End If
End Sub


