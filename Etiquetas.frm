VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"Etiquetas.frx":0000
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Etiquetas.frx":00E3
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   14430
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Lista de produtos e quantidades de etiquetas para impressão"
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
      Height          =   3315
      Left            =   60
      TabIndex        =   16
      Top             =   4650
      Width           =   14310
      Begin VB.CommandButton cmdGerarEtiquetas 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ir para a tela que Formata suas etiquetas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11790
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Já esta com a sua lista pronta?  Então clique aqui para Formatar suas etiquetas"
         Top             =   2640
         Width           =   2445
      End
      Begin SSDataWidgets_B.SSDBGrid grdEtiquetasLista 
         Height          =   2820
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   11580
         ScrollBars      =   2
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
         Col.Count       =   5
         BevelColorFrame =   -2147483633
         BevelColorHighlight=   -2147483634
         BevelColorShadow=   -2147483634
         BevelColorFace  =   -2147483634
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   2
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorOdd    =   12640511
         RowHeight       =   423
         ExtraHeight     =   185
         Columns.Count   =   5
         Columns(0).Width=   4339
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "CodProd"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).Case =   2
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(0).PromptChar=   32
         Columns(1).Width=   8573
         Columns(1).Caption=   "Nome Produto"
         Columns(1).Name =   "NomeProd"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   2646
         Columns(2).Caption=   "Tamanho"
         Columns(2).Name =   "TamProd"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   3
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   2408
         Columns(3).Caption=   "Cor"
         Columns(3).Name =   "CorProd"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   3
         Columns(3).Locked=   -1  'True
         Columns(4).Width=   1376
         Columns(4).Caption=   "Qtde"
         Columns(4).Name =   "QtdeEtiq"
         Columns(4).Alignment=   1
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   6
         _ExtentX        =   20426
         _ExtentY        =   4974
         _StockProps     =   79
         Caption         =   "::"
         BackColor       =   -2147483634
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
      Begin VB.Label lbl_numeroEtiquetasDaLista 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "0 etiquetas na lista"
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
         Left            =   9420
         TabIndex        =   23
         Top             =   210
         Width           =   2265
      End
   End
   Begin VB.CommandButton cmd_limpaTela 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Limpar a tela"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   11820
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Limpar a tela para uma nova busca de produtos"
      Top             =   60
      Width           =   2535
   End
   Begin VB.Data datSubclasse 
      Caption         =   "datSubclasse"
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
      Left            =   4710
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM [Sub Classes] ORDER BY Código"
      Top             =   7710
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data datInst 
      Caption         =   "Funcionário"
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
      Height          =   420
      Left            =   7050
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Apelido"
      Top             =   7680
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Data datClasse 
      Caption         =   "Classe"
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
      Height          =   420
      Left            =   9180
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   7710
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Buscar produtos para indicar a quantidade de etiquetas a ser impressa"
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
      Height          =   4080
      Left            =   60
      TabIndex        =   3
      Top             =   510
      Width           =   14310
      Begin VB.CommandButton cmd_pesquisarProduto 
         Height          =   405
         Left            =   4590
         Picture         =   "Etiquetas.frx":4EA3D
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_pesquisarProdutos 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Buscar produtos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Buscar produtos pelos critérios de pesquisa"
         Top             =   330
         Width           =   2445
      End
      Begin VB.CommandButton cmd_salvarQtde 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salvar na lista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Salvar na sua lista os produtos que você selecionou e digitou a Quantidade de etiquetas que deseja imprimir"
         Top             =   1410
         Width           =   2445
      End
      Begin VB.CommandButton cmd_limparTabelaDeEtiquetas 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Apagar a sua lista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Apagar a sua lista de etiquetas para impressão"
         Top             =   3510
         Width           =   2445
      End
      Begin VB.CommandButton cmd_ApagarUmProdutoDaLista 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Excluir selecionado da lista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Tirar/Excluir da sua lista de etiquetas o produto selecionado"
         Top             =   2910
         Width           =   2445
      End
      Begin VB.TextBox txt_sequenciaEntrada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   10290
         TabIndex        =   10
         ToolTipText     =   "Digite a sequência de entrada"
         Top             =   630
         Width           =   1380
      End
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1230
         TabIndex        =   1
         ToolTipText     =   "Digite o código ou parte do código principal"
         Top             =   270
         Width           =   3315
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "Etiquetas.frx":4F67F
         Height          =   345
         Left            =   6000
         TabIndex        =   0
         Top             =   270
         Width           =   870
         DataFieldList   =   "Código"
         ListAutoValidate=   0   'False
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
         Columns(0).Width=   8837
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1984
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1535
         _ExtentY        =   609
         _StockProps     =   93
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
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubclasse 
         Bindings        =   "Etiquetas.frx":4F697
         Height          =   345
         Left            =   6000
         TabIndex        =   2
         Top             =   675
         Width           =   870
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
         BackColorOdd    =   16777152
         Columns(0).Width=   3200
         _ExtentX        =   1535
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
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
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBGrid grdEtiquetas 
         Height          =   2880
         Left            =   90
         TabIndex        =   14
         Top             =   1080
         Width           =   11580
         ScrollBars      =   2
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
         Col.Count       =   5
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   2
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   3
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   185
         Columns.Count   =   5
         Columns(0).Width=   4339
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "CodProd"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).Case =   2
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(0).PromptChar=   32
         Columns(1).Width=   8573
         Columns(1).Caption=   "Nome Produto"
         Columns(1).Name =   "NomeProd"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   2646
         Columns(2).Caption=   "Tamanho"
         Columns(2).Name =   "TamProd"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   3
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   2408
         Columns(3).Caption=   "Cor"
         Columns(3).Name =   "CorProd"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   3
         Columns(3).Locked=   -1  'True
         Columns(4).Width=   1376
         Columns(4).Caption=   "Qtde"
         Columns(4).Name =   "QtdeEtiq"
         Columns(4).Alignment=   1
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   6
         _ExtentX        =   20426
         _ExtentY        =   5080
         _StockProps     =   79
         Caption         =   "::"
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
      Begin VB.Label lbl_nomeProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   690
         Width           =   4965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sequência Entrada"
         Height          =   195
         Left            =   10290
         TabIndex        =   9
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lblNomeSubclasse 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   6915
         TabIndex        =   8
         Top             =   675
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Subclasse"
         Height          =   195
         Left            =   5220
         TabIndex        =   7
         Top             =   750
         Width           =   705
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Produto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   330
         Width           =   1110
      End
      Begin VB.Label Nome_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   6915
         TabIndex        =   5
         Top             =   270
         Width           =   3135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Classe"
         Height          =   195
         Left            =   5220
         TabIndex        =   4
         Top             =   330
         Width           =   465
      End
   End
   Begin SSDataWidgets_B.SSDBCombo cboInst 
      Bindings        =   "Etiquetas.frx":4F6B2
      Height          =   345
      Left            =   990
      TabIndex        =   11
      Top             =   60
      Width           =   1080
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   4921
      Columns(0).Caption=   "Apelido"
      Columns(0).Name =   "Apelido"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Apelido"
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
      _ExtentX        =   1905
      _ExtentY        =   609
      _StockProps     =   93
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
   End
   Begin VB.Label Nome_func 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2130
      TabIndex        =   13
      Top             =   60
      Width           =   4800
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Funcionário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   60
      TabIndex        =   12
      Top             =   135
      Width           =   915
   End
End
Attribute VB_Name = "frmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gsCaption As String
Private gnCodInst As Long
Private gsCodClasse As String
Private gbChanged As Boolean
Private gsOrder As String
Private gnColuna As Integer
Private gbFindTextActived As Boolean

Private Type FindEtiqueta
  sCodigo As String
  snome As String
End Type

Private tabEtiqueta() As FindEtiqueta
Private gnCount As Long

Private Sub cboInst_LostFocus()
    LoadGridEtiquetasLista
End Sub

Private Sub cboSubClasse_CloseUp()
  cboSubClasse.Text = cboSubClasse.Columns(0).Text
  cboSubClasse_LostFocus
End Sub

Private Sub cboSubClasse_LostFocus()
  Dim rstSubClasse As Recordset
  
  lblNomeSubClasse.Caption = ""

  If Not IsNumeric(cboSubClasse.Text) Then Exit Sub
  
  Set rstSubClasse = db.OpenRecordset("SELECT Nome FROM [Sub Classes] WHERE Código = " & CInt(cboSubClasse.Text), dbOpenDynaset)
  
  With rstSubClasse
    If Not (.BOF And .EOF) Then
      lblNomeSubClasse.Caption = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  
  Set rstSubClasse = Nothing
  
End Sub

Private Sub cboClasse_CloseUp()
  cboClasse.DataFieldList = "Código"
  cboClasse.Text = cboClasse.Columns("Código").Text
  Call cboClasse_Validate(True)
End Sub

Private Sub cboClasse_DropDown()
  cboClasse.DataFieldList = "Código"
End Sub

Private Sub cboClasse_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboClasse_KeyPress(KeyAscii As Integer)
  Dim bCancel As Boolean

  If cboClasse.DroppedDown Then
    cboClasse.DataFieldList = "Nome"
  End If
  If KeyAscii = vbKeyReturn Then
    If Nome_func.Caption = "" Then
      cboInst.SetFocus
    Else
      Call cboClasse_Validate(bCancel)
      If Not bCancel Then
        'Pesquisar
        Call SearchEtiquetas
      End If
    End If
  ElseIf Len(cboClasse.Text) >= 4 Then
    If KeyAscii <> vbKeyBack And Not cboClasse.DroppedDown Then
      Beep
      KeyAscii = 0
      Exit Sub
    End If
  End If
End Sub

Private Sub cboClasse_LostFocus()
  Call StatusMsg("")
End Sub

Private Sub cboClasse_Validate(Cancel As Boolean)
  If cboClasse.DroppedDown Then
    cboClasse.DroppedDown = False
  End If
  If cboClasse.Text = "" Or cboClasse.Text = "0" Then
    Nome_Classe.Caption = ""
  Else
    If IsNumeric(cboClasse.Text) Then
      datClasse.Recordset.FindFirst "Código = " & CInt(cboClasse.Text)
      If Not datClasse.Recordset.NoMatch Then
        Nome_Classe.Caption = datClasse.Recordset(0).Value
      Else
        Beep
        cboClasse.Text = ""
        Nome_Classe.Caption = ""
        Cancel = True
      End If
    Else
      Beep
      cboClasse.Text = ""
      Nome_Classe.Caption = ""
      Cancel = True
    End If
  End If
End Sub

Private Sub cboInst_CloseUp()
  Dim bCancel As Boolean
  
  cboInst.DataFieldList = "Código"
  cboInst.Text = cboInst.Columns("Código").Text
  Call cboInst_Validate(bCancel)
  If Not bCancel Then
    cboClasse.SetFocus
  End If
End Sub

Private Sub cboInst_DropDown()
  cboInst.DataFieldList = "Código"
End Sub

Private Sub cboInst_KeyPress(KeyAscii As Integer)
  If cboInst.DroppedDown Then
    cboInst.DataFieldList = "Apelido"
  End If
  If Len(cboInst.Text) >= 8 Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
      Exit Sub
    End If
  End If
End Sub

Private Sub ClearScreen()
  If gbChanged = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja inicializar a tela sem gravar atualizações?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Exit Sub
    End If
  End If
'''  gnCodInst = 0
  cboInst.Enabled = True
'''  cboInst.Text = ""
'''  Nome_func.Caption = ""
  txt_sequenciaEntrada.Text = ""
  lbl_numeroEtiquetasDaLista.Caption = "0 etiquetas na lista"
  lbl_NomeProduto.Caption = ""
  
  gsCodClasse = "<> -1"
  cboClasse.Text = ""
  cboClasse.Enabled = True
  Nome_Classe.Caption = ""
  '13/06/2005 - Daniel
  'Tratamento para o campo Subclasse
  cboSubClasse.Text = ""
  cboSubClasse.Enabled = True
  lblNomeSubClasse.Caption = ""
  '---------------------------------
  txt_sequenciaEntrada.Enabled = True
  txtProduto.Text = ""
  txtProduto.Enabled = True
  grdEtiquetas.RemoveAll
  grdEtiquetas.Caption = ""
  '''ActiveBar1.Tools("miOpFindText").Text = ""
  '''cboInst.SetFocus
  txtProduto.SetFocus
  gbChanged = False

'''  grdEtiquetasLista.RemoveAll
'''  grdEtiquetasLista.Caption = ""

End Sub

Private Sub SearchEtiquetas()
  If gbChanged = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja pesquisar sem gravar atualizações?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Exit Sub
    End If
  End If
  Call LoadGridEtiquetas
  Call RefreshQtde
End Sub

Private Sub cboInst_Validate(Cancel As Boolean)
  If cboInst.DroppedDown Then
    cboInst.DroppedDown = False
  End If
  If cboInst.Text <> "" Then
    If IsNumeric(cboInst.Text) Then
      datInst.Recordset.FindFirst "Código = " & CLng(cboInst.Text)
      If Not datInst.Recordset.NoMatch Then
        Nome_func.Caption = datInst.Recordset(0).Value
        gnCodInst = CLng(cboInst.Text)
      Else
        Beep
        cboInst.Text = ""
        Nome_func.Caption = ""
        gnCodInst = 0#
        Cancel = True
      End If
    Else
      Beep
      cboInst.Text = ""
      Nome_func.Caption = ""
      gnCodInst = 0#
      Cancel = True
    End If
  Else
    Nome_func.Caption = ""
    gnCodInst = 0#
  End If
End Sub

Private Sub cmd_ApagarUmProdutoDaLista_Click()
    Call DeleteRecord
End Sub

Private Sub cmd_limparTabelaDeEtiquetas_Click()
On Error GoTo Erro
  
  If MsgBox("Deseja realmente apagar a lista de produtos e quantidades de etiquetas para impressão que você criou?", vbQuestion + vbYesNo) = vbYes Then
      Call StatusMsg("Apagando a tabela com produtos e quantidades de etiquetas para impressão...")
      MousePointer = vbHourglass
      
      If Len(Nome_func.Caption) > 0 Then
          db.Execute "DELETE * FROM [Etiquetas - Tempo] WHERE Funcionario = " & CInt(cboInst.Text)
          db.Execute "DELETE * FROM Etiquetas WHERE Funcionário = " & CInt(cboInst.Text)
          MsgBox "Tabela apagada com sucesso", vbInformation, "Sucesso"
          
          Call ClearScreen
      Else
          If MsgBox("Você esta optou por apagar a lista com produtos para todos os usuários. Deseja continuar e apagar?", vbQuestion + vbYesNo) = vbYes Then
              db.Execute "DELETE * FROM [Etiquetas - Tempo] "
              db.Execute "DELETE * FROM Etiquetas"
          
              MsgBox "Tabela apagada com sucesso", vbInformation, "Sucesso"
              Call ClearScreen
          End If
      End If
      
      Call StatusMsg("")
      MousePointer = vbDefault
  End If

  LoadGridEtiquetasLista
  
  Exit Sub
Erro:
    MsgBox "Erro ao realizar a limpeza/exclusão da tabela de etiquetas para impressão " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cmd_limpaTela_Click()
    Call ClearScreen
    
    txtProduto.BackColor = &HC0FFFF
    txt_sequenciaEntrada.BackColor = &HC0FFFF
    Label1.ForeColor = &H0&
    Label5.ForeColor = &H0&
    Label3.ForeColor = &H0&
    Label4.ForeColor = &H0&
    Nome_Classe.BackColor = &H80FFFF
    lblNomeSubClasse.BackColor = &H80FFFF
    cboClasse.BackColor = &HC0FFFF
    cboSubClasse.BackColor = &HC0FFFF
End Sub

Private Sub cmd_pesquisarProduto_Click()
    nChamaConsulta = 7
    frmPesquisaProduto.Show
End Sub

Private Sub cmd_pesquisarProdutos_Click()
    SearchEtiquetas
End Sub

Private Sub cmd_salvarQtde_Click()
    Call WriteGridEtiquetas
    LoadGridEtiquetasLista
End Sub

Private Sub cmdGerarEtiquetas_Click()
On Error GoTo Erro
  Dim objTela As frmImprimeEtiq
  
  Set objTela = New frmImprimeEtiq
  
  If cboInst.Text <> "" Then
      objTela.ParamCodigoUsuario = cboInst.Text
  End If
  objTela.Show
  
  Exit Sub
Erro:
  MsgBox "Erro no acionamento deste botão, descrição " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
      'Call objHelp.Show(strfile, "QuickStore10Help")
      Call objHelp.Show(strfile, "QuickStore10Help", 10008)
      Set objHelp = Nothing
  Else
      Call HandleKeyDown(KeyCode, Shift)
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  Dim rsParametros As Recordset
  
  KeyPreview = True
  
  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  gsOrder = "ORDER BY [Código Ordenação]"
  
'''  With ActiveBar1.Tools("miOpOrdem")
'''    .CBList.Clear
'''    .CBList.InsertItem 0, "Por Código"
'''    .CBList.InsertItem 1, "Por Nome"
'''    .Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
'''  End With
'''  ActiveBar1.Tools("miOpFindText").ToolTipText = "Digite o código completo do produto"
'''  ActiveBar1.RecalcLayout
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial")
  'rsParametros.Index = "PK_ParâmetrosFilial"
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  'rsParametros.MoveFirst
  'rsParametros.Find "[Filial] = " & gnCodFilial
  gbGrade = rsParametros("Usar Grade") = True
  rsParametros.Close
  Set rsParametros = Nothing
  
  datInst.DatabaseName = gsQuickDBFileName
  datClasse.DatabaseName = gsQuickDBFileName
  '08/06/2005 - Daniel
  'Adição do filtro Subclasse
  datSubClasse.DatabaseName = gsQuickDBFileName
  
  gnCodInst = 0
  gsCodClasse = "<> -1"
  gsCaption = "Número Atual de Etiquetas para este Funcionário = "
  
  ReDim tabEtiqueta(0)
  
  Screen.MousePointer = vbDefault
  
'''  cboInst.Text = gnUserCode
End Sub

Private Sub RefreshQtde()
  Dim sSql As String
  Dim rsEtiquetas As Recordset
  If gnCodInst <= 0 Then
    Exit Sub
  End If
  sSql = "SELECT Sum(Qtde) As ValQtde FROM Etiquetas "
  sSql = sSql & " WHERE Funcionário = " & CStr(gnCodInst)
  Set rsEtiquetas = db.OpenRecordset(sSql, dbOpenSnapshot)
  If rsEtiquetas.EOF And rsEtiquetas.BOF Then
    grdEtiquetas.Caption = gsCaption & "0"
  Else
    grdEtiquetas.Caption = gsCaption & rsEtiquetas("ValQtde")
  End If
  rsEtiquetas.Close
  Set rsEtiquetas = Nothing
End Sub

Private Sub DeleteRecord()
  Dim Resposta As Integer
  Dim Num_Registro2 As Variant
  If grdEtiquetas.SelBookmarks.Count = 0 Then
    DisplayMsg "Selecione linhas para apagar."
    Exit Sub
  End If
  grdEtiquetas.DeleteSelected
  Call WriteGridEtiquetas
  Call LoadGridEtiquetas
  Call RefreshQtde
  LoadGridEtiquetasLista
End Sub

'''Private Sub DeleteAllRecord()
'''  Dim sSql As String
'''  Dim rsEtiquetas As Recordset
'''  Dim bCancel As Boolean
'''
'''  Call cboInst_Validate(bCancel)
'''  If bCancel Or gnCodInst <= 0 Then
'''    DisplayMsg "Entre com um funcionário."
'''    cboInst.SetFocus
'''    Exit Sub
'''  End If
'''
'''  If gnUserCode <> gnCodInst Then
'''    If Not frmGerente.gbSenhaGerente Then
'''      Exit Sub
'''    End If
'''  End If
'''
'''  sSql = "SELECT * FROM Etiquetas "
'''  sSql = sSql & " WHERE Funcionário = " & CStr(gnCodInst)
'''  Set rsEtiquetas = db.OpenRecordset(sSql, dbOpenSnapshot)
'''  If rsEtiquetas.EOF And rsEtiquetas.BOF Then
'''    DisplayMsg "Não existe registros de etiquetas para apagar."
'''    Exit Sub
'''  End If
'''
'''  gsTitle = LoadResString(201)
'''  gsMsg = "Deseja realmente apagar TODAS as etiquetas do funcionário atual?"
'''  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
'''  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'''  If gnResponse = vbYes Then
'''    Screen.MousePointer = vbHourglass
'''    grdEtiquetas.RemoveAll
'''    sSql = "DELETE * FROM Etiquetas WHERE Funcionário = " & CStr(gnCodInst)
'''    db.Execute (sSql)
'''
'''
'''
'''    '----------------------------------------------------------------------
'''    '03/12/2002 - mpdea
'''    'Modificado atualização da tela após apagar todas as etiquetas do funcionário
'''    grdEtiquetas.Caption = gsCaption & "0"
'''    Call ClearScreen
''''    Call WriteGridEtiquetas
''''    Call LoadGridEtiquetas
''''    Call RefreshQtde
'''    '----------------------------------------------------------------------
'''
'''
'''    Screen.MousePointer = vbDefault
'''    DisplayMsg "Todas as etiquetas foram apagadas do funcionário " & Nome_func
'''  End If
'''
'''End Sub

Private Sub LoadGridEtiquetas()
  Dim rsProdutos As Recordset
  Dim rsGrade As Recordset
  Dim rsEtiquetas As Recordset
  Dim sRecord As String
  Dim sSql As String
  Dim sCodInst As String
  Dim sCodClasse As String
  Dim bAllow As Boolean
  Dim sCodAnt As String
  Dim nTam As Integer
  Dim nCor As Integer
  Dim nEtiquetas As Integer
  Dim bCancel As Boolean
  '08/06/2005 - Daniel
  'Adição do filtro Subclasse
  Dim strSubClasse As String
  
  On Error GoTo ErrHandler
  
  Call cboInst_Validate(bCancel)
  If bCancel Or gnCodInst <= 0 Then
    DisplayMsg "Entre com um funcionário."
    cboInst.SetFocus
    Exit Sub
  End If
  
  Call cboClasse_Validate(bCancel)
  If bCancel Then
    DisplayMsg "Classe incorreta."
    cboClasse.SetFocus
    Exit Sub
  End If
  
  If Val(cboClasse.Text) > 0 Then
    gsCodClasse = "= " & CStr(CInt(cboClasse.Text))
  Else
    gsCodClasse = "<> -1"
  End If
  
  Call StatusMsg("Aguarde...")
  
  grdEtiquetas.CancelUpdate
  
  bAllow = grdEtiquetas.AllowAddNew
  grdEtiquetas.AllowAddNew = True
  grdEtiquetas.AllowUpdate = True
  
  '08/06/2005 - Daniel
  'Adição do filtro Subclasse através da variável strSubclasse
  strSubClasse = ""
  If Len(lblNomeSubClasse.Caption) > 0 Then strSubClasse = " AND [Sub Classe] = " & CInt(cboSubClasse.Text) & " "
  
  If txt_sequenciaEntrada.Text <> "" Then
      sSql = "SELECT E.Código, E.Qtde, E.[Código sem Grade], P.Nome FROM [Entradas - Produtos] E, Produtos P "
      sSql = sSql & " WHERE E.Filial = " & gnCodFilial
      sSql = sSql & " AND E.Sequência = " & txt_sequenciaEntrada.Text
      sSql = sSql & " AND E.[Código sem Grade] = P.Código "
  Else
      sSql = "SELECT Código, Nome FROM Produtos WHERE Classe " & gsCodClasse
      sSql = sSql & " AND Código <> '0' "
      
      If Trim(txtProduto.Text) <> "" Then
          sSql = sSql & " AND Código LIKE '" & Trim(txtProduto.Text) & "*'"
      End If
      
      sSql = sSql & " AND Desativado = False " & strSubClasse & gsOrder
  End If
  
  Set rsProdutos = db.OpenRecordset(sSql, dbOpenSnapshot)
  
  If rsProdutos.EOF And rsProdutos.BOF Then
    DisplayMsg "Nenhum produto cadastrado com as informações fornecidas."
    rsProdutos.Close
    Set rsProdutos = Nothing
    Call StatusMsg("")
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  txtProduto.BackColor = &HE0E0E0
  txt_sequenciaEntrada.BackColor = &HE0E0E0
  Label1.ForeColor = &H808080
  Label5.ForeColor = &H808080
  Label3.ForeColor = &H808080
  Label4.ForeColor = &H808080
  Nome_Classe.BackColor = &HE0E0E0
  lblNomeSubClasse.BackColor = &HE0E0E0
  cboClasse.BackColor = &HE0E0E0
  cboSubClasse.BackColor = &HE0E0E0
  
  cboInst.Enabled = False
  cboClasse.Enabled = False
  '13/06/2005 - Daniel
  'Tratamento para Subclasse
  cboSubClasse.Enabled = False
  '---------------------------
  txtProduto.Enabled = False
  txt_sequenciaEntrada.Enabled = False
  
  sSql = "SELECT * FROM Etiquetas"
  sSql = sSql & " WHERE Funcionário = " & CStr(gnCodInst)
  Set rsEtiquetas = db.OpenRecordset(sSql, dbOpenSnapshot)
  
  grdEtiquetas.RemoveAll
  grdEtiquetas.Redraw = False
  
  ReDim tabEtiqueta(0)
  gnCount = 0
  
  If txt_sequenciaEntrada.Text <> "" Then
      If Not rsProdutos.EOF Then
        With rsProdutos
          .MoveLast
          .MoveFirst
          Do Until .EOF
            If gbHasGrade(![Código Sem grade]) Then
              sSql = "SELECT * FROM [Códigos da Grade]"
              sSql = sSql & " WHERE [Código Original] = '" & CStr(![Código Sem grade]) & "'"
              Set rsGrade = db.OpenRecordset(sSql, dbOpenSnapshot)
              If rsGrade.RecordCount > 0 Then
                Do Until rsGrade.EOF
                  
                  '18/08/2005 - mpdea
                  'Corrigido obtenção do código da grade do produto (tamanho fixo = 6:
                  'tamanho = 3 + cor = 3)
                  'A função comentada errava em casos como o código 1001001001
                  'em que o código principal aparecia no código da grade e por vez
                  'era substituído
                  '
                  'sCodAnt = Replace(rsGrade("Código"), rsGrade("Código Original"), "")
                  If (!Código) = rsGrade("Código") Then
                      sCodAnt = Right(rsGrade("Código"), 6)
                  Else
                      sCodAnt = ""
                  End If
                  '''sCodAnt = Right(rsGrade("Código"), 6)
                  
                  If sCodAnt <> "" Then
                      nTam = Left(sCodAnt, 3)
                      nCor = Right(sCodAnt, 3)
                      rsEtiquetas.FindFirst "Produto = '" & rsGrade("Código Original") & _
                        "' AND Tamanho = " & nTam & " AND Cor = " & nCor
                      nEtiquetas = (!Qtde)
'''                      If Not rsEtiquetas.NoMatch Then
'''                        nEtiquetas = rsEtiquetas("Qtde")
'''                      Else
'''                        nEtiquetas = 0
'''                      End If
                      sRecord = rsGrade("Código original") & vbTab & !Nome & vbTab & Format(nTam, "000") & _
                        " - " & gsGetNameTamanho(nTam) & vbTab & Format(nCor, "000") & _
                        " - " & gsGetNameCor(nCor) & vbTab & nEtiquetas
                      grdEtiquetas.AddItem sRecord
                      ReDim Preserve tabEtiqueta(gnCount)
                      With tabEtiqueta(gnCount)
                        .sCodigo = rsGrade!Código
                        .snome = rsProdutos!Nome
                      End With
                      gnCount = gnCount + 1
                      rsGrade.MoveNext
                  Else
                      rsGrade.MoveNext
                  End If
                Loop
              End If
            Else
              rsEtiquetas.FindFirst "Produto = '" & .Fields("Código") & "'"
              nEtiquetas = (!Qtde)
'''              If Not rsEtiquetas.NoMatch Then
'''                nEtiquetas = rsEtiquetas("Qtde")
'''              Else
'''                nEtiquetas = 0
'''              End If
              sRecord = .Fields("Código") & vbTab & .Fields("nome") & vbTab & "" & _
                vbTab & "" & vbTab & nEtiquetas
              grdEtiquetas.AddItem sRecord
              ReDim Preserve tabEtiqueta(gnCount)
              With tabEtiqueta(gnCount)
                .sCodigo = rsProdutos!Código
                .snome = rsProdutos!Nome
              End With
              gnCount = gnCount + 1
            End If
            .MoveNext
          Loop
          .MoveFirst
        End With
        grdEtiquetas.Scroll -99, -99
      End If
  
  Else
      If Not rsProdutos.EOF Then
        With rsProdutos
          .MoveLast
          .MoveFirst
          Do Until .EOF
            If gbHasGrade(!Código) Then
              sSql = "SELECT * FROM [Códigos da Grade]"
              sSql = sSql & " WHERE [Código Original] = '" & CStr(!Código) & "'"
              Set rsGrade = db.OpenRecordset(sSql, dbOpenSnapshot)
              If rsGrade.RecordCount > 0 Then
                Do Until rsGrade.EOF
                  
                  '18/08/2005 - mpdea
                  'Corrigido obtenção do código da grade do produto (tamanho fixo = 6:
                  'tamanho = 3 + cor = 3)
                  'A função comentada errava em casos como o código 1001001001
                  'em que o código principal aparecia no código da grade e por vez
                  'era substituído
                  '
                  'sCodAnt = Replace(rsGrade("Código"), rsGrade("Código Original"), "")
                  sCodAnt = Right(rsGrade("Código"), 6)
                  
                  nTam = Left(sCodAnt, 3)
                  nCor = Right(sCodAnt, 3)
                  rsEtiquetas.FindFirst "Produto = '" & rsGrade("Código Original") & _
                    "' AND Tamanho = " & nTam & " AND Cor = " & nCor
                  If Not rsEtiquetas.NoMatch Then
                    nEtiquetas = rsEtiquetas("Qtde")
                  Else
                    nEtiquetas = 0
                  End If
                  sRecord = !Código & vbTab & !Nome & vbTab & Format(nTam, "000") & _
                    " - " & gsGetNameTamanho(nTam) & vbTab & Format(nCor, "000") & _
                    " - " & gsGetNameCor(nCor) & vbTab & nEtiquetas
                  grdEtiquetas.AddItem sRecord
                  ReDim Preserve tabEtiqueta(gnCount)
                  With tabEtiqueta(gnCount)
                    .sCodigo = rsGrade!Código
                    .snome = rsProdutos!Nome
                  End With
                  gnCount = gnCount + 1
                  rsGrade.MoveNext
                Loop
              End If
            Else
              rsEtiquetas.FindFirst "Produto = '" & .Fields("Código") & "'"
              If Not rsEtiquetas.NoMatch Then
                nEtiquetas = rsEtiquetas("Qtde")
              Else
                nEtiquetas = 0
              End If
              sRecord = .Fields("Código") & vbTab & .Fields("nome") & vbTab & "" & _
                vbTab & "" & vbTab & nEtiquetas
              grdEtiquetas.AddItem sRecord
              ReDim Preserve tabEtiqueta(gnCount)
              With tabEtiqueta(gnCount)
                .sCodigo = rsProdutos!Código
                .snome = rsProdutos!Nome
              End With
              gnCount = gnCount + 1
            End If
            .MoveNext
          Loop
          .MoveFirst
        End With
        grdEtiquetas.Scroll -99, -99
      End If
  End If
  
  
  With grdEtiquetas
    .Redraw = True
    .AllowAddNew = bAllow
    .AllowUpdate = bAllow
    .Refresh
  End With
  
  gbChanged = False
  
  rsEtiquetas.Close
  Set rsEtiquetas = Nothing
  rsProdutos.Close
  Set rsProdutos = Nothing
  rsGrade.Close
  Set rsGrade = Nothing
  
  
  Screen.MousePointer = vbDefault
  
  Call StatusMsg("")
  
  If txt_sequenciaEntrada.Text <> "" Then
      MsgBox "Está DE ACORDO com a quantidade de etiquetas a serem 'impressas' conforme listado na grade?" & vbCrLf & vbCrLf & "Caso sim, clique em SALVAR", vbInformation, "Atenção"
  End If
  
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Exit Sub

End Sub

Private Sub LoadGridEtiquetasLista()
  Dim rsGrade As Recordset
  Dim rsEtiquetas As Recordset
  Dim sRecord As String
  Dim sSql As String
  Dim sCodInst As String
  Dim sCodClasse As String
  Dim bAllow As Boolean
  Dim sCodAnt As String
  Dim nTam As Integer
  Dim nCor As Integer
  Dim nEtiquetas As Integer
  Dim nEtiquetasTotal As Integer
  Dim bCancel As Boolean
  Dim strSubClasse As String
  
  On Error GoTo ErrHandler
  
  Call cboInst_Validate(bCancel)
  If bCancel Or gnCodInst <= 0 Then
    DisplayMsg "Entre com um funcionário."
    cboInst.SetFocus
    Exit Sub
  End If
  
  nEtiquetasTotal = 0
  
  Call StatusMsg("Aguarde...")
  grdEtiquetasLista.CancelUpdate

'''  bAllow = grdEtiquetasLista.AllowAddNew
'''  grdEtiquetasLista.AllowAddNew = True
'''  grdEtiquetasLista.AllowUpdate = True

  Screen.MousePointer = vbHourglass

  sSql = "SELECT E.Produto, E.Tamanho, E.Cor, E.Qtde, P.Nome, P.Tipo FROM Etiquetas E, Produtos P "
  sSql = sSql & " WHERE E.Funcionário = " & CStr(gnCodInst) & " AND E.Qtde > 0 "
  sSql = sSql & " AND E.Produto = P.Código"
  Set rsEtiquetas = db.OpenRecordset(sSql, dbOpenSnapshot)
  
  grdEtiquetasLista.RemoveAll
  grdEtiquetasLista.Redraw = False

  If Not (rsEtiquetas.EOF And rsEtiquetas.BOF) Then
      rsEtiquetas.MoveLast
      rsEtiquetas.MoveFirst
      While Not rsEtiquetas.EOF
          If rsEtiquetas.Fields("Tipo") = "G" Then
              sSql = "SELECT * FROM [Códigos da Grade]"
              sSql = sSql & " WHERE [Código Original] = '" & rsEtiquetas.Fields("Produto") & "'"
              Set rsGrade = db.OpenRecordset(sSql, dbOpenSnapshot)
              If rsGrade.RecordCount > 0 Then
                  Do Until rsGrade.EOF
                      sCodAnt = Right(rsGrade("Código"), 6)
                      nTam = Left(sCodAnt, 3)
                      nCor = Right(sCodAnt, 3)
                      If rsEtiquetas.Fields("Produto").Value = rsGrade("Código Original") And _
                            rsEtiquetas.Fields("Tamanho").Value = nTam And rsEtiquetas.Fields("Cor").Value = nCor Then
                          nEtiquetas = rsEtiquetas("Qtde")

                          sRecord = rsEtiquetas.Fields("Produto") & vbTab & rsEtiquetas.Fields("Nome") & vbTab & Format(nTam, "000") & _
                            " - " & gsGetNameTamanho(nTam) & vbTab & Format(nCor, "000") & _
                            " - " & gsGetNameCor(nCor) & vbTab & nEtiquetas
                          grdEtiquetasLista.AddItem sRecord
                      End If
                      nEtiquetasTotal = nEtiquetasTotal + nEtiquetas
                      nEtiquetas = 0
                      rsGrade.MoveNext
                  Loop
              End If
              rsGrade.Close
              Set rsGrade = Nothing
          Else
              '''rsEtiquetas.FindFirst "Produto = '" & rsEtiquetas.Fields("Produto") & "'"
              '''If Not rsEtiquetas.NoMatch Then
                  nEtiquetas = rsEtiquetas("Qtde")
              '''Else
              '''    nEtiquetas = 0
              '''End If
              sRecord = rsEtiquetas.Fields("Produto") & vbTab & rsEtiquetas.Fields("nome") & vbTab & "" & _
                vbTab & "" & vbTab & nEtiquetas
              grdEtiquetasLista.AddItem sRecord
              gnCount = gnCount + 1
              
              nEtiquetasTotal = nEtiquetasTotal + nEtiquetas
              nEtiquetas = 0
              'rsGrade.MoveNext
          End If
          rsEtiquetas.MoveNext
          
      Wend
      rsEtiquetas.MoveFirst

      grdEtiquetasLista.Scroll -99, -99
  Else
      lbl_numeroEtiquetasDaLista.Caption = "0 etiquetas na lista"
  End If
  
  lbl_numeroEtiquetasDaLista.Caption = nEtiquetasTotal & " etiquetas na lista"

  With grdEtiquetasLista
    .Redraw = True
    .AllowAddNew = bAllow
    .AllowUpdate = bAllow
    .Refresh
  End With
  
  rsEtiquetas.Close
  Set rsEtiquetas = Nothing
  
'''  rsGrade.Close
'''  Set rsGrade = Nothing
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Exit Sub
End Sub

Private Sub WriteGridEtiquetas()
  Dim sCodInst As String
  Dim sCodClasse As String
  Dim sCodigo As String
  Dim nQtde As Long
  Dim bm As Variant
  Dim nRow As Long
  Dim sSql As String
  Dim rsEtiquetas As Recordset
  
  On Error GoTo ErrHandler
  
  If grdEtiquetas.Rows <= 0 Then
    DisplayMsg "Não há registro para atualizar."
    Exit Sub
  End If
  
  grdEtiquetas.Update
  
  Call ws.BeginTrans
  
  Call StatusMsg("Gravando ...")
    
  sSql = "SELECT * FROM Etiquetas INNER JOIN Produtos ON " & _
    "Produtos.Código = Etiquetas.Produto WHERE Produtos.Código <> '0' " & _
    "AND Código LIKE '" & txtProduto.Text & "*' AND Produtos.Classe " & _
    gsCodClasse & " AND Etiquetas.Funcionário = " & CStr(gnCodInst)
    
  Set rsEtiquetas = db.OpenRecordset(sSql, dbOpenDynaset)

  With rsEtiquetas
    'Apaga os registros
    Do Until .EOF
      .Delete
      .MoveNext
    Loop
    'Grava os novos registros
    For nRow = 0 To grdEtiquetas.Rows - 1
    
      bm = grdEtiquetas.AddItemBookmark(nRow)
      
      sCodigo = grdEtiquetas.Columns("CodProd").CellValue(bm)
      nQtde = CLng(gsHandleNull(grdEtiquetas.Columns("QtdeEtiq").CellValue(bm)))
      
      If nQtde > 0& Then
        .AddNew
        .Fields("Funcionário") = gnCodInst
        .Fields("Produto") = sCodigo
        If gbHasGrade(sCodigo) Then
          .Fields("Tamanho") = CInt(gsHandleNull(Left(grdEtiquetas.Columns("TamProd").CellValue(bm), 3)))
          .Fields("Cor") = CInt(gsHandleNull(Left(grdEtiquetas.Columns("CorProd").CellValue(bm), 3)))
        Else
          .Fields("Tamanho") = 0
          .Fields("Cor") = 0
        End If
        .Fields("Qtde") = nQtde
        .Update
      End If
      
    Next nRow
  End With
  
  Call RefreshQtde
  
  rsEtiquetas.Close
  Set rsEtiquetas = Nothing
  
  Call ws.CommitTrans
  MsgBox "Sua lista de etiquetas foi atualizada com sucesso !", vbInformation, "Sucesso"
  Call StatusMsg("")
  gbChanged = False

  
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Atualizar Cadastro de Etiquetas."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If gbChanged = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja sair sem gravar atualizações?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Cancel = True
    End If
  End If
End Sub

Private Sub grdEtiquetas_AfterColUpdate(ByVal ColIndex As Integer)
  If ColIndex = 4 And grdEtiquetas.Columns("Qtde").Text = "" Then
    grdEtiquetas.Columns("Qtde").Text = "0"
  End If
End Sub

Private Sub grdEtiquetas_AfterUpdate(RtnDispErrMsg As Integer)
  Dim nRow As Long
  Dim nQtde As Long
  Dim bm As Variant
  gbChanged = True
  nQtde = 0&
  For nRow = 0& To grdEtiquetas.Rows - 1
    bm = grdEtiquetas.AddItemBookmark(nRow)
    nQtde = nQtde + gsHandleNull(grdEtiquetas.Columns("Qtde").CellValue(bm))
  Next nRow
  grdEtiquetas.Caption = gsCaption & CStr(nQtde)
End Sub

Private Sub grdEtiquetas_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  Dim nX As Integer
  Dim vBook As Variant
  
  DispPromptMsg = False
  If Len(Trim(grdEtiquetas.ActiveCell.Text)) = 0 Then 'grdEtiquetas.SelBookmarks.Count = 0
    If bGridBeforeDelete() = True Then
      gbChanged = True
      Call RefreshQtde
      'Limpa o registro do produto na lista
      For nX = 0 To (grdEtiquetas.SelBookmarks.Count - 1)
        vBook = grdEtiquetas.SelBookmarks(nX)
        Call nFindRecordInTable(True, _
          grdEtiquetas.Columns("CodProd").CellValue(vBook) & _
          Left(grdEtiquetas.Columns("TamProd").CellValue(vBook), 3) & _
          Left(grdEtiquetas.Columns("CorProd").CellValue(vBook), 3), True)
      Next nX
      Cancel = False
    Else
      Cancel = True
    End If
  Else
    Cancel = True
  End If
End Sub

Private Sub grdEtiquetas_BeforeInsert(Cancel As Integer)
  Cancel = True
End Sub

Private Sub grdEtiquetas_InitColumnProps()
  With grdEtiquetas
    .Columns("TamProd").Visible = gbGrade
    .Columns("CorProd").Visible = gbGrade
    If Not gbGrade Then
      .Columns("NomeProd").Width = .Columns("NomeProd").Width + _
        .Columns("TamProd").Width + .Columns("CorProd").Width
    End If
  End With
End Sub

Private Sub grdEtiquetas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If grdEtiquetas.DroppedDown Then
      grdEtiquetas.DroppedDown = False
    Else
      ''Call MoveNext
      SendKeys "{Tab}"
    End If
  End If
End Sub

Private Sub grdEtiquetas_KeyPress(KeyAscii As Integer)
  If grdEtiquetas.Col = 4 Then 'Qtde
    KeyAscii = gnSomenteNumero(KeyAscii)
  End If
End Sub

Private Sub grdEtiquetas_Validate(Cancel As Boolean)
  If grdEtiquetas.RowChanged Then
    grdEtiquetas.Update
  End If
End Sub

Private Function nFindRecordInTable(ByVal bCode As Boolean, ByVal sKEY As String, Optional ByVal bDelete As Boolean) As Long
  Dim nX As Long
  Dim nC As Long
  
  For nX = LBound(tabEtiqueta) To UBound(tabEtiqueta)
    With tabEtiqueta(nX)
      If .sCodigo <> "" Then
        nC = nC + 1
        If (bCode And .sCodigo = sKEY) Or UCase(.snome) Like UCase(sKEY) & "*" Then
          nFindRecordInTable = nC
          If bDelete Then
            .sCodigo = ""
            .snome = ""
          End If
          Exit Function
        End If
      End If
    End With
  Next nX
  nFindRecordInTable = -1
End Function

Private Sub txtProduto_LostFocus()
    If Trim(txtProduto.Text) = "" Then
        lbl_NomeProduto.Caption = ""
    End If
End Sub
