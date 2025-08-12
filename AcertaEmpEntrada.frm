VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAcertaEmpEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerto de Empr�stimos (Entrada)"
   ClientHeight    =   6315
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   11295
   Icon            =   "AcertaEmpEntrada.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11295
   Begin VB.Frame Frame3 
      Caption         =   "Ordem"
      Height          =   795
      Left            =   8280
      TabIndex        =   44
      Top             =   0
      Width           =   1455
      Begin VB.OptionButton optOrdemProduto 
         Caption         =   "Produto"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optOrdemSequencia 
         Caption         =   "Sequ�ncia"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton B_Confirma_Mov 
      Caption         =   "&Confirmar "
      Height          =   400
      Left            =   9765
      TabIndex        =   22
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton B_Cancela_Mov 
      Caption         =   "Cancelar"
      Height          =   400
      Left            =   9765
      TabIndex        =   23
      Top             =   4500
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Verificar Datas"
      Height          =   400
      Left            =   90
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1335
   End
   Begin Threed.SSPanel L_Estoque 
      Height          =   285
      Left            =   6720
      TabIndex        =   42
      Top             =   5955
      Width           =   4560
      _Version        =   65536
      _ExtentX        =   8043
      _ExtentY        =   503
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Frame Frame4 
      Caption         =   "Detalhes"
      Height          =   1965
      Left            =   75
      TabIndex        =   41
      Top             =   3210
      Width           =   6525
      Begin SSDataWidgets_B.SSDBGrid Grade2 
         Bindings        =   "AcertaEmpEntrada.frx":058A
         Height          =   1560
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   6315
         _Version        =   196617
         RowHeight       =   423
         Columns.Count   =   7
         Columns(0).Width=   1402
         Columns(0).Caption=   "Data Opera��o"
         Columns(0).Name =   "Data Opera��o"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Data Opera��o"
         Columns(0).DataType=   7
         Columns(0).FieldLen=   256
         Columns(1).Width=   1640
         Columns(1).Caption=   "Ordem"
         Columns(1).Name =   "Ordem"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Ordem"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         Columns(2).Width=   1349
         Columns(2).Caption=   "Saldo Anterior"
         Columns(2).Name =   "Saldo Anterior"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Saldo Anterior"
         Columns(2).DataType=   3
         Columns(2).FieldLen=   256
         Columns(3).Width=   1508
         Columns(3).Caption=   "Vendas Cliente"
         Columns(3).Name =   "Vendas Cliente"
         Columns(3).Alignment=   1
         Columns(3).CaptionAlignment=   1
         Columns(3).DataField=   "Vendas Cliente"
         Columns(3).DataType=   3
         Columns(3).FieldLen=   256
         Columns(4).Width=   1746
         Columns(4).Caption=   "Devolu��o"
         Columns(4).Name =   "Devolu��o"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "Devolu��o"
         Columns(4).DataType=   3
         Columns(4).FieldLen=   256
         Columns(5).Width=   2302
         Columns(5).Caption=   "Novo Empr�stimo"
         Columns(5).Name =   "Novo Empr�stimo"
         Columns(5).Alignment=   1
         Columns(5).CaptionAlignment=   1
         Columns(5).DataField=   "Novo Empr�stimo"
         Columns(5).DataType=   3
         Columns(5).FieldLen=   256
         Columns(6).Width=   1693
         Columns(6).Caption=   "Saldo Atual"
         Columns(6).Name =   "Saldo Atual"
         Columns(6).Alignment=   1
         Columns(6).CaptionAlignment=   1
         Columns(6).DataField=   "Saldo Atual"
         Columns(6).DataType=   3
         Columns(6).FieldLen=   256
         _ExtentX        =   11139
         _ExtentY        =   2752
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
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2   'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1   'Dynaset
      RecordSource    =   ""
      Top             =   8340
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CheckBox O_Mostra_Detalhe 
      Caption         =   "Mostrar detalhes para cada linha"
      Height          =   225
      Left            =   90
      TabIndex        =   13
      Top             =   5235
      Value           =   1  'Checked
      Width           =   3165
   End
   Begin VB.Frame Frame_Mov 
      Caption         =   "Movimenta��o do Produto"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   6705
      TabIndex        =   35
      Top             =   3210
      Width           =   2625
      Begin VB.TextBox Emp_Prod 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1470
         MaxLength       =   100
         TabIndex        =   19
         Top             =   945
         Width           =   960
      End
      Begin VB.TextBox Dev_Prod 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1470
         MaxLength       =   100
         TabIndex        =   18
         Top             =   630
         Width           =   960
      End
      Begin VB.TextBox Vendas_Prod 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1470
         MaxLength       =   100
         TabIndex        =   17
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Saldo_Prod 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   1470
         TabIndex        =   40
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label Label9 
         Caption         =   "Saldo :"
         Height          =   225
         Left            =   105
         TabIndex        =   39
         Top             =   1365
         Width           =   1065
      End
      Begin VB.Label Label8 
         Caption         =   "Novo Empr�stimo :"
         Height          =   225
         Left            =   105
         TabIndex        =   38
         Top             =   990
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Devolu��es :"
         Height          =   225
         Left            =   105
         TabIndex        =   37
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Vendas :"
         Height          =   225
         Left            =   105
         TabIndex        =   36
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "&Imprimir Tabela"
      Height          =   400
      Left            =   1485
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Data_Acerto 
      Height          =   315
      Left            =   8040
      TabIndex        =   20
      ToolTipText     =   "Pressione F2 para Calend�rio"
      Top             =   4950
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.CommandButton B_Atualiza 
      Caption         =   "&Atualizar Total"
      Height          =   400
      Left            =   2880
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton B_Atualiza_Tudo 
      Caption         =   "&Gerar Entrada"
      Height          =   400
      Left            =   8175
      TabIndex        =   24
      ToolTipText     =   "Atualizar os Empr�stimos e Gerar Entrada com as Compras"
      Top             =   5415
      Width           =   1335
   End
   Begin VB.CommandButton B_Atualiza_Empr�stimos 
      Caption         =   "Atualizar &Empr�stimos"
      Height          =   400
      Left            =   9765
      TabIndex        =   25
      Top             =   5415
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2   'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2475
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1   'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   8355
      Visible         =   0   'False
      Width           =   1875
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Height          =   2235
      Left            =   75
      TabIndex        =   11
      Top             =   960
      Width           =   11160
      _Version        =   196617
      DataMode        =   1
      UseGroups       =   -1  'True
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   65280
      RowHeight       =   423
      Groups(0).Width =   18680
      Groups(0).Caption=   "Empr�stimos"
      Groups(0).Columns.Count=   14
      Groups(0).Columns(0).Width=   1217
      Groups(0).Columns(0).Caption=   "Seq."
      Groups(0).Columns(0).Name=   "Sequ�ncia"
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   8
      Groups(0).Columns(0).FieldLen=   256
      Groups(0).Columns(0).Locked=   -1  'True
      Groups(0).Columns(1).Width=   2117
      Groups(0).Columns(1).Caption=   "Produto"
      Groups(0).Columns(1).Name=   "Produto"
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   256
      Groups(0).Columns(1).Locked=   -1  'True
      Groups(0).Columns(2).Width=   3731
      Groups(0).Columns(2).Caption=   "Nome"
      Groups(0).Columns(2).Name=   "Nome"
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   8
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(2).Locked=   -1  'True
      Groups(0).Columns(3).Width=   1111
      Groups(0).Columns(3).Caption=   "Tam"
      Groups(0).Columns(3).Name=   "Tamanho"
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   2
      Groups(0).Columns(3).FieldLen=   256
      Groups(0).Columns(3).Locked=   -1  'True
      Groups(0).Columns(4).Width=   1058
      Groups(0).Columns(4).Caption=   "Cor"
      Groups(0).Columns(4).Name=   "Cor"
      Groups(0).Columns(4).DataField=   "Column 4"
      Groups(0).Columns(4).DataType=   2
      Groups(0).Columns(4).FieldLen=   256
      Groups(0).Columns(4).Locked=   -1  'True
      Groups(0).Columns(5).Width=   1032
      Groups(0).Columns(5).Caption=   "Ed."
      Groups(0).Columns(5).Name=   "Edi��o"
      Groups(0).Columns(5).DataField=   "Column 5"
      Groups(0).Columns(5).DataType=   3
      Groups(0).Columns(5).FieldLen=   256
      Groups(0).Columns(5).Locked=   -1  'True
      Groups(0).Columns(6).Width=   1323
      Groups(0).Columns(6).Caption=   "Ordem"
      Groups(0).Columns(6).Name=   "Ordem"
      Groups(0).Columns(6).DataField=   "Column 6"
      Groups(0).Columns(6).DataType=   3
      Groups(0).Columns(6).FieldLen=   256
      Groups(0).Columns(6).Locked=   -1  'True
      Groups(0).Columns(7).Width=   1852
      Groups(0).Columns(7).Caption=   "Data"
      Groups(0).Columns(7).Name=   "Data Opera��o"
      Groups(0).Columns(7).DataField=   "Column 7"
      Groups(0).Columns(7).DataType=   7
      Groups(0).Columns(7).FieldLen=   256
      Groups(0).Columns(7).Locked=   -1  'True
      Groups(0).Columns(8).Width=   1349
      Groups(0).Columns(8).Caption=   "$ Unit."
      Groups(0).Columns(8).Name=   "Pre�o Unit�rio"
      Groups(0).Columns(8).Alignment=   1
      Groups(0).Columns(8).DataField=   "Column 8"
      Groups(0).Columns(8).DataType=   5
      Groups(0).Columns(8).NumberFormat=   "##,##0.00"
      Groups(0).Columns(8).FieldLen=   256
      Groups(0).Columns(8).Locked=   -1  'True
      Groups(0).Columns(9).Width=   1402
      Groups(0).Columns(9).Caption=   "Saldo Ant."
      Groups(0).Columns(9).Name=   "Saldo_Final"
      Groups(0).Columns(9).Alignment=   1
      Groups(0).Columns(9).DataField=   "Column 9"
      Groups(0).Columns(9).DataType=   3
      Groups(0).Columns(9).NumberFormat=   "###,##0"
      Groups(0).Columns(9).FieldLen=   256
      Groups(0).Columns(9).Locked=   -1  'True
      Groups(0).Columns(10).Width=   2487
      Groups(0).Columns(10).Caption=   "Novo Saldo"
      Groups(0).Columns(10).Name=   "Saldo_Prod"
      Groups(0).Columns(10).Alignment=   1
      Groups(0).Columns(10).DataField=   "Column 10"
      Groups(0).Columns(10).DataType=   8
      Groups(0).Columns(10).FieldLen=   256
      Groups(0).Columns(10).Locked=   -1  'True
      Groups(0).Columns(11).Width=   1191
      Groups(0).Columns(11).Visible=   0   'False
      Groups(0).Columns(11).Caption=   "Vendas_Prod"
      Groups(0).Columns(11).Name=   "Vendas_Prod"
      Groups(0).Columns(11).DataField=   "Column 11"
      Groups(0).Columns(11).DataType=   8
      Groups(0).Columns(11).FieldLen=   256
      Groups(0).Columns(12).Width=   2328
      Groups(0).Columns(12).Visible=   0   'False
      Groups(0).Columns(12).Caption=   "Dev_Prod"
      Groups(0).Columns(12).Name=   "Dev_Prod"
      Groups(0).Columns(12).DataField=   "Column 12"
      Groups(0).Columns(12).DataType=   8
      Groups(0).Columns(12).FieldLen=   256
      Groups(0).Columns(13).Width=   3493
      Groups(0).Columns(13).Visible=   0   'False
      Groups(0).Columns(13).Caption=   "Emp_Prod"
      Groups(0).Columns(13).Name=   "Emp_Prod"
      Groups(0).Columns(13).DataField=   "Column 13"
      Groups(0).Columns(13).DataType=   8
      Groups(0).Columns(13).FieldLen=   256
      _ExtentX        =   19685
      _ExtentY        =   3942
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton B_Monta 
      Caption         =   "&Pesquisar"
      Height          =   400
      Left            =   9840
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "AcertaEmpEntrada.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
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
      Columns(0).Width=   9472
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2037
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   795
      Left            =   6600
      TabIndex        =   27
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton O_Conclu�da 
         Caption         =   "&Conclu�das"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1200
      End
      Begin VB.OptionButton O_Aberto 
         Caption         =   "Em &Aberto"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   795
      Left            =   1560
      TabIndex        =   26
      Top             =   0
      Width           =   4935
      Begin MSMask.MaskEdBox Data_Ace 
         Height          =   315
         Left            =   3480
         TabIndex        =   4
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   195
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Emp 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   195
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton O_Todas_Datas 
         Caption         =   "&Todos"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton O_Acerto 
         Caption         =   "&Acerto"
         Height          =   225
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   840
      End
      Begin VB.OptionButton O_Empr�stimo 
         Caption         =   "&Empr�stimo"
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1140
      End
   End
   Begin MSMask.MaskEdBox Valor_Prod 
      Height          =   285
      Left            =   9825
      TabIndex        =   21
      Top             =   3450
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   14
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label Label10 
      Caption         =   "Valor do Produto :"
      Height          =   225
      Left            =   9720
      TabIndex        =   43
      Top             =   3225
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Pr�ximo acerto :"
      Height          =   195
      Left            =   6780
      TabIndex        =   34
      Top             =   4995
      Width           =   1140
   End
   Begin VB.Label Valor_Total 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5325
      TabIndex        =   33
      Top             =   5820
      Width           =   1065
   End
   Begin VB.Label Valor_Linha 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5325
      TabIndex        =   32
      Top             =   5490
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   4440
      TabIndex        =   31
      Top             =   5835
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Linha atual"
      Height          =   195
      Left            =   4440
      TabIndex        =   30
      Top             =   5535
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "Valor de Compras "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   29
      Top             =   5235
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor :"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmAcertaEmpEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsClientes As Recordset
Private rsProdutos As Recordset
Private rsEmprestimos As Recordset
Private rsEstoque As Recordset
Private rsEstoque_Final As Recordset
Private rsResumo_Di�rio As Recordset
Private rsParametros As Recordset
Private rsEntradas As Recordset
Private rsEntradas_Prod As Recordset

Private Type Tab_Emp
  Sequ�ncia As Long
  Produto As String
  Nome As String
  Tamanho As Integer
  Cor As Integer
  Edi��o As Long
  Ordem As Long
  Data As Date
  Saldo_Ant As Long
  Valor_Unit As Double
  
  Novo_Emp As Long
  Nova_Venda As Long
  Nova_Devol As Long
  Novo_Saldo As Long
  
  '27/08/2003 - mpdea
  'Exibi��o por ordem de c�digo do produto
  Ordenacao As String
  
End Type

'02/10/2003 - mpdea
'Redimensionado tamanho m�ximo do array (1000 -> 5000)
Private Const EMP_ARRAY_SIZE As Integer = 5000
Private Empr�stimos(EMP_ARRAY_SIZE) As Tab_Emp

Private Linha As Integer


Private Function Acha_Reg_Estoque(Filial As Integer, Dia As Date, Produto As String, Tamanho As Integer, Cor As Integer, Edi��o As Long) As Variant
  Dim Est_Final As Single
  Dim Erro As Boolean
   
  Est_Final = 0
  rsEstoque.Index = "Produto"
  rsEstoque.Seek "=", Filial, Dia, Produto, Tamanho, Cor, Edi��o
                        
  If Not rsEstoque.NoMatch Then
    Acha_Reg_Estoque = rsEstoque.Bookmark
    Exit Function
  End If
       
       
  rsProdutos.Index = "C�digo"
  rsProdutos.Seek "=", Produto
       
           
  'N�o tem registro no dia atual
  rsEstoque.Index = "Data2"
  Erro = False
  rsEstoque.Seek ">", Filial, Produto, Tamanho, Cor, Edi��o, Dia
   
  If rsEstoque.NoMatch Then Erro = True
  If Erro = False Then If rsEstoque("Filial") <> Filial Then Erro = True
  If Erro = False Then If rsEstoque("Produto") <> Produto Then Erro = True
  If Erro = False Then If rsEstoque("Tamanho") <> Tamanho Then Erro = True
  If Erro = False Then If rsEstoque("Cor") <> Cor Then Erro = True
  If Erro = False Then If rsEstoque("Edi��o") <> Edi��o Then Erro = True
  If Erro = False Then  'j� tinha em outro dia
     Est_Final = rsEstoque("Estoque Final")
     rsEstoque.AddNew
       rsEstoque("Filial") = gnCodFilial
       rsEstoque("Data") = Dia
       rsEstoque("Produto") = Produto
       rsEstoque("Tamanho") = Tamanho
       rsEstoque("Cor") = Cor
       rsEstoque("Edi��o") = Edi��o
       rsEstoque("Classe") = rsProdutos("Classe")
       rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
       rsEstoque("Estoque Anterior") = Est_Final
     rsEstoque.Update
     rsEstoque.Index = "Produto"
     rsEstoque.Seek "=", Filial, Dia, Produto, Tamanho, Cor, Edi��o
     Acha_Reg_Estoque = rsEstoque.Bookmark
     Exit Function
   End If
   
   If Erro = True Then  'Cria novo para o produto
     rsEstoque.AddNew
       rsEstoque("Filial") = gnCodFilial
       rsEstoque("Data") = Dia
       rsEstoque("Produto") = Produto
       rsEstoque("Tamanho") = Tamanho
       rsEstoque("Cor") = Cor
       rsEstoque("Edi��o") = Edi��o
       rsEstoque("Classe") = rsProdutos("Classe")
       rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
       rsEstoque("Estoque Anterior") = 0
     rsEstoque.Update
     rsEstoque.Index = "Produto"
     rsEstoque.Seek "=", Filial, Dia, Produto, Tamanho, Cor, Edi��o
     Acha_Reg_Estoque = rsEstoque.Bookmark
     Exit Function
   End If

End Function

'02/10/2003 - mpdea
'Implementado tratamento de erro e transa��o
Private Function Atu_Empr�stimo() As Integer
  Dim i As Integer
  Dim Qtde_Dev As Long
  Dim Qtde_Emp As Long
  Dim Qtde_Vendas As Long
  Dim Erro As Boolean
  Dim Est_Final As Single
  Dim Num_Reg As Variant
  Dim Tot_Vendas As Long
  Dim Tot_Devolu��o As Long
  Dim Tot_Empr�stimos As Long
  
  Dim blnInTransaction As Boolean
  
  
  On Error GoTo ErrHandler
  
  
  Qtde_Dev = 0
  Qtde_Emp = 0
  Qtde_Vendas = 0
  
  Call StatusMsg("")
 
  For i = 0 To Grade1.Rows
    If Empr�stimos(i).Nova_Devol <> 0 Then Qtde_Dev = Qtde_Dev + Empr�stimos(i).Nova_Devol
    If Empr�stimos(i).Novo_Emp <> 0 Then Qtde_Emp = Qtde_Emp + Empr�stimos(i).Novo_Emp
    If Empr�stimos(i).Nova_Venda <> 0 Then Qtde_Vendas = Qtde_Vendas + Empr�stimos(i).Nova_Venda
  Next i
  
  If Qtde_Dev = 0 And Qtde_Emp = 0 And Qtde_Vendas = 0 Then
    DisplayMsg "N�o existe nenhum movimento de vendas, devolu��es ou empr�stimos."
    Atu_Empr�stimo = 1
    Exit Function
  End If
  
  If Not IsDate(Data_Acerto.Text) Then
    DisplayMsg "Digite a data para o pr�ximo acerto."
    Data_Acerto.SetFocus
    Atu_Empr�stimo = 1
    Exit Function
  End If
  
  If CDate(Data_Acerto.Text) < CDate(Data_Atual) Then
    DisplayMsg "Data de acerto inv�lida, verifique."
    Data_Acerto.SetFocus
    Atu_Empr�stimo = 1
    Exit Function
  End If
 
  
  '02/10/2003 - mpdea
  'Inicia transa��o
  ws.BeginTrans
  blnInTransaction = True
  
  
  rsProdutos.Index = "C�digo"
  For i = 0 To Grade1.Rows
    If Empr�stimos(i).Nova_Devol <> 0 Or Empr�stimos(i).Novo_Emp <> 0 Or Empr�stimos(i).Nova_Venda <> 0 Then
   
      rsProdutos.Seek "=", Empr�stimos(i).Produto
      
      Rem Posiciona no registro do estoque
      Num_Reg = Acha_Reg_Estoque(gnCodFilial, CDate(Data_Atual), _
        Empr�stimos(i).Produto, Empr�stimos(i).Tamanho, _
        Empr�stimos(i).Cor, Empr�stimos(i).Edi��o)
   
      rsEstoque.Bookmark = Num_Reg
      
      Rem Neste ponto tem o registro de estoque no buffer
      rsEstoque.Edit
        rsEstoque("Empre Sa�da") = rsEstoque("Empre Sa�da") + Empr�stimos(i).Nova_Devol
        rsEstoque("Valor Empre Sa�da") = rsEstoque("Valor Empre Sa�da") + (Empr�stimos(i).Nova_Devol * Empr�stimos(i).Valor_Unit)
        rsEstoque("Empre Entra") = rsEstoque("Empre Entra") + Empr�stimos(i).Novo_Emp
        rsEstoque("Valor Empre Entra") = rsEstoque("Valor Empre Entra") + (Empr�stimos(i).Novo_Emp * Empr�stimos(i).Valor_Unit)
        rsEstoque("Estoque Final") = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") _
           + rsEstoque("Compras") - rsEstoque("Transf Sa�da") + rsEstoque("Transf Entra") _
           - rsEstoque("Ajuste Sa�da") + rsEstoque("Ajuste Entra") - rsEstoque("Gr�tis Sa�da") _
           + rsEstoque("Gr�tis Entra") - rsEstoque("Quebras") - rsEstoque("Empre Sa�da") _
           + rsEstoque("Empre Entra")
        Est_Final = rsEstoque("Estoque Final")
      rsEstoque.Update
   
   
      Rem Acerta Estoque Final
      Grava_Estoque_Final gnCodFilial, Empr�stimos(i).Produto, _
            Empr�stimos(i).Tamanho, Empr�stimos(i).Cor, Empr�stimos(i).Edi��o, _
            Est_Final, CDate(Data_Atual)
   
   
      Rem Acerta o resumo di�rio
      rsResumo_Di�rio.Index = "Data"
      rsResumo_Di�rio.Seek "=", gnCodFilial, Data_Atual
      If rsResumo_Di�rio.NoMatch Then
        rsResumo_Di�rio.AddNew
        rsResumo_Di�rio("Filial") = gnCodFilial
        rsResumo_Di�rio("Data") = Data_Atual
        rsResumo_Di�rio("Valor E Sa�da") = 0
        rsResumo_Di�rio("Valor E Entrada") = 0
      Else
        rsResumo_Di�rio.Edit
      End If
      rsResumo_Di�rio("Valor E Sa�da") = rsResumo_Di�rio("Valor E Sa�da") + (Empr�stimos(i).Nova_Devol * Empr�stimos(i).Valor_Unit)
      rsResumo_Di�rio("Valor E Entrada") = rsResumo_Di�rio("Valor E Entrada") + (Empr�stimos(i).Novo_Emp * Empr�stimos(i).Valor_Unit)
      rsResumo_Di�rio.Update
   
   
   
      Rem Grava Status OK para linha atual
      rsEmprestimos.Index = "Cliente"
      rsEmprestimos.Seek "=", gnCodFilial, Empr�stimos(i).Sequ�ncia, (Combo_Cliente.Text), Empr�stimos(i).Produto, Empr�stimos(i).Tamanho, Empr�stimos(i).Cor, Empr�stimos(i).Edi��o, Empr�stimos(i).Ordem
      If rsEmprestimos.NoMatch Then
        
        '02/10/2003 - mpdea
        'Desfaz transa��o
        If blnInTransaction Then ws.Rollback
        
        MsgBox "Erro ao encontrar empr�stimo."
        Atu_Empr�stimo = 1
        Exit Function
      End If
      
      Tot_Vendas = 0
      Tot_Devolu��o = 0
      Tot_Empr�stimos = 0
   
      With rsEmprestimos
        .Edit
         .Fields("Conclu�do") = True
         Est_Final = .Fields("Saldo Atual")
        .Update
        .AddNew
        .Fields("Filial").Value = gnCodFilial
        .Fields("Sequ�ncia").Value = Empr�stimos(i).Sequ�ncia
        .Fields("Fornecedor").Value = Val(Combo_Cliente.Text)
        .Fields("Produto").Value = Empr�stimos(i).Produto
        .Fields("Tamanho").Value = Empr�stimos(i).Tamanho
        .Fields("Cor").Value = Empr�stimos(i).Cor
        .Fields("Edi��o").Value = Empr�stimos(i).Edi��o
        .Fields("Ordem").Value = Empr�stimos(i).Ordem + 1
        .Fields("Data Opera��o").Value = Data_Atual
        .Fields("Saldo Anterior").Value = Est_Final
        .Fields("Pre�o Unit�rio").Value = Empr�stimos(i).Valor_Unit
        
        .Fields("Vendas").Value = Empr�stimos(i).Nova_Venda
        .Fields("Empr�stimo Recebido").Value = Empr�stimos(i).Novo_Emp
        .Fields("Devolu��o").Value = Empr�stimos(i).Nova_Devol
        
        .Fields("Saldo Atual").Value = Est_Final - Empr�stimos(i).Nova_Venda - Empr�stimos(i).Nova_Devol + Empr�stimos(i).Novo_Emp
        .Fields("Data Cobran�a").Value = Data_Acerto.Text
        .Fields("Data Altera��o").Value = Format(Date, "dd/mm/yyyy")
        If .Fields("Saldo Atual").Value = 0 Then .Fields("Conclu�do").Value = True
        .Update
      End With
    End If
  Next i

  '02/10/2003 - mpdea
  'Finaliza transa��o
  ws.CommitTrans
  blnInTransaction = False
 
  Atu_Empr�stimo = 0

  Exit Function
  
ErrHandler:
  '02/10/2003 - mpdea
  'Desfaz transa��o
  If blnInTransaction Then ws.Rollback
  
  Atu_Empr�stimo = -1
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

Private Sub Atualiza_Detalhes()
  Dim Rec_SQL As Recordset
  Dim sSql As String
  
  sSql = "Select [Data Opera��o], Ordem, [Saldo Anterior], [Vendas], [Devolu��o], [Empr�stimo Recebido] , [Saldo Atual]"
  sSql = sSql + " From [Consigna��o Entrada] "
  sSql = sSql + " Where [Consigna��o Entrada].Filial = " + str(gnCodFilial)
  sSql = sSql + " And [Consigna��o Entrada].Sequ�ncia = " + str(Grade1.Columns(0).Text)
  sSql = sSql + " And [Consigna��o Entrada].Produto = '" + Grade1.Columns(1).Text + "'"
  sSql = sSql + " And [Consigna��o Entrada].Tamanho = " + Grade1.Columns(3).Text
  sSql = sSql + " And [Consigna��o Entrada].Cor = " + Grade1.Columns(4).Text
  sSql = sSql + " And [Consigna��o Entrada].Edi��o = " + Grade1.Columns(5).Text
  sSql = sSql + " Order By -Ordem"
  
  Set Rec_SQL = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade2.DataMode = 1
  
  Set Data2.Recordset = Rec_SQL
  
  Grade2.DataMode = 0
  
  Grade2.ReBind
    
    
  Grade2.Columns(0).Width = 1100
  Grade2.Columns(0).Locked = True
  
  Grade2.Columns(1).Width = 600
  Grade2.Columns(1).Locked = True
    
  Grade2.Columns(2).Width = 900
  Grade2.Columns(2).Caption = "Saldo Ant."
  Grade2.Columns(2).Locked = True
  
  Grade2.Columns(3).Width = 700
  Grade2.Columns(3).Locked = True
  
  Grade2.Columns(4).Width = 600
  Grade2.Columns(4).Caption = "Dev."
  Grade2.Columns(4).Locked = True
  
  Grade2.Columns(5).Width = 600
  Grade2.Columns(5).Caption = "Empr."
  Grade2.Columns(5).Locked = True
  
  Grade2.Columns(6).Width = 1100
  Grade2.Columns(6).Caption = "Saldo Final"
  Grade2.Columns(6).Locked = True
  
  Grade2.MoveLast
  Grade2.MoveFirst
  
End Sub

Private Sub Mostra_Estoque()
  Dim Est As Single
  Dim Dif As Single
  Dim Erro As Integer
  
  Est = Acha_Estoque(gnCodFilial, Grade1.Columns(1).Text, Grade1.Columns(3).Text, Grade1.Columns(4).Text, Grade1.Columns(5).Text, Erro)
  
  If Erro <> 0 Then Est = 0
    
  Dif = Grade1.Columns(9).Text - Est
   
  L_Estoque.Caption = "Estoque :" + str(Est) + "       Diferen�a : " + str(Dif)
 
End Sub

Private Sub Recalcula_Saldo()
 Dim Saldo As Double

 Saldo = 0
 If IsNumeric(Grade1.Columns(10).Text) Then
   Saldo = Grade1.Columns(10).Text
   Saldo = Saldo - Val(Replace(Vendas_Prod.Text, ",", "."))
   Saldo = Saldo - Val(Replace(Dev_Prod.Text, ",", "."))
   Saldo = Saldo + Val(Replace(Emp_Prod.Text, ",", "."))
   Saldo_Prod.Caption = Saldo
 End If
End Sub

'02/10/2003 - mpdea
'Modificado for next para at� o limite do array Empr�stimos
Private Sub B_Atualiza_Click()
  Dim i As Integer
  Dim Aux_Dbl As Double
  
  Aux_Dbl = 0
  For i = 0 To UBound(Empr�stimos)
    If Empr�stimos(i).Nova_Venda <> 0 Then
      Aux_Dbl = Aux_Dbl + (Empr�stimos(i).Nova_Venda * Empr�stimos(i).Valor_Unit)
    End If
  Next i
  
  Valor_Total.Caption = Format(Aux_Dbl, "###,###,###,##0.00")

End Sub

Private Sub B_Atualiza_Empr�stimos_Click()
  Dim Resp As Integer
  
  Resp = Atu_Empr�stimo
  
  If Resp = 0 Then B_Monta_Click
   
  Vendas_Prod.Text = 0
  Dev_Prod.Text = 0
  Emp_Prod.Text = 0
  Saldo_Prod.Caption = 0
  Valor_Prod.Text = 0
  
  Grade1.Refresh
End Sub

'02/10/2003 - mpdea
'Implementado tratamento de erro e transa��o
Private Sub B_Atualiza_Tudo_Click()
  Dim i As Integer
  Dim Qtde_Vendas As Long
  Dim Erro As Boolean
  Dim Est_Final As Single
  Dim Num_Reg As Variant
  Dim Tot_Vendas As Long
  Dim Tot_Devolu��o As Long
  Dim Tot_Empr�stimos As Long
  Dim Mov As Long
  Dim Linha As Integer
  Dim Produto As String
  Dim Tamanho As String
  Dim Cor As String
  Dim Prod_Sem_Grade As String
  Dim Total As Double
  Dim Texto As String
 
  Dim blnInTransaction As Boolean

  
  On Error GoTo ErrHandler


  Qtde_Vendas = 0
  
  Call StatusMsg("")
  
  For i = 0 To Grade1.Rows
   If Empr�stimos(i).Nova_Venda <> 0 Then Qtde_Vendas = Qtde_Vendas + Empr�stimos(i).Nova_Venda
  Next i
  
  If Qtde_Vendas = 0 Then
    DisplayMsg "N�o existem vendas. Use o bot�o Atualiza Empr�stimo."
    Exit Sub
  End If
  
  i = Atu_Empr�stimo
  If i <> 0 Then Exit Sub
 
 
  '02/10/2003 - mpdea
  'Inicia transa��o
  ws.BeginTrans
  blnInTransaction = True
  
  
  With rsParametros
    .Index = "Filial"
    .Seek "=", gnCodFilial
    If .NoMatch Then
      '02/10/2003 - mpdea
      'Desfaz transa��o
      If blnInTransaction Then ws.Rollback
      
      MsgBox ("Erro ao encontrar par�metros")
      Exit Sub
    End If
    .Edit
    .Fields("�ltima Movimenta��o").Value = gnGetNextSequencia(gnCodFilial)
    Mov = .Fields("�ltima Movimenta��o").Value
    .Update
  End With
  
  Total = 0
 
  Linha = 1
  For i = 0 To Grade1.Rows
    If Empr�stimos(i).Nova_Venda <> 0 Then
      Produto = Empr�stimos(i).Produto
      Prod_Sem_Grade = Produto
      If Empr�stimos(i).Tamanho <> 0 Then
        Tamanho = "000" + Trim(str(Empr�stimos(i).Tamanho))
        Tamanho = Right(Tamanho, 3)
        
        Cor = "000" + Trim(str(Empr�stimos(i).Cor))
        Cor = Right(Cor, 3)
        Produto = Produto + Tamanho + Cor
      End If
      
      With rsEntradas_Prod
        .AddNew
        .Fields("Filial").Value = gnCodFilial
        .Fields("Sequ�ncia").Value = Mov
        .Fields("Linha").Value = Linha
        .Fields("C�digo").Value = Produto
        .Fields("Qtde").Value = Empr�stimos(i).Nova_Venda
        .Fields("Pre�o").Value = Empr�stimos(i).Valor_Unit
        .Fields("Desconto").Value = 0
        .Fields("ICM").Value = gvGetValueInTable("Produtos", "[Percentual ICM]", ftNumero, "C�digo", ftTexto, Produto)
        .Fields("IPI").Value = gvGetValueInTable("Produtos", "[Percentual IPI]", ftNumero, "C�digo", ftTexto, Produto)
        .Fields("Pre�o Final").Value = .Fields("Qtde").Value * .Fields("Pre�o").Value
        .Fields("Etiqueta").Value = False
        .Fields("C�digo Sem Grade").Value = Prod_Sem_Grade
        Total = Total + .Fields("Pre�o Final").Value
        .Update
      End With
      Linha = Linha + 1
    End If
  Next i
 
  Rem Grava Entrada
  With rsEntradas
    .AddNew
    .Fields("Filial").Value = gnCodFilial
    .Fields("Data").Value = Data_Atual
    .Fields("Sequ�ncia").Value = Mov
    .Fields("Opera��o").Value = 0
    .Fields("Digitador").Value = 0
    .Fields("Fornecedor").Value = Val(Combo_Cliente.Text)
    .Fields("Observa��es").Value = ""
    .Fields("Produtos").Value = Total
    .Fields("Total").Value = Total
    .Fields("Efetivada").Value = False
    .Update
  End With
  
  '02/10/2003 - mpdea
  'Finaliza transa��o
  ws.CommitTrans
  blnInTransaction = False
  
  Texto = "A entrada " + str(Mov) + " foi criada."
  Texto = Texto & vbCrLf & Chr(13)
  Texto = Texto + "Voc� DEVE entrar na tela de ENTRADAS e verificar a movimenta��o, os valores, impostos e quantidades de produtos. Se os produtos vendidos tem ICM ou IPI verifique tamb�m estes impostos."
  MsgBox (Texto), vbInformation, "Infoma��o"
 
  '-------------------------------------------------------------------------
  '02/10/2003 - mpdea
  'Atualiza informa��es na tela
  Call B_Monta_Click
  
  Vendas_Prod.Text = 0
  Dev_Prod.Text = 0
  Emp_Prod.Text = 0
  Saldo_Prod.Caption = 0
  Valor_Prod.Text = 0
  
  Grade1.Refresh
  '-------------------------------------------------------------------------
  
  Exit Sub
  
ErrHandler:
  '02/10/2003 - mpdea
  'Desfaz transa��o
  If blnInTransaction Then ws.Rollback
  
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub B_Cancela_Mov_Click()
  Vendas_Prod.Text = 0
  Dev_Prod.Text = 0
  Emp_Prod.Text = 0
End Sub

Private Sub B_Confirma_Mov_Click()
  
  Call StatusMsg("")
  
  If Grade1.SelBookmarks.Count = 0 Then
    DisplayMsg "Selecione uma linha antes."
    Exit Sub
  End If
  
  If CDbl(Saldo_Prod.Caption) < 0 Then
    DisplayMsg "Estoque n�o pode ficar negativo."
    Exit Sub
  End If
  
  Grade1.Columns(11).Text = Val(Vendas_Prod.Text)
  Grade1.Columns(12).Text = Val(Dev_Prod.Text)
  Grade1.Columns(13).Text = Val(Emp_Prod.Text)
  Grade1.Columns(10).Text = Val(Saldo_Prod.Caption)
  Grade1.Columns(8).Text = CDbl(Valor_Prod.Text)
  
  Grade1.Update
  
  B_Atualiza_Click
End Sub

Private Sub B_Imprime_Click()
  Grade1.PrintData ssPrintAllRows, True, True
End Sub

Private Sub B_Monta_Click()
  Dim Aux_Cliente As Long
  Dim Aux_Produto As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edi��o As Integer
  Dim Aux_ordem As Long
  Dim Aux_Seq As Long

  Call StatusMsg("")
 
  If Len(Combo_Cliente.Text) = 0 Then
    DisplayMsg "Fornecedor incorreto."
    Combo_Cliente.SetFocus
  End If

  If O_Acerto.Value = True And Not IsDate(Data_Ace.Text) Then
    DisplayMsg "Data inv�lida, verifique."
    Data_Ace.SetFocus
    Exit Sub
  End If
  
  If O_Empr�stimo.Value = True And Not IsDate(Data_Emp.Text) Then
    DisplayMsg "Data inv�lida, verifique."
    Data_Emp.SetFocus
    Exit Sub
  End If

  '02/10/2003 - mpdea
  'Status
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde...")
  
  Aux_Cliente = 0
  Aux_Produto = 0
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edi��o = 0
  Aux_ordem = 0
  Aux_Seq = 0
  
  Linha = 0
  
  Erase Empr�stimos
  
  '02/10/2003 - mpdea
  'Zerado as linhas e modificado a atualiza��o do grid
  With Grade1
    .Rows = 0
    .MoveLast
    .MoveFirst
    .Refresh
    .Redraw = False
  End With
  
  rsEmprestimos.Index = "Cliente"
  rsProdutos.Index = "C�digo"
Lp1:
  rsEmprestimos.Seek ">", gnCodFilial, Aux_Seq, Aux_Cliente, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_ordem
If rsEmprestimos.NoMatch Then GoTo Fim_Lp1
  Aux_Seq = rsEmprestimos("Sequ�ncia")
  Aux_Produto = rsEmprestimos("Produto")
  Aux_Tamanho = rsEmprestimos("Tamanho")
  Aux_Cor = rsEmprestimos("Cor")
  Aux_Edi��o = rsEmprestimos("Edi��o")
  Aux_ordem = rsEmprestimos("Ordem")
  Aux_Cliente = rsEmprestimos("Fornecedor")
  
  If rsEmprestimos("Filial") <> gnCodFilial Then GoTo Fim_Lp1
  If rsEmprestimos("Fornecedor") <> Val(Combo_Cliente.Text) Then GoTo Lp1
    
  If O_Aberto.Value = True And rsEmprestimos("Conclu�do") = True Then GoTo Lp1
  If O_Aberto.Value = False And rsEmprestimos("Conclu�do") = False Then GoTo Lp1
  
  If O_Empr�stimo.Value = True Then
    If CDate(Data_Emp.Text) <> CDate(rsEmprestimos("Data Opera��o")) Then GoTo Lp1
  End If
  
  If O_Acerto.Value = True Then
    If CDate(Data_Ace.Text) <> CDate(rsEmprestimos("Data Cobran�a")) Then GoTo Lp1
  End If
  
  rsProdutos.Seek "=", Aux_Produto
  
  '-----------------------------------------------------------------------------
  '27/08/2003 - mpdea
  'Otimizado c�digo e adicionado ordena��o
  With Empr�stimos(Linha)
    .Sequ�ncia = Aux_Seq
    .Produto = Aux_Produto
  
    If rsProdutos.NoMatch Then
      .Nome = "Produto n�o encontrado"
      .Ordenacao = ""
    Else
      .Nome = rsProdutos.Fields("Nome").Value & ""
      .Ordenacao = rsProdutos.Fields("C�digo Ordena��o").Value & ""
    End If
    
    .Tamanho = Aux_Tamanho
    .Cor = Aux_Cor
    .Edi��o = Aux_Edi��o
    .Ordem = Aux_ordem
    .Data = rsEmprestimos.Fields("Data Opera��o").Value
    .Saldo_Ant = rsEmprestimos.Fields("Saldo Atual").Value
    .Valor_Unit = rsEmprestimos.Fields("Pre�o Unit�rio").Value
  
    .Nova_Devol = 0
    .Nova_Venda = 0
    .Novo_Emp = 0
    .Novo_Saldo = rsEmprestimos.Fields("Saldo Atual").Value
  End With
  '-----------------------------------------------------------------------------
  
  Linha = Linha + 1
  
  GoTo Lp1
  
Fim_Lp1:
  
  '28/08/2003 - mpdea
  'Ordena a tabela por c�digo
  If optOrdemProduto.Value Then Call OrderByCode
  
  '02/10/2003 - mpdea
  'Modificado a atualiza��o do grid
  With Grade1
    '28/10/2003 - Maikel
    .Rows = Linha
    
    .MoveLast
    .MoveFirst
    .Refresh
    .Redraw = True
  End With
  
  
  '-----------------------------------------------------------------------------
  '27/08/2003 - mpdea
  'Otimizado c�digo

  Grade1.Columns(3).Visible = gbGrade
  Grade1.Columns(4).Visible = gbGrade
  Grade1.Columns(5).Visible = gbEdicao
  
  Frame_Mov.Enabled = O_Aberto.Value
  '-----------------------------------------------------------------------------

  '02/10/2003 - mpdea
  'Status
  Screen.MousePointer = vbDefault
  Call StatusMsg("Pronto")

End Sub

Private Sub B_Mostra_Click()
  Call StatusMsg("")
   
  If Grade1.SelBookmarks.Count < 1 Then
    DisplayMsg "Selecione a linha antes."
    Exit Sub
  End If
End Sub

Private Sub Combo_Cliente_CloseUp()
  Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
  Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()

  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Combo_Cliente.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 1 Then Exit Sub
  
  rsClientes.Index = "C�digo"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then Exit Sub
  Call StatusMsg(rsClientes("Nome") & "")
  
End Sub

Private Sub Command1_Click()
  With frmVerificaDatas
    .Tipo = "ENTRADA"
    .Show
    .WindowState = vbNormal
  End With
End Sub

Private Sub Data_Ace_LostFocus()
  Data_Ace.Text = Ajusta_Data(Data_Ace.Text)
End Sub

Private Sub Data_Ace_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ace.Text = frmCalendario.gsDateCalender(Data_Ace.Text)
  End Select
End Sub

Private Sub Data_Acerto_LostFocus()
  Data_Acerto.Text = Ajusta_Data(Data_Acerto.Text)
End Sub

Private Sub Data_Acerto_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Acerto.Text = frmCalendario.gsDateCalender(Data_Acerto.Text)
  End Select
End Sub

Private Sub Data_Emp_LostFocus()
  Data_Emp.Text = Ajusta_Data(Data_Emp.Text)
End Sub

Private Sub Data_Emp_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Emp.Text = frmCalendario.gsDateCalender(Data_Emp.Text)
  End Select
End Sub

Private Sub Dev_Prod_Change()
  Recalcula_Saldo
End Sub

Private Sub Dev_Prod_GotFocus()
  Dev_Prod.SelStart = 0
  Dev_Prod.SelLength = Len(Dev_Prod.Text)
End Sub

Private Sub Dev_Prod_LostFocus()
If Not IsNumeric(Dev_Prod.Text) Then
      DisplayMsg "Quantidade incorreta."
      Dev_Prod.SetFocus
End If
End Sub

Private Sub Emp_Prod_Change()
  Recalcula_Saldo
End Sub

Private Sub Emp_Prod_GotFocus()
  Emp_Prod.SelStart = 0
  Emp_Prod.SelLength = Len(Emp_Prod.Text)
End Sub

Private Sub Emp_Prod_LostFocus()
	If Not IsNumeric(Emp_Prod.Text) Then
		  DisplayMsg "Quantidade incorreta."
		  Emp_Prod.SetFocus
	End If
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)

  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsEmprestimos = db.OpenRecordset("Consigna��o Entrada")
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque")
  Set rsEstoque_Final = db.OpenRecordset("Estoque Final")
  Set rsResumo_Di�rio = db.OpenRecordset("Resumo Di�rio")
  Set rsParametros = db.OpenRecordset("Par�metros Filial")
  Set rsEntradas = db.OpenRecordset("Entradas")
  Set rsEntradas_Prod = db.OpenRecordset("Entradas - Produtos")
  
  Grade1.Columns(13).NumberFormat = Formato_Pre�o
  
End Sub

Private Sub Grade1_Click()
  Grade1.SelBookmarks.RemoveAll
  Grade1.SelBookmarks.Add Grade1.Bookmark
End Sub

Private Sub Grade1_LostFocus()
  If Grade1.RowChanged = True Then
    Grade1.Update
  End If
End Sub

Private Sub Grade1_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  Dim Aux_Dbl As Double
 
  On Error GoTo Erro
  
  Aux_Dbl = Val(Grade1.Columns(11).Text) * CDbl(Grade1.Columns(8).Text)
  Valor_Linha.Caption = Format(Aux_Dbl, "###,###,###,##0.00")
  
  Vendas_Prod.Text = Grade1.Columns(11).Text
  Dev_Prod.Text = Grade1.Columns(12).Text
  Emp_Prod.Text = Grade1.Columns(13).Text
  Saldo_Prod.Caption = Grade1.Columns(10).Text
  Valor_Prod.Text = Grade1.Columns(8).Text
  
  If O_Mostra_Detalhe.Value = 1 Then
    Atualiza_Detalhes
  End If
  
  Call Mostra_Estoque
  
Erro:
  Exit Sub
  
End Sub

Private Sub Grade1_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Long
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade1.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p

End Sub

Private Sub Grade1_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Dim r As Long, i As Long, p As Long

  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      p = Grade1.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
    If ReadPriorRows Then
      p = p - 1
    Else
      p = p + 1
    End If
  End If

  For i = 0 To RowBuf.RowCount - 1
    If p < 0 Or p >= Grade1.Rows Then Exit For
    
    RowBuf.Value(i, 0) = Empr�stimos(p).Sequ�ncia
    RowBuf.Value(i, 1) = Empr�stimos(p).Produto
    RowBuf.Value(i, 2) = Empr�stimos(p).Nome
    RowBuf.Value(i, 3) = Empr�stimos(p).Tamanho
    RowBuf.Value(i, 4) = Empr�stimos(p).Cor
    RowBuf.Value(i, 5) = Empr�stimos(p).Edi��o
    RowBuf.Value(i, 6) = Empr�stimos(p).Ordem
    RowBuf.Value(i, 7) = Empr�stimos(p).Data
    RowBuf.Value(i, 8) = Empr�stimos(p).Valor_Unit
    RowBuf.Value(i, 9) = Empr�stimos(p).Saldo_Ant

    'campos invis�veis
    RowBuf.Value(i, 10) = Empr�stimos(p).Novo_Saldo
    RowBuf.Value(i, 11) = Empr�stimos(p).Nova_Venda
    RowBuf.Value(i, 12) = Empr�stimos(p).Nova_Devol
    RowBuf.Value(i, 13) = Empr�stimos(p).Novo_Emp
       
    RowBuf.Bookmark(i) = p
    If ReadPriorRows Then
      p = p - 1
    Else
      p = p + 1
    End If
    
    r = r + 1
  Next i
 
  RowBuf.RowCount = r

End Sub


Private Sub Grade1_UnboundWriteData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, WriteLocation As Variant)
On Error GoTo Erro
  Dim Linha1 As Integer
  
  If IsNull(WriteLocation) Then Exit Sub
  
  Linha1 = WriteLocation
  If Linha1 = -1 Then Linha1 = 0
  
  With Empr�stimos(Linha1)
    .Sequ�ncia = Grade1.Columns(0).Text
    .Produto = Grade1.Columns(1).Text
    .Nome = Grade1.Columns(2).Text
    .Tamanho = Grade1.Columns(3).Text
    .Cor = Grade1.Columns(4).Text
    .Edi��o = Grade1.Columns(5).Text
    .Ordem = Grade1.Columns(6).Text
    .Data = Grade1.Columns(7).Text
    .Valor_Unit = Grade1.Columns(8).Text
    .Saldo_Ant = Grade1.Columns(9).Text
    'campos invis�veis
    .Novo_Saldo = Grade1.Columns(10).Text
    .Nova_Venda = Grade1.Columns(11).Text
    .Nova_Devol = Grade1.Columns(12).Text
    .Novo_Emp = Grade1.Columns(13).Text
  End With
  
  Exit Sub
Erro:
  MsgBox "Algo deu errado. Repita o procedimento.", vbInformation, "Aten��o"
End Sub

Private Sub O_Mostra_Detalhe_Click()
 If O_Mostra_Detalhe.Value = 1 Then
   Grade2.Visible = True
 Else
   Grade2.Visible = False
 End If
End Sub

Private Sub Valor_Prod_GotFocus()
  Valor_Prod.SelStart = 0
  Valor_Prod.SelLength = Len(Valor_Prod.Text)
End Sub

Private Sub Valor_Prod_LostFocus()
If Not IsNumeric(Valor_Prod.Text) Then
      DisplayMsg "Valor incorreto."
      Valor_Prod.SetFocus
End If
End Sub

Private Sub Vendas_Prod_Change()
  Recalcula_Saldo
End Sub

Private Sub Vendas_Prod_GotFocus()
  Vendas_Prod.SelStart = 0
  Vendas_Prod.SelLength = Len(Vendas_Prod.Text)
End Sub

'08/10/2003 - mpdea
'Corrigido esquema da ordena��o (compara��o com vazios)

'27-28/08/2003 - mpdea
'Ordena��o da lista por c�digo
Private Sub OrderByCode()
  Dim TEMP_Emprestimos() As Tab_Emp
  Dim intX1 As Integer
  Dim intX2 As Integer
  Dim strCodigoOrdenacao As String
  Dim intMinPos As Integer
  
  '02/10/2003 - mpdea
  'Igualado o redimensionado
  ReDim TEMP_Emprestimos(UBound(Empr�stimos)) As Tab_Emp
  
  For intX1 = LBound(Empr�stimos) To UBound(Empr�stimos)
    
    '�ndice do primeiro item preenchido
    intMinPos = m_intFirstRegister()
    
    'N�o h� mais itens
    If intMinPos = -1 Then Exit For
    
    'C�digo m�nimo a ser testado para ordena��o
    strCodigoOrdenacao = Empr�stimos(intMinPos).Ordenacao
    
    For intX2 = LBound(Empr�stimos) To UBound(Empr�stimos)
      'Compara ordem
      If Empr�stimos(intX2).Ordenacao <> "" And Empr�stimos(intX2).Ordenacao < strCodigoOrdenacao Then
        intMinPos = intX2
        strCodigoOrdenacao = Empr�stimos(intMinPos).Ordenacao
      End If
    Next intX2
    
    'Copia registro
    With TEMP_Emprestimos(intX1)
      .Sequ�ncia = Empr�stimos(intMinPos).Sequ�ncia
      .Produto = Empr�stimos(intMinPos).Produto
      .Nome = Empr�stimos(intMinPos).Nome
      .Tamanho = Empr�stimos(intMinPos).Tamanho
      .Cor = Empr�stimos(intMinPos).Cor
      .Edi��o = Empr�stimos(intMinPos).Edi��o
      .Ordem = Empr�stimos(intMinPos).Ordem
      .Data = Empr�stimos(intMinPos).Data
      .Saldo_Ant = Empr�stimos(intMinPos).Saldo_Ant
      .Valor_Unit = Empr�stimos(intMinPos).Valor_Unit
      .Novo_Emp = Empr�stimos(intMinPos).Novo_Emp
      .Nova_Venda = Empr�stimos(intMinPos).Nova_Venda
      .Nova_Devol = Empr�stimos(intMinPos).Nova_Devol
      .Novo_Saldo = Empr�stimos(intMinPos).Novo_Saldo
      .Ordenacao = Empr�stimos(intMinPos).Ordenacao
    End With
    
    'Zera ordena��o do registro copiado
    Empr�stimos(intMinPos).Ordenacao = ""
    
  Next intX1
  
  'Realiza a c�pia dos dados em ordem de c�digo
  Erase Empr�stimos
  intX2 = 0
  
  For intX1 = LBound(TEMP_Emprestimos) To UBound(TEMP_Emprestimos)
    If TEMP_Emprestimos(intX1).Ordenacao <> "" Then
      With Empr�stimos(intX2)
        .Sequ�ncia = TEMP_Emprestimos(intX1).Sequ�ncia
        .Produto = TEMP_Emprestimos(intX1).Produto
        .Nome = TEMP_Emprestimos(intX1).Nome
        .Tamanho = TEMP_Emprestimos(intX1).Tamanho
        .Cor = TEMP_Emprestimos(intX1).Cor
        .Edi��o = TEMP_Emprestimos(intX1).Edi��o
        .Ordem = TEMP_Emprestimos(intX1).Ordem
        .Data = TEMP_Emprestimos(intX1).Data
        .Saldo_Ant = TEMP_Emprestimos(intX1).Saldo_Ant
        .Valor_Unit = TEMP_Emprestimos(intX1).Valor_Unit
        .Novo_Emp = TEMP_Emprestimos(intX1).Novo_Emp
        .Nova_Venda = TEMP_Emprestimos(intX1).Nova_Venda
        .Nova_Devol = TEMP_Emprestimos(intX1).Nova_Devol
        .Novo_Saldo = TEMP_Emprestimos(intX1).Novo_Saldo
        .Ordenacao = TEMP_Emprestimos(intX1).Ordenacao
      End With
      intX2 = intX2 + 1
    End If
  Next intX1
  
End Sub

'28/08/2003 - mpdea
'Obt�m o primeiro registro preenchido da tabela
Private Function m_intFirstRegister() As Integer
  Dim intX As Integer
  
  For intX = LBound(Empr�stimos) To UBound(Empr�stimos)
    If Empr�stimos(intX).Ordenacao <> "" Then
      'Primeiro registro ocupado
      m_intFirstRegister = intX
      Exit Function
    End If
  Next intX
  
  'N�o h� registro
  m_intFirstRegister = -1
  
End Function

Private Sub Vendas_Prod_LostFocus()
If Not IsNumeric(Vendas_Prod.Text) Then
      DisplayMsg "Quantidade incorreta."
      Vendas_Prod.SetFocus
End If
End Sub
