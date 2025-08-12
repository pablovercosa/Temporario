VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosCalculoVenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo do Preço de Venda"
   ClientHeight    =   6315
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   11280
   Icon            =   "CalculaVenda.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11280
   Begin VB.Data datPrecos 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   180
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Preços ORDER BY Tabela"
      Top             =   7875
      Width           =   1875
   End
   Begin VB.ComboBox cmbArredonda 
      Height          =   315
      ItemData        =   "CalculaVenda.frx":058A
      Left            =   1545
      List            =   "CalculaVenda.frx":05A3
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1890
      Width           =   3060
   End
   Begin VB.Frame Frame6 
      Caption         =   "Data da alteração do estoque - Considerar os produtos"
      Height          =   705
      Left            =   4800
      TabIndex        =   51
      Top             =   1110
      Width           =   4590
      Begin MSMask.MaskEdBox Data_Estoque 
         Height          =   315
         Left            =   2835
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label13 
         Caption         =   "Com estoque alterado após :"
         Height          =   255
         Left            =   600
         TabIndex        =   52
         Top             =   345
         Width           =   2265
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2175
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Sub_Classe"
      Top             =   7470
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   7485
      Visible         =   0   'False
      Width           =   1800
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Sub 
      Bindings        =   "CalculaVenda.frx":063C
      DataSource      =   "Data3"
      Height          =   315
      Left            =   2190
      TabIndex        =   2
      ToolTipText     =   "Use 0 para todas"
      Top             =   735
      Width           =   855
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
      Columns(0).Width=   8784
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1588
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data de cadastro - Considerar os produtos"
      Height          =   690
      Left            =   165
      TabIndex        =   45
      Top             =   1110
      Width           =   4485
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         Caption         =   "Cadastrados/Alterados após :"
         Height          =   240
         Left            =   375
         TabIndex        =   46
         Top             =   300
         Width           =   2190
      End
   End
   Begin VB.CheckBox O_Percentual 
      Caption         =   "Calcular todos os preços mantendo a Margem de Lucro, independente do Cadastro de Produtos."
      Height          =   240
      Left            =   165
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4455
      Width           =   7680
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordem"
      Height          =   720
      Left            =   9495
      TabIndex        =   44
      Top             =   1110
      Width           =   1740
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome"
         Height          =   225
         Left            =   90
         TabIndex        =   6
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Caption         =   "Código"
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Registros"
      Height          =   900
      Left            =   6525
      TabIndex        =   41
      Top             =   45
      Width           =   2865
      Begin VB.OptionButton O_Todos 
         Caption         =   "Ajustar todos"
         Height          =   225
         Left            =   105
         TabIndex        =   43
         Top             =   570
         Width           =   2115
      End
      Begin VB.OptionButton O_Marcados 
         Caption         =   "Ajustar somente os marcados"
         Height          =   330
         Left            =   105
         TabIndex        =   42
         Top             =   225
         Value           =   -1  'True
         Width           =   2445
      End
   End
   Begin VB.CommandButton B_Ajusta2 
      Caption         =   "Atualizar Custos de Produtos"
      Height          =   400
      Left            =   3870
      TabIndex        =   11
      ToolTipText     =   "Ajusta Preços na Tabela e Atualiza Pasta Cálculos de Custo no Cadastro de Produtos"
      Top             =   5775
      Width           =   1335
   End
   Begin VB.CommandButton B_Ajusta 
      Caption         =   "Atualizar Preços"
      Height          =   400
      Left            =   1365
      TabIndex        =   10
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tamanho da Letra"
      Height          =   915
      Left            =   9480
      TabIndex        =   17
      Top             =   45
      Width           =   1725
      Begin VB.OptionButton O_Pequena 
         Caption         =   "Pequena"
         Height          =   225
         Left            =   135
         TabIndex        =   4
         Top             =   585
         Width           =   1170
      End
      Begin VB.OptionButton O_normal 
         Caption         =   "Normal"
         Height          =   225
         Left            =   135
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4290
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7485
      Visible         =   0   'False
      Width           =   2010
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Bindings        =   "CalculaVenda.frx":0650
      Height          =   2115
      Left            =   120
      TabIndex        =   15
      Top             =   2265
      Width           =   11055
      _Version        =   196617
      DataMode        =   1
      RowHeight       =   423
      Columns.Count   =   13
      Columns(0).Width=   1931
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3598
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1614
      Columns(2).Caption=   "Preço Custo Anterior"
      Columns(2).Name =   "Preço Custo Anterior"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   4
      Columns(2).FieldLen=   256
      Columns(3).Width=   1323
      Columns(3).Caption=   "Preço Custo Atual"
      Columns(3).Name =   "Preço Custo Atual"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   4
      Columns(3).FieldLen=   256
      Columns(4).Width=   1270
      Columns(4).Caption=   "Preço Custo Calc Ant"
      Columns(4).Name =   "Preço Custo Calc Ant"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   4
      Columns(4).FieldLen=   256
      Columns(5).Width=   1191
      Columns(5).Caption=   "Preço Custo Calc Atu"
      Columns(5).Name =   "Preço Custo Calc Atu"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   4
      Columns(5).FieldLen=   256
      Columns(6).Width=   1191
      Columns(6).Caption=   "Preço Venda Anterior"
      Columns(6).Name =   "Preço Venda Anterior"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   1
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(7).Width=   1376
      Columns(7).Caption=   "Lucro Anterior"
      Columns(7).Name =   "Lucro Anterior"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   1
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   1191
      Columns(8).Caption=   "Lucro Anterior Perc"
      Columns(8).Name =   "Lucro Anterior Perc"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   1
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   4
      Columns(8).FieldLen=   256
      Columns(9).Width=   1270
      Columns(9).Caption=   "Preço Venda Atual"
      Columns(9).Name =   "Preço Venda Atual"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   1
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   4
      Columns(9).FieldLen=   256
      Columns(10).Width=   1032
      Columns(10).Caption=   "Lucro Atual"
      Columns(10).Name=   "Lucro Atual"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   1
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   1323
      Columns(11).Caption=   "Lucro Atual Perc"
      Columns(11).Name=   "Lucro Atual Perc"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   1
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   4
      Columns(11).FieldLen=   256
      Columns(12).Width=   1032
      Columns(12).Caption=   "Alterar"
      Columns(12).Name=   "Alterar"
      Columns(12).Alignment=   1
      Columns(12).CaptionAlignment=   1
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   11
      Columns(12).FieldLen=   256
      _ExtentX        =   19500
      _ExtentY        =   3731
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
   Begin VB.CommandButton B_Monta 
      Caption         =   "&Pesquisar"
      Height          =   400
      Left            =   9840
      TabIndex        =   12
      Top             =   1860
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "CalculaVenda.frx":0664
      DataSource      =   "Data2"
      Height          =   315
      Left            =   2190
      TabIndex        =   1
      ToolTipText     =   "Use 0 para todas"
      Top             =   405
      Width           =   855
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
      Columns(0).Width=   8176
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1693
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "CalculaVenda.frx":0678
      Height          =   315
      Left            =   2190
      TabIndex        =   0
      Top             =   30
      Width           =   1935
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Tabela"
   End
   Begin VB.Label Label14 
      Caption         =   "Arredondamento:"
      Height          =   210
      Left            =   240
      TabIndex        =   53
      Top             =   1935
      Width           =   1260
   End
   Begin VB.Label Nome_Sub 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   50
      Top             =   735
      Width           =   2535
   End
   Begin VB.Label Nome_Classe 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3120
      TabIndex        =   49
      Top             =   405
      Width           =   2535
   End
   Begin VB.Label Label12 
      Caption         =   "Subclasse"
      Height          =   225
      Left            =   195
      TabIndex        =   48
      Top             =   825
      Width           =   2325
   End
   Begin VB.Label Label11 
      Caption         =   "Classe:"
      Height          =   225
      Left            =   210
      TabIndex        =   47
      Top             =   465
      Width           =   1905
   End
   Begin VB.Label L_V_Di 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10035
      TabIndex        =   40
      Top             =   5985
      Width           =   1065
   End
   Begin VB.Label L_V_At 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8985
      TabIndex        =   39
      Top             =   5985
      Width           =   960
   End
   Begin VB.Label L_V_An 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7830
      TabIndex        =   38
      Top             =   5985
      Width           =   1065
   End
   Begin VB.Label Label9 
      Caption         =   "Preço Venda :"
      Height          =   225
      Left            =   6630
      TabIndex        =   37
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Label L_LP_Di 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10035
      TabIndex        =   36
      Top             =   5670
      Width           =   1065
   End
   Begin VB.Label L_L_Di 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10035
      TabIndex        =   35
      Top             =   5355
      Width           =   1065
   End
   Begin VB.Label L_CT_Di 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10035
      TabIndex        =   34
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label L_C_Di 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10035
      TabIndex        =   33
      Top             =   4725
      Width           =   1065
   End
   Begin VB.Label L_LP_At 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8985
      TabIndex        =   32
      Top             =   5670
      Width           =   960
   End
   Begin VB.Label L_L_At 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8985
      TabIndex        =   31
      Top             =   5355
      Width           =   960
   End
   Begin VB.Label L_CT_At 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8985
      TabIndex        =   30
      Top             =   5040
      Width           =   960
   End
   Begin VB.Label L_C_At 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8985
      TabIndex        =   29
      Top             =   4725
      Width           =   960
   End
   Begin VB.Label L_LP_An 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7830
      TabIndex        =   28
      Top             =   5670
      Width           =   1065
   End
   Begin VB.Label L_L_An 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7830
      TabIndex        =   27
      Top             =   5355
      Width           =   1065
   End
   Begin VB.Label L_CT_An 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7830
      TabIndex        =   26
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label L_C_An 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7830
      TabIndex        =   25
      Top             =   4725
      Width           =   1065
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Dif."
      Height          =   225
      Left            =   10140
      TabIndex        =   24
      Top             =   4515
      Width           =   750
   End
   Begin VB.Label Label7 
      Caption         =   "Lucro em %"
      Height          =   225
      Left            =   6630
      TabIndex        =   23
      Top             =   5700
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "Lucro em $"
      Height          =   225
      Left            =   6645
      TabIndex        =   22
      Top             =   5400
      Width           =   945
   End
   Begin VB.Label Label5 
      Caption         =   "Custo Total"
      Height          =   225
      Left            =   6630
      TabIndex        =   21
      Top             =   5070
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "Custo "
      Height          =   225
      Left            =   6630
      TabIndex        =   20
      Top             =   4785
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Atual"
      Height          =   225
      Left            =   9090
      TabIndex        =   19
      Top             =   4515
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Anterior"
      Height          =   225
      Left            =   7935
      TabIndex        =   18
      Top             =   4515
      Width           =   750
   End
   Begin VB.Label Explica 
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   210
      TabIndex        =   16
      Top             =   4875
      Width           =   6105
   End
   Begin VB.Label Label1 
      Caption         =   "Tabela de Preços Visada:"
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPrecosCalculoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTempo As Recordset
Dim rsProdutos As Recordset
Dim rsPreços As Recordset
Dim rsClasses As Recordset
Dim rsSub_Classes As Recordset
Dim rsEstoque_Final As Recordset

Private gbToCancel As Boolean
Private gbRunning As Boolean

Private gsArredonda As String

Dim Fixa_ICM_V_Perc As Boolean
Dim Fixa_IPI_V_Perc As Boolean
Dim Fixa_Imp_V_Perc As Boolean
Dim Fixa_Outros_V_Perc As Boolean


Dim Custo_Preço_Valor As Double
Dim Custo_Desconto_Perc As Double
Dim Custo_Desconto_Valor As Double
Dim Custo_Desconto_Fixo As String
Dim Custo_Frete_Perc As Double
Dim Custo_Frete_Valor As Double
Dim Custo_Frete_Fixo As String
Dim Custo_ICM_Compra_Perc As Double
Dim Custo_ICM_Compra_Valor As Double
Dim Custo_ICM_Compra_Fixo As String
Dim Custo_IPI_Compra_Perc As Double
Dim Custo_IPI_Compra_Valor As Double
Dim Custo_IPI_Compra_Fixo As String
Dim Custo_Custo_Compra_Perc As Double
Dim Custo_Custo_Compra_Valor As Double
Dim Custo_Custo_Compra_Fixo As String
Dim Custo_Outros_Compra_Perc As Double
Dim Custo_Outros_Compra_Valor As Double
Dim Custo_Outros_Compra_Fixo As String
Dim Custo_Perc_Compra_Sem As Double
Dim Custo_Custo_Calculado As Double
Dim Custo_Preço_Venda As Double
Dim Custo_ICM_Venda_Perc As Double
Dim Custo_ICM_Venda_Valor As Double
Dim Custo_ICM_Venda_Fixo As String
Dim Custo_IPI_Venda_Perc As Double
Dim Custo_IPI_Venda_Valor As Double
Dim Custo_IPI_Venda_Fixo As String
Dim Custo_Impostos_Venda_Perc As Double
Dim Custo_Impostos_Venda_Valor As Double
Dim Custo_Impostos_Venda_Fixo As String
Dim Custo_Outros_Venda_Perc As Double
Dim Custo_Outros_Venda_Valor As Double
Dim Custo_Outros_Venda_Fixo As String
Dim Custo_Perc_Venda_Sem As Double
Dim Custo_Lucro_Perc As Double
Dim Custo_Lucro_Valor As Double
Dim Custo_Manter As String

  


Sub Calc_Lucro_Perc()
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
 
 
  If Custo_ICM_Venda_Fixo = "P" Then
     Fixa_ICM_V_Perc = True
  Else
     Fixa_ICM_V_Perc = False
  End If
  
  If Custo_IPI_Venda_Fixo = "P" Then
     Fixa_IPI_V_Perc = True
  Else
     Fixa_IPI_V_Perc = False
  End If
  
  If Custo_Impostos_Venda_Fixo = "P" Then
     Fixa_Imp_V_Perc = True
  Else
     Fixa_Imp_V_Perc = False
  End If
  
  If Custo_Outros_Venda_Fixo = "P" Then
     Fixa_Outros_V_Perc = True
  Else
     Fixa_Outros_V_Perc = False
  End If


 
 
 Max = CDbl(Custo_Preço_Valor)
 Desejado = CDbl(Custo_Lucro_Perc)
 Max = Max * 500
 If Max < 1000 Then Max = 10000
 
 Call StatusMsg("")
 
 Fim = False
 
 
 
 Atual = (min + Max) / 2
 
 
 Diferença = 0.0001
 
 
 i = 0
 
 Do
  Rem Seta variáveis
 
 
  Venda = Atual

  Calcula_Lucro Venda, Custo_ICM_Venda_Perc, Custo_ICM_Venda_Valor, _
  Custo_IPI_Venda_Perc, Custo_IPI_Venda_Valor, Custo_Impostos_Venda_Perc, _
  Custo_Impostos_Venda_Valor, Custo_Outros_Venda_Valor, Custo_Outros_Venda_Perc, _
  Custo_Perc_Venda_Sem, Custo_Perc_Compra_Sem, Fixa_ICM_V_Perc, Fixa_IPI_V_Perc, _
  Fixa_Imp_V_Perc, Fixa_Outros_V_Perc, Lucro, Custo_Preço_Valor, Custo_Desconto_Valor, _
  Custo_Frete_Valor, Custo_Custo_Compra_Valor, Custo_Outros_Compra_Valor, _
  Custo_ICM_Compra_Valor, Custo_IPI_Compra_Valor
  
  
   
  Lucro_Perc = (Lucro / (CDbl(Custo_Preço_Valor) - CDbl(Custo_Desconto_Valor))) * CDbl(100)
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
      Venda = 0
      'Call StatusMsg("Não foi possível calcular o preço."
      'Exit Sub
    End If
  End If
 Loop Until Fim = True
 
 
' C_Venda_Valor.Text = Format(Atual, "#########0.00")
 
 
' Call StatusMsg("Preço calculado em " + Str(I) + " tentativas."
 
 
' B_Calc_Valor_Click
  Venda = Atual
  Custo_Preço_Venda = Atual
End Sub

Sub Calc_Lucro_Val()


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
 
 
 
  If Custo_ICM_Venda_Fixo = "P" Then
     Fixa_ICM_V_Perc = True
  Else
     Fixa_ICM_V_Perc = False
  End If
  
  If Custo_IPI_Venda_Fixo = "P" Then
     Fixa_IPI_V_Perc = True
  Else
     Fixa_IPI_V_Perc = False
  End If
  
  If Custo_Impostos_Venda_Fixo = "P" Then
     Fixa_Imp_V_Perc = True
  Else
     Fixa_Imp_V_Perc = False
  End If
  
  If Custo_Outros_Venda_Fixo = "P" Then
     Fixa_Outros_V_Perc = True
  Else
     Fixa_Outros_V_Perc = False
  End If




 
 Max = Custo_Lucro_Valor
 Desejado = Max
 Max = Max * 200
 
 
 
 Fim = False
 
 
 Atual = (min + Max) / 2
 
 
 Diferença = 0.0000001
 
 
 i = 0
 
 Do

    
  Venda = Atual


  Calcula_Lucro Venda, Custo_ICM_Venda_Perc, Custo_ICM_Venda_Valor, _
  Custo_IPI_Venda_Perc, Custo_IPI_Venda_Valor, Custo_Impostos_Venda_Perc, _
  Custo_Impostos_Venda_Valor, Custo_Outros_Venda_Valor, Custo_Outros_Venda_Perc, _
  Custo_Perc_Venda_Sem, Custo_Perc_Compra_Sem, Fixa_ICM_V_Perc, Fixa_IPI_V_Perc, _
  Fixa_Imp_V_Perc, Fixa_Outros_V_Perc, Lucro, Custo_Preço_Valor, Custo_Desconto_Valor, _
  Custo_Frete_Valor, Custo_Custo_Compra_Valor, Custo_Outros_Compra_Valor, _
  Custo_ICM_Compra_Valor, Custo_IPI_Compra_Valor
  
  

 ' Calcula_Lucro Venda, Venda_ICM_P, Venda_ICM_V, _
 ' Venda_IPI_P, Venda_IPI_V, Venda_Imp_P, Venda_Imp_V, _
 ' Venda_Outros_V, Venda_Outros_P, C_Venda_Sem_Nota.Text, C_Compra_Sem_Nota.Text, _
 ' -Fixa_ICM_V_Perc.Value, -Fixa_IPI_V_Perc.Value, -Fixa_Imp_V_Perc.Value, _
 ' -Fixa_Outros_V_Perc.Value, Lucro, Compra_Valor, Compra_Desc_V, _
 ' Compra_Frete_V, Compra_Finan_V, Compra_Outros_V, Compra_ICM_V, Compra_IPI_V
 
'  C_Venda_ICM_V.Caption = Venda_ICM_V
'  C_Venda_ICM_P.Text = Venda_ICM_P
'  C_Venda_IPI_V.Caption = Venda_IPI_V
'   C_Venda_IPI_P.Text = Venda_IPI_P
'  C_Venda_Imp_V.Text = Venda_Imp_V
'  C_Venda_Imp_P.Text = Venda_Imp_P
'  C_Venda_Outros_V.Text = Venda_Outros_V
'  C_Venda_Outros_P.Text = Venda_Outros_P


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
      'Call StatusMsg("Não foi possível calcular o preço."
      Venda = 0
      Exit Sub
      
    End If
  End If
 Loop Until Fim = True
 
 
 'C_Venda_Valor.Text = Format(Atual, "#########0.00")
 
 
 
' Call StatusMsg("Preço calculado em " + Str(I) + " tentativas."
 
 Venda = Atual
 Custo_Preço_Venda = Atual
 
  'B_Calc_Valor_Click
 
 End Sub
 
'-----------------------------------------------------------------------------------
'05/07/2002 - mpdea
'Implementado o suporte a transação com tratamento a erro
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub B_Ajusta_Click()
 Dim Código As String
 
 If MsgBox("Atenção, esta operação não poderá ser desfeita. Deseja continuar ?", vbQuestion + vbYesNo) = vbNo Then
'   DisplayMsg "Alterações NÃO efetivadas."
   Exit Sub
 End If
 
 B_Ajusta.Enabled = False
 B_Ajusta2.Enabled = False
 
 On Error GoTo ErrHandler
 ws.BeginTrans
 
 rsTempo.Index = "Código"
 rsPreços.Index = "Produto"
 Código = ""
Lp1:
 rsTempo.Seek ">", Código
 If rsTempo.NoMatch Then GoTo Fim
 Código = rsTempo("Código")
 
 Call StatusMsg("Aguarde, verificando produto " + Código)
 DoEvents
 
 
 If O_Marcados.Value = True Then
   If rsTempo("Alterar") = False Then GoTo Lp1
 End If
 
 rsPreços.Seek "=", Código, cboLista.Text
 If rsPreços.NoMatch Then GoTo Lp1
 
 rsPreços.Edit
   rsPreços("Preço") = Format(rsTempo("Preço Venda Atual"), "############0.00")
   rsPreços("Data Alteração") = Format(Data_Atual, "dd/mm/yyyy")
 rsPreços.Update
 
  'Atualiza o sincronismo para o produto WEB alterado
  Call WEB_SynchronizeProduct(Código)
 
 GoTo Lp1
 
Fim:
  ws.CommitTrans
  DisplayMsg "Alterações finalizadas."
  Exit Sub
  
ErrHandler:
  ws.Rollback
  MsgBox "Erro [" & Err.Number & "] - " & Err.Description, vbCritical, "Erro"
 
End Sub

'-----------------------------------------------------------------------------------
'05/07/2002 - mpdea
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub B_Ajusta2_Click()
 Dim Resp As Integer
 Dim Código As String
 
 Resp = MsgBox("Atenção, esta operação não poderá ser desfeita. Deseja continuar ?", vbOKCancel)
 If Resp = vbCancel Then
   DisplayMsg "Alterações NÃO efetivadas."
   Exit Sub
 End If
 
 On Error GoTo ErrHandle
 
 B_Ajusta.Enabled = False
 B_Ajusta2.Enabled = False
  
  Call ws.BeginTrans
  
 rsTempo.Index = "Código"
 rsPreços.Index = "Produto"
 rsProdutos.Index = "Código"
 Código = ""
Lp1:
 rsTempo.Seek ">", Código
 If rsTempo.NoMatch Then GoTo Fim
 Código = rsTempo("Código")
 
 Call StatusMsg("Aguarde, verificando produto " + Código)
 DoEvents
 
 If O_Marcados.Value = True Then
   If rsTempo("Alterar") = False Then GoTo Lp1
 End If
 
 rsPreços.Seek "=", Código, cboLista.Text
 If rsPreços.NoMatch Then GoTo Lp1
 
 rsProdutos.Seek "=", Código
 If rsProdutos.NoMatch Then GoTo Lp1
 
 rsPreços.Edit
  rsPreços("Preço") = Format(rsTempo("Preço Venda Atual"), "############0.00")
  rsPreços("Data Alteração") = Format(Data_Atual, "dd/mm/yyyy")
 rsPreços.Update
 
 rsProdutos.Edit
   rsProdutos("Custo Preço Valor") = Format(rsTempo("Preço Custo Atual"), "#########0.00")
   rsProdutos("Custo Custo Calculado") = Format(rsTempo("Preço Custo Calc Atu"), "#########0.00")
   rsProdutos("Custo Preço Venda") = Format(rsTempo("Preço Venda Atual"), "##########0.00")
   rsProdutos("Custo Lucro Valor") = Format(rsTempo("Lucro Atual"), "###########0.00")
   rsProdutos("Custo Lucro Perc") = Format(rsTempo("Lucro Atual Perc"), "############0.00")
 rsProdutos.Update
 
  'Atualiza o sincronismo para o produto WEB alterado
  Call WEB_SynchronizeProduct(Código)
 
 GoTo Lp1
 
Fim:
  Call ws.CommitTrans
  
  DisplayMsg "Alterações finalizadas."
  
  Exit Sub
  
ErrHandle:
  Call ws.Rollback
  MsgBox "Erro [" & Err.Number & "] - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub B_Monta_Click()
  Dim Produto As String
  Dim sSql As String
  Dim Preço_Venda As Double
  Dim Preço_Custo As Double
  Dim Custo As Double
  Dim Lucro As Double
  Dim Margem As Double
  Dim Val2 As Double
  Dim Rec_Preços As Recordset
'  Dim gsArredonda As String
  Dim dtDataAlter As Date
  Dim Fixa_ICM_C_Perc As Boolean
  
  If cboLista.Text = "" Then
    DisplayMsg "Informe a tabela de preços antes."
    Exit Sub
  End If

  If Not IsDate(Data.Text) Then
    dtDataAlter = "01/01/1990"
  Else
    dtDataAlter = CDate(Data.Text)
  End If

  Call StatusMsg("Aguarde, preparando arquivo temporário...")
  
  sSql = "Delete * From [Preço Custo]"
  dbTemp.Execute sSql
  Call StatusMsg("")

  Produto = ""
  
  gbToCancel = False
  gbRunning = True
  
  rsPreços.Index = "Produto"
  rsProdutos.Index = "Código"
  rsEstoque_Final.Index = "Produto"
Lp1:
  rsProdutos.Seek ">", Produto
  If rsProdutos.NoMatch Then GoTo Fim
  
  Produto = rsProdutos("Código")

  Call StatusMsg("Aguarde, verificando produto " + Produto)
  DoEvents
  If gbToCancel = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Rotina de Pesquisa em andamento. Deseja interrompê-la?"
    gnStyle = vbYesNo + vbQuestion
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbYes Then
      GoTo Fim
    End If
    gbToCancel = False
  End If

  If Produto = "0" Then GoTo Lp1

  If Not IsDate(rsProdutos("Data Alteração")) Then GoTo Lp1
  If CDate(rsProdutos("Data Alteração")) < dtDataAlter Then GoTo Lp1
  
  If Nome_Classe.Caption <> "" Then
    If rsProdutos("Classe") <> Val(Combo_Classe.Text) Then GoTo Lp1
  End If
  
  If Nome_Sub.Caption <> "" Then
    If rsProdutos("Sub Classe") <> Val(Combo_Sub.Text) Then GoTo Lp1
  End If
  
  
  If IsDate(Data_Estoque.Text) Then
    If CDate(Data_Estoque.Text) > CDate("01/01/1990") Then
      rsEstoque_Final.Seek ">", gnCodFilial, Produto, 0, 0, -1
        If rsEstoque_Final.NoMatch Then GoTo Lp1
        If rsEstoque_Final("Filial") <> gnCodFilial Then GoTo Lp1
        If rsEstoque_Final("Produto") <> Produto Then GoTo Lp1
        If Not IsDate(rsEstoque_Final("Última Data")) Then GoTo Lp1
        If CDate(rsEstoque_Final("Última Data") <= CDate(Data_Estoque.Text)) Then GoTo Lp1
    End If
  End If

  Preço_Venda = 0
  rsPreços.Seek "=", Produto, cboLista.Text
  If Not rsPreços.NoMatch Then
    Preço_Venda = rsPreços("Preço")
  End If
  
  Preço_Custo = 0
  rsPreços.Seek "=", Produto, "CUSTO"
  If rsPreços.NoMatch Then GoTo Lp1
  If rsPreços("Preço") = 0 Then GoTo Lp1
  
  Preço_Custo = rsPreços("Preço")
  


  GoSub Atualiza_Variáveis

  Custo = 0
  If Preço_Custo <> 0 Then
    Calcula_Custo Custo, Custo_Desconto_Fixo, Custo_Desconto_Valor, _
    Custo_Desconto_Perc, Preço_Custo, Custo_Frete_Fixo, Custo_Frete_Valor, _
    Custo_Frete_Perc, Custo_ICM_Compra_Fixo, Custo_ICM_Compra_Valor, _
    Custo_ICM_Compra_Perc, Custo_IPI_Compra_Fixo, Custo_IPI_Compra_Valor, _
    Custo_IPI_Compra_Perc, Custo_Custo_Compra_Fixo, Custo_Custo_Compra_Valor, _
    Custo_Custo_Compra_Perc, Custo_Outros_Compra_Fixo, Custo_Outros_Compra_Valor, _
    Custo_Outros_Compra_Perc
  End If

  rsTempo.AddNew
    rsTempo("Código") = Produto
    rsTempo("Nome") = rsProdutos("Nome")
    rsTempo("Preço Custo Anterior") = rsProdutos("Custo Preço Valor")
    rsTempo("Preço Custo Atual") = Preço_Custo
    
    Rem Alterado em 19/1/99
    If rsTempo("Preço Custo Anterior") = 0 Then rsTempo("Preço Custo Anterior") = Preço_Custo
        
    
    rsTempo("Preço Custo Calc Ant") = rsProdutos("Custo Custo Calculado")
    
    rsTempo("Preço Custo Calc Atu") = Custo
    rsTempo("Preço Venda Anterior") = rsProdutos("Custo Preço Venda")
    
    
    rsTempo("Lucro Anterior") = rsProdutos("Custo Lucro Valor")
    rsTempo("Lucro Anterior Perc") = rsProdutos("Custo Lucro Perc")


    Preço_Venda = Arredonda_Valor(Preço_Venda, gsArredonda)
    rsTempo("Preço Venda Atual") = Preço_Venda
    
    If rsProdutos("Custo Manter") = "P" And O_Percentual.Value = 0 Then
      Custo_Preço_Venda = Preço_Venda
      GoSub Calc_Lucro
      rsTempo("Lucro Atual") = Lucro
   
      Val2 = Preço_Custo - Custo_Desconto_Valor
      If Val2 = 0 Then Val2 = 0.01
      Margem = Lucro / Val2
      Margem = Margem * 100
      rsTempo("Lucro Atual Perc") = Format(Margem, "###,##0.00")
    End If
     
    If rsProdutos("Custo Manter") = "L" And O_Percentual.Value = 0 Then 'mantem valor lucro
      If Custo_Preço_Valor = 0 Then
        rsTempo("Lucro Atual") = 0
        rsTempo("Lucro Atual Perc") = 0
        rsTempo("Preço Venda Atual") = 0
      Else
        Custo_Preço_Venda = Preço_Venda
        Custo_Preço_Valor = Preço_Custo
        Calc_Lucro_Val
        
        Preço_Venda = Custo_Preço_Venda
        GoSub Calc_Lucro
        Preço_Venda = Arredonda_Valor(Preço_Venda, gsArredonda)
        rsTempo("Preço Venda Atual") = Preço_Venda
        
        rsTempo("Lucro Atual") = Lucro
        
        Val2 = Preço_Custo - Custo_Desconto_Valor
        If Val2 = 0 Then Val2 = 0.01
        Margem = Lucro / Val2
        Margem = Margem * 100
        rsTempo("Lucro Atual Perc") = Format(Margem, "###,##0.00")
        
        
      End If
    End If
    
    If rsProdutos("Custo Manter") = "U" Or O_Percentual.Value = 1 Then  'manter perc lucro
      If Custo_Preço_Valor = 0 Then Custo_Preço_Valor = Preço_Custo
      If Custo_Preço_Valor = 0 Then
        rsTempo("Lucro Atual") = 0
        rsTempo("Lucro Atual Perc") = 0
        rsTempo("Preço Venda Atual") = 0
      Else
        Custo_Preço_Venda = Preço_Venda
        Custo_Preço_Valor = Preço_Custo
        Calc_Lucro_Perc
        
        Custo_Preço_Venda = Arredonda_Valor(Custo_Preço_Venda, gsArredonda)
        rsTempo("Preço Venda Atual") = Custo_Preço_Venda
        Preço_Venda = Custo_Preço_Venda
        GoSub Calc_Lucro
                
        rsTempo("Lucro Atual") = Lucro
        
        Val2 = Preço_Custo - Custo_Desconto_Valor
        If Val2 = 0 Then Val2 = 0.01
        Margem = Lucro / Val2
        Margem = Margem * 100
        rsTempo("Lucro Atual Perc") = Format(Margem, "###,##0.00")
        
        
        
      End If
    End If
    
   rsTempo.Update


   GoTo Lp1


Fim:

  If O_Nome.Value = True Then sSql = "Select * From [Preço Custo] order by Nome"
  If O_Código.Value = True Then sSql = "Select * From [Preço Custo] order by Código"
  Set Rec_Preços = dbTemp.OpenRecordset(sSql, dbOpenDynaset)
 
  Grade1.DataMode = 1
  Set Data1.Recordset = Rec_Preços
  Grade1.Visible = False
  Grade1.DataMode = 0
  Grade1.ReBind

  Grade1.Columns(0).Width = 1100
  Grade1.Columns(0).Locked = True
  
  Grade1.Columns(1).Width = 2000
  Grade1.Columns(1).Locked = True
  
  Grade1.Columns(2).Width = 730
  Grade1.Columns(2).NumberFormat = "########0.00"
  Grade1.Columns(2).Caption = "PCAnt"
  Grade1.Columns(2).Locked = True
  
  Grade1.Columns(3).Width = 730
  Grade1.Columns(3).Caption = "PCAtu"
  Grade1.Columns(3).Locked = True
  Grade1.Columns(3).NumberFormat = "########0.00"
  
  Grade1.Columns(4).Width = 730
  Grade1.Columns(4).Caption = "PCCAnt"
  Grade1.Columns(4).Locked = True
  Grade1.Columns(4).NumberFormat = "########0.00"
  
  Grade1.Columns(5).Width = 730
  Grade1.Columns(5).Caption = "PCCAtu"
  Grade1.Columns(5).Locked = True
  Grade1.Columns(5).NumberFormat = "########0.00"
  
  Grade1.Columns(6).Width = 730
  Grade1.Columns(6).Caption = "PVAnt"
  Grade1.Columns(6).Locked = True
  Grade1.Columns(6).NumberFormat = "########0.00"

  Grade1.Columns(7).Width = 730
  Grade1.Columns(7).Caption = "LAnt"
  Grade1.Columns(7).Locked = True
  Grade1.Columns(7).NumberFormat = "########0.00"
  
  Grade1.Columns(8).Width = 730
  Grade1.Columns(8).Caption = "LAntP"
  Grade1.Columns(8).Locked = True
  Grade1.Columns(9).NumberFormat = "########0.00"
  
  Grade1.Columns(9).Width = 730
  Grade1.Columns(9).Caption = "PVAtu"
  Grade1.Columns(9).NumberFormat = "########0.00"
  
  Grade1.Columns(10).Width = 730
  Grade1.Columns(10).Caption = "LAtu"
  Grade1.Columns(10).Locked = True
  Grade1.Columns(10).NumberFormat = "########0.00"
  
  Grade1.Columns(11).Width = 700
  Grade1.Columns(11).Caption = "LAtuP"
  Grade1.Columns(11).Locked = True
  Grade1.Columns(11).NumberFormat = "########0.00"
  
  Grade1.Columns(12).Width = 350
  Grade1.Columns(12).Caption = "A"
    
  Grade1.Columns(12).Style = ssStyleCheckBox
  
  Grade1.Visible = True
  
  gbRunning = False
  gbToCancel = False
  
  Call StatusMsg("")
  
  Exit Sub

Atualiza_Variáveis:
  Custo_Preço_Valor = rsProdutos("Custo Preço Valor")
  Custo_Desconto_Perc = rsProdutos("Custo Desconto Perc")
  Custo_Desconto_Valor = rsProdutos("Custo Desconto Valor")
  If IsNull(rsProdutos("Custo Desconto Fixo")) Then
    Custo_Desconto_Fixo = "P"
  Else
    Custo_Desconto_Fixo = rsProdutos("Custo Desconto Fixo")
  End If
  
  Custo_Frete_Perc = rsProdutos("Custo Frete Perc")
  Custo_Frete_Valor = rsProdutos("Custo Frete Valor")
  
  If IsNull(rsProdutos("Custo Frete Fixo")) Then
    Custo_Frete_Fixo = "P"
  Else
    Custo_Frete_Fixo = rsProdutos("Custo Frete Fixo")
  End If
  
  If IsNull(rsProdutos("Custo ICM Compra Perc")) Then
    Custo_ICM_Compra_Perc = "P"
  Else
    Custo_ICM_Compra_Perc = rsProdutos("Custo ICM Compra Perc")
  End If
  
  Custo_ICM_Compra_Valor = rsProdutos("Custo ICM Compra Valor")
  
  If IsNull(rsProdutos("Custo ICM Compra Fixo")) Then
    Custo_ICM_Compra_Fixo = "P"
  Else
    Custo_ICM_Compra_Fixo = rsProdutos("Custo ICM Compra Fixo")
  End If
  Custo_IPI_Compra_Perc = rsProdutos("Custo IPI Compra Perc")
  Custo_IPI_Compra_Valor = rsProdutos("Custo IPI Compra Valor")
  
  If IsNull(rsProdutos("Custo IPI Compra Fixo")) Then
    Custo_IPI_Compra_Fixo = "P"
  Else
    Custo_IPI_Compra_Fixo = rsProdutos("Custo IPI Compra Fixo")
  End If
  
  Custo_Custo_Compra_Perc = rsProdutos("Custo Custo Finan Perc")
  Custo_Custo_Compra_Valor = rsProdutos("Custo Custo Finan Valor")
  
  If IsNull(rsProdutos("Custo Custo Finan Fixo")) Then
    Custo_Custo_Compra_Fixo = "P"
  Else
    Custo_Custo_Compra_Fixo = rsProdutos("Custo Custo Finan Fixo")
  End If
  
  Custo_Outros_Compra_Perc = rsProdutos("Custo Outros Compra Perc")
  Custo_Outros_Compra_Valor = rsProdutos("Custo Outros Compra Valor")
  
  If IsNull(rsProdutos("Custo Outros Compra Fixo")) Then
    Custo_Outros_Compra_Fixo = "P"
  Else
    Custo_Outros_Compra_Fixo = rsProdutos("Custo Outros Compra Fixo")
  End If
    
  Custo_Perc_Compra_Sem = rsProdutos("Custo Perc Compra Sem")
  Custo_Custo_Calculado = rsProdutos("Custo Custo Calculado")
  Custo_Preço_Venda = rsProdutos("Custo Preço Venda")
  Custo_ICM_Venda_Perc = rsProdutos("Custo ICM Venda Perc")
  Custo_ICM_Venda_Valor = rsProdutos("Custo ICM Venda Valor")
  
  If IsNull(rsProdutos("Custo ICM Venda Fixo")) Then
    Custo_ICM_Venda_Fixo = "P"
  Else
    Custo_ICM_Venda_Fixo = rsProdutos("Custo ICM Venda Fixo")
  End If
    
  Custo_IPI_Venda_Perc = rsProdutos("Custo IPI Venda Perc")
  Custo_IPI_Venda_Valor = rsProdutos("Custo IPI Venda Valor")
  
  If IsNull(rsProdutos("Custo IPI Venda Fixo")) Then
    Custo_IPI_Venda_Fixo = "P"
  Else
    Custo_IPI_Venda_Fixo = rsProdutos("Custo IPI Venda Fixo")
  End If
    
  Custo_Impostos_Venda_Perc = rsProdutos("Custo Impostos Perc")
  Custo_Impostos_Venda_Valor = rsProdutos("Custo Impostos Valor")
  
  If IsNull(rsProdutos("Custo Impostos Fixo")) Then
    Custo_Impostos_Venda_Fixo = "P"
  Else
    Custo_Impostos_Venda_Fixo = rsProdutos("Custo Impostos Fixo")
  End If
  
  Custo_Outros_Venda_Perc = rsProdutos("Custo Outros Venda Perc")
  Custo_Outros_Venda_Valor = rsProdutos("Custo Outros Venda Valor")
  
  If IsNull(rsProdutos("Custo Outros Venda Fixo")) Then
    Custo_Outros_Venda_Fixo = "P"
  Else
    Custo_Outros_Venda_Fixo = rsProdutos("Custo Outros Venda Fixo")
  End If
    
  Custo_Perc_Venda_Sem = rsProdutos("Custo Perc Venda Sem")
  Custo_Lucro_Perc = rsProdutos("Custo Lucro Perc")
  Custo_Lucro_Valor = rsProdutos("Custo Lucro Valor")
  Custo_Manter = "P"
  If Not IsNull(rsProdutos("Custo Manter")) Then Custo_Manter = rsProdutos("Custo Manter")
  
 If Custo_ICM_Compra_Fixo = "P" Then
     Fixa_ICM_C_Perc = True
  Else
     Fixa_ICM_C_Perc = False
  End If
  
  If Custo_IPI_Venda_Fixo = "P" Then
     Fixa_IPI_V_Perc = True
  Else
     Fixa_IPI_V_Perc = False
  End If
  
  If Custo_Impostos_Venda_Fixo = "P" Then
     Fixa_Imp_V_Perc = True
  Else
     Fixa_Imp_V_Perc = False
  End If
  
  If Custo_Outros_Venda_Fixo = "P" Then
     Fixa_Outros_V_Perc = True
  Else
     Fixa_Outros_V_Perc = False
  End If
  
  Return
  
  
Calc_Lucro:
  If Custo_ICM_Venda_Fixo = "P" Then
     Fixa_ICM_V_Perc = True
  Else
     Fixa_ICM_V_Perc = False
  End If
  
  If Custo_IPI_Venda_Fixo = "P" Then
     Fixa_IPI_V_Perc = True
  Else
     Fixa_IPI_V_Perc = False
  End If
  
  If Custo_Impostos_Venda_Fixo = "P" Then
     Fixa_Imp_V_Perc = True
  Else
     Fixa_Imp_V_Perc = False
  End If
  
  If Custo_Outros_Venda_Fixo = "P" Then
     Fixa_Outros_V_Perc = True
  Else
     Fixa_Outros_V_Perc = False
  End If


'  Calcula_Lucro Custo_Preço_Venda, Custo_ICM_Venda_Perc, Custo_ICM_Venda_Valor, _
'   Custo_IPI_Venda_Perc, Custo_IPI_Venda_Valor, Custo_Impostos_Venda_Perc, _
'   Custo_Impostos_Venda_Valor, Custo_Outros_Venda_Valor, _
'   Custo_Outros_Venda_Perc, Custo_Perc_Venda_Sem, Custo_Perc_Compra_Sem, _
'   Fixa_ICM_V_Perc, Fixa_IPI_V_Perc, Fixa_Imp_V_Perc, _
'   Fixa_Outros_V_Perc, Lucro, Custo_Preço_Valor, Custo_Desconto_Valor, _
'   Custo_Frete_Valor, Custo_Custo_Compra_Valor, Custo_Outros_Compra_Valor, _
'   Custo_ICM_Compra_Valor, Custo_IPI_Compra_Valor
  
  
  Custo_ICM_Venda_Valor = Preço_Venda * Custo_ICM_Venda_Perc / 100
  Custo_IPI_Venda_Valor = Preço_Venda * Custo_IPI_Venda_Perc / 100
  Custo_Impostos_Venda_Valor = Preço_Venda * Custo_Impostos_Venda_Perc / 100
  Custo_Outros_Venda_Valor = Preço_Venda * Custo_Outros_Venda_Perc / 100
  
  Calcula_Lucro Preço_Venda, Custo_ICM_Venda_Perc, Custo_ICM_Venda_Valor, _
   Custo_IPI_Venda_Perc, Custo_IPI_Venda_Valor, Custo_Impostos_Venda_Perc, _
   Custo_Impostos_Venda_Valor, Custo_Outros_Venda_Valor, _
   Custo_Outros_Venda_Perc, Custo_Perc_Venda_Sem, Custo_Perc_Compra_Sem, _
   Fixa_ICM_V_Perc, Fixa_IPI_V_Perc, Fixa_Imp_V_Perc, _
   Fixa_Outros_V_Perc, Lucro, Preço_Custo, Custo_Desconto_Valor, _
   Custo_Frete_Valor, Custo_Custo_Compra_Valor, Custo_Outros_Compra_Valor, _
   Custo_ICM_Compra_Valor, Custo_IPI_Compra_Valor
  
  Return
  
End Sub


Private Sub cmbArredonda_Click()
  
  gsArredonda = "000"
  Select Case cmbArredonda.ListIndex
    Case 1
      gsArredonda = "005"
    Case 2
      gsArredonda = "010"
    Case 3
      gsArredonda = "050"
    Case 4
      gsArredonda = "100"
    Case 5
      gsArredonda = "500"
    Case 6
      gsArredonda = "1000"
  End Select

End Sub

Private Sub Combo_Classe_CloseUp()

 Combo_Classe.Text = Combo_Classe.Columns(1).Text
 Combo_Classe_LostFocus

End Sub

Private Sub Combo_Classe_KeyPress(KeyAscii As Integer)
  If Not Combo_Classe.DroppedDown Then
    KeyAscii = gnLimitKeyPress(Combo_Classe, 4, KeyAscii, True)
  End If
End Sub

Private Sub Combo_Classe_LostFocus()
  Nome_Classe.Caption = ""
  rsClasses.Index = "Código"
  
  If IsNull(Combo_Classe.Text) Then Exit Sub
  If Combo_Classe.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Classe.Text) Then Exit Sub
  If Val(Combo_Classe.Text) < 1 Then Exit Sub
  
  rsClasses.Seek "=", Val(Combo_Classe.Text)
  If rsClasses.NoMatch Then Exit Sub
  
  Nome_Classe.Caption = rsClasses("Nome") & ""

End Sub

Private Sub Combo_Sub_CloseUp()

 Combo_Sub.Text = Combo_Sub.Columns(1).Text
 Combo_Sub_LostFocus


End Sub

Private Sub Combo_Sub_KeyPress(KeyAscii As Integer)
  If Not Combo_Sub.DroppedDown Then
    KeyAscii = gnLimitKeyPress(Combo_Sub, 4, KeyAscii, True)
  End If
End Sub

Private Sub Combo_Sub_LostFocus()

  Nome_Sub.Caption = ""
  rsSub_Classes.Index = "Código"
  
  If IsNull(Combo_Sub.Text) Then Exit Sub
  If Combo_Sub.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Sub.Text) Then Exit Sub
  If Val(Combo_Sub.Text) < 1 Then Exit Sub
  
  rsSub_Classes.Seek "=", Val(Combo_Sub.Text)
  If rsSub_Classes.NoMatch Then Exit Sub
  
  Nome_Sub.Caption = rsSub_Classes("Nome") & ""

End Sub

Private Sub Data_Estoque_LostFocus()
 Data_Estoque.Text = Ajusta_Data(Data_Estoque.Text)
End Sub

Private Sub Data_Estoque_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Estoque.Text = frmCalendario.gsDateCalender(Data_Estoque.Text)
  End Select
End Sub

Private Sub Data_LostFocus()
  Data.Text = Ajusta_Data(Data.Text)
End Sub

Private Sub Data_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data.Text = frmCalendario.gsDateCalender(Data.Text)
  End Select
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)
  
  Set rsTempo = dbTemp.OpenRecordset("Preço Custo")
  Set rsProdutos = db.OpenRecordset("Produtos")
  Set rsPreços = db.OpenRecordset("Preços")
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsEstoque_Final = db.OpenRecordset("Estoque Final", , dbReadOnly)
  
  Data1.DatabaseName = gsTempDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  datPrecos.DatabaseName = gsQuickDBFileName
  
  Data.Mask = ""
  Data.Text = ""
  Data.Mask = "##/##/####"
  
  Data_Estoque.Mask = ""
  Data_Estoque.Text = ""
  Data_Estoque.Mask = "##/##/####"
  
  cmbArredonda.ListIndex = 0
'  cmbArredonda.Text = cmbArredonda.List(0)
  
  gbRunning = False
  gbToCancel = False
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If gbRunning = True Then
    gbToCancel = True
    Cancel = True
    Exit Sub
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsTempo.Close
  rsProdutos.Close
  rsPreços.Close
  rsClasses.Close
  rsSub_Classes.Close
  rsEstoque_Final.Close
  Set rsTempo = Nothing
  Set rsProdutos = Nothing
  Set rsPreços = Nothing
  Set rsClasses = Nothing
  Set rsSub_Classes = Nothing
  Set rsEstoque_Final = Nothing

End Sub

Private Sub Grade1_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
 Dim Coluna As Integer
 Dim Aux As Double
 
 Coluna = Grade1.Col
 If Coluna = 2 Then
   Explica.Caption = "Preço de Custo Anterior - Este é o preço que está gravado no campo Custo do Cadastro de Produtos."
 End If
 
 If Coluna = 3 Then
   Explica.Caption = "Preço de Custo Atual - Este é o preço de custo de que está na tabela CUSTO. Ele pode ter sido digitado ou gravado automaticamente no momento da última compra."
 End If
 
 If Coluna = 4 Then
   Explica.Caption = "Preço de Custo Anterior Calculado - Este é o preço de custo final já com impostos, frete e os outros campos da pasta Custo do Cadastro de Produtos."
 End If
 
 If Coluna = 5 Then
   Explica.Caption = "Preço de Custo Atual Calculado - Este é o preço de custo final já com impostos, frete e os outros campos da pasta Custo do Cadastro de Produtos."
 End If
 
 If Coluna = 6 Then
   Explica.Caption = "Preço de Venda Anterior - Este é o preço de venda digitado no campo Preço de Venda do Cadastro de Produtos."
 End If
 
 If Coluna = 7 Then
   Explica.Caption = "Lucro Anterior - É o lucro em reais com o preço de venda e custo anterior."
 End If
 
 If Coluna = 8 Then
   Explica.Caption = "Lucro Anterior Percentual - É o lucro percentual com o preço de venda e custo anterior."
 End If
 
 If Coluna = 9 Then
   Explica.Caption = "Preço de Venda Atual - É o novo preço de venda calculado pelo sistema. Você pode alterar este preço caso deseje."
 End If
 
 If Coluna = 10 Then
   Explica.Caption = "Lucro Atual - É o lucro em reais com o preço de venda e custo atuais."
 End If
 
 If Coluna = 11 Then
   Explica.Caption = "Lucro Percentual - É o lucro percentual com o preço de venda e custo atuais."
 End If
 
 If Coluna = 12 Then
   Explica.Caption = "Alterar - Marque o campo para os preços que desejar alterar."
 End If
 
 
 L_C_An.Caption = Format(Grade1.Columns(2).Text, "###,###,##0.00")
 L_CT_An.Caption = Format(Grade1.Columns(4).Text, "###,###,##0.00")
 L_L_An.Caption = Format(Grade1.Columns(7).Text, "###,###,##0.00")
 L_LP_An.Caption = Format(Grade1.Columns(8).Text, "###,###,##0.00")
 L_V_An.Caption = Format(Grade1.Columns(6).Text, "###,###,##0.00")
 
 L_C_At.Caption = Format(Grade1.Columns(3).Text, "###,###,##0.00")
 L_CT_At.Caption = Format(Grade1.Columns(5).Text, "###,###,##0.00")
 L_L_At.Caption = Format(Grade1.Columns(10).Text, "###,###,##0.00")
 L_LP_At.Caption = Format(Grade1.Columns(11).Text, "###,###,##0.00")
 L_V_At.Caption = Format(Grade1.Columns(9).Text, "###,###, ##0.00")
 
 Aux = Retorna_Valor(L_C_At.Caption) - Retorna_Valor(L_C_An.Caption)
 L_C_Di.Caption = Format(Aux, "###,###,##0.00")
 
 Aux = Retorna_Valor(L_CT_At.Caption) - Retorna_Valor(L_CT_An.Caption)
 L_CT_Di.Caption = Format(Aux, "###,###,##0.00")
 
 Aux = Retorna_Valor(L_L_At.Caption) - Retorna_Valor(L_L_An.Caption)
 L_L_Di.Caption = Format(Aux, "###,###,##0.00")

 Aux = Retorna_Valor(L_LP_At.Caption) - Retorna_Valor(L_LP_An.Caption)
 L_LP_Di.Caption = Format(Aux, "###,###,##0.00")

 Aux = Retorna_Valor(L_V_At.Caption) - Retorna_Valor(L_V_An.Caption)
 L_V_Di.Caption = Format(Aux, "###,###,##0.00")


  
 'Rem Preto  &H80000012&
 'REm Vermelho &H000000FF&
 'Rem Azul  &H00FF0000&
 
 L_C_Di.ForeColor = &H80000012
 If Retorna_Valor(L_C_Di.Caption) > 0 Then L_C_Di.ForeColor = &HFF0000
 If Retorna_Valor(L_C_Di.Caption) < 0 Then L_C_Di.ForeColor = &HFF&
 
 L_CT_Di.ForeColor = &H80000012
 If Retorna_Valor(L_CT_Di.Caption) > 0 Then L_CT_Di.ForeColor = &HFF0000
 If Retorna_Valor(L_CT_Di.Caption) < 0 Then L_CT_Di.ForeColor = &HFF&
 
 L_L_Di.ForeColor = &H80000012
 If Retorna_Valor(L_L_Di.Caption) > 0 Then L_L_Di.ForeColor = &HFF0000
 If Retorna_Valor(L_L_Di.Caption) < 0 Then L_L_Di.ForeColor = &HFF&
 
 L_LP_Di.ForeColor = &H80000012
 If Retorna_Valor(L_LP_Di.Caption) > 0 Then L_LP_Di.ForeColor = &HFF0000
 If Retorna_Valor(L_LP_Di.Caption) < 0 Then L_LP_Di.ForeColor = &HFF&
 
 L_V_Di.ForeColor = &H80000012
 If Retorna_Valor(L_V_Di.Caption) > 0 Then L_V_Di.ForeColor = &HFF0000
 If Retorna_Valor(L_V_Di.Caption) < 0 Then L_V_Di.ForeColor = &HFF&
 
End Sub

Private Sub O_Normal_Click()
 Grade1.Font.Size = 8
End Sub

Private Sub O_Pequena_Click()
 Grade1.Font.Size = 7
End Sub

