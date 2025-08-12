VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Realizar baixa de Cheques Pré-Datados"
   ClientHeight    =   7125
   ClientLeft      =   165
   ClientTop       =   1425
   ClientWidth     =   12840
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
   Icon            =   "ManCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7125
   ScaleWidth      =   12840
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
      Left            =   7500
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Banco"
      Height          =   840
      Left            =   3840
      TabIndex        =   35
      Top             =   75
      Width           =   3285
      Begin SSDataWidgets_B.SSDBCombo cboBanco 
         Bindings        =   "ManCheques.frx":4E95A
         Height          =   315
         Left            =   105
         TabIndex        =   2
         ToolTipText     =   "Escolha a Administradora ou deixe em branco para todas"
         Top             =   330
         Width           =   3045
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
         _ExtentX        =   5371
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdCheques 
      Height          =   2640
      Left            =   60
      TabIndex        =   10
      Top             =   1890
      Width           =   12690
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
      AllowUpdate     =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      BackColorOdd    =   16777152
      RowHeight       =   423
      ExtraHeight     =   53
      Columns.Count   =   10
      Columns(0).Width=   873
      Columns(0).Caption=   "Filial"
      Columns(0).Name =   "Filial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1270
      Columns(1).Caption=   "Cód."
      Columns(1).Name =   "Cód."
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2910
      Columns(2).Caption=   "Banco"
      Columns(2).Name =   "NomeBanco"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2805
      Columns(3).Caption=   "Cheque"
      Columns(3).Name =   "Cheque"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2355
      Columns(4).Caption=   "Data Vcto"
      Columns(4).Name =   "Data Vcto"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   10
      Columns(4).Mask =   "##/##/####"
      Columns(4).PromptInclude=   -1  'True
      Columns(4).PromptChar=   32
      Columns(5).Width=   2805
      Columns(5).Caption=   "Valor"
      Columns(5).Name =   "Valor"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1746
      Columns(6).Caption=   "Depositado"
      Columns(6).Name =   "Depositado"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   11
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(7).Width=   2064
      Columns(7).Caption=   "Cliente"
      Columns(7).Name =   "CodCli"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3281
      Columns(8).Caption=   "Nome"
      Columns(8).Name =   "NomeCli"
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
      _ExtentX        =   22384
      _ExtentY        =   4657
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
   Begin VB.Frame Frame3 
      Caption         =   "Seleção de Cheques"
      Height          =   1755
      Left            =   60
      TabIndex        =   30
      Top             =   4980
      Width           =   12690
      Begin VB.CommandButton cmdBaixar 
         BackColor       =   &H00C0FFC0&
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
         Height          =   465
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   285
         Width           =   4305
      End
      Begin VB.Frame Quadro1 
         Caption         =   "Depositar em :"
         Height          =   855
         Left            =   2205
         TabIndex        =   32
         Top             =   765
         Visible         =   0   'False
         Width           =   2265
         Begin VB.OptionButton O_Não_Determinado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "&Não determinado"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   12
            Top             =   225
            Width           =   1785
         End
         Begin VB.OptionButton O_Conta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "&Conta Corrente"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   13
            Top             =   540
            Value           =   -1  'True
            Width           =   1575
         End
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
         Left            =   165
         Picture         =   "ManCheques.frx":4E96E
         ScaleHeight     =   705
         ScaleWidth      =   1770
         TabIndex        =   31
         Top             =   885
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.CommandButton B_Deposita 
         BackColor       =   &H00C0FFC0&
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
         Height          =   435
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Deposita Cheques"
         Top             =   510
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Seleciona_Tudo 
         BackColor       =   &H00C0FFC0&
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
         Height          =   435
         Left            =   9945
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Seleciona Todos os Cheques"
         Top             =   255
         Width           =   2625
      End
      Begin VB.CommandButton B_Apaga 
         BackColor       =   &H00C0FFC0&
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
         Height          =   435
         Left            =   9945
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Apaga os Cheques Selecionados"
         Top             =   1185
         Width           =   2625
      End
      Begin VB.CommandButton B_Cancela_Sel 
         BackColor       =   &H00C0FFC0&
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
         Height          =   435
         Left            =   9945
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Cancela atual seleção"
         Top             =   720
         Width           =   2625
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
         Bindings        =   "ManCheques.frx":503B8
         DataSource      =   "Data2"
         Height          =   345
         Left            =   5280
         TabIndex        =   14
         Top             =   540
         Visible         =   0   'False
         Width           =   975
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
         _ExtentX        =   1720
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin VB.Label Label6 
         Caption         =   "Conta"
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Nome_Conta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   5295
         TabIndex        =   33
         Top             =   1020
         Visible         =   0   'False
         Width           =   3375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Período de Vencimento"
      Height          =   840
      Left            =   75
      TabIndex        =   27
      Top             =   75
      Width           =   3720
      Begin MSMask.MaskEdBox Vcto_Final 
         Height          =   345
         Left            =   2340
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   315
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
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
         Height          =   345
         Left            =   600
         TabIndex        =   0
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   315
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
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
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   29
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1935
         TabIndex        =   28
         Top             =   375
         Width           =   435
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
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   6900
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   840
      Left            =   7170
      TabIndex        =   20
      Top             =   75
      Width           =   2865
      Begin VB.OptionButton O_Não_Depositados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Não depositado"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   4
         Top             =   540
         Width           =   1605
      End
      Begin VB.OptionButton O_Depositados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Depositados"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   270
         Width           =   1365
      End
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1800
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   795
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
      Left            =   1845
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM Bancos ORDER BY Nome"
      Top             =   6900
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
      Height          =   435
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1380
      Width           =   12645
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      ForeColor       =   &H00FF0000&
      Height          =   840
      Left            =   10080
      TabIndex        =   19
      Top             =   75
      Width           =   2670
      Begin VB.OptionButton O_Valor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Valor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   375
         Width           =   735
      End
      Begin VB.OptionButton O_Data 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Data"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1020
         TabIndex        =   7
         Top             =   375
         Width           =   675
      End
      Begin VB.OptionButton O_Banco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Banco"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   375
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "ManCheques.frx":503CC
      DataSource      =   "datFilial"
      Height          =   345
      Left            =   540
      TabIndex        =   36
      Top             =   960
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
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label lblFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1620
      TabIndex        =   38
      Top             =   960
      Width           =   5835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   1005
      Width           =   300
   End
   Begin VB.Label Tot_Selec 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   11400
      TabIndex        =   26
      Top             =   4590
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Não Depositado Selecionado"
      Height          =   255
      Left            =   9090
      TabIndex        =   25
      Top             =   4620
      Width           =   2265
   End
   Begin VB.Label Label4 
      Caption         =   "Total Não Depositado"
      Height          =   255
      Left            =   4530
      TabIndex        =   24
      Top             =   4620
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   255
      Left            =   60
      TabIndex        =   23
      Top             =   4620
      Width           =   435
   End
   Begin VB.Label Tot_Não 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6270
      TabIndex        =   22
      Top             =   4590
      Width           =   1335
   End
   Begin VB.Label Tot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   540
      TabIndex        =   21
      Top             =   4560
      Width           =   1335
   End
End
Attribute VB_Name = "frmManCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsContas As Recordset
Dim rsLançamentos As Recordset
Dim gnSomaValSelected As Single
Private gbBaixar As Boolean
Private gnCodBanco As Integer

Sub Soma_Seleção()
  Dim nRow As Long
  Dim bm As Variant

  gnSomaValSelected = 0
  For nRow = 0 To (grdCheques.SelBookmarks.Count - 1)
    bm = grdCheques.SelBookmarks(nRow)
    grdCheques.Bookmark = bm
    If grdCheques.Columns("Depositado").Value = False Then
      gnSomaValSelected = gnSomaValSelected + grdCheques.Columns("Valor").Value
    End If
  Next nRow

  Tot_Selec.Caption = Format(gnSomaValSelected, "###,###,##0.00")

End Sub

Private Sub B_Apaga_Click()
  Dim rsCheques As Recordset
  Dim nRow As Long
  Dim bm As Variant
  
  Call StatusMsg("")
  If grdCheques.SelBookmarks.Count = 0 Then
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
  
  Set rsCheques = db.OpenRecordset("SELECT * FROM [Contas a Receber] WHERE Tipo = 'C'", dbOpenDynaset)
  
  Call ws.BeginTrans
  
  For nRow = 0 To (grdCheques.SelBookmarks.Count - 1)
    bm = grdCheques.SelBookmarks(nRow)
    grdCheques.Bookmark = bm
    rsCheques.FindFirst "Contador = " & grdCheques.Columns("Contador").Value
    If Not rsCheques.NoMatch Then
    
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
        "Cli:" & rsCheques("Cliente") & "- Seq:" & rsCheques("Sequência") & "- NF:" & rsCheques("Nota") & "- Venc:" & rsCheques("Vencimento") & "- Valor:" & rsCheques("Valor"), _
        "frmManCheques_B_Apaga_Click", _
        "Contas a Receber", g_strArquivoSystemLog
      End If
      
      rsCheques.Delete
    End If
  Next nRow
  
  Call ws.CommitTrans
  
  rsCheques.Close
  Set rsCheques = Nothing
  
  grdCheques.Update
  
  Call B_Monta_Click

End Sub

Private Sub B_Cancela_Sel_Click()
  grdCheques.SelBookmarks.RemoveAll
  Call Soma_Seleção
  gbBaixar = True
  Call cmdBaixar_Click
End Sub

Private Sub B_Deposita_Click()
  Dim rsCheques As Recordset
  Dim Saldo_Ant As Double
  Dim nRow As Integer
  Dim bm As Variant
  
  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked As Integer
  
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

  Set rsCheques = db.OpenRecordset("SELECT * FROM [Contas a Receber] WHERE Tipo = 'C'", dbOpenDynaset)
  
  Call ws.BeginTrans
  blnInTransaction = True
  
  With rsCheques
    For nRow = 0 To (grdCheques.SelBookmarks.Count - 1)
      bm = grdCheques.SelBookmarks(nRow)
      grdCheques.Bookmark = bm
      .FindFirst "Contador = " & grdCheques.Columns("Contador").Value
      If Not .NoMatch Then
        If .Fields("Processado").Value = False Then
          '10/09/2007 - Anderson
          'Gera arquivo log do sistema
          If g_bolSystemLog Then
            SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
            "Cli:" & rsCheques("Cliente") & "- Seq:" & rsCheques("Sequência") & "- NF:" & rsCheques("Nota") & "- Venc:" & rsCheques("Vencimento") & "- Valor:" & rsCheques("Valor"), _
            "frmManCheques_B_Deposita_Click", _
            "Contas a Receber", g_strArquivoSystemLog
          End If
          .LockEdits = True
          .Edit
          .Fields("Processado").Value = True
          .Fields("Valor Recebido").Value = .Fields("Valor").Value
          .Fields("Data Recebimento").Value = Data_Atual
          .Fields("Data Alteração").Value = Data_Atual
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
  rsLançamentos("Descrição") = "Cheques pré depositados"
  rsLançamentos("Saldo Anterior") = Saldo_Ant
  rsLançamentos("Crédito") = CDbl(gnSomaValSelected)
  rsLançamentos("Saldo Atual") = Saldo_Ant + CDbl(gnSomaValSelected)
  rsLançamentos.Update
  
  Call ws.CommitTrans
  blnInTransaction = False
  rsCheques.Close
  Set rsCheques = Nothing
  
  DisplayMsg "Cheques depositados."
  
  grdCheques.Enabled = True
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Cnt:" & Combo_Conta.Text & " SldAn:" & Saldo_Ant & " SldAt:" & rsLançamentos("Saldo Atual") & " Nm:" & O_Conta.Value, 80) & "', 'CNT_REC: baixar-cheq')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
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
  Call LoadGridCheques
  B_Monta.Enabled = True
  
End Sub


Private Sub cboBanco_CloseUp()
  Dim bm As Variant
  bm = cboBanco.GetBookmark(0)
  gnCodBanco = cboBanco.Columns("Código").CellValue(bm)
End Sub

Private Sub cboBanco_LostFocus()
  If Len(Trim(cboBanco.Text)) = 0 Then
    gnCodBanco = 0
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

Private Sub cmdBaixar_Click()
  If gbBaixar = False Then
    Call StatusMsg("")
    If gnSomaValSelected = 0 Then
      DisplayMsg "Selecione antes um ou mais cheques não depositados para a operação."
      Exit Sub
    End If
    grdCheques.Enabled = False
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
    grdCheques.Enabled = True
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

Private Sub Form_Load()
 
  Call CenterForm(Me)
  
  gbBaixar = False
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários")
  
  Call GetSettings
  
  cboFilial.Text = gnCodFilial
  cboFilial_LostFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Call SaveSetting("QuickStore", "CRMan", "Data1", Vcto_Inicial.Text)
  Call SaveSetting("QuickStore", "CRMan", "Data2", Vcto_Final.Text)
  
  rsContas.Close
  rsLançamentos.Close
  Set rsContas = Nothing
  Set rsLançamentos = Nothing
  
End Sub

Private Sub GetSettings()
  Vcto_Final.Text = GetSetting("QuickStore", "CRMan", "Data2", CDate(Date))
  Vcto_Inicial.Text = GetSetting("QuickStore", "CRMan", "Data1", CDate(Date))
End Sub

Private Sub grdCheques_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
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

Private Sub grdCheques_InitColumnProps()
  grdCheques.Columns("Valor").NumberFormat = "##,###,##0.00"
End Sub

Private Sub grdCheques_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

  Call Soma_Seleção

End Sub

Private Sub Seleciona_Tudo_Click()
  Dim nRow As Integer
  
  grdCheques.MoveFirst
  
  grdCheques.SelBookmarks.RemoveAll
   
  For nRow = 0 To (grdCheques.Rows - 1)
    grdCheques.SelBookmarks.Add grdCheques.Bookmark
    grdCheques.MoveNext
  Next nRow
  grdCheques.MoveFirst
  
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

Private Sub LoadGridCheques()
  Dim rsCheques As Recordset
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
  
  bAllow = grdCheques.AllowAddNew
  grdCheques.AllowAddNew = True
  grdCheques.AllowUpdate = True
  
  sSql = "SELECT Filial, Banco as CodBanco, Bancos.Nome as NomeBanco, Cheque, Vencimento, Valor, "
  sSql = sSql + "Processado, Cli_For.Código as CodCli, Cli_For.Nome as NomeCli, [Valor Recebido], Contador "
  sSql = sSql & "FROM ([Contas a Receber]"
  sSql = sSql + " INNER JOIN Bancos ON [Contas a Receber].Banco = Bancos.Código) "
  sSql = sSql & " INNER JOIN Cli_For ON [Contas a Receber].Cliente = Cli_For.Código"
  sSql = sSql + " WHERE Filial " & sFilial & " AND Vencimento >= " + Data_Ini
  sSql = sSql + " And Vencimento <= " + Data_Fim + " AND [Contas a Receber].Tipo = 'C'"
  
  If O_Depositados.Value = True Then
    sSql = sSql + " AND Processado = True"
  Else
    If O_Não_Depositados.Value = True Then
      sSql = sSql + " AND Processado = False"
    End If
  End If
  
  If gnCodBanco > 0 Then
    sSql = sSql & " And [Contas a Receber].Banco = " & gnCodBanco
  End If
  
  If O_Banco.Value = True Then sSql = sSql + " ORDER BY Bancos.Nome"
  If O_Data.Value = True Then sSql = sSql + " ORDER BY Vencimento"
  If O_Valor.Value = True Then sSql = sSql + " ORDER BY Valor"
  
  Call StatusMsg("Aguarde, montando tabela...")
  
  Set rsCheques = db.OpenRecordset(sSql, dbOpenDynaset)

  grdCheques.RemoveAll
  grdCheques.Redraw = False
  
  If Not rsCheques.EOF Then
    With rsCheques
      .MoveFirst
      Do While Not .EOF
        DoEvents
        sRecord = .Fields("Filial") & vbTab & _
          .Fields("CodBanco") & vbTab & _
          .Fields("NomeBanco") & vbTab & _
          .Fields("Cheque") & vbTab & _
          .Fields("Vencimento") & vbTab & _
          .Fields("Valor") & vbTab & _
          .Fields("Processado") & vbTab & _
          .Fields("CodCli") & vbTab & _
          .Fields("NomeCli") & vbTab & _
          .Fields("Contador")
        grdCheques.AddItem sRecord
        .MoveNext
      Loop
      .MoveFirst
    End With
    
    'Verifica Totais
    Tot_Não_Depositado = 0
    Total = 0
    rsCheques.MoveFirst
    Do Until rsCheques.EOF
      If rsCheques("Processado") = False Then
        Tot_Não_Depositado = Tot_Não_Depositado + rsCheques("Valor").Value
      End If
      Total = Total + rsCheques("Valor").Value
      rsCheques.MoveNext
    Loop
    
    Tot.Caption = Format(Total, "###,###,##0.00")
    Tot_Não.Caption = Format(Tot_Não_Depositado, "###,###,##0.00")
    Tot_Selec.Caption = "0,00"
    
  Else
  
    DisplayMsg "Nenhuma conta de cheques pré-datados encontrada segundo os critérios fornecidos."
  
  End If

  grdCheques.Scroll -99, -99
  grdCheques.Redraw = True
  
  grdCheques.AllowAddNew = bAllow
  grdCheques.AllowUpdate = bAllow

  rsCheques.Close
  Set rsCheques = Nothing
  Call StatusMsg("")
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao ler registros do Contas a Receber."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub


