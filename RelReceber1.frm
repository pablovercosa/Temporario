VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRelReceber1 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Contas a Receber por Data de Vencimento"
   ClientHeight    =   8535
   ClientLeft      =   3015
   ClientTop       =   2760
   ClientWidth     =   16095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "RelReceber1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8535
   ScaleWidth      =   16095
   Begin VB.Data Data4 
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
      Height          =   345
      Left            =   8310
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cli_For"
      Top             =   8130
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmd_imprimirOperDet 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8070
      Width           =   2115
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de busca de dados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   60
      TabIndex        =   28
      Top             =   60
      Width           =   15975
      Begin VB.OptionButton opt_especifica 
         Appearance      =   0  'Flat
         Caption         =   "Busca Específica"
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
         Height          =   225
         Left            =   8310
         TabIndex        =   30
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton opt_generico 
         Appearance      =   0  'Flat
         Caption         =   "Busca Genérica"
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
         Height          =   225
         Left            =   4110
         TabIndex        =   29
         Top             =   360
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.CheckBox chk_pendencia 
      Appearance      =   0  'Flat
      Caption         =   "Parcela com Pendência "
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
      Height          =   240
      Left            =   13830
      TabIndex        =   27
      Top             =   1050
      Visible         =   0   'False
      Width           =   2175
   End
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
      Height          =   390
      Left            =   9330
      Picture         =   "RelReceber1.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
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
      Height          =   390
      Left            =   11670
      Picture         =   "RelReceber1.frx":4F23C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   465
   End
   Begin VB.CommandButton cmd_pesquisar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar"
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
      Left            =   7905
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   8130
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
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   8130
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Recebimento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7920
      TabIndex        =   16
      Top             =   1380
      Width           =   8130
      Begin SSDataWidgets_B.SSDBCombo Combo_Banco 
         Bindings        =   "RelReceber1.frx":4FB1E
         DataSource      =   "Data2"
         Height          =   315
         Left            =   4500
         TabIndex        =   21
         Top             =   225
         Width           =   750
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
         BackColorOdd    =   8454143
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   6376
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3704
         Columns(1).Caption=   "Conta"
         Columns(1).Name =   "Conta"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Conta"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1720
         Columns(2).Caption=   "Código"
         Columns(2).Name =   "Código"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Código"
         Columns(2).DataType=   2
         Columns(2).FieldLen=   256
         _ExtentX        =   1323
         _ExtentY        =   556
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
         Enabled         =   0   'False
      End
      Begin VB.OptionButton O_Banco1 
         Appearance      =   0  'Flat
         Caption         =   "Banco"
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
         Height          =   225
         Left            =   3690
         TabIndex        =   20
         Top             =   270
         Width           =   855
      End
      Begin VB.OptionButton O_Carnet 
         Appearance      =   0  'Flat
         Caption         =   "Carnet"
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
         Height          =   225
         Left            =   2520
         TabIndex        =   19
         Top             =   270
         Width           =   1065
      End
      Begin VB.OptionButton O_Carteira 
         Appearance      =   0  'Flat
         Caption         =   "Carteira"
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
         Height          =   225
         Left            =   1320
         TabIndex        =   18
         Top             =   270
         Width           =   1065
      End
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
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
         Height          =   225
         Left            =   270
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Nome_Banco 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   5280
         TabIndex        =   22
         Top             =   225
         Width           =   2745
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3060
      TabIndex        =   12
      Top             =   1380
      Width           =   4740
      Begin VB.OptionButton O_Banco 
         Appearance      =   0  'Flat
         Caption         =   "Para Banco"
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
         Height          =   255
         Left            =   1770
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton O_Completo 
         Appearance      =   0  'Flat
         Caption         =   "Completo"
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
         Height          =   255
         Left            =   3420
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton O_Resumido 
         Appearance      =   0  'Flat
         Caption         =   "Resumido"
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
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Width           =   7740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   60
      TabIndex        =   9
      Top             =   1380
      Width           =   2955
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
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
         Height          =   255
         Left            =   1590
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
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
         Height          =   255
         Left            =   270
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
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
      Left            =   4710
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   8130
      Visible         =   0   'False
      Width           =   1725
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   4260
      Top             =   8130
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
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelReceber1.frx":4FB32
      DataSource      =   "Data1"
      Height          =   315
      Left            =   510
      TabIndex        =   1
      Top             =   1005
      Width           =   885
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
      Columns(0).Width=   8520
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1455
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1561
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox Data_Fim 
      Height          =   285
      Left            =   10530
      TabIndex        =   7
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   1020
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Height          =   285
      Left            =   8175
      TabIndex        =   4
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   1020
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin MSFlexGridLib.MSFlexGrid gridLancamentos 
      Height          =   5310
      Left            =   60
      TabIndex        =   25
      Top             =   2700
      Width           =   15990
      _ExtentX        =   28205
      _ExtentY        =   9366
      _Version        =   393216
      Rows            =   1
      Cols            =   13
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454143
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16250871
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBCombo cboCliente 
      Bindings        =   "RelReceber1.frx":4FB46
      DataSource      =   "Data4"
      Height          =   330
      Left            =   690
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   1755
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      MaxDropDownItems=   16
      BevelType       =   0
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
      BorderStyle     =   0
      BevelColorFace  =   15066597
      CheckBox3D      =   0   'False
      ForeColorEven   =   0
      BackColorEven   =   15066597
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   9075
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1746
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   847
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4339
      Columns(3).Caption=   "Cidade"
      Columns(3).Name =   "Cidade"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "Cidade"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1191
      Columns(4).Caption=   "Estado"
      Columns(4).Name =   "Estado"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "Estado"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   3096
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   0
      BackColor       =   12648447
   End
   Begin VB.Label lbl_cliente 
      Appearance      =   0  'Flat
      Caption         =   "Cliente"
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
      Height          =   225
      Left            =   60
      TabIndex        =   34
      Top             =   1620
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2490
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label lbl_totalRegistros 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Total de registros: 0"
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
      Left            =   90
      TabIndex        =   26
      Top             =   8070
      Width           =   2385
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Até"
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
      Height          =   255
      Left            =   10170
      TabIndex        =   6
      Top             =   1035
      Width           =   345
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "De"
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
      Height          =   255
      Left            =   7890
      TabIndex        =   3
      Top             =   1035
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
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
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   1050
      Width           =   375
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1005
      Width           =   4650
   End
End
Attribute VB_Name = "frmRelReceber1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros            As Recordset
Dim rsContas                As Recordset
Dim arrayContas()           As Variant
Dim contador_arrayContas    As Long
Dim rsCliFor                As Recordset

Private Function AchaContato(pCliente As Long) As String
  Dim J As Long
  AchaContato = ""
  For J = 0 To contador_arrayContas - 1
      If arrayContas(J, 0) = pCliente Then
          AchaContato = AchaContato & ", " & arrayContas(J, 1)
      End If
  Next
End Function

Private Sub B_Imprime_Click()
On Error GoTo Erro

    Dim Val1, Val2, Erro As Integer
    Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
    Dim Str_Rel As String
    Dim Data1 As Variant
    
    Call StatusMsg("")
    
    Rem Verifica empresa
    If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
      DisplayMsg "Escolha a filial."
      Combo.SetFocus
      Exit Sub
    End If

    If Filial_Liberada <> 0 Then
      If Val(Combo.Text) <> Filial_Liberada Then
        DisplayMsg "Funcionário não tem acesso a esta filial."
        Exit Sub
      End If
    End If
    
    
    Rem Verifica Data
    Erro = False
    If IsNull(Data_Ini.Text) Then Erro = True
    If Not Erro Then If Not IsDate(Data_Ini.Text) Then Erro = True
    If Erro = True Then
      DisplayMsg "Data incorreta, verifique."
      Data_Ini.SetFocus
      Exit Sub
    End If
 
    Rem Verifica Data Final
    Erro = False
    If IsNull(Data_Fim.Text) Then Erro = True
    If Not Erro Then If Not IsDate(Data_Fim.Text) Then Erro = True
    If Erro = True Then
      DisplayMsg "Data incorreta, verifique."
      Data_Fim.SetFocus
      Exit Sub
    End If
    
    
    If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
      DisplayMsg "Data inicial deve ser menor ou igual a data final."
      Data_Ini.SetFocus
      Exit Sub
    End If
    
    Rem  Nome do BD
    Str1 = gsQuickDBFileName
    Rel1.DataFiles(0) = Str1

    Rem Saída
    If B_Vídeo = True Then Rel1.Destination = 0
    If B_Impressora = True Then Rel1.Destination = 1
    
    Rem Nome do arquivo .rpt
    If O_Resumido.Value = True Then Str1 = gsReportPath & "RECEBE1R.RPT"
    If O_Completo.Value = True Then Str1 = gsReportPath & "RECEBE1C.RPT"
    If O_Banco.Value = True Then Str1 = gsReportPath & "RECEBE1B.RPT"
    
    Rel1.ReportFileName = Str1
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel1
    
    Rem Seleção
    Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
    Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
    
    Str_Rel = "{Contas a Receber.Filial} =" + Combo.Text
    Str_Rel = Str_Rel + " And {Contas a Receber.Vencimento} >="
    Str_Rel = Str_Rel + Str_Data1
    Str_Rel = Str_Rel + " And {Contas a Receber.Vencimento} <=" + Str_Data2
    Str_Rel = Str_Rel + " And {Contas a Receber.Valor Recebido} = 0"
    Str_Rel = Str_Rel + " And {Contas a Receber.Tipo} = 'R'"
 
    If O_Carteira.Value = True Then
      Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'C'"
    End If
    
    If O_Carnet.Value = True Then
      Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'T'"
    End If
    
    If O_Banco1.Value = True Then
      Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'B'"
      If Nome_Banco.Caption <> "" Then
        Str_Rel = Str_Rel + " And {Contas a Receber.Conta Boleto} = " + str(Combo_Banco.Text)
      End If
    End If

    Rel1.SelectionFormula = Str_Rel
    
    Str_Rel = "nome_empresa = '"
    Str_Rel = Str_Rel + gsNomeEmpresa + "'"

    Rel1.Formulas(0) = Str_Rel
    
    Str_Rel = "nome_filial = '"
    Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
    Rel1.Formulas(1) = Str_Rel
    
    Rem data inicial
    Str_Rel = "data_ini = '"
    Str_Rel = Str_Rel + Data_Ini.Text + "'"
    Rel1.Formulas(2) = Str_Rel
    
    Rem data final
    Str_Rel = "data_fim = '"
    Str_Rel = Str_Rel + Data_Fim.Text + "'"
    Rel1.Formulas(3) = Str_Rel
    
    Call StatusMsg("Aguarde, imprimindo...")
    MousePointer = vbHourglass
    
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel1)
    
    Rel1.Action = 1
    
    Call StatusMsg("")
    MousePointer = vbDefault

    Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cboCliente_Click()
  cboCliente.Text = cboCliente.Columns(1).Text
End Sub

Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(1).Text
  cboCliente_LostFocus
End Sub

Private Sub cboCliente_LostFocus()
  Dim Aux As Variant
  
  Call StatusMsg("")
  
  Nome_Cliente.Caption = ""
  
  Aux = cboCliente.Text
  If IsNull(Aux) Then Exit Sub
  If Aux = "" Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) < 1 Then Exit Sub
  If Val(Aux) > 99999999 Then Exit Sub
  
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Val(Aux)
  If rsCliFor.NoMatch Then
      DisplayMsg "Cliente incorreto."
      cboCliente.SetFocus
      Exit Sub
  End If

  Nome_Cliente.Caption = rsCliFor("Nome") & ""
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
  
  If lngCasasDecimais = 2 Then
      If Len(FormataValorTexto) = 7 Then  ' 9999.99     = 9.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 6)
      ElseIf Len(FormataValorTexto) = 8 Then ' 99999.99    = 99.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 6)
      ElseIf Len(FormataValorTexto) = 9 Then ' 999999.99   = 999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 6)
      ElseIf Len(FormataValorTexto) = 10 Then ' 9999999.99   = 9.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 3) + "." + Mid(FormataValorTexto, 5, 6)
      ElseIf Len(FormataValorTexto) = 11 Then ' 99999999.99   = 99.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 3) + "." + Mid(FormataValorTexto, 6, 6)
      ElseIf Len(FormataValorTexto) = 12 Then ' 999999999.99   = 999.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 3) + "." + Mid(FormataValorTexto, 7, 6)
      End If
  End If
  
End Function

Private Sub cmd_imprimirOperDet_Click()
  On Error GoTo Erro
  
  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome As String
  Dim strNomeLPT As String
  Dim strPortaLPT As String
  Dim intX As Integer
  Dim i As Integer
  
  strNome = "REL"
  strNomeLPT = "NOME IMPRESSORA REL"
  strPortaLPT = "PORTA IMPRESSORA REL"

  strImpressora = GetSetting("QuickStore", "ConfigLPT", strNomeLPT, "")
  strPorta = GetSetting("QuickStore", "ConfigLPT", strPortaLPT, "")
      
  If Len(Trim(strImpressora)) > 0 And Len(Trim(strPorta)) > 0 Then
      For Each objPrinter In Printers
        If objPrinter.DeviceName = strImpressora And objPrinter.Port = strPorta Then
            Set Printer = objPrinter
            Exit For
        End If
      Next objPrinter
  End If

  Dim nRow As Long
  Dim sLinha As String
  
  Printer.Font = "LUCIDA CONSOLE"
  
  Printer.Print ""
  sLinha = "                                 Quick Store 10 - Soluções Comerciais inteligentes"
  
  Printer.Print ""

  sLinha = "                           Contas a Receber de parcelas com data de vencimento no período "
  Printer.Print sLinha

  Printer.Print ""

  sLinha = "   Filial             : " & Combo & " - " & Nome_Empresa
  Printer.Print sLinha
  sLinha = "   Data Vencimento de : " & Data_Ini.Text
  Printer.Print sLinha
  sLinha = "                  até : " & Data_Fim.Text
  Printer.Print sLinha

  If Nome_Cliente.Caption <> "" Then
      sLinha = "   Cliente            : " & cboCliente.Text & " - " & Nome_Cliente.Caption
      Printer.Print sLinha
  End If

  Printer.Print ""

  sLinha = "   Código   Nome                                                          Crediário  Descrição"
  Printer.Print sLinha
  sLinha = "   Vencimento      R$ Valor      Parcela          Nota          Cartão           Fatura       ID"
  Printer.Print sLinha
  sLinha = "   Contatos"
  Printer.Print sLinha

  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""

  Dim sCodigo         As String
  Dim sNome           As String
  Dim sCrediario      As String
  Dim sDescricao      As String
  Dim sVencimento     As String
  Dim sValor          As String
  Dim sParcela        As String
  Dim sNota           As String
  Dim sCartao         As String
  Dim sFatura         As String
  Dim sID             As String
  Dim sContados       As String

  With gridLancamentos
      For nRow = 1 To .Rows - 1
          ' ************************** ATENÇÃO ***********************************
          ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
          ' De preferência com o mesmo nome da impressora !!!

          sCodigo = gridLancamentos.TextMatrix(nRow, 1)
          If Len(sCodigo) < 10 Then
            For i = Len(sCodigo) To 9
                sCodigo = sCodigo & " "
            Next
          End If

          sNome = gridLancamentos.TextMatrix(nRow, 2)
          If Len(sNome) < 50 Then
              For i = Len(sNome) To 49
                sNome = sNome & " "
              Next
          Else
              sNome = Mid(sNome, 1, 50)
          End If

          sCrediario = gridLancamentos.TextMatrix(nRow, 3)
          If Len(sCrediario) < 16 Then
            For i = Len(sCrediario) To 15
                sCrediario = " " & sCrediario
            Next
          End If

          sDescricao = gridLancamentos.TextMatrix(nRow, 5)

          sLinha = sCodigo
          sLinha = sLinha & "  " & sNome
          sLinha = sLinha & "  " & sCrediario
          sLinha = sLinha & "  " & sDescricao
          Printer.Print "   " & sLinha


          sVencimento = gridLancamentos.TextMatrix(nRow, 4)

          sValor = gridLancamentos.TextMatrix(nRow, 6)
          If Len(sValor) < 13 Then
            For i = Len(sValor) To 12
                sValor = " " & sValor
            Next
          End If
          
          sParcela = gridLancamentos.TextMatrix(nRow, 7)
          If Len(sParcela) < 13 Then
            For i = Len(sParcela) To 12
                sParcela = " " & sParcela
            Next
          End If
          
          sNota = gridLancamentos.TextMatrix(nRow, 8)
          If Len(sNota) < 13 Then
            For i = Len(sNota) To 12
                sNota = " " & sNota
            Next
          End If
          
          sCartao = gridLancamentos.TextMatrix(nRow, 9)
          If Len(sCartao) < 13 Then
            For i = Len(sCartao) To 12
                sCartao = " " & sCartao
            Next
          End If
          
          sFatura = gridLancamentos.TextMatrix(nRow, 10)
          If Len(sFatura) < 13 Then
            For i = Len(sFatura) To 12
                sFatura = " " & sFatura
            Next
          End If
          
          sID = gridLancamentos.TextMatrix(nRow, 11)
          If Len(sID) < 10 Then
            For i = Len(sID) To 9
                sID = " " & sID
            Next
          End If
  
          sLinha = sVencimento
          sLinha = sLinha & "  " & sValor
          sLinha = sLinha & "  " & sParcela
          sLinha = sLinha & "  " & sNota
          sLinha = sLinha & "  " & sCartao
          sLinha = sLinha & "  " & sFatura
          sLinha = sLinha & "  " & sID
              
          Printer.Print "   " & sLinha
          
          sContados = gridLancamentos.TextMatrix(nRow, 12)
          sLinha = sContados
          Printer.Print "   FONE: " & sLinha

          Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      Next nRow
  End With
      
  Printer.Print ""
    
  Printer.Print "   " & lbl_totalRegistros.Caption
  Printer.Print "   -----------------------------------------------------------------------------------------------------------------"

  Printer.EndDoc

  Exit Sub
Erro:
    MsgBox "Erro na impressão da grade " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_pesquisar_Click()
On Error GoTo Erro
  Dim sSql                As String
  Dim rsContasReceber     As Recordset
  Dim rsContatos          As Recordset
  Dim lngContadorRegGrid  As Long
  Dim sFatura             As String
  Dim sParcela            As String
  Dim sCartao             As String
  Dim sNota               As String
  Dim sDescricao          As String
  Dim sSituacaoCrediario  As String
  Dim iCol                As Integer
  Dim sCodigoAnterior     As String
  Dim sCorAnterior        As String
  Dim lContador           As Long
  Dim sFones              As String
  
  gridLancamentos.Rows = 1
  gridLancamentos.Row = 0
  
  ' Verifica empresa
  If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
    DisplayMsg "Escolha a filial."
    Combo.SetFocus
    Exit Sub
  End If

  If Filial_Liberada <> 0 Then
    If Val(Combo.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If
  
  If IsNull(Data_Ini.Text) Then
    DisplayMsg "Data incorreta, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Data incorreta, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If

  If IsNull(Data_Fim.Text) Then
    DisplayMsg "Data incorreta, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
    
  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Data incorreta, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
    
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual a data final."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Trim(Nome_Cliente.Caption) = "" Then
    If (CDate(Data_Fim.Text) - CDate(Data_Ini.Text)) > 90 Then
        MsgBox "Sem selecionar um Cliente, o período máximo para pesquisa é de até 90 dias.", vbInformation, "Atenção"
        Exit Sub
    End If
  End If
  
  ' ==========================================================================
  ' Obter todos os contatos da base
  If Trim(Nome_Cliente.Caption) <> "" Then
      Set rsContatos = db.OpenRecordset("Select Cliente, Ramal from Contatos where Cliente= " & cboCliente.Text, dbOpenDynaset, dbReadOnly)
  Else
      Set rsContatos = db.OpenRecordset("Select Cliente, Ramal from Contatos order by Cliente", dbOpenDynaset, dbReadOnly)
  End If
  
  lContador = 0
  If Not (rsContatos.EOF And rsContatos.BOF) Then
      rsContatos.MoveLast
      rsContatos.MoveFirst
      
      ReDim arrayContas(rsContatos.RecordCount, 2)
      contador_arrayContas = rsContatos.RecordCount
      While Not rsContatos.EOF
          arrayContas(lContador, 0) = rsContatos.Fields("Cliente").Value
          arrayContas(lContador, 1) = rsContatos.Fields("Ramal").Value
          lContador = lContador + 1
          rsContatos.MoveNext
      Wend
  End If
  rsContatos.Close
  Set rsContatos = Nothing
  ' ==========================================================================
  
  sSql = "Select R.Cliente, C.Nome, R.Valor, R.Fatura, R.Parcela, C.Faturado, "
  sSql = sSql & " R.Descrição, R.Nota, R.Cartão, R.Vencimento, C.[Fone 1] as Fone1, "
  sSql = sSql & " C.[Fone 2] as Fone2, C.Fax, R.Contador "
  sSql = sSql & " From [Contas a Receber] R, Cli_For C "
  sSql = sSql & " Where R.Filial = " & Combo.Text & " And "
  sSql = sSql & " R.Vencimento >= CDATE('" & Data_Ini & " 00:00:00') and "
  sSql = sSql & " R.Vencimento <= CDATE('" & Data_Fim & " 23:59:59') and "
  
  If Trim(Nome_Cliente.Caption) <> "" Then
      sSql = sSql & " R.Cliente = " & cboCliente.Text & " and "
  End If

  sSql = sSql & " R.[Valor Recebido] = 0 And "
  
  If chk_pendencia.Value = 1 Then
      sSql = sSql & " R.Pendencia = -1 And "
  End If
  
  sSql = sSql & " R.Cliente = C.Código "
  sSql = sSql & " Order by R.Vencimento, C.Nome, R.Parcela "
  
  Screen.MousePointer = vbHourglass
  
  Set rsContasReceber = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
  
  lngContadorRegGrid = 0
  
  If Not (rsContasReceber.EOF And rsContasReceber.BOF) Then
    rsContasReceber.MoveFirst
  End If
  
  sCorAnterior = vbWhite
  
  While Not rsContasReceber.EOF
  
      sFatura = ""
      sParcela = ""
      sCartao = ""
      sNota = ""
      sDescricao = ""
      sFones = ""
      
      If Not IsNull(rsContasReceber.Fields("Fatura").Value) Then
          sFatura = rsContasReceber.Fields("Fatura").Value
      End If
  
      If Not IsNull(rsContasReceber.Fields("Parcela").Value) Then
          sParcela = rsContasReceber.Fields("Parcela").Value
      End If
  
      If Not IsNull(rsContasReceber.Fields("Cartão").Value) Then
          sCartao = rsContasReceber.Fields("Cartão").Value
      End If
  
      If Not IsNull(rsContasReceber.Fields("Descrição").Value) Then
          sDescricao = rsContasReceber.Fields("Descrição").Value
      End If
  
      If Not IsNull(rsContasReceber.Fields("Nota").Value) Then
          sNota = rsContasReceber.Fields("Nota").Value
      End If
  
      '=========================================
      ' Fones
      If Not IsNull(rsContasReceber.Fields("Fone1").Value) And Trim(rsContasReceber.Fields("Fone1").Value) <> "" Then
          sFones = rsContasReceber.Fields("Fone1").Value
      End If
      
      If Not IsNull(rsContasReceber.Fields("Fone2").Value) And Trim(rsContasReceber.Fields("Fone2").Value) <> "" Then
          sFones = sFones & ", " & rsContasReceber.Fields("Fone2").Value
      End If
      
      If Not IsNull(rsContasReceber.Fields("Fax").Value) And Trim(rsContasReceber.Fields("Fax").Value) <> "" Then
          sFones = sFones & ", " & rsContasReceber.Fields("Fax").Value
      End If
      
      sFones = sFones & AchaContato(rsContasReceber.Fields("Cliente").Value)
  
      sSituacaoCrediario = ""
      If rsContasReceber.Fields("Faturado").Value = True Then
          sSituacaoCrediario = "Habilitado"
      Else
          sSituacaoCrediario = "Suspenso"
      End If
      
      gridLancamentos.AddItem vbTab & rsContasReceber.Fields("Cliente").Value & vbTab & _
                          rsContasReceber.Fields("Nome").Value & vbTab & _
                          sSituacaoCrediario & vbTab & _
                          rsContasReceber.Fields("Vencimento").Value & vbTab & _
                          sDescricao & vbTab & _
                          FormataValorTexto(rsContasReceber.Fields("Valor").Value, 2) & vbTab & _
                          sParcela & vbTab & _
                          sNota & vbTab & _
                          sCartao & vbTab & _
                          sFatura & vbTab & _
                          rsContasReceber.Fields("Contador").Value & vbTab & _
                          sFones

      lngContadorRegGrid = lngContadorRegGrid + 1

      ' tratar cores da grid
      If sCodigoAnterior = rsContasReceber.Fields("Cliente").Value Then
          For iCol = 0 To gridLancamentos.Cols - 1
              gridLancamentos.Col = iCol
              gridLancamentos.Row = lngContadorRegGrid
              'alterna a cor das linhas do grid
              gridLancamentos.CellBackColor = sCorAnterior
          Next
      Else
          If sCorAnterior = vbWhite Then
              sCorAnterior = vbButtonFace
          Else
              sCorAnterior = vbWhite
          End If
      
          For iCol = 0 To gridLancamentos.Cols - 1
              gridLancamentos.Col = iCol
              gridLancamentos.Row = lngContadorRegGrid
              'alterna a cor das linhas do grid
              gridLancamentos.CellBackColor = sCorAnterior
          Next
      End If
      sCodigoAnterior = rsContasReceber.Fields("Cliente").Value
      ' fim cores

      rsContasReceber.MoveNext
  Wend
  rsContasReceber.Close
  Set rsContasReceber = Nothing
  
  lbl_totalRegistros.Caption = "Total de registros: " & lngContadorRegGrid
  
  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If
  
  MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Combo_Banco_CloseUp()
  Combo_Banco.Text = Combo_Banco.Columns(2).Text
  Combo_Banco_LostFocus
End Sub

Private Sub Combo_Banco_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Banco_LostFocus()
  Call StatusMsg("")
  Nome_Banco.Caption = ""
  
  If IsNull(Combo_Banco.Text) Then Exit Sub
  If Combo_Banco.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Banco.Text) Then Exit Sub
  If Val(Combo_Banco.Text) > 9999 Then Exit Sub
  If Val(Combo_Banco.Text) < 1 Then Exit Sub
 
  rsContas.Index = "Código"
  rsContas.Seek "=", Val(Combo_Banco.Text)
  If rsContas.NoMatch Then Exit Sub
  
  Nome_Banco.Caption = rsContas("Descrição") & ""
End Sub

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(1).Text
  Combo_LostFocus
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
    
    Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)

    Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
    Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
    
    Data1.DatabaseName = gsQuickDBFileName
    Data2.DatabaseName = gsQuickDBFileName
    Data4.DatabaseName = gsQuickDBFileName

    Combo.Text = gnCodFilial
    
    gridLancamentos.ColWidth(0) = 0
    gridLancamentos.ColWidth(1) = 900
    gridLancamentos.ColWidth(2) = 3800
    gridLancamentos.ColWidth(3) = 950
    gridLancamentos.ColWidth(4) = 950
    gridLancamentos.ColWidth(5) = 2100
    gridLancamentos.ColWidth(6) = 1000
    gridLancamentos.ColWidth(7) = 600
    gridLancamentos.ColWidth(8) = 800
    gridLancamentos.ColWidth(9) = 1200
    gridLancamentos.ColWidth(10) = 1200
    gridLancamentos.ColWidth(11) = 900
    gridLancamentos.ColWidth(12) = 4500
  
    gridLancamentos.Row = 0
    gridLancamentos.TextMatrix(0, 1) = "Código"
    gridLancamentos.TextMatrix(0, 2) = "Nome"
    gridLancamentos.TextMatrix(0, 3) = "Crediário"
    gridLancamentos.TextMatrix(0, 4) = "Vencimento"
    gridLancamentos.TextMatrix(0, 5) = "Descrição"
    gridLancamentos.TextMatrix(0, 6) = "R$ Valor"
    gridLancamentos.TextMatrix(0, 7) = "Parcela"
    gridLancamentos.TextMatrix(0, 8) = "Nota"
    gridLancamentos.TextMatrix(0, 9) = "Cartão"
    gridLancamentos.TextMatrix(0, 10) = "Fatura"
    gridLancamentos.TextMatrix(0, 11) = "ID"
    gridLancamentos.TextMatrix(0, 12) = "Contatos"
    'gridLancamentos.ColAlignment(11) = vbleft
    
    B_Imprime.Width = 15990
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsCliFor.Close
    Set rsCliFor = Nothing
End Sub

Private Sub O_Banco1_Click()
    Combo_Banco.Enabled = True
    Nome_Banco.Enabled = True
End Sub

Private Sub O_Carnet_Click()
    Combo_Banco.Enabled = False
    Nome_Banco.Enabled = False
End Sub

Private Sub O_Carteira_Click()
    Combo_Banco.Enabled = False
    Nome_Banco.Enabled = False
End Sub

Private Sub O_Todos_Click()
    Combo_Banco.Enabled = False
    Nome_Banco.Enabled = False
End Sub


Private Sub opt_especifica_Click()
  If opt_especifica.Value = True Then
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    chk_pendencia.Visible = True
    cmd_pesquisar.Visible = True
    gridLancamentos.Rows = 1
    gridLancamentos.Enabled = True
    lbl_totalRegistros.Caption = "Total de registros: 0"
    cmd_pesquisar.Left = 60
    cmd_pesquisar.Width = 15990
    B_Imprime.Visible = False
    chk_pendencia.Value = 0
    
    lbl_cliente.Visible = True
    cboCliente.Visible = True
    Nome_Cliente.Visible = True
  Else
  End If

End Sub

Private Sub opt_generico_Click()
  If opt_generico.Value = True Then
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = True
    chk_pendencia.Visible = False
    cmd_pesquisar.Visible = False
    gridLancamentos.Rows = 1
    gridLancamentos.Enabled = False
    lbl_totalRegistros.Caption = "Total de registros: 0"
    B_Imprime.Width = 15990
    B_Imprime.Visible = True
    cmd_pesquisar.Visible = False
    chk_pendencia.Value = 0
    
    lbl_cliente.Visible = False
    cboCliente.Visible = False
    Nome_Cliente.Visible = False
    
  Else
  End If
End Sub
