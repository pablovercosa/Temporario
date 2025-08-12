VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmLancaCReceber 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Lançamentos/Manutenção de Contas a Receber"
   ClientHeight    =   5535
   ClientLeft      =   2130
   ClientTop       =   1155
   ClientWidth     =   14115
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
   HelpContextID   =   1350
   Icon            =   "LancaCReceber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   14115
   Begin VB.CheckBox chk_pendencia 
      Appearance      =   0  'Flat
      Caption         =   "Pendência"
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
      Height          =   345
      Left            =   12660
      TabIndex        =   46
      Top             =   4290
      Width           =   1335
   End
   Begin VB.TextBox txt_ID 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12555
      MaxLength       =   9
      TabIndex        =   4
      Top             =   75
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações sobre boleto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   6420
      TabIndex        =   30
      Top             =   2790
      Width           =   3510
      Begin VB.TextBox txtNossoNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   32
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Nosso Número"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1230
      End
   End
   Begin SSDataWidgets_B.SSDBCombo cboTipo 
      Height          =   360
      Left            =   1755
      TabIndex        =   12
      Top             =   1425
      Width           =   1845
      DataFieldList   =   "Column 0"
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
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).Name =   "Column0"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   3254
      _ExtentY        =   635
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
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<   &Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4245
      Width           =   2010
   End
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
      Left            =   5370
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   7275
      Visible         =   0   'False
      Width           =   1905
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
      Bindings        =   "LancaCReceber.frx":4E95A
      DataSource      =   "Data4"
      Height          =   360
      Left            =   6990
      TabIndex        =   26
      Top             =   2325
      Width           =   1125
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
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   6879
      Columns(0).Caption=   "Descrição"
      Columns(0).Name =   "Descrição"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Descrição"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3016
      Columns(1).Caption=   "Conta"
      Columns(1).Name =   "Conta"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Conta"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1508
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   256
      _ExtentX        =   1984
      _ExtentY        =   635
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
   Begin VB.CheckBox Carnet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Carnet Impresso"
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
      Left            =   10155
      TabIndex        =   34
      Top             =   3360
      Width           =   1740
   End
   Begin VB.CheckBox Boleto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Boleto Impresso"
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
      Left            =   10155
      TabIndex        =   33
      Top             =   2940
      Width           =   1695
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
      Left            =   2055
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Vendedor"
      Top             =   6855
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Descrição 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10155
      MaxLength       =   30
      TabIndex        =   20
      Top             =   1890
      Width           =   3855
   End
   Begin VB.TextBox Fatura 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6990
      MaxLength       =   10
      TabIndex        =   18
      Top             =   1890
      Width           =   2175
   End
   Begin VB.TextBox Nota 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      MaxLength       =   9
      TabIndex        =   14
      Top             =   1890
      Width           =   1845
   End
   Begin VB.TextBox Sequência 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4815
      MaxLength       =   9
      TabIndex        =   16
      Top             =   1890
      Width           =   1335
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
      Left            =   4020
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   6855
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   195
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   6870
      Visible         =   0   'False
      Width           =   1725
   End
   Begin MSMask.MaskEdBox Valor_Pago 
      Height          =   360
      Left            =   1755
      TabIndex        =   40
      Top             =   4260
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Acréscimo 
      Height          =   360
      Left            =   1755
      TabIndex        =   38
      Top             =   3675
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Desconto 
      Height          =   360
      Left            =   1755
      TabIndex        =   36
      Top             =   3240
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   360
      Left            =   1755
      TabIndex        =   29
      Top             =   2790
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "LancaCReceber.frx":4E96E
      DataSource      =   "Data2"
      Height          =   360
      Left            =   1755
      TabIndex        =   9
      Top             =   975
      Width           =   1845
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
      Columns(0).Width=   8017
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1773
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   3254
      _ExtentY        =   635
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "LancaCReceber.frx":4E982
      DataSource      =   "Data3"
      Height          =   360
      Left            =   1755
      TabIndex        =   6
      Top             =   525
      Width           =   1845
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
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1958
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   3254
      _ExtentY        =   635
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "LancaCReceber.frx":4E996
      DataSource      =   "Data1"
      Height          =   360
      Left            =   1755
      TabIndex        =   1
      Top             =   75
      Width           =   1845
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
      Columns(0).Width=   6826
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   3254
      _ExtentY        =   635
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
   Begin MSMask.MaskEdBox Data_Pagto 
      Height          =   360
      Left            =   1755
      TabIndex        =   43
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   4680
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   360
      Left            =   4815
      TabIndex        =   24
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   2325
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      ForeColor       =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
   Begin MSMask.MaskEdBox Emissão 
      Height          =   360
      Left            =   1755
      TabIndex        =   22
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   2325
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   12240
      TabIndex        =   3
      Top             =   135
      Width           =   345
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "* ATENÇÃO: Caso efetue o recebimento nesta tela...o CAIXA NÃO será sensibilizado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3870
      TabIndex        =   44
      Top             =   4770
      Width           =   7200
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1755
      X2              =   3600
      Y1              =   4125
      Y2              =   4110
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   270
      Top             =   5340
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
      Bands           =   "LancaCReceber.frx":4E9AA
   End
   Begin VB.Label Nome_Conta 
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
      Height          =   360
      Left            =   8190
      TabIndex        =   27
      Top             =   2325
      Width           =   5805
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Conta Boleto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6405
      TabIndex        =   25
      Top             =   2400
      Width           =   585
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Recebimento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   1500
      Width           =   1485
   End
   Begin VB.Label Nome_Vendedor 
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
      Height          =   360
      Left            =   3615
      TabIndex        =   10
      Top             =   975
      Width           =   6240
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   1035
      Width           =   975
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Data Recebimento"
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
      Left            =   150
      TabIndex        =   42
      Top             =   4740
      Width           =   1620
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Recebido (=)"
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
      Left            =   150
      TabIndex        =   39
      Top             =   4320
      Width           =   1710
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Acréscimo  (+)"
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
      Left            =   150
      TabIndex        =   37
      Top             =   3735
      Width           =   1215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto   (-)"
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
      Left            =   150
      TabIndex        =   35
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Left            =   150
      TabIndex        =   28
      Top             =   2850
      Width           =   975
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Seqüência"
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
      Left            =   3765
      TabIndex        =   15
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
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
      Left            =   3765
      TabIndex        =   23
      Top             =   2385
      Width           =   975
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão"
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
      Left            =   150
      TabIndex        =   21
      Top             =   2385
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
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
      Left            =   9330
      TabIndex        =   19
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Fatura :"
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
      Left            =   6405
      TabIndex        =   17
      Top             =   1950
      Width           =   600
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Nota"
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
      Left            =   150
      TabIndex        =   13
      Top             =   1950
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   585
      Width           =   735
   End
   Begin VB.Label Nome_Cliente 
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
      Height          =   360
      Left            =   3615
      TabIndex        =   7
      Top             =   525
      Width           =   6240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   135
      Width           =   855
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
      Height          =   360
      Left            =   3615
      TabIndex        =   2
      Top             =   75
      Width           =   6240
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFA324&
      Height          =   915
      Left            =   90
      TabIndex        =   45
      Top             =   4200
      Width           =   3615
   End
End
Attribute VB_Name = "frmLancaCReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sValorOriginal As String
Dim sDtVencOriginal As String

Dim Num_Registro As Variant
Dim rsParametros As Recordset
Dim rsClientes As Recordset
Dim rsCR As Recordset
Dim rsFuncionarios As Recordset
Dim rsContas_Correntes As Recordset
Dim Conta As Long

Private gsSql As String
Private gsWhere As String
Private gsOrder As String

Private Sub DeleteRecord()
  Dim Resposta As Integer
  
  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Não existe registro para apagar !"
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente apagar esta conta?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
  
    '11/07/2007 - Anderson
    'Criação de log para registro de exclusão de registro
    'Efetua registro do Log
  
    'LOG *****************
    db.Execute "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '" & Left("Usu:" & _
      gnUserCode & " Cli:" & rsCR("Cliente") & " Seq:" & rsCR("Sequência") & " NF:" & rsCR("Nota") & " Venc:" & rsCR("Vencimento") & " Vlr:" & rsCR("Valor"), 80) & "', 'DEL CONTA RECEBER')", dbFailOnError
    'fim *******************

    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
      "Cli:" & rsCR("Cliente") & "- Seq:" & rsCR("Sequência") & "- NF:" & rsCR("Nota") & "- Venc:" & rsCR("Vencimento") & "- Valor:" & rsCR("Valor"), _
      "frmLancaCReceber_DeleteRecord", _
      "Contas a Receber", g_strArquivoSystemLog
    End If
    rsCR.Delete
    rsCR.MovePrevious
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  Dim Valor_Correto As Double
  Dim Contador As Long
  Dim sTexto As String
  
  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked As Integer
  
  Dim rstContasReceber As Recordset '05/06/2007 - Anderson
  Dim bolErroNossoNumero As Boolean '05/06/2007 - Anderson
  
  Call StatusMsg("")
  
  On Error GoTo Trata_Erro:
  
  Rem Verifica Empresa
  If Nome_Empresa.Caption = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Filial inválida, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Combo_Empresa.SetFocus
    Exit Sub
  End If
  
  If IsNull(Sequência.Text) Then Sequência.Text = 0
  If Not IsNumeric(Sequência.Text) Then Sequência.Text = 0
  If Val(Sequência.Text) < 0 Then Sequência.Text = 0
  
  If Nome_Cliente.Caption = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Cliente inválido, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Combo_Cliente.SetFocus
    Exit Sub
  End If
  
  If IsNull(Sequência.Text) Then Sequência.Text = 0
  If Not IsNumeric(Sequência.Text) Then Sequência.Text = 0
  
  If IsNull(Nota.Text) Then Nota.Text = 0
  If Nota.Text = "" Then Nota.Text = 0
  
  
  If cboTipo.Text <> "CARTEIRA" And cboTipo.Text <> "CARNET" And cboTipo.Text <> "BANCO - BOLETO" Then
    gsTitle = LoadResString(201)
    gsMsg = "Tipo de recebimento inválido. Verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    cboTipo.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Emissão.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = "Data de emissão inválida."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Emissão.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Vencimento.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = "Data de vencimento inválida."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Vencimento.SetFocus
    Exit Sub
  End If
  
  If CDate(Vencimento.Text) < CDate(Emissão.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = "Data de vencimento não pode ser anterior à data de emissão."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Vencimento.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If Not IsNumeric(Valor.Text) Then Valor.Text = 0
  If Erro = False Then If CDbl(Valor.Text) <= 0 Then Erro = True
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Valor incorreto."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Valor.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If Not IsNumeric(Desconto.Text) Then Desconto.Text = 0
  If Erro = False Then If CDbl(Desconto.Text) < 0 Then Erro = True
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Desconto incorreto."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Desconto.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If Not IsNumeric(Acréscimo.Text) Then Acréscimo.Text = 0
  If CDbl(Acréscimo.Text) < 0 Then Erro = True
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Acréscimo incorreto."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Acréscimo.SetFocus
    Exit Sub
  End If
  
  
  If Not IsDate(Data_Pagto.Text) Then
    If Not IsNumeric(Valor_Pago.Text) Then Valor_Pago.Text = 0
    If CDbl(Valor_Pago.Text) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "Data de pagamento incorreta."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Data_Pagto.SetFocus
      Exit Sub
    End If
  End If
  
  If IsDate(Data_Pagto.Text) Then
    If Not IsNumeric(Valor_Pago.Text) Then Valor_Pago.Text = 0
    If CDbl(Valor_Pago.Text) <> 0 Then
      Rem verifica soma
      Valor_Correto = CDbl(Valor.Text) - CDbl(Desconto.Text) + CDbl(Acréscimo.Text)
      If Abs((Valor_Correto - CDbl(Valor_Pago.Text))) > 0.001 Then
        gsTitle = LoadResString(201)
        gsMsg = "Valor pago incorreto, valor correto seria " + str$(Valor_Correto)
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Valor_Pago.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  '05/06/2007 - Anderson
  'Verifica se o Nosso Número já foi emitido em outro boleto para evitar duplicidade.
  'Solicitado pelo cliente Agrotama
  If g_blnInformarNossoNumero Then

    '12/11/2007 - Anderson
    'Tratamento de verificação do Nosso Número repetido para evitar que o sistema posicione o cursor em um registro inexistente
    
    If Trim(txtNossoNumero.Text) <> "" Or Val("0" & txtNossoNumero.Text) <> 0 Then
    
      'Abre registro para evitar duplicidade em nosso número
      Set rstContasReceber = db.OpenRecordset("SELECT CNAB_NossoNumero, Filial, Cliente, Vendedor, Sequência, Nota, [Data Emissão], Vencimento, Valor FROM [Contas a Receber] Where CNAB_NossoNumero='" & Right(String(11, "0") & txtNossoNumero.Text, 11) & "' AND Contador<>" & rsCR.Fields("Contador"))
  
      'Informa que não existe problemas com Nosso Numero
      bolErroNossoNumero = False
  
      'Verifica se existe Nosso número no banco de dados
      If Not rstContasReceber.EOF Then
        MsgBox "Já existe um título com o Nosso Número: " & Right(String(11, "0") & txtNossoNumero.Text, 11) & " informado em outro boleto." & Chr(13) & _
               "Favor verificar o título com os dados abaixo: " & Chr(13) & Chr(13) & _
               "Nosso Número: " & rstContasReceber("CNAB_NossoNumero") & Chr(13) & _
               "Filial: " & rstContasReceber("Filial") & Chr(13) & _
               "Cliente: " & rstContasReceber("Cliente") & Chr(13) & _
               "Vendedor: " & rstContasReceber("Vendedor") & Chr(13) & _
               "Sequência: " & rstContasReceber("Sequência") & Chr(13) & _
               "Nota: " & rstContasReceber("Nota") & Chr(13) & _
               "Data Emissão: " & rstContasReceber("Data Emissão") & Chr(13) & _
               "Vencimento: " & rstContasReceber("Vencimento") & Chr(13) & _
               "Valor: " & rstContasReceber("Valor"), vbOKOnly + vbInformation, "Impressão de Boletos"
  
        'Informa que existe um título com o mesmo Nosso Numero
        bolErroNossoNumero = True
      End If
  
      'Fecha tabela de contas a receber
      rstContasReceber.Close
      Set rstContasReceber = Nothing
  
      'Se houver duplicidade em Nosso Número, o sistema encerra.
      If bolErroNossoNumero Then
        Exit Sub
      End If
      
    End If

  End If
  
  Call StatusMsg("Gravando ...")
  
  With rsCR
  
    ws.BeginTrans
    blnInTransaction = True
    
    If IsNull(Num_Registro) Then
      .AddNew
      sTexto = "Lançamento efetuado."
    Else
      .LockEdits = True
      .Edit
      sTexto = "Lançamento alterado."
    End If
    
    Conta = .Fields("Contador")
    .Fields("Tipo") = "R"
    
    .Fields("Filial") = Combo_Empresa.Text
    .Fields("Cliente") = Combo_Cliente.Text
    .Fields("Vendedor") = Val(Combo_Vendedor.Text)
    .Fields("Sequência") = Sequência.Text
    
    If cboTipo.Text = "CARTEIRA" Then .Fields("Tipo Parcelamento") = "C"
    If cboTipo.Text = "CARNET" Then .Fields("Tipo Parcelamento") = "T"
    If cboTipo.Text = "BANCO - BOLETO" Then .Fields("Tipo Parcelamento") = "B"
    
    .Fields("Conta Boleto") = 0
    If Nome_Conta.Caption <> "" Then .Fields("Conta Boleto") = Val(Combo_Conta.Text)
    
    .Fields("Nota") = Nota.Text
    .Fields("Fatura") = Fatura.Text
    .Fields("Descrição") = Descrição.Text
    .Fields("Data Emissão") = Emissão.Text
    .Fields("Vencimento") = Vencimento.Text
    .Fields("Valor") = CDbl(Valor.Text)
    .Fields("Desconto") = CDbl(Desconto.Text)
    .Fields("Acréscimo") = CDbl(Acréscimo.Text)
    .Fields("Valor Recebido") = CDbl(gsHandleNull(Valor_Pago.Text & ""))
    If Not IsDate(Data_Pagto.Text) Then
      .Fields("Data Recebimento") = Null
    Else
      .Fields("Data Recebimento") = Data_Pagto.Text
    End If
    
    .Fields("Impresso") = Boleto.Value
    .Fields("Carnet Impresso") = Carnet.Value
    .Fields("CNAB_NossoNumero") = txtNossoNumero.Text
    .Fields("Data Alteração") = Format(Date, "dd/mm/yyyy")
    
    '16/05/2007 - Anderson
    'Se Número de série Agrotama, informar nosso número para boletos pré-impressos
    If g_blnInformarNossoNumero Then
      txtNossoNumero.Text = Right(String(11, "0") & txtNossoNumero.Text, 11)
      .Fields("CNAB_NossoNumero") = Right(String(11, "0") & txtNossoNumero.Text, 11)
      .Fields("CNAB_DigitoVerificador") = GetDigitoVerificador_NossoNumero(txtNossoNumero.Text, Bradesco)
      .Fields("CNAB_Carteira") = "9"
    End If
    
    '10/09/2007 - Anderson
    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
      "Cli:" & rsCR("Cliente") & "- Seq:" & rsCR("Sequência") & "- NF:" & rsCR("Nota") & "- Venc:" & rsCR("Vencimento") & "- Valor:" & rsCR("Valor"), _
      "frmLancaCReceber_UpdateRecord", _
      "Contas a Receber", g_strArquivoSystemLog
    End If
    
    If chk_pendencia.Value = 1 Then
        .Fields("Pendencia") = -1
    Else
        .Fields("Pendencia") = 0
    End If
    
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
       
    ws.CommitTrans
    blnInTransaction = False
    
  End With
   
  Call StatusMsg("")
  
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & Sequência.Text & " Cli:" & Combo_Cliente.Text & " VrOr:" & sValorOriginal & " VrAtu:" & Valor.Text & " DtVcOr:" & sDtVencOriginal & " DtVcAtu:" & Data_Pagto.Text, 80) & "', 'CNT_REC: novo-atu')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
  MsgBox "Salvo com sucesso", vbInformation, "Sucesso"
  
  Exit Sub
  
Trata_Erro:
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

Private Sub PrintBoleto()
  Dim Resp As Integer
  Dim Nome_Arq As String
  Dim F As Form
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Ache ou grave uma conta antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  '16/05/2007 - Anderson
  'Se Número de série Agrotama, informar nosso número para boletos pré-impressos
  '16/05/2007 - Anderson
  'Se Número de série Agrotama, informar nosso número para boletos pré-impressos
  If CheckSerialCaseMod("QS73070-894") And txtNossoNumero.Text = "" Then
    MsgBox "Informe o Nosso Número para a impressão do boleto.", vbInformation, "Impressão de Boletos"
    Exit Sub
  End If
  
  If CheckSerialCaseMod("QS73070-894") And txtNossoNumero.Text <> "" Then
    If Not IsNumeric(txtNossoNumero.Text) Then
      MsgBox "A informação digitada no campo Nosso Número não é válida!", vbExclamation, "Impressão de Boletos"
      Exit Sub
    End If
  End If
  
  Set F = New frmObsDoc
  F.Caption = "Impressão de Boletos"
  F.gsFileExt = ".CBB"
  F.Show vbModal
  Set F = Nothing
    
  If gsRetornoDoc <> "OK" Then
'    gsTitle = LoadResString(201)
'    gsMsg = "Impressão cancelada."
'    gnStyle = vbOKOnly + vbExclamation
'    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Nome_Arq = gsConfigPath & gsDocFileName & ".CBB"
  If Dir(Nome_Arq) = "" Then
    DisplayMsg "Arquivo """ & Nome_Arq & """ não encontrado."
    Exit Sub
  End If
    
  Resp = Imprime_Boleto("R", rsCR("Filial"), rsCR("Vencimento"), rsCR("Contador"), Nome_Arq)
  
  If Resp <> 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Houve o erro " + str(Resp) + " na emissão do boleto."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Boleto impresso."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Boleto.Value = 1
    Call UpdateRecord
  End If
 
End Sub

Public Sub ClearScreen()

  Call StatusMsg("")
  Combo_Empresa.Text = ""
  Nome_Empresa.Caption = ""
  Combo_Cliente.Text = ""
  Nome_Cliente.Caption = ""
  Combo_Vendedor.Text = ""
  Nome_Vendedor.Caption = ""
  cboTipo.Text = ""
  Combo_Conta.Text = ""
  Nome_Conta.Caption = ""
  Sequência.Text = ""
  Nota.Text = ""
  Fatura.Text = ""
  Descrição.Text = ""
  Emissão.Mask = ""
  Emissão.Text = ""
  Emissão.Mask = "##/##/####"
  Vencimento.Mask = ""
  Vencimento.Text = ""
  Vencimento.Mask = "##/##/####"
  Data_Pagto.Mask = ""
  Data_Pagto.Text = ""
  Data_Pagto.Mask = "##/##/####"
  Valor.Text = 0
  Desconto.Text = 0
  Acréscimo.Text = 0
  Valor_Pago.Text = 0
  Boleto.Value = 0
  Carnet.Value = 0
  txt_ID.Text = ""
  chk_pendencia.Value = 0
  
  txtNossoNumero.Text = ""
  
  If Not rsCR.EOF Then
    On Error Resume Next
    rsCR.MoveFirst
    rsCR.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  
  Combo_Empresa.SetFocus
  
End Sub

Private Sub Acréscimo_GotFocus()
  Acréscimo.SelStart = 0
  Acréscimo.SelLength = 16
End Sub

Private Sub Acréscimo_LostFocus()
'  Valor_Pago.Text = CCur(gsHandleNull(Valor.Text)) - CCur(gsHandleNull(Desconto.Text)) + CCur(gsHandleNull(Acréscimo.Text))
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
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
    Case "miOpSearch"
      Call SearchRecord
    Case "miOpPrintBoleto"
      Call PrintBoleto
  End Select
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  gsOrder = ""
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case -1, 0 '"Por Filial, Vencimento"
          gsOrder = "ORDER BY Filial, Vencimento, Cliente"
        Case 1 '"Por Filial, Cliente"
          gsOrder = "ORDER BY Filial, Cliente, Vencimento"
        Case 2 '"Por Filial, Data Recebimento"
          gsOrder = "ORDER BY Filial, [Data Recebimento], Cliente"
        Case 3 '"Por Nota, Cliente"
          gsOrder = "ORDER BY Nota, Cliente"
        Case 4 '"Por ID"
          gsOrder = "ORDER BY Contador"
      End Select
  End Select
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsCR
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
  With rsCR
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
  With rsCR
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
  With rsCR
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Private Sub SearchRecord()

  If Not IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Apague todos os campos da tela com o botão NOVO."
    gsMsg = gsMsg & vbCrLf & "Selecione a Ordem de Pesquisa na lista e preencha com dados iniciais os campos respectivos."
    gsMsg = gsMsg & vbCrLf & "Pressione novamente este botão PROCURAR."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  gsWhere = ""
  
  If Len(Trim(Combo_Empresa.Text)) = 0 Then
    Combo_Empresa.Text = "0"
  End If
  
  Select Case ActiveBar1.Tools("miOpOrdem").CBListIndex
    Case -1, 0  '"Por Filial, Vencimento"
      If Not IsDate(Vencimento.Text) Then
        Vencimento.Text = Date - 3
      End If
      gsWhere = "WHERE Tipo = 'R' AND Filial >= " & Combo_Empresa.Text & " AND Vencimento >= #" & Format(Vencimento.Text, "mm/dd/yyyy") & "#"
    Case 1  '"Por Filial, Cliente"
      If Len(Trim(Combo_Cliente.Text)) = 0 Then
        Combo_Cliente.Text = "0"
      End If
      gsWhere = "WHERE Tipo = 'R' AND Filial >= " & Combo_Empresa.Text & " AND Cliente >= " & Combo_Cliente.Text
    Case 2  '"Por Filial, Data Recebimento"
      If Not IsDate(Data_Pagto.Text) Then
        Data_Pagto.Text = Date - 3
      End If
      gsWhere = "WHERE Tipo = 'R' AND Filial >= " & Combo_Empresa.Text & " AND [Data Recebimento] >= #" & Format(Data_Pagto.Text, "mm/dd/yyyy") & "#"
    Case 3  '"Por Nota, Cliente"
      If Len(Trim(Nota.Text)) = 0 Then
        Nota.Text = "0"
      End If
      If Len(Trim(Combo_Cliente.Text)) = 0 Then
        Combo_Cliente.Text = "0"
      End If
      gsWhere = "WHERE Tipo = 'R' AND Nota >= " & Trim(Nota.Text) & " AND Cliente >= " & Combo_Cliente.Text
    Case 4 '"Por ID"
        If Not IsNumeric(txt_ID.Text) Then
            MsgBox "Digite um ID válido.", vbInformation, "Atenção"
            txt_ID.SetFocus
            Exit Sub
        End If
    
        If Len(Trim(txt_ID.Text)) = 0 Then
            txt_ID.Text = "0"
        End If
        gsWhere = " Where CONTADOR = " & Trim(txt_ID.Text)
  End Select
  
  Set rsCR = db.OpenRecordset(gsSql & " " & gsWhere & " " & gsOrder, dbOpenDynaset)
  If Not rsCR.EOF Then
    Call ShowRecord
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em função dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  
End Sub

Private Sub cboTipo_Click()
  cboTipo.Text = cboTipo.Columns(0).Text
End Sub

Private Sub cboTipo_CloseUp()
  cboTipo.Text = cboTipo.Columns(0).Text
End Sub

Private Sub cmdTotal_Click()
  Valor_Pago.Text = Format(CDbl(gsHandleNull(Valor.Text)) - CCur(gsHandleNull(Desconto.Text)) + CCur(gsHandleNull(Acréscimo.Text)), "##,###,##0.00")
End Sub

Private Sub Combo_Cliente_CloseUp()
 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 0 Or Val(Combo_Cliente.Text) > 99999999 Then Exit Sub

  rsClientes.Index = "Código"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then Exit Sub
  Nome_Cliente.Caption = rsClientes("Nome")

End Sub

Private Sub Combo_Conta_CloseUp()
  Combo_Conta.Text = Combo_Conta.Columns(2).Text
  Combo_Conta_LostFocus
End Sub

Private Sub Combo_Conta_LostFocus()

  Call StatusMsg("")
  Nome_Conta.Caption = ""
  If IsNull(Combo_Conta.Text) Then Exit Sub
  If Combo_Conta.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Conta.Text) Then Exit Sub
  If Val(Combo_Conta.Text) < 1 Then Exit Sub
  '28/11/2006 - Anderson
  'Alteração do número de contas bancárias
  'Solicitado por: 2227883 - SANTA FÉ DO ARAGUAIA PREFEITURA MUNICIPAL
  If Val(Combo_Conta.Text) > 255 Then Exit Sub
  
  rsContas_Correntes.Index = "Código"
  rsContas_Correntes.Seek "=", Val(Combo_Conta.Text)
  If rsContas_Correntes.NoMatch Then Exit Sub
  
  Nome_Conta.Caption = rsContas_Correntes("Descrição") & ""

End Sub

Private Sub Combo_Empresa_CloseUp()
  Combo_Empresa.Text = Combo_Empresa.Columns(1).Text
  Combo_Empresa_LostFocus
End Sub

Private Sub Combo_Empresa_LostFocus()
  Nome_Empresa.Caption = ""
  If IsNull(Combo_Empresa.Text) Then Exit Sub
  If Not IsNumeric(Combo_Empresa.Text) Then Exit Sub
  If Val(Combo_Empresa.Text) < 0 Or Val(Combo_Empresa.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Empresa.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")
End Sub

Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(1).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_LostFocus()
  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) < 0 Or Val(Combo_Vendedor.Text) > 9999 Then Exit Sub

  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Combo_Vendedor.Text)
  If rsFuncionarios.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsFuncionarios("Nome")

End Sub

Private Sub Data_Pagto_LostFocus()
  Data_Pagto.Text = Ajusta_Data(Data_Pagto.Text)
End Sub

Private Sub Data_Pagto_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Pagto.Text = frmCalendario.gsDateCalender(Data_Pagto.Text)
  End Select
End Sub

Private Sub Desconto_GotFocus()
  Desconto.SelStart = 0
  Desconto.SelLength = 16
End Sub

Private Sub Desconto_LostFocus()
'  Valor_Pago.Text = CCur(gsHandleNull(Valor.Text)) - CCur(gsHandleNull(Desconto.Text)) + CCur(gsHandleNull(Acréscimo.Text))
End Sub

Private Sub Emissão_LostFocus()
  Emissão.Text = Ajusta_Data(Emissão.Text)
End Sub

Private Sub Emissão_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Emissão.Text = frmCalendario.gsDateCalender(Emissão.Text)
  End Select
End Sub

Private Sub Valor_GotFocus()
  Valor.SelStart = 0
  Valor.SelLength = 16
End Sub

Private Sub Valor_LostFocus()
'  Valor_Pago.Text = CCur(gsHandleNull(Valor.Text)) - CCur(gsHandleNull(Desconto.Text)) + CCur(gsHandleNull(Acréscimo.Text))
End Sub

Private Sub Valor_Pago_GotFocus()
  Valor_Pago.SelStart = 0
  Valor_Pago.SelLength = 16
End Sub

Private Sub Vencimento_LostFocus()
  Vencimento.Text = Ajusta_Data(Vencimento.Text)
End Sub

Private Sub Vencimento_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Vencimento.Text = frmCalendario.gsDateCalender(Vencimento.Text)
  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()

  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  ActiveBar1.Tools("miOpOrdem").CBList.Clear
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 0, "Por Filial, Vencimento"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 1, "Por Filial, Cliente"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 2, "Por Filial, Data Recebimento"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 3, "Por Nota, Cliente"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 4, "Por ID"
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  
  '22/04/2005 - Daniel
  'Otimizado rotina para abrir a tela de lançamentos de contas
  'com a conta selecionada
  '
  'Solicitante: Consultor Carlos (Petrópolis - RJ)
  If frmManContasReceber.g_blnFind Then
    'Carregamos o CR com um único registro escolhido
    Set rsCR = db.OpenRecordset(frmManContasReceber.g_strQuery, dbOpenDynaset)
  Else
    gsSql = "SELECT * FROM [Contas a Receber] "
    gsOrder = "ORDER BY Filial, Vencimento, Cliente"
    Set rsCR = db.OpenRecordset(gsSql & " WHERE Tipo = 'R' " & gsOrder, dbOpenDynaset)
  End If
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsContas_Correntes = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  
  Me.Show
  DoEvents
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName

  Call ActiveBarLoadToolTips(Me)
  
  cboTipo.DataMode = ssDataModeAddItem
  cboTipo.RemoveAll
  cboTipo.AddItem "CARTEIRA"
  cboTipo.AddItem "CARNET"
  cboTipo.AddItem "BANCO - BOLETO"
  cboTipo.Text = "CARTEIRA"
  cboTipo.DataFieldList = "Column0"
  
  Call ClearScreen
  
  '06/06/2005 - Daniel
  'Carregar automaticamente a Filial corrente
  'e a data atual para a Data de Emissão
  Combo_Empresa.Text = gnCodFilial
  Combo_Empresa_LostFocus
  
  Emissão.Text = Format(Data_Atual, "DD/MM/YYYY")
  Combo_Empresa.SetFocus
  '----------------------------------------------
  
  '22/04/2005 - Daniel
  'Exibição da conta a partir da tela de
  'manutenções
  If frmManContasReceber.g_blnFind Then
    Call MoveFirst
    frmManContasReceber.g_blnFind = False
  End If
  
  Screen.MousePointer = vbDefault

  'Guarda para o zzzLog apenas
  sValorOriginal = Valor.Text
  sDtVencOriginal = Vencimento.Text

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsClientes.Close
  rsCR.Close
  rsFuncionarios.Close
  rsContas_Correntes.Close
  Set rsParametros = Nothing
  Set rsClientes = Nothing
  Set rsCR = Nothing
  Set rsFuncionarios = Nothing
  Set rsContas_Correntes = Nothing
End Sub

Private Sub ShowRecord()
  Combo_Empresa.Text = rsCR("Filial")
  Combo_Empresa_LostFocus
  Combo_Cliente.Text = rsCR("Cliente")
  Combo_Cliente_LostFocus
  Combo_Vendedor.Text = rsCR("Vendedor")
  Combo_Vendedor_LostFocus
  
  cboTipo.Text = ""
  If rsCR("Tipo Parcelamento") = "C" Then cboTipo.Text = "CARTEIRA"
  If rsCR("Tipo Parcelamento") = "B" Then cboTipo.Text = "BANCO - BOLETO"
  If rsCR("Tipo Parcelamento") = "T" Then cboTipo.Text = "CARNET"
  
  Combo_Conta.Text = rsCR("Conta Boleto") & ""
  Combo_Conta_LostFocus
  
  Sequência.Text = rsCR("Sequência")
  Nota.Text = rsCR("Nota") & ""
  Fatura.Text = rsCR("Fatura") & ""
  Descrição.Text = rsCR("Descrição") & ""
  If IsDate(rsCR("Data Emissão")) Then
    Emissão.Text = gsFormatDate(rsCR("Data Emissão"))
  Else
    Emissão.Mask = ""
    Emissão.Text = ""
    Emissão.Mask = "##/##/####"
  End If
  If IsDate(rsCR("Vencimento")) Then
    Vencimento.Text = gsFormatDate(rsCR("Vencimento"))
  Else
    Vencimento.Mask = ""
    Vencimento.Text = ""
    Vencimento.Mask = "##/##/####"
  End If
  Valor.Text = rsCR("Valor")
  Desconto.Text = rsCR("Desconto")
  Acréscimo.Text = rsCR("Acréscimo")
  Valor_Pago.Text = rsCR("Valor Recebido")
  If IsDate(rsCR("Data Recebimento")) Then
    Data_Pagto.Text = gsFormatDate(rsCR("Data Recebimento"))
  Else
    Data_Pagto.Mask = ""
    Data_Pagto.Text = ""
    Data_Pagto.Mask = "##/##/####"
  End If
  Conta = rsCR("Contador")
  Boleto.Value = -rsCR("Impresso")
  Carnet.Value = -rsCR("Carnet Impresso")
  
  txtNossoNumero.Text = rsCR.Fields("CNAB_NossoNumero") & ""
  
  If rsCR.Fields("Pendencia").Value = True Then
      chk_pendencia.Value = 1
  Else
      chk_pendencia.Value = 0
  End If
  
  Num_Registro = rsCR.Bookmark
  
End Sub

Private Sub Nota_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

