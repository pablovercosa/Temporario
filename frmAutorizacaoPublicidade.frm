VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAutorizacaoPublicidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro da Autorização de Publicidade"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutorizacaoPublicidade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11760
   Begin VB.Frame fraFicha 
      Caption         =   "Ficha de Contrato"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   4920
      Width           =   11415
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
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
         Left            =   3840
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdProgramacoes 
         BackColor       =   &H0000C0C0&
         Caption         =   "&Programações"
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
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   $"frmAutorizacaoPublicidade.frx":058A
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblDica 
         AutoSize        =   -1  'True
         Caption         =   "Para cada Autorização de Publicidade, é permitido gerar até 03 Programações: Mês 1, Mês 2 e Mês 3."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   7545
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Observações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   3285
      Width           =   11415
      Begin VB.TextBox txtObservacoes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1200
         MaxLength       =   1200
         MultiLine       =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Número máximo de caracteres 1200."
         Top             =   840
         Width           =   9975
      End
      Begin VB.TextBox txtPatrocinio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1200
         MaxLength       =   1200
         MultiLine       =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Número máximo de caracteres 1200."
         Top             =   240
         Width           =   9975
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Patrocínio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Observações"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1020
         Width           =   945
      End
   End
   Begin VB.Frame fraX 
      Caption         =   "Contrato"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   11415
      Begin VB.TextBox txtVlTotContrato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6765
         MaxLength       =   8
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
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
         Height          =   345
         Left            =   8475
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'C' ORDER BY Código"
         Top             =   600
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox txtNomeCliente 
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
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   645
         Width           =   4455
      End
      Begin VB.TextBox txtComissao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         MaxLength       =   8
         TabIndex        =   7
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtVendedor 
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
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Data datVendedor 
         Caption         =   "datVendedor"
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
         Left            =   8445
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Funcionários ORDER BY Código"
         Top             =   2280
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Data datTipoComercial 
         Caption         =   "datTipoComercial"
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
         Left            =   8445
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Descricao FROM TipoComercial ORDER BY Código"
         Top             =   1860
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox txtTipoComercial 
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
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1875
         Width           =   4455
      End
      Begin VB.Data datRadio 
         Caption         =   "datRadio"
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
         Left            =   8445
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Radio ORDER BY Código"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox txtRadio 
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
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1455
         Width           =   4455
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
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1050
         Width           =   4455
      End
      Begin VB.Data datFornecedor 
         Caption         =   "datFornecedor"
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
         Left            =   8445
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' AND AgenciaPublicidade ORDER BY Código"
         Top             =   1020
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox txtNumAutorizacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmAutorizacaoPublicidade.frx":0612
         Height          =   315
         Left            =   2205
         TabIndex        =   3
         Top             =   1035
         Width           =   1095
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
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         Columns(0).Width=   3200
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboRadio 
         Bindings        =   "frmAutorizacaoPublicidade.frx":062E
         Height          =   315
         Left            =   2205
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
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
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         Columns(0).Width=   3200
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboTipoComercial 
         Bindings        =   "frmAutorizacaoPublicidade.frx":0645
         Height          =   315
         Left            =   2205
         TabIndex        =   5
         Top             =   1845
         Width           =   1095
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
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         Columns(0).Width=   3200
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin MSMask.MaskEdBox mskDataAssinatura 
         Height          =   315
         Left            =   6645
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para carregar o calendário"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin SSDataWidgets_B.SSDBCombo cboVendedor 
         Bindings        =   "frmAutorizacaoPublicidade.frx":0664
         Height          =   315
         Left            =   2205
         TabIndex        =   6
         Top             =   2235
         Width           =   1095
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
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         Columns(0).Width=   3200
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCliente 
         Bindings        =   "frmAutorizacaoPublicidade.frx":067E
         Height          =   315
         Left            =   2205
         TabIndex        =   2
         Top             =   645
         Width           =   1095
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
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total do Contrato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   32
         Top             =   2700
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   31
         Top             =   705
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comissão do Vendedor (%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   2700
         Width           =   1965
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1395
         TabIndex        =   28
         Top             =   2295
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agência de Publicidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   450
         TabIndex        =   27
         Top             =   1095
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rádio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   26
         Top             =   1500
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Autorização"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1005
         TabIndex        =   25
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Comercial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1050
         TabIndex        =   24
         Top             =   1905
         Width           =   1035
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Data de Assinatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5205
         TabIndex        =   22
         Top             =   300
         Width           =   1380
      End
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   120
      Top             =   6960
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
      Bands           =   "frmAutorizacaoPublicidade.frx":0698
   End
End
Attribute VB_Name = "frmAutorizacaoPublicidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'22/01/2004 - Daniel
'26/07/2004 - Alterações II Fase
'Case: STC Sistema Trídio de Comunicação
Private varNumRegistro  As Variant
Private rstContrato     As Recordset
Private m_intDia(30)    As Integer
Private m_strAuxi       As String

Public g_lngNumAutorizacao As Long
Public g_intVendedor       As Integer

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstContrato
    .MoveFirst
    
    If .BOF Then Beep
    If Not .BOF Then Call ShowRecord
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  
  With rstContrato
    .MoveLast
    
    If .EOF Then Beep
    If Not .EOF Then Call ShowRecord
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  
  With rstContrato
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
  
  With rstContrato
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Private Sub DeleteRecord()
  Dim intResposta   As Integer
  Dim strAux        As String
  
  If IsNull(varNumRegistro) Then
    Beep
    DisplayMsg "Não existe registro para apagar."
    Exit Sub
  End If
  
  If ExisteProgramacao Then Exit Sub
  
  strAux = "Deseja realmente apagar este Contrato?"
  intResposta = MsgBox(strAux, 20, "ATENÇÃO")
  
  If intResposta = 6 Then
  
    'Primeiro é necessário destruir as Programações com
    'este Num Autorização para depois excluir o Contrato
    Call DeleteProgramacao(rstContrato.Fields("Num Autorizacao").Value)
  
    rstContrato.Delete
    varNumRegistro = Null
    Call ClearScreen
  End If
End Sub

Private Sub UpdateRecord()
  Dim blnErro   As Boolean
  Dim intDia    As Integer
    
  On Error GoTo Processa_Erro
  
  If Not IsNumeric(txtNumAutorizacao.Text) Then
    MsgBox "Número de Autorização incorreto, verifique.", vbExclamation, "Quick Store"
    txtNumAutorizacao.SetFocus
    Exit Sub
  End If
  
  If blnVerificaCampos Then
    MsgBox "O campo " & m_strAuxi & " não foi preenchido corretamente, verifique.", vbExclamation, "Quick Store"
    
    Select Case m_strAuxi
      Case "Data de Assinatura"
        mskDataAssinatura.SetFocus
      Case "Cliente"
        cboCliente.SetFocus
      Case "Agência de Publicidade"
        cboFornecedor.SetFocus
      Case "Rádio"
        cboRadio.SetFocus
      Case "Tipo Comercial"
        cboTipoComercial.SetFocus
      Case "Vendedor"
        cboVendedor.SetFocus
      Case "Comissão"
        txtComissao.SetFocus
      Case "Valor Total do Contrato"
        txtVlTotContrato.SetFocus
    End Select
    
    Exit Sub
  End If
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
  With rstContrato
    If IsNull(varNumRegistro) Then
      .AddNew
      .Fields("Num Autorizacao") = CLng(txtNumAutorizacao.Text)
      .Fields("Cod Cliente") = CLng(cboCliente.Text)
    Else
      .Edit
    End If
    
    .Fields("Cod Radio") = CInt(cboRadio.Text)
    .Fields("Cod Fornecedor") = CLng(cboFornecedor.Text)
    .Fields("Patrocinio") = txtPatrocinio.Text & ""
    .Fields("Observacoes") = txtObservacoes.Text & ""
    .Fields("Cod TipoComercial") = CInt(cboTipoComercial.Text)
    .Fields("Data Assinatura") = CDate(mskDataAssinatura.Text)
    .Fields("Cod Vendedor").Value = CInt(cboVendedor.Text)
    .Fields("Comissao").Value = CDbl(txtComissao.Text)
    .Fields("VlTotContrato").Value = CDbl(txtVlTotContrato.Text)
    
    .Update
    varNumRegistro = .LastModified
    .Bookmark = varNumRegistro
  End With 'With rstContrato
  
  Call StatusMsg("")
  
  'Call ClearScreen
  
  Exit Sub
  
Processa_Erro:
  MsgBox "Erro (" & Err.Number & ") " & Err.Description, vbCritical, "Quick Store"
End Sub

Public Sub ClearScreen()
  Dim ctlControl As Control
  
  Call StatusMsg("")
  
  txtNumAutorizacao.Text = ""
  mskDataAssinatura.Mask = ""
  mskDataAssinatura.Text = ""
  mskDataAssinatura.Mask = "##/##/####"
  cboCliente.Text = ""
  txtNomeCliente.Text = ""
  cboFornecedor.Text = ""
  txtNomeFornecedor.Text = ""
  cboRadio.Text = ""
  txtRadio.Text = ""
  cboTipoComercial.Text = ""
  txtTipoComercial.Text = ""
  cboVendedor.Text = ""
  txtVendedor.Text = ""
  txtComissao.Text = ""
  txtPatrocinio.Text = ""
  txtObservacoes.Text = ""
  txtVlTotContrato.Text = ""
  
  varNumRegistro = Null
  
  If Not rstContrato.EOF Then
    On Error Resume Next
    rstContrato.MoveFirst
    rstContrato.MovePrevious
    On Error GoTo 0
  End If
  
  SelectAllText txtNumAutorizacao, True
  
End Sub

Sub ShowRecord()
  Dim intDia     As Integer
  Dim dblRet     As Double
  
  With rstContrato
    txtNumAutorizacao.Text = .Fields("Num Autorizacao")
    mskDataAssinatura.Text = .Fields("Data Assinatura").Value & ""
    cboCliente.Text = .Fields("Cod Cliente").Value
    cboCliente_LostFocus
    cboFornecedor.Text = .Fields("Cod Fornecedor").Value & ""
    cboFornecedor_LostFocus
    cboRadio.Text = .Fields("Cod Radio").Value & ""
    cboRadio_LostFocus
    cboTipoComercial.Text = .Fields("Cod TipoComercial")
    cboTipoComercial_LostFocus
    cboVendedor.Text = .Fields("Cod Vendedor").Value
    cboVendedor_LostFocus
    txtComissao.Text = .Fields("Comissao").Value
    txtPatrocinio.Text = .Fields("Patrocinio").Value & ""
    txtObservacoes.Text = .Fields("Observacoes").Value & ""
    txtVlTotContrato.Text = .Fields("VlTotContrato").Value
    
    varNumRegistro = .Bookmark
  End With
End Sub

Private Sub SearchRecord()
  frmPesquisarAutorizacao.Show
End Sub
Private Sub Report()
  frmRelAutorizacaoPublicidade.Show
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
    Case "miOpReport"
      Call Report
  End Select
End Sub

Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(0).Text
  cboCliente_LostFocus
End Sub

Private Sub cboCliente_LostFocus()
  Dim rstClientes As Recordset

  txtNomeCliente.Text = ""
  If Not IsNumeric(cboCliente.Text) Then Exit Sub

  Set rstClientes = db.OpenRecordset("SELECT Código, Nome FROM Cli_For WHERE Código = " & cboCliente.Text, dbOpenDynaset)

  With rstClientes
    If Not (.BOF And .EOF) Then
      txtNomeCliente.Text = .Fields("Nome") & ""
    End If
  End With

  rstClientes.Close
  Set rstClientes = Nothing

End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(0).Text
  cboFornecedor_LostFocus
End Sub

Private Sub cboFornecedor_LostFocus()
  Dim rstFornecedores As Recordset

  txtNomeFornecedor.Text = ""
  If Not IsNumeric(cboFornecedor) Then Exit Sub

  Set rstFornecedores = db.OpenRecordset("SELECT Código, Nome FROM Cli_For WHERE Código = " & cboFornecedor.Text, dbOpenDynaset)

  With rstFornecedores
    If Not (.BOF And .EOF) Then
      txtNomeFornecedor.Text = .Fields("Nome") & ""
    End If
  End With

  rstFornecedores.Close
  Set rstFornecedores = Nothing
End Sub

Private Sub cboRadio_CloseUp()
  cboRadio.Text = cboRadio.Columns(0).Text
  cboRadio_LostFocus
End Sub

Private Sub cboRadio_LostFocus()
  Dim rstRadio As Recordset

  txtRadio.Text = ""
  If Not IsNumeric(cboRadio) Then Exit Sub

  Set rstRadio = db.OpenRecordset("SELECT Código, Nome FROM Radio WHERE Código = " & cboRadio.Text, dbOpenDynaset)

  With rstRadio
    If Not (.BOF And .EOF) Then
      txtRadio.Text = .Fields("Nome") & ""
    End If
  End With

  rstRadio.Close
  Set rstRadio = Nothing
End Sub

Private Sub cboTipoComercial_CloseUp()
  cboTipoComercial.Text = cboTipoComercial.Columns(0).Text
  cboTipoComercial_LostFocus
End Sub

Private Sub cboTipoComercial_LostFocus()
  Dim rstTipoComercial As Recordset

  txtTipoComercial.Text = ""
  If Not IsNumeric(cboTipoComercial) Then Exit Sub

  Set rstTipoComercial = db.OpenRecordset("SELECT Código, Descricao FROM TipoComercial WHERE Código = " & cboTipoComercial.Text, dbOpenDynaset)

  With rstTipoComercial
    If Not (.BOF And .EOF) Then
      txtTipoComercial.Text = .Fields("Descricao") & ""
    End If
  End With

  rstTipoComercial.Close
  Set rstTipoComercial = Nothing

End Sub

Private Sub cboVendedor_CloseUp()
  cboVendedor.Text = cboVendedor.Columns(0).Text
  cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_LostFocus()
  Dim rstFuncionarios As Recordset

  txtVendedor.Text = ""
  If Not IsNumeric(cboVendedor.Text) Then Exit Sub

  Set rstFuncionarios = db.OpenRecordset("SELECT Código, Nome FROM Funcionários WHERE Código = " & cboVendedor.Text & " ORDER BY Código ", dbOpenDynaset)

  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      txtVendedor.Text = .Fields("Nome") & ""
    End If
  End With

  rstFuncionarios.Close
  Set rstFuncionarios = Nothing

End Sub

Private Sub cmdProgramacoes_Click()
  Dim blnVerifica As Boolean
  
  If Not IsNull(varNumRegistro) Then
    
    If Not IsNumeric(txtNumAutorizacao) Then
      blnVerifica = True
    End If
    
    blnVerifica = Len(txtNomeFornecedor.Text) <= 0 Or Len(txtRadio.Text) <= 0 Or Len(txtTipoComercial.Text) <= 0 Or Len(txtVendedor.Text) <= 0 Or Len(txtComissao.Text) <= 0
    
    If blnVerifica Then
      MsgBox "Impossível criar a Programação, algum campo da Autorização deve estar inválido ou não preenchido, verifique.", vbExclamation, "Quick Store"
      Exit Sub
    Else
      g_lngNumAutorizacao = CLng(txtNumAutorizacao.Text)
      g_intVendedor = CInt(cboVendedor.Text)
      frmProgramacao.Show
    End If
  Else
      MsgBox "Salve primeiro o Contrato para em seguida criar suas respectivas Programações.", vbExclamation, "Quick Store"
      Exit Sub
  End If
  
End Sub

Private Sub cmdSair_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
  datClientes.DatabaseName = gsQuickDBFileName
  datFornecedor.DatabaseName = gsQuickDBFileName
  datRadio.DatabaseName = gsQuickDBFileName
  datTipoComercial.DatabaseName = gsQuickDBFileName
  datVendedor.DatabaseName = gsQuickDBFileName
  
  Call CenterForm(Me)
  
  Set rstContrato = db.OpenRecordset("Contrato", dbOpenDynaset)
  
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Liberar o Recordset
  rstContrato.Close
  Set rstContrato = Nothing
End Sub

Private Sub mskDataAssinatura_KeyDown(KeyCode As Integer, Shift As Integer)
'A tecla está pressionada para baixo
  If KeyCode = vbKeyF2 Then
    mskDataAssinatura.Text = frmCalendario.gsDateCalender(mskDataAssinatura.Text)
  End If
End Sub

Private Sub mskDataAssinatura_LostFocus()
  mskDataAssinatura.Text = Ajusta_Data(mskDataAssinatura.Text)
End Sub

Public Sub txtNumAutorizacao_LostFocus()
  If Not IsNumeric(txtNumAutorizacao.Text) Then Exit Sub
  
  rstContrato.FindFirst "[Num Autorizacao] = " & CLng(txtNumAutorizacao.Text)
  If Not rstContrato.NoMatch Then
    Call ShowRecord
  Else
    varNumRegistro = Null
  End If
End Sub

Private Function blnVerificaCampos() As Boolean

  m_strAuxi = ""
  
  If Not IsDate(mskDataAssinatura.Text) Then
    blnVerificaCampos = True
    m_strAuxi = "Data de Assinatura"
  End If
  
  If Len(txtNomeCliente.Text) <= 0 Then
    blnVerificaCampos = True
    m_strAuxi = "Cliente"
  End If
  
  If Len(txtNomeFornecedor.Text) <= 0 Then
    blnVerificaCampos = True
    m_strAuxi = "Agência de Publicidade"
  End If
  
  If Len(txtRadio.Text) <= 0 Then
    blnVerificaCampos = True
    m_strAuxi = "Rádio"
  End If
  
  If Len(txtTipoComercial.Text) <= 0 Then
    blnVerificaCampos = True
    m_strAuxi = "Tipo Comercial"
  End If
  
  If Len(txtVendedor.Text) <= 0 Then
    blnVerificaCampos = True
    m_strAuxi = "Vendedor"
  End If
  
  If Len(txtComissao.Text) <= 0 Or Not IsNumeric(txtComissao.Text) Then
    blnVerificaCampos = True
    m_strAuxi = "Comissão"
  End If
  
  If Not IsNumeric(txtVlTotContrato.Text) Then
    blnVerificaCampos = True
    m_strAuxi = "Valor Total do Contrato"
  End If
  
End Function

Private Sub DeleteProgramacao(ByVal NumAutorizacao As Long)
  Dim rstProgramacao As Recordset
  
  Set rstProgramacao = db.OpenRecordset("SELECT * FROM Programacao WHERE [Num Autorizacao] = " & NumAutorizacao, dbOpenDynaset)

  With rstProgramacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        .Delete
        
        .MoveNext
      Loop
      
    End If
  End With
  
  rstProgramacao.Close
  Set rstProgramacao = Nothing

End Sub

Private Function ExisteProgramacao() As Boolean
  Dim rstProgramacao As Recordset
  Dim strQuery       As String

  strQuery = "SELECT [Num Autorizacao] "
  strQuery = strQuery & " FROM Programacao "
  strQuery = strQuery & " WHERE [Num Autorizacao] = " & CLng(txtNumAutorizacao.Text)

  Set rstProgramacao = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstProgramacao.RecordCount = 0 Then Exit Function
  
  With rstProgramacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      ExisteProgramacao = True
    End If
    .Close
  End With

  Set rstProgramacao = Nothing
  
  MsgBox "Impossível excluir este Contrato, existem Programações amarradas nele.", vbExclamation, "Quick Store"

End Function

