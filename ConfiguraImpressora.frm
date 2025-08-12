VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmConfiguraImpressora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Configuração de Impressoras"
   ClientHeight    =   5490
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConfiguraImpressora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   7830
   Begin VB.CommandButton cmd_impressorNFCe 
      BackColor       =   &H0063D503&
      Caption         =   "Impressor NFCe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1110
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Frame Frame2 
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
      Height          =   1215
      Left            =   -120
      TabIndex        =   33
      Top             =   -120
      Width           =   7995
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"ConfiguraImpressora.frx":4E95A
         ForeColor       =   &H00808080&
         Height          =   675
         Left            =   1560
         TabIndex        =   35
         Top             =   420
         Width           =   6015
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Configuração do sistema de impressão"
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   180
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "ConfiguraImpressora.frx":4EA0C
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Fecha a caixa de diálogo sem salvar qualquer alteração que tenha sido feita."
      Top             =   4920
      Width           =   7575
   End
   Begin TabDlg.SSTab sstOptions 
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "&Impressora padrão"
      TabPicture(0)   =   "ConfiguraImpressora.frx":50706
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Relatórios"
      TabPicture(1)   =   "ConfiguraImpressora.frx":50722
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picSheet"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkRelZebrados"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdChangeColor"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cdgMensagem"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "shpReportColor"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Opçõ&es"
      TabPicture(2)   =   "ConfiguraImpressora.frx":5073E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraImpressoraSaida"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkConfigCompressionPrinter"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fraImpressoraSaida 
         Caption         =   "Saída de impressão"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   7095
         Begin VB.CommandButton cmdSelecionarPasta 
            Height          =   390
            Left            =   6480
            Picture         =   "ConfiguraImpressora.frx":5075A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   960
            Width           =   435
         End
         Begin VB.TextBox txtPastaSaidaArquivoImpressao 
            Appearance      =   0  'Flat
            BackColor       =   &H00F7F7F7&
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   960
            Width           =   6135
         End
         Begin VB.CheckBox chkUtilizarSistemaRemotoImpressao 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Uitlizar sistema remoto de impressão"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   3675
         End
         Begin VB.Label lblTitlePastaSaidaArquivoImpressao 
            AutoSize        =   -1  'True
            Caption         =   "Pasta de saída para arquivo de impressão"
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   720
            Width           =   3000
         End
      End
      Begin VB.CheckBox chkConfigCompressionPrinter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   $"ConfiguraImpressora.frx":508A4
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   6615
      End
      Begin VB.PictureBox picSheet 
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
         Height          =   2175
         Left            =   -69600
         ScaleHeight     =   2115
         ScaleWidth      =   1680
         TabIndex        =   31
         Top             =   720
         Width           =   1735
         Begin VB.Image imgFundoRel 
            Height          =   1425
            Left            =   135
            Picture         =   "ConfiguraImpressora.frx":50937
            Top             =   500
            Width           =   1395
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Relatório"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   45
            Width           =   1455
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   120
            X2              =   1560
            Y1              =   2000
            Y2              =   2000
         End
         Begin VB.Shape shpLine 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   135
            Index           =   5
            Left            =   120
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Shape shpLine 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   135
            Index           =   4
            Left            =   120
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Shape shpLine 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   135
            Index           =   3
            Left            =   120
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Shape shpLine 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   135
            Index           =   2
            Left            =   120
            Top             =   960
            Width           =   1455
         End
         Begin VB.Shape shpLine 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   720
            Width           =   1455
         End
         Begin VB.Shape shpLine 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   480
            Width           =   1455
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   120
            X2              =   1560
            Y1              =   375
            Y2              =   375
         End
      End
      Begin VB.CheckBox chkRelZebrados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Relatórios devem sair zebrados"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Value           =   1  'Checked
         Width           =   3510
      End
      Begin VB.CommandButton cmdChangeColor 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Escolher cor..."
         Height          =   375
         Left            =   -72120
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Grava Configurações"
         Top             =   1787
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   7350
         Begin SSDataWidgets_B.SSDBCombo Combo_Impressora 
            Height          =   345
            Index           =   4
            Left            =   1320
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1770
            Width           =   3495
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
            BackColorOdd    =   16777152
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   7964
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Porta"
            Columns(1).Name =   "Porta"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   6165
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   12648447
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Impressora 
            Height          =   345
            Index           =   3
            Left            =   1320
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1410
            Width           =   3495
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
            BackColorOdd    =   16777152
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   7964
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Porta"
            Columns(1).Name =   "Porta"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   6165
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   12648447
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Impressora 
            Height          =   345
            Index           =   2
            Left            =   1320
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1050
            Width           =   3495
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
            BackColorOdd    =   16777152
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   7964
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Porta"
            Columns(1).Name =   "Porta"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   6165
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   12648447
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Impressora 
            Height          =   345
            Index           =   1
            Left            =   1320
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   690
            Width           =   3495
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
            BackColorOdd    =   16777152
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   7964
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Porta"
            Columns(1).Name =   "Porta"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   6165
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   12648447
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Impressora 
            Height          =   345
            Index           =   0
            Left            =   1320
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   330
            Width           =   3495
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
            BackColorOdd    =   16777152
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   7964
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Porta"
            Columns(1).Name =   "Porta"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   6165
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   12648447
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Impressora 
            Height          =   345
            Index           =   5
            Left            =   1320
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2130
            Width           =   3495
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
            BackColorOdd    =   16777152
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   7964
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Porta"
            Columns(1).Name =   "Porta"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   6165
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   12648447
         End
         Begin VB.Label Label3 
            Caption         =   "Carnês"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "Relatórios"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "Nota Fiscal"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Tickets"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label32 
            Caption         =   "Cheques"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label49 
            Caption         =   "Boletos"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Porta_Impressora 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   4920
            TabIndex        =   23
            Top             =   2130
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Porta"
            Height          =   255
            Left            =   4920
            TabIndex        =   22
            Top             =   105
            Width           =   1455
         End
         Begin VB.Label Porta_Impressora 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   4920
            TabIndex        =   21
            Top             =   1770
            Width           =   2295
         End
         Begin VB.Label Porta_Impressora 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   4920
            TabIndex        =   20
            Top             =   1410
            Width           =   2295
         End
         Begin VB.Label Porta_Impressora 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   4920
            TabIndex        =   19
            Top             =   1050
            Width           =   2295
         End
         Begin VB.Label Porta_Impressora 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   4920
            TabIndex        =   18
            Top             =   690
            Width           =   2295
         End
         Begin VB.Label Porta_Impressora 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   4920
            TabIndex        =   17
            Top             =   330
            Width           =   2295
         End
      End
      Begin MSComDlg.CommonDialog cdgMensagem 
         Left            =   -75000
         Top             =   2670
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Shape shpReportColor 
         BackColor       =   &H00F7F7F7&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   375
         Left            =   -74400
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cor a ser utilizada no fundo dos relatórios (a cor da letra não mudará)"
         Height          =   390
         Left            =   -74400
         TabIndex        =   30
         Top             =   1290
         Width           =   3270
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Fecha a caixa de diálogo e salva qualquer alteração que tenha sido feita."
      Top             =   4410
      Width           =   7575
   End
End
Attribute VB_Name = "frmConfiguraImpressora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'31/03/2010 - mpdea
'Comentado códigos de utilização do sistema remoto de impressão
'por não ser mais necesário

Private rsReports As Recordset
Private rsReportsTemp As Recordset

Private Declare Function WTSGetActiveConsoleSessionId Lib "Kernel32.dll" () As Long
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetCurrentProcessId Lib "Kernel32" () As Long

Private Type WTS_PROCESS_INFO
    SessionID As Long
    ProcessId As Long
    pProcessName As Long
    pUserSid As Long
End Type

Function TerminalServerSessionId() As String
    Dim lRetVal As Long, lCount As Long, lThisProcess As Long, lThisProcessId  As Long
    Dim lpBuffer As Long, lp As Long, udtProcessInfo As WTS_PROCESS_INFO
    Const WTS_CURRENT_SERVER_HANDLE = 0&
    
    On Error GoTo ErrNotTerminalServer
    'Set Default Value
    TerminalServerSessionId = "0"
    lThisProcessId = GetCurrentProcessId
    lRetVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, lCount)
    If lRetVal Then
        'Successful
        lp = lpBuffer
        For lThisProcess = 1 To lCount
            CopyMemory udtProcessInfo, ByVal lp, LenB(udtProcessInfo)
            If lThisProcessId = udtProcessInfo.ProcessId Then
                TerminalServerSessionId = CStr(udtProcessInfo.SessionID)
                Exit For
            End If
            lp = lp + LenB(udtProcessInfo)
        Next
        'Free memory buffer
        WTSFreeMemory lpBuffer
    End If
    
    Exit Function
    
ErrNotTerminalServer:
    'The machine is not a Terminal Server
    On Error GoTo 0
End Function


Private Sub cmd_impressorNFCe_Click()
On Error GoTo Erro
  Dim strCaminhoExecutavel As String
  strCaminhoExecutavel = App.Path

  Shell strCaminhoExecutavel & "\NFCe\ImpressaoNFCe.exe"

  Exit Sub
Erro:
  MsgBox "Erro na chamada do aplicativo de Impressão de NFCe QuickStore para IX. " & Err.Description, vbInformation, "Atenção"
End Sub

'Private Sub chkUtilizarSistemaRemotoImpressao_Click()
'  Dim bln_enabled As Boolean
'
'  bln_enabled = chkUtilizarSistemaRemotoImpressao.Value = vbChecked
'  txtPastaSaidaArquivoImpressao.Enabled = bln_enabled
'  cmdSelecionarPasta.Enabled = bln_enabled
'End Sub

Private Sub cmdOK_Click()
  Dim sNomeLPT(6) As String
  Dim sPortaLPT(6) As String
  Dim nI As Integer
  Dim nRed As Byte
  Dim nGreen As Byte
  Dim nBlue As Byte
  
  Call StatusMsg("")
  
'  '12/02/2010 - mpdea
'  'Validação
'  If chkUtilizarSistemaRemotoImpressao.Value = vbChecked Then
'    If Trim(txtPastaSaidaArquivoImpressao.Text) = "" Then
'      sstOptions.Tab = 2
'      DisplayMsg "Informe a pasta de saída para o arquivo de impressão."
'      txtPastaSaidaArquivoImpressao.SetFocus
'      Exit Sub
'    End If
'  End If
  

  gSetPrinterName_jaChamou_REL = 0
  gSetPrinterName_jaChamou_NOTA = 0
  gSetPrinterName_jaChamou_TICKET = 0
  gSetPrinterName_jaChamou_CHEQUE = 0
  gSetPrinterName_jaChamou_BOLETO = 0
  gSetPrinterName_jaChamou_CARNE = 0
  
  
  'Impressora padrão
  sNomeLPT(0) = "NOME IMPRESSORA REL"
  sPortaLPT(0) = "PORTA IMPRESSORA REL"
  sNomeLPT(1) = "NOME IMPRESSORA NOTA"
  sPortaLPT(1) = "PORTA IMPRESSORA NOTA"
  sNomeLPT(2) = "NOME IMPRESSORA TICKET"
  sPortaLPT(2) = "PORTA IMPRESSORA TICKET"
  sNomeLPT(3) = "NOME IMPRESSORA CHEQUE"
  sPortaLPT(3) = "PORTA IMPRESSORA CHEQUE"
  sNomeLPT(4) = "NOME IMPRESSORA BOLETO"
  sPortaLPT(4) = "PORTA IMPRESSORA BOLETO"
  sNomeLPT(5) = "NOME IMPRESSORA CARNÊ"
  sPortaLPT(5) = "PORTA IMPRESSORA CARNÊ"
  
  For nI = 0 To 5
    Call UpdateArqConfig("ConfigLPT", sNomeLPT(nI), Combo_Impressora(nI).Text)
    Call UpdateArqConfig("ConfigLPT", sPortaLPT(nI), Porta_Impressora(nI).Caption)
  Next nI
  
  'Relatórios
  Call db.Execute("DELETE * FROM Reports")
  Call dbTemp.Execute("DELETE * FROM Reports")
  Call ConvertRGB(shpReportColor.BackColor, nRed, nGreen, nBlue)
  With rsReports
    .AddNew
    !InRelZebrados = (chkRelZebrados.Value = vbChecked)
    !nColorRed = nRed
    !nColorGreen = nGreen
    !nColorBlue = nBlue
    .Update
  End With
  With rsReportsTemp
    .AddNew
    !InRelZebrados = (chkRelZebrados.Value = vbChecked)
    !nColorRed = nRed
    !nColorGreen = nGreen
    !nColorBlue = nBlue
    .Update
  End With
    
  'Opções
  SaveSetting "QuickStore", "ConfigLPT", "ConfigCompressionPrinter", chkConfigCompressionPrinter.Value
  
'  '12/02/2010 - mpdea
'  SaveSetting "QuickStore", "ConfigLPT", "ModoRemoto", chkUtilizarSistemaRemotoImpressao.Value
'  SaveSetting "QuickStore", "ConfigLPT", "PastaSaidaModoRemoto", txtPastaSaidaArquivoImpressao.Text
  
  Unload Me
  
End Sub

Private Sub chkRelZebrados_Click()
  If chkRelZebrados.Value = vbChecked Then
    cmdChangeColor.Enabled = True
  Else
    cmdChangeColor.Enabled = False
  End If
  Call RefreshRel
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdChangeColor_Click()
  On Error GoTo ErrHandler
  With cdgMensagem
    .CancelError = True
    .ShowColor
    shpReportColor.BackColor = .Color
  End With
  Call RefreshRel
ErrHandler:
End Sub

Private Sub cmdSelecionarPasta_Click()
  Dim sRet As String
  sRet = sFindDir("", Me.hwnd)
  If sRet <> "" Then
    txtPastaSaidaArquivoImpressao.Text = sRet
  End If
End Sub

Private Sub Combo_Impressora_CloseUp(Index As Integer)
  Porta_Impressora(Index).Caption = Combo_Impressora(Index).Columns(1).Text
End Sub

Private Sub Combo_Impressora_LostFocus(Index As Integer)
  If Combo_Impressora(Index).Text = "" Then
    Porta_Impressora(Index).Caption = ""
  End If
End Sub


Private Sub Form_Load()
On Error GoTo Erro

  Dim X As Printer
  Dim Aux As String
  Dim Resp As Integer
  Dim nI As Integer
  Dim sNomeLPT(6) As String
  Dim sPortaLPT(6) As String
  
  Dim sTS_SessionId As String
  
 
  Call CenterForm(Me)
  
  '' 02/03/2023 - Pablo
  '' Verificar se o usuário está executando a aplicação via RDP no servidor A3
  'If giQuick_viaRDP = 1 Then
  If gIsRDP = True Then
      sTS_SessionId = TerminalServerSessionId
      sTS_SessionId = "(" & sTS_SessionId & " redirecionada)"
  End If
  
  For Each X In Printers
    Aux = X.DeviceName
    Aux = Aux + Chr$(9) + X.Port
   
    '' 02/03/2023 - Pablo
    '' Verificar se o usuário está executando a aplicação via RDP no servidor A3
    'If giQuick_viaRDP = 1 Then
    If gIsRDP = True Then
        If InStr(1, Aux, sTS_SessionId) > 0 Then
            Combo_Impressora(0).AddItem Aux
            Combo_Impressora(1).AddItem Aux
            Combo_Impressora(2).AddItem Aux
            Combo_Impressora(3).AddItem Aux
            Combo_Impressora(4).AddItem Aux
            Combo_Impressora(5).AddItem Aux
        End If
    Else
        Combo_Impressora(0).AddItem Aux
        Combo_Impressora(1).AddItem Aux
        Combo_Impressora(2).AddItem Aux
        Combo_Impressora(3).AddItem Aux
        Combo_Impressora(4).AddItem Aux
        Combo_Impressora(5).AddItem Aux
    End If
  Next

  Call StatusMsg("")
  
  sNomeLPT(0) = "NOME IMPRESSORA REL"
  sPortaLPT(0) = "PORTA IMPRESSORA REL"
  sNomeLPT(1) = "NOME IMPRESSORA NOTA"
  sPortaLPT(1) = "PORTA IMPRESSORA NOTA"
  sNomeLPT(2) = "NOME IMPRESSORA TICKET"
  sPortaLPT(2) = "PORTA IMPRESSORA TICKET"
  sNomeLPT(3) = "NOME IMPRESSORA CHEQUE"
  sPortaLPT(3) = "PORTA IMPRESSORA CHEQUE"
  sNomeLPT(4) = "NOME IMPRESSORA BOLETO"
  sPortaLPT(4) = "PORTA IMPRESSORA BOLETO"
  sNomeLPT(5) = "NOME IMPRESSORA CARNÊ"
  sPortaLPT(5) = "PORTA IMPRESSORA CARNÊ"
  
  For nI = 0 To 5
    Aux = GetSetting("QuickStore", "ConfigLPT", sNomeLPT(nI), "")
    If Len(Trim(Aux)) > 0 Then
      Combo_Impressora(nI).Text = Aux
    End If
    Aux = GetSetting("QuickStore", "ConfigLPT", sPortaLPT(nI), "")
    If Len(Trim(Aux)) > 0 Then
      Porta_Impressora(nI).Caption = Aux
    End If
  Next nI
  
  Set rsReports = db.OpenRecordset("SELECT * FROM Reports", dbOpenDynaset)
  Set rsReportsTemp = dbTemp.OpenRecordset("SELECT * FROM Reports", dbOpenDynaset)
  
  chkRelZebrados.Value = 1
  With rsReports
    If Not .EOF Then
      If .Fields("InRelZebrados").Value = False Then
        chkRelZebrados.Value = 0
      End If
      shpReportColor.BackColor = RGB(IIf(IsNull(!nColorRed), 0, !nColorRed), IIf(IsNull(!nColorGreen), 0, !nColorGreen), IIf(IsNull(!nColorBlue), 0, !nColorBlue))
    End If
  End With
  
  Call RefreshRel
  
  chkConfigCompressionPrinter.Value = GetSetting("QuickStore", "ConfigLPT", "ConfigCompressionPrinter", vbChecked)
  
'  '12/02/2010 - mpdea
'  chkUtilizarSistemaRemotoImpressao.Value = GetSetting("QuickStore", "ConfigLPT", "ModoRemoto", vbChecked)
'  txtPastaSaidaArquivoImpressao.Text = GetSetting("QuickStore", "ConfigLPT", "PastaSaidaModoRemoto", "")
'  chkUtilizarSistemaRemotoImpressao_Click
  
  ' ============================================
  ' Tratamento habilitar botão para IX
  Dim strCaminhoExecutavel As String
  Dim iIndice_CaminhoExecutavel As Integer
  
  strCaminhoExecutavel = App.Path
  iIndice_CaminhoExecutavel = InStr(1, strCaminhoExecutavel, "InfoparIX")
  If iIndice_CaminhoExecutavel > -1 Then
      cmd_impressorNFCe.Visible = True
  End If
  ' ============================================
    
  
  Exit Sub
Erro:
    MsgBox "Inconsistência na carga da tela..." & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsReports.Close
  rsReportsTemp.Close
  Set rsReports = Nothing
  Set rsReportsTemp = Nothing
End Sub

Private Sub RefreshRel()
  Dim nX As Byte
  Dim bVisible As Boolean
  
  If chkRelZebrados.Value = vbChecked Then
    bVisible = True
  Else
    bVisible = False
  End If
  
  For nX = 0 To 5
    shpLine(nX).Visible = bVisible
    shpLine(nX).BackColor = shpReportColor.BackColor
  Next nX
End Sub

