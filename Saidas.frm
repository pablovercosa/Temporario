VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSaidas 
   Appearance      =   0  'Flat
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Sa�das"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   330
   ClientWidth     =   15315
   DrawMode        =   8  'Xor Pen
   FillColor       =   &H00666666&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F7F7F7&
   HelpContextID   =   1390
   Icon            =   "Saidas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   158.221
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   270.14
   Begin VB.CommandButton btnComandaVendas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      Height          =   375
      Left            =   14790
      Picture         =   "Saidas.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   413
      Width           =   405
   End
   Begin VB.TextBox txtComanda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   13395
      MaxLength       =   13
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   420
      Width           =   1361
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E5E5E5&
      Height          =   735
      Left            =   930
      TabIndex        =   120
      Top             =   1740
      Width           =   6405
      Begin VB.TextBox txt_chave10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5220
         MaxLength       =   4
         TabIndex        =   131
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5790
         MaxLength       =   4
         TabIndex        =   130
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4650
         MaxLength       =   4
         TabIndex        =   129
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3510
         MaxLength       =   4
         TabIndex        =   128
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   127
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2370
         MaxLength       =   4
         TabIndex        =   126
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2940
         MaxLength       =   4
         TabIndex        =   125
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1230
         MaxLength       =   4
         TabIndex        =   124
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   123
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   660
         MaxLength       =   4
         TabIndex        =   122
         Top             =   270
         Width           =   525
      End
      Begin VB.TextBox txt_chave1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         MaxLength       =   4
         TabIndex        =   121
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.CommandButton cmd_incluirChave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Incluir"
      Height          =   315
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   585
   End
   Begin VB.CommandButton cmd_excluirChave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Excluir"
      Height          =   315
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2130
      Width           =   585
   End
   Begin VB.CommandButton cmd_gerarNFe 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10440
      Picture         =   "Saidas.frx":4EE2F
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   7560
      Width           =   915
   End
   Begin VB.Frame frm_produtoSemPrecoNaGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   4890
      TabIndex        =   115
      Top             =   7140
      Visible         =   0   'False
      Width           =   3675
      Begin VB.CommandButton cmd_fecharFrameProdutoSemPrecoNaGrade 
         BackColor       =   &H008080FF&
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Produto sem pre�o na grade.   Se ok, ignore."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   90
         TabIndex        =   117
         Top             =   30
         Width           =   3705
      End
   End
   Begin VB.CommandButton cmd_acataUsuarioLogadoComoOperador 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   11790
      Picture         =   "Saidas.frx":50141
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Acata o Usu�rio Logado no sistema como Operador"
      Top             =   780
      Width           =   525
   End
   Begin VB.CommandButton cmd_devolucaoProdutos 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cliente devolve produtos"
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
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   7950
      Width           =   2295
   End
   Begin VB.CommandButton cmd_tabelaDePrecos 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tabela de Pre�os"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13590
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   6720
      Width           =   1695
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
      Left            =   9870
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Servi�os"
      Top             =   8460
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox L_Tot_Prod 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H00666666&
      Height          =   285
      Left            =   6540
      Locked          =   -1  'True
      TabIndex        =   111
      Top             =   8970
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CheckBox chk_freteNaoSomaPercentual 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   540
      MaskColor       =   &H00FFA324&
      OLEDropMode     =   1  'Manual
      TabIndex        =   110
      ToolTipText     =   "Soma ou n�o o frete no valor total da NFe para o descritivo de Estimativa de Impostos"
      Top             =   7260
      Value           =   1  'Checked
      Width           =   240
   End
   Begin VB.TextBox L_Tot_Desc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   285
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   109
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox L_Tot_IPI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   108
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox L_Frete 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   285
      Left            =   60
      TabIndex        =   107
      Top             =   7500
      Width           =   975
   End
   Begin VB.TextBox L_Base_ICM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   285
      Left            =   5820
      Locked          =   -1  'True
      TabIndex        =   106
      Top             =   8700
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox L_Valor_ICM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   285
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   105
      Top             =   8820
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox L_Base_ICM_Subs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   285
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   104
      Top             =   7500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox L_Valor_ICM_Subs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   103
      Top             =   7500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox L_Tot_Serv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H00666666&
      Height          =   285
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   102
      Top             =   8790
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtNrTerminal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   101
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtImpostosSobreServicos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   285
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   100
      ToolTipText     =   "Este campo � soma de Impostos Sobre o Faturamento [ CSLL, COFINS, PIS e IRRF ]"
      Top             =   8790
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSeguro 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   285
      Left            =   1080
      TabIndex        =   99
      ToolTipText     =   "O Seguro tem origem em vendas realizadas atrav�s da Loja Virtual (web)"
      Top             =   7500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox L_Tot_ICMS_Deson 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   285
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   97
      Top             =   8670
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox L_Tot_ISS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   96
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   13395
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   30
      Width           =   1815
   End
   Begin VB.TextBox txtDescSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   13455
      Locked          =   -1  'True
      TabIndex        =   80
      Text            =   "0"
      Top             =   7665
      Width           =   1815
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   13455
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   7275
      Width           =   1815
   End
   Begin VB.TextBox L_Tot_Pagar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   13455
      Locked          =   -1  'True
      TabIndex        =   78
      Text            =   "0"
      Top             =   8055
      Width           =   1815
   End
   Begin VB.Frame F_Empr�stimo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Caption         =   "Data Acerto Empr�stimo"
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   4290
      TabIndex        =   72
      Top             =   7230
      Visible         =   0   'False
      Width           =   2040
      Begin MSMask.MaskEdBox Data_Acerto 
         Height          =   285
         Left            =   390
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   12648447
         ForeColor       =   0
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
   End
   Begin VB.TextBox Obs 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00666666&
      Height          =   300
      Left            =   945
      MaxLength       =   70
      TabIndex        =   14
      Top             =   1110
      Width           =   8475
   End
   Begin VB.TextBox txtRef 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00666666&
      Height          =   300
      Left            =   10365
      MaxLength       =   10
      TabIndex        =   15
      Top             =   1110
      Width           =   1965
   End
   Begin VB.TextBox Senha 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   13395
      MaxLength       =   8
      PasswordChar    =   "�"
      TabIndex        =   13
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtNF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   60
      MaxLength       =   10
      TabIndex        =   25
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox txtNrSerieNF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   26
      ToolTipText     =   "Entre com o n�mero de s�rie da nota"
      Top             =   8040
      Width           =   975
   End
   Begin VB.ComboBox cboConsumidorFinal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Saidas.frx":506CB
      Left            =   6915
      List            =   "Saidas.frx":506D5
      TabIndex        =   17
      Text            =   "1=Sim"
      Top             =   1455
      Width           =   915
   End
   Begin VB.ComboBox cboPresencaComprador 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Saidas.frx":506E7
      Left            =   8700
      List            =   "Saidas.frx":506FD
      TabIndex        =   18
      Text            =   "1 =Opera��o presencial"
      Top             =   1455
      Width           =   3660
   End
   Begin VB.ComboBox cboFinalidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Saidas.frx":507D1
      Left            =   945
      List            =   "Saidas.frx":507E1
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1455
      Width           =   4740
   End
   Begin VB.TextBox textAliqDestino 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2100
      TabIndex        =   28
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbAliquotaOrigem 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Saidas.frx":50836
      Left            =   3120
      List            =   "Saidas.frx":50843
      TabIndex        =   27
      Top             =   8040
      Visible         =   0   'False
      Width           =   1530
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4725
      Left            =   45
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2505
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   8334
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   15066597
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Produtos"
      TabPicture(0)   =   "Saidas.frx":508B6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblQtdeTotal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSubTotal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDescSubTotal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTitleQtdeTotal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Grade1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DropDown1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&Servi�os"
      TabPicture(1)   =   "Saidas.frx":508D2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DropDown2"
      Tab(1).Control(1)=   "Combo_T�cnico"
      Tab(1).Control(2)=   "Prometido_Para"
      Tab(1).Control(3)=   "Or�amento_Aprovado"
      Tab(1).Control(4)=   "B_Servi�os_Conc"
      Tab(1).Control(5)=   "Grade_Serv"
      Tab(1).Control(6)=   "Label23"
      Tab(1).Control(7)=   "Nome_T�cnico"
      Tab(1).Control(8)=   "L_Aprovado"
      Tab(1).Control(9)=   "Label21"
      Tab(1).ControlCount=   10
      Begin SSDataWidgets_B.SSDBDropDown DropDown2 
         Bindings        =   "Saidas.frx":508EE
         Height          =   1980
         Left            =   -71940
         TabIndex        =   112
         Top             =   2010
         Width           =   10965
         DataFieldList   =   "Nome"
         ListAutoValidate=   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadFont3D      =   2
         BevelColorShadow=   14671839
         BevelColorFace  =   15066597
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   15724527
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   185
         Columns.Count   =   3
         Columns(0).Width=   10689
         Columns(0).Caption=   "Descri��o"
         Columns(0).Name =   "Descri��o"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descri��o"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3360
         Columns(1).Caption=   "C�digo"
         Columns(1).Name =   "C�digo"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "C�digo"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         Columns(2).Width=   4366
         Columns(2).Caption=   "Pre�o"
         Columns(2).Name =   "Pre�o"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Pre�o"
         Columns(2).DataType=   5
         Columns(2).FieldLen=   256
         _ExtentX        =   19341
         _ExtentY        =   3492
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16250871
      End
      Begin SSDataWidgets_B.SSDBDropDown DropDown1 
         Bindings        =   "Saidas.frx":50902
         Height          =   2355
         Left            =   990
         TabIndex        =   36
         Top             =   1575
         Width           =   10890
         DataFieldList   =   "Nome"
         ListAutoValidate=   0   'False
         MaxDropDownItems=   16
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadFont3D      =   2
         BevelColorFrame =   -2147483632
         BevelColorShadow=   -2147483633
         BevelColorFace  =   15066597
         CheckBox3D      =   0   'False
         BackColorEven   =   15724527
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   185
         Columns.Count   =   5
         Columns(0).Width=   8229
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "C�digo"
         Columns(1).Name =   "C�digo"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "C�digo"
         Columns(1).DataType=   5
         Columns(1).FieldLen=   256
         Columns(2).Width=   2910
         Columns(2).Caption=   "Estoque"
         Columns(2).Name =   "Estoque"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Estoque"
         Columns(2).DataType=   11
         Columns(2).FieldLen=   256
         Columns(3).Width=   1984
         Columns(3).Caption=   "Pre�o"
         Columns(3).Name =   "Pre�o"
         Columns(3).Alignment=   1
         Columns(3).CaptionAlignment=   1
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3096
         Columns(4).Caption=   "Fabricante"
         Columns(4).Name =   "Fabricante"
         Columns(4).Alignment=   1
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         _ExtentX        =   19209
         _ExtentY        =   4154
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16250871
      End
      Begin SSDataWidgets_B.SSDBGrid Grade1 
         Height          =   4050
         Left            =   45
         TabIndex        =   30
         Top             =   360
         Width           =   15165
         _Version        =   196617
         DataMode        =   1
         GroupHeaders    =   0   'False
         GroupHeadLines  =   0
         stylesets.count =   4
         stylesets(0).Name=   "Bold"
         stylesets(0).ForeColor=   0
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "Saidas.frx":50916
         stylesets(1).Name=   "Normal"
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "Saidas.frx":50932
         stylesets(2).Name=   "Total"
         stylesets(2).ForeColor=   32768
         stylesets(2).HasFont=   -1  'True
         BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(2).Picture=   "Saidas.frx":5094E
         stylesets(3).Name=   "Grid_Saidas"
         stylesets(3).ForeColor=   -2147483640
         stylesets(3).HasFont=   -1  'True
         BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(3).Picture=   "Saidas.frx":5096A
         UseGroups       =   -1  'True
         HeadFont3D      =   2
         BevelColorFrame =   14671839
         BevelColorHighlight=   16777215
         BevelColorShadow=   15724527
         BevelColorFace  =   15724527
         CheckBox3D      =   0   'False
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   0
         SelectByCell    =   -1  'True
         StyleSet        =   "Grid_Saidas"
         ForeColorEven   =   0
         BackColorEven   =   16250871
         BackColorOdd    =   12648447
         RowHeight       =   503
         ExtraHeight     =   53
         Groups(0).Width =   27331
         Groups(0).Caption=   "Produtos"
         Groups(0).AllowSizing=   0   'False
         Groups(0).HasHeadForeColor=   -1  'True
         Groups(0).HasHeadBackColor=   -1  'True
         Groups(0).HeadBackColor=   13619151
         Groups(0).Columns.Count=   21
         Groups(0).Columns(0).Width=   2514
         Groups(0).Columns(0).Caption=   "C�digo"
         Groups(0).Columns(0).Name=   "C�digo"
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).DataType=   8
         Groups(0).Columns(0).Case=   2
         Groups(0).Columns(0).FieldLen=   40
         Groups(0).Columns(0).HasHeadForeColor=   -1  'True
         Groups(0).Columns(0).HeadBackColor=   13619151
         Groups(0).Columns(1).Width=   1429
         Groups(0).Columns(1).Caption=   "Qtde"
         Groups(0).Columns(1).Name=   "Qtde"
         Groups(0).Columns(1).Alignment=   2
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).DataType=   5
         Groups(0).Columns(1).FieldLen=   256
         Groups(0).Columns(2).Width=   5080
         Groups(0).Columns(2).Caption=   "Nome"
         Groups(0).Columns(2).Name=   "Nome"
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   8
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(2).Locked=   -1  'True
         Groups(0).Columns(2).HeadBackColor=   -2147483633
         Groups(0).Columns(2).BackColor=   12632256
         Groups(0).Columns(3).Width=   767
         Groups(0).Columns(3).Caption=   "UN"
         Groups(0).Columns(3).Name=   "Unidade"
         Groups(0).Columns(3).Alignment=   1
         Groups(0).Columns(3).CaptionAlignment=   2
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   8
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(3).Locked=   -1  'True
         Groups(0).Columns(4).Width=   1138
         Groups(0).Columns(4).Caption=   "Pre�o Unit."
         Groups(0).Columns(4).Name=   "Pre�o Unit�rio"
         Groups(0).Columns(4).Alignment=   1
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(5).Width=   1349
         Groups(0).Columns(5).Caption=   "SubTotal"
         Groups(0).Columns(5).Name=   "Pre�o Total"
         Groups(0).Columns(5).Alignment=   1
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   8
         Groups(0).Columns(5).NumberFormat=   "###,###,##0.00"
         Groups(0).Columns(5).FieldLen=   20
         Groups(0).Columns(5).Locked=   -1  'True
         Groups(0).Columns(5).StyleSet=   "Total"
         Groups(0).Columns(6).Width=   873
         Groups(0).Columns(6).Caption=   "%-$"
         Groups(0).Columns(6).Name=   "Desconto"
         Groups(0).Columns(6).Alignment=   1
         Groups(0).Columns(6).CaptionAlignment=   2
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   8
         Groups(0).Columns(6).FieldLen=   256
         Groups(0).Columns(7).Width=   714
         Groups(0).Columns(7).Caption=   "ICM "
         Groups(0).Columns(7).Name=   "ICM"
         Groups(0).Columns(7).Alignment=   1
         Groups(0).Columns(7).DataField=   "Column 7"
         Groups(0).Columns(7).DataType=   8
         Groups(0).Columns(7).FieldLen=   256
         Groups(0).Columns(7).Locked=   -1  'True
         Groups(0).Columns(8).Width=   794
         Groups(0).Columns(8).Caption=   "IPI "
         Groups(0).Columns(8).Name=   "IPI"
         Groups(0).Columns(8).Alignment=   1
         Groups(0).Columns(8).DataField=   "Column 8"
         Groups(0).Columns(8).DataType=   8
         Groups(0).Columns(8).FieldLen=   256
         Groups(0).Columns(9).Width=   820
         Groups(0).Columns(9).Caption=   "CFOP"
         Groups(0).Columns(9).Name=   "CFOP_Produto"
         Groups(0).Columns(9).Alignment=   2
         Groups(0).Columns(9).CaptionAlignment=   2
         Groups(0).Columns(9).DataField=   "Column 9"
         Groups(0).Columns(9).DataType=   8
         Groups(0).Columns(9).FieldLen=   14
         Groups(0).Columns(10).Width=   741
         Groups(0).Columns(10).Caption=   "Etiq"
         Groups(0).Columns(10).Name=   "Etiqueta"
         Groups(0).Columns(10).Alignment=   2
         Groups(0).Columns(10).AllowSizing=   0   'False
         Groups(0).Columns(10).DataField=   "Column 10"
         Groups(0).Columns(10).DataType=   11
         Groups(0).Columns(10).FieldLen=   1
         Groups(0).Columns(11).Width=   1349
         Groups(0).Columns(11).Caption=   "Total"
         Groups(0).Columns(11).Name=   "Pre�o Final"
         Groups(0).Columns(11).Alignment=   1
         Groups(0).Columns(11).CaptionAlignment=   2
         Groups(0).Columns(11).DataField=   "Column 11"
         Groups(0).Columns(11).DataType=   8
         Groups(0).Columns(11).NumberFormat=   "###,###,##0.00"
         Groups(0).Columns(11).FieldLen=   20
         Groups(0).Columns(11).Locked=   -1  'True
         Groups(0).Columns(11).StyleSet=   "Total"
         Groups(0).Columns(12).Width=   344
         Groups(0).Columns(12).Visible=   0   'False
         Groups(0).Columns(12).Caption=   "Base_ICM"
         Groups(0).Columns(12).Name=   "Base_ICM"
         Groups(0).Columns(12).DataField=   "Column 12"
         Groups(0).Columns(12).DataType=   8
         Groups(0).Columns(12).FieldLen=   256
         Groups(0).Columns(13).Width=   159
         Groups(0).Columns(13).Visible=   0   'False
         Groups(0).Columns(13).Caption=   "Valor_ICM"
         Groups(0).Columns(13).Name=   "Valor_ICM"
         Groups(0).Columns(13).DataField=   "Column 13"
         Groups(0).Columns(13).DataType=   8
         Groups(0).Columns(13).FieldLen=   256
         Groups(0).Columns(14).Width=   7990
         Groups(0).Columns(14).Visible=   0   'False
         Groups(0).Columns(14).Caption=   "Valor_Base_Unit"
         Groups(0).Columns(14).Name=   "Valor_Base_Unit"
         Groups(0).Columns(14).DataField=   "Column 14"
         Groups(0).Columns(14).DataType=   8
         Groups(0).Columns(14).FieldLen=   256
         Groups(0).Columns(15).Width=   2143
         Groups(0).Columns(15).Caption=   "Redu��o_ICM"
         Groups(0).Columns(15).Name=   "Redu��o_ICM"
         Groups(0).Columns(15).DataField=   "Column 15"
         Groups(0).Columns(15).DataType=   8
         Groups(0).Columns(15).FieldLen=   256
         Groups(0).Columns(16).Width=   9260
         Groups(0).Columns(16).Visible=   0   'False
         Groups(0).Columns(16).Caption=   "Tipo_ICM"
         Groups(0).Columns(16).Name=   "Tipo_ICM"
         Groups(0).Columns(16).DataField=   "Column 16"
         Groups(0).Columns(16).DataType=   8
         Groups(0).Columns(16).FieldLen=   256
         Groups(0).Columns(17).Width=   1826
         Groups(0).Columns(17).Caption=   "Despesas"
         Groups(0).Columns(17).Name=   "Desp_Acessorias"
         Groups(0).Columns(17).Alignment=   2
         Groups(0).Columns(17).DataField=   "Column 17"
         Groups(0).Columns(17).DataType=   5
         Groups(0).Columns(17).FieldLen=   256
         Groups(0).Columns(18).Width=   847
         Groups(0).Columns(18).Visible=   0   'False
         Groups(0).Columns(18).Caption=   "ICMS Deson"
         Groups(0).Columns(18).Name=   "Valor Desonerado"
         Groups(0).Columns(18).DataField=   "Column 18"
         Groups(0).Columns(18).DataType=   8
         Groups(0).Columns(18).FieldLen=   256
         Groups(0).Columns(19).Width=   1773
         Groups(0).Columns(19).Caption=   "% Diferimento"
         Groups(0).Columns(19).Name=   "% Diferimento"
         Groups(0).Columns(19).DataField=   "Column 19"
         Groups(0).Columns(19).DataType=   8
         Groups(0).Columns(19).FieldLen=   256
         Groups(0).Columns(20).Width=   4022
         Groups(0).Columns(20).Caption=   " Adicional"
         Groups(0).Columns(20).Name=   "Descri��o Adicional"
         Groups(0).Columns(20).DataField=   "Column 20"
         Groups(0).Columns(20).DataType=   8
         Groups(0).Columns(20).FieldLen=   50
         UseDefaults     =   0   'False
         _ExtentX        =   26749
         _ExtentY        =   7144
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   12566463
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_T�cnico 
         Bindings        =   "Saidas.frx":50986
         DataSource      =   "Data7"
         Height          =   285
         Left            =   -67080
         TabIndex        =   34
         Top             =   4005
         Width           =   750
         DataFieldList   =   "Nome"
         ListAutoPosition=   0   'False
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
         BevelColorShadow=   15066597
         BevelColorFace  =   15066597
         ForeColorEven   =   0
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   6879
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Apelido"
         Columns(1).Name =   "Apelido"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Apelido"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1799
         Columns(2).Caption=   "C�digo"
         Columns(2).Name =   "C�digo"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "C�digo"
         Columns(2).DataType=   3
         Columns(2).FieldLen=   256
         _ExtentX        =   1323
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483630
         BackColor       =   16250871
      End
      Begin VB.TextBox Prometido_Para 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -72915
         MaxLength       =   50
         TabIndex        =   32
         Top             =   4005
         Width           =   4950
      End
      Begin VB.TextBox Or�amento_Aprovado 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -72915
         MaxLength       =   50
         TabIndex        =   33
         Top             =   4320
         Width           =   4950
      End
      Begin VB.CommandButton B_Servi�os_Conc 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Caption         =   "TODOS OS SERVI�OS Ok"
         Height          =   285
         Left            =   -62310
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4005
         Width           =   2235
      End
      Begin SSDataWidgets_B.SSDBGrid Grade_Serv 
         Height          =   3615
         Left            =   -74910
         TabIndex        =   31
         Top             =   360
         Width           =   14910
         _Version        =   196617
         DataMode        =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semilight"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets.count =   3
         stylesets(0).Name=   "Normal"
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "Saidas.frx":5099A
         stylesets(1).Name=   "Total"
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "Saidas.frx":509B6
         stylesets(2).Name=   "Font12"
         stylesets(2).HasFont=   -1  'True
         BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(2).Picture=   "Saidas.frx":509D2
         UseGroups       =   -1  'True
         BevelColorFrame =   6710886
         BevelColorHighlight=   16250871
         BevelColorShadow=   15066597
         BevelColorFace  =   15066597
         CheckBox3D      =   0   'False
         MultiLine       =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         ForeColorEven   =   16250871
         ForeColorOdd    =   12648447
         BackColorEven   =   16250871
         BackColorOdd    =   12648447
         RowHeight       =   450
         ExtraHeight     =   265
         Groups(0).Width =   25003
         Groups(0).Caption=   "Servi�os"
         Groups(0).Columns.Count=   9
         Groups(0).Columns(0).Width=   2117
         Groups(0).Columns(0).Caption=   "C�digo"
         Groups(0).Columns(0).Name=   "C�digo"
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).DataType=   8
         Groups(0).Columns(0).FieldLen=   20
         Groups(0).Columns(0).HasForeColor=   -1  'True
         Groups(0).Columns(0).StyleSet=   "Normal"
         Groups(0).Columns(1).Width=   12330
         Groups(0).Columns(1).Caption=   "Descri��o"
         Groups(0).Columns(1).Name=   "Descri��o"
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).DataType=   8
         Groups(0).Columns(1).FieldLen=   60
         Groups(0).Columns(1).HasForeColor=   -1  'True
         Groups(0).Columns(1).ForeColor=   4144959
         Groups(0).Columns(1).StyleSet=   "Normal"
         Groups(0).Columns(2).Width=   2355
         Groups(0).Columns(2).Caption=   "Qtde"
         Groups(0).Columns(2).Name=   "Qtde"
         Groups(0).Columns(2).Alignment=   1
         Groups(0).Columns(2).CaptionAlignment=   1
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   8
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(2).HasForeColor=   -1  'True
         Groups(0).Columns(2).ForeColor=   4144959
         Groups(0).Columns(2).StyleSet=   "Normal"
         Groups(0).Columns(3).Width=   1455
         Groups(0).Columns(3).Caption=   "CFOP"
         Groups(0).Columns(3).Name=   "CFOP"
         Groups(0).Columns(3).Alignment=   2
         Groups(0).Columns(3).CaptionAlignment=   2
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   8
         Groups(0).Columns(3).FieldLen=   14
         Groups(0).Columns(3).HasForeColor=   -1  'True
         Groups(0).Columns(3).ForeColor=   4144959
         Groups(0).Columns(3).StyleSet=   "Normal"
         Groups(0).Columns(4).Width=   2487
         Groups(0).Columns(4).Caption=   "Pre�o"
         Groups(0).Columns(4).Name=   "Pre�o"
         Groups(0).Columns(4).Alignment=   1
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).NumberFormat=   "#,###,##0.00"
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(4).HasForeColor=   -1  'True
         Groups(0).Columns(4).ForeColor=   4144959
         Groups(0).Columns(4).StyleSet=   "Normal"
         Groups(0).Columns(5).Width=   2566
         Groups(0).Columns(5).Caption=   "Total"
         Groups(0).Columns(5).Name=   "Total"
         Groups(0).Columns(5).Alignment=   1
         Groups(0).Columns(5).CaptionAlignment=   1
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   8
         Groups(0).Columns(5).NumberFormat=   "#,###,##0.00"
         Groups(0).Columns(5).FieldLen=   256
         Groups(0).Columns(5).HasForeColor=   -1  'True
         Groups(0).Columns(5).ForeColor=   4144959
         Groups(0).Columns(5).StyleSet=   "Normal"
         Groups(0).Columns(6).Width=   688
         Groups(0).Columns(6).Caption=   "Ok"
         Groups(0).Columns(6).Name=   "Completo"
         Groups(0).Columns(6).Alignment=   2
         Groups(0).Columns(6).CaptionAlignment=   2
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   8
         Groups(0).Columns(6).FieldLen=   256
         Groups(0).Columns(6).Style=   2
         Groups(0).Columns(6).HasForeColor=   -1  'True
         Groups(0).Columns(6).ForeColor=   4144959
         Groups(0).Columns(7).Width=   3200
         Groups(0).Columns(7).Visible=   0   'False
         Groups(0).Columns(7).Caption=   "iss"
         Groups(0).Columns(7).Name=   "iss"
         Groups(0).Columns(7).DataField=   "Column 7"
         Groups(0).Columns(7).DataType=   8
         Groups(0).Columns(7).FieldLen=   256
         Groups(0).Columns(8).Width=   1005
         Groups(0).Columns(8).Caption=   "CST"
         Groups(0).Columns(8).Name=   "CST"
         Groups(0).Columns(8).Alignment=   1
         Groups(0).Columns(8).CaptionAlignment=   2
         Groups(0).Columns(8).DataField=   "Column 8"
         Groups(0).Columns(8).DataType=   8
         Groups(0).Columns(8).FieldLen=   20
         Groups(0).Columns(8).HasForeColor=   -1  'True
         Groups(0).Columns(8).ForeColor=   5197647
         Groups(0).Columns(8).StyleSet=   "Normal"
         UseDefaults     =   0   'False
         _ExtentX        =   26300
         _ExtentY        =   6376
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   16250871
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.Label lblTitleQtdeTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantidade de produtos"
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
         Height          =   255
         Left            =   45
         TabIndex        =   75
         Top             =   4410
         Width           =   2325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sequ�ncia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   345
         Left            =   11700
         TabIndex        =   67
         Top             =   7635
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   345
         Left            =   12420
         TabIndex        =   66
         Top             =   7200
         Width           =   735
      End
      Begin VB.Label lblDescSubTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto no SubTotal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   195
         Left            =   11430
         TabIndex        =   65
         Top             =   6855
         Width           =   1845
      End
      Begin VB.Label lblSubTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubTotal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   195
         Left            =   12525
         TabIndex        =   64
         Top             =   6435
         Width           =   750
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Impostos s/ Fat."
         ForeColor       =   &H00666666&
         Height          =   195
         Left            =   4800
         TabIndex        =   63
         Top             =   6945
         Width           =   1185
      End
      Begin VB.Label lblQtdeTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   4410
         Width           =   1245
      End
      Begin VB.Label Label23 
         Caption         =   "T�cnico :"
         BeginProperty Font 
            Name            =   "WeblySleek UI Light"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   285
         Left            =   -67800
         TabIndex        =   40
         Top             =   4005
         Width           =   750
      End
      Begin VB.Label Nome_T�cnico 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   -66270
         TabIndex        =   39
         Top             =   4005
         Width           =   3810
      End
      Begin VB.Label L_Aprovado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Or�amento aprovado por :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   285
         Left            =   -74865
         TabIndex        =   38
         Top             =   4320
         Width           =   1905
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Prometido para : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   285
         Left            =   -74865
         TabIndex        =   37
         Top             =   4005
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdInsertItens 
      Caption         =   "255!"
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
      Left            =   4200
      TabIndex        =   42
      Top             =   9000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
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
      Left            =   7320
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Pre�os WHERE Tabela <> ""CUSTO"" ORDER BY Tabela"
      Top             =   8550
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Left            =   13545
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Funcion�rios"
      Top             =   8955
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   12300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Caixas"
      Top             =   8550
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   8280
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Produtos"
      Top             =   8550
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, C�digo FROM Funcion�rios WHERE Liberado = TRUE AND Ativo AND isPrestServ = FALSE ORDER BY Nome"
      Top             =   8970
      Visible         =   0   'False
      Width           =   1815
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   1650
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
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
      Left            =   9210
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cli_For"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   12150
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, C�digo FROM Funcion�rios WHERE Liberado = TRUE AND Ativo AND isPrestServ = FALSE ORDER BY Nome"
      Top             =   8970
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   11100
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Op_Sa�da"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSMask.MaskEdBox mskValidade 
      Height          =   330
      Left            =   13890
      TabIndex        =   24
      ToolTipText     =   "Pressione F2 para obter calend�rio"
      Top             =   1830
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      ForeColor       =   6710886
      Enabled         =   0   'False
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Pre�o 
      Bindings        =   "Saidas.frx":509EE
      Height          =   330
      Left            =   945
      TabIndex        =   4
      Top             =   780
      Width           =   4740
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
      BevelType       =   0
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   15066597
      ForeColorOdd    =   4210752
      BackColorOdd    =   12648447
      RowHeight       =   503
      Columns(0).Width=   3200
      Columns(0).Caption=   "Tabela"
      Columns(0).Name =   "Tabela"
      Columns(0).DataField=   "Tabela"
      Columns(0).FieldLen=   256
      _ExtentX        =   8361
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   6710886
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
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Operador 
      Bindings        =   "Saidas.frx":50A02
      Height          =   330
      Left            =   6705
      TabIndex        =   9
      Top             =   780
      Width           =   1005
      DataFieldList   =   "Nome"
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
      BevelColorHighlight=   15066597
      BevelColorFace  =   15066597
      ForeColorEven   =   0
      BackColorEven   =   15066597
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   6244
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2408
      Columns(1).Caption=   "Apelido"
      Columns(1).Name =   "Apelido"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Apelido"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2064
      Columns(2).Caption=   "C�digo"
      Columns(2).Name =   "C�digo"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "C�digo"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   256
      _ExtentX        =   1773
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   -2147483630
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Caixa 
      Bindings        =   "Saidas.frx":50A16
      DataSource      =   "Data6"
      Height          =   330
      Left            =   13080
      TabIndex        =   19
      Top             =   1470
      Width           =   795
      DataFieldList   =   "Descri��o"
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
      BevelColorHighlight=   15066597
      BevelColorFace  =   15066597
      ForeColorEven   =   0
      BackColorEven   =   14737632
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9975
      Columns(0).Caption=   "Descri��o"
      Columns(0).Name =   "Descri��o"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Descri��o"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1455
      Columns(1).Caption=   "Caixa"
      Columns(1).Name =   "Caixa"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Caixa"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1402
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   -2147483630
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboCliente 
      Bindings        =   "Saidas.frx":50A2A
      DataSource      =   "Data4"
      Height          =   330
      Left            =   6705
      TabIndex        =   5
      Top             =   0
      Width           =   1005
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
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
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
      _ExtentX        =   1773
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   0
      BackColor       =   12648384
   End
   Begin SSDataWidgets_B.SSDBCombo cboDigitador 
      Bindings        =   "Saidas.frx":50A3E
      Height          =   330
      Left            =   6705
      TabIndex        =   7
      Top             =   390
      Width           =   1005
      DataFieldList   =   "Nome"
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
      BevelColorHighlight=   15066597
      BevelColorFace  =   15066597
      ForeColorEven   =   0
      BackColorEven   =   15066597
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1773
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboOper 
      Bindings        =   "Saidas.frx":50A52
      DataSource      =   "Data2"
      Height          =   315
      Left            =   945
      TabIndex        =   2
      Top             =   390
      Width           =   915
      DataFieldList   =   "Nome"
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
      BevelColorFace  =   15066597
      ForeColorEven   =   0
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   6959
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1879
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1614
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483630
      BackColor       =   12648384
   End
   Begin MSMask.MaskEdBox mskDataEmissaoNotaManual 
      Height          =   255
      Left            =   1860
      TabIndex        =   98
      Top             =   8730
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   15066597
      ForeColor       =   6710886
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
   Begin MSFlexGridLib.MSFlexGrid gridChaves 
      Height          =   645
      Left            =   8010
      TabIndex        =   23
      Top             =   1800
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   1138
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   12648447
      BackColorSel    =   12648384
      ForeColorSel    =   -2147483641
      BackColorBkg    =   12648447
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblComanda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comanda"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   12480
      TabIndex        =   118
      Top             =   480
      Width           =   675
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00666666&
      Height          =   225
      Left            =   4440
      TabIndex        =   95
      Top             =   8580
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Descont. �tens"
      ForeColor       =   &H00666666&
      Height          =   195
      Left            =   4080
      TabIndex        =   94
      Top             =   8550
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Base ICM"
      ForeColor       =   &H00666666&
      Height          =   195
      Left            =   2550
      TabIndex        =   93
      Top             =   8895
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ICMS"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00666666&
      Height          =   225
      Left            =   4110
      TabIndex        =   92
      Top             =   8670
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Base ICMS ST"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2100
      TabIndex        =   91
      Top             =   7260
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS ST"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3120
      TabIndex        =   90
      Top             =   7260
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Frete"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   89
      Top             =   7260
      Width           =   390
   End
   Begin VB.Label lblTotServ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Servi�os"
      ForeColor       =   &H00666666&
      Height          =   195
      Left            =   4050
      TabIndex        =   88
      Top             =   8700
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblTotISS 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ISS"
      ForeColor       =   &H00666666&
      Height          =   195
      Left            =   2460
      TabIndex        =   87
      Top             =   8610
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblNrTerminal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nr Terminal"
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   60
      TabIndex        =   86
      Top             =   8400
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Impostos s/ Fat."
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00666666&
      Height          =   225
      Left            =   3210
      TabIndex        =   85
      Top             =   8760
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   84
      Top             =   7260
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblDataEmissaoNotaManual 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emiss�o NF Manual"
      ForeColor       =   &H00666666&
      Height          =   195
      Left            =   2490
      TabIndex        =   83
      Top             =   8730
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblICMSDesonerado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ICMS Deson."
      ForeColor       =   &H00666666&
      Height          =   255
      Left            =   2400
      TabIndex        =   82
      Top             =   8610
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Nota_Cancelada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nota Cancelada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   10260
      TabIndex        =   81
      Top             =   8010
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblMovPendencia 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Entregas Pendentes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   7200
      TabIndex        =   77
      Top             =   7245
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label L_Efetivada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opera��o Efetivada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   7200
      TabIndex        =   76
      Top             =   7635
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label Movimenta��o_Desfeita 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Movimenta��o Desfeita"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   7200
      TabIndex        =   74
      Top             =   8025
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sequ�ncia"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   12480
      TabIndex        =   71
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label48 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   12705
      TabIndex        =   70
      Top             =   8070
      Width           =   615
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto no SubTotal"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11745
      TabIndex        =   69
      Top             =   7785
      Width           =   1575
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   12690
      TabIndex        =   68
      Top             =   7380
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filial"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      TabIndex        =   62
      Top             =   90
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3945
      TabIndex        =   61
      Top             =   45
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opera��o"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      TabIndex        =   60
      Top             =   435
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5715
      TabIndex        =   59
      Top             =   435
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente/Forn"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5730
      TabIndex        =   58
      Top             =   45
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observa��o"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      TabIndex        =   57
      Top             =   1170
      Width           =   870
   End
   Begin VB.Label L_Dia 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4350
      TabIndex        =   1
      Top             =   0
      Width           =   1320
   End
   Begin VB.Label Nome_Opera��o 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   1875
      TabIndex        =   3
      Top             =   360
      Width           =   3795
   End
   Begin VB.Label Nome_Digitador 
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
      Left            =   7725
      TabIndex        =   8
      Top             =   390
      Width           =   4605
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   7725
      TabIndex        =   6
      Top             =   0
      Width           =   4605
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab. Pre�os"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      TabIndex        =   56
      Top             =   810
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref."
      Height          =   195
      Left            =   10005
      TabIndex        =   55
      Top             =   1170
      Width           =   315
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   54
      Top             =   7815
      Width           =   345
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   12510
      TabIndex        =   53
      Top             =   1530
      Width           =   405
   End
   Begin VB.Label Nome_Caixa 
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
      Left            =   13890
      TabIndex        =   20
      Top             =   1440
      Width           =   1320
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5715
      TabIndex        =   52
      Top             =   795
      Width           =   690
   End
   Begin VB.Label Nome_Operador 
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
      Left            =   7725
      TabIndex        =   10
      Top             =   780
      Width           =   4035
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   12510
      TabIndex        =   51
      Top             =   900
      Width           =   450
   End
   Begin VB.Label lblValidade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Validade"
      Height          =   195
      Left            =   13200
      TabIndex        =   50
      Top             =   1890
      Width           =   600
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S�rie"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   49
      Top             =   7815
      Width           =   360
   End
   Begin VB.Label lblConsumidorFinal 
      BackColor       =   &H00F7F7F7&
      BackStyle       =   0  'Transparent
      Caption         =   "Consumidor Final"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5685
      TabIndex        =   48
      Top             =   1500
      Width           =   1305
   End
   Begin VB.Label lblPresencaComprador 
      BackColor       =   &H00F7F7F7&
      BackStyle       =   0  'Transparent
      Caption         =   "Presencial?"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7845
      TabIndex        =   47
      ToolTipText     =   "Indicador de Presen�a do Comprador"
      Top             =   1500
      Width           =   870
   End
   Begin VB.Label lblFinalidade 
      BackStyle       =   0  'Transparent
      Caption         =   "Finalidade"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   46
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lblChave 
      BackStyle       =   0  'Transparent
      Caption         =   "Chave"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   45
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Al�q. Inter"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3120
      TabIndex        =   44
      ToolTipText     =   "Al�quota Interestadual"
      Top             =   7815
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Al�q. Destino"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2100
      TabIndex        =   43
      ToolTipText     =   "Al�quota UF Destino"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Nome_Filial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   945
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   2190
      Top             =   8760
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
      Bands           =   "Saidas.frx":50A66
   End
End
Attribute VB_Name = "frmSaidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type Tabela
  C�digo As String
  Nome As String
  Unidade As String
  Pre�o_Total As Currency
  Pre�o_Final As Currency
  Nada As String
  Informa As String
  Qtde As Single
  Desp_Acessorias As Double
  
  '04/05/2004 - Daniel
  'Case Embalavi formata��o de 5 casas ap�s
  'a "," quando Currency permitia apenas 4 casas
  'Pre�o As Currency
  Pre�o As Double
  PesoLiquido As Double
  PesoBruto As Double
  Desconto As Currency
  IPI As Double
  ICM As Double
  Base_ICM As Double
  Valor_ICM As Double
  Valor_Base_Unit As Double
  Redu��o_ICM As Double
  Tipo_ICM As String
  Etiqueta As String
  Descr_Adicional As String
  CFOP_Produto As String '19/12/2006 - Anderson - Registro de CFOP por Produto
  Valor_Desonerado As Double
  Total_Valor_Desonerado As Double
  Percentual_Diferimento As Double
  
End Type
Dim Tabe(255) As Tabela

Private Type Tabela_Serv
  C�digo    As Double
  Descri��o As String
  Tempo     As String
  Pre�o     As Double
  Total     As Double
  Completo  As Integer
  ISS       As Single
  '26/07/2005 - Daniel
  'C.S.T. (C�digo de Situa��o Tribut�ria)
  'Refere-se �s personaliza��es para a empresa J.R. Hidroqu�mica
  CST      As String
  CFOP_Servico As String '19/12/2006 - Anderson - Registro de CFOP por Servi�o
End Type
Dim Tabe_Serv(255) As Tabela_Serv

Dim Num_Registro As Variant

Dim Tamanho As Integer
Dim Cor As Integer
Dim Edicao As Long
Dim Tipo As Integer
Dim Erro As Integer

Dim sOPERACAO_APPQuick01 As String
Dim sOPERACAO_APPQuick02 As String
Dim rsVerificaOperacaoERP_APP As Recordset
Dim rsProdutos2 As Recordset
Dim rsServi�os As Recordset
Dim rsParametros As Recordset
Dim rsOp_Sa�da As Recordset
Dim rsFuncionarios As Recordset
Dim rsT�cnicos As Recordset
Dim rsCliFor As Recordset
Dim rsPre�os As Recordset
Dim rsGrade As Recordset
Dim rsSaidas As Recordset
Dim rsSaidas_Prod As Recordset
Dim rsSaidas_Serv As Recordset
Dim rsMovi_Parcelas As Recordset
Dim rsSa�da_Parcelas As Recordset
Dim rsSa�da_Cheques As Recordset
Dim rsUsu�rios As Recordset
Dim rsTabelas As Recordset
Dim rsCota��es As Recordset
Dim rsContas_Receber As Recordset
Dim rsEstados As Recordset
Dim rsCaixas As Recordset
Dim rsOperadores As Recordset
Dim rsLog As Recordset
'20/12/2006 - Anderson - Altera��o realizada para o registro do CFOP por produto e servi�o
Dim rsProdutoCFOP As Recordset
Dim rsServicoCFOP As Recordset
'04/12/2007 - Anderson
'Verifica se deve somar produtos ao total da nota
Dim blnSomarProdutosTotalNota As Boolean

'11/11/2008 - mpdea
'Verifica se deve somar o ICMS Retido ao total da nota
Dim m_blnSomaIcmsRetidoTotalNota As Boolean

Private gsSql As String
Private gsWhere As String
Private gsOrder As String

Private gbBaseICMSomadoIPI As Boolean

Dim Total_Pagar As Double

Dim Total_Desconto As Double
Private gcDescInTotal As Currency

Dim Total_Servi�os As Double
Dim Desconto_Cli As Double
Dim Erro_Data As Integer
Dim gbLogError As Boolean
Dim Calcula_ICM As Integer
Dim Calcula_IPI As Integer
Dim Linhas_Grade As Integer
Dim Linhas_Servi�o As Integer
Dim Alterar_Servi�os As Integer
Dim sSql As String
Dim Estado As String
Dim Calcula_IPI_TOT As Integer
Dim PercIcmsFrete As Integer
Dim Calcula_ICM_Frete As Boolean
Dim Soma_Frete As Boolean
'12/04/2005 - Daniel
'Tratamento para a soma do seguro
'ao total a receber
Dim Soma_Seguro As Boolean

'20/09/2002 - mpdea
'Desconto no SubTotal
Private mcurDescontoSubTotal As Currency
'Flag para for�ar a atualiza��o do registro
Private mblnForceUpdate      As Boolean

'30/04/2003 - mpdea
'Desconto rateado
Private m_blnDescontoRateado As Boolean

'01/10/2002 - mpdea
'Flag para indicar exibi��o do registro
Private mblnInShowRecord     As Boolean

'27/02/2004 - Daniel
'Flag de indica��o que � o Cliente PSV
Private m_blnPSV             As Boolean

'04/05/2004 - Daniel
'Flag de indica��o que � o Cliente Embalavi
'realizar� a��es personalizadas para este Cliente
Private m_blnEmbalavi As Boolean

'26/07/2005 - Daniel
'Flag de identifica��o para �s personaliza��es da empresa
'W.V. Hidroan�lise Ltda (J.R. Hidroqu�mica)
Private m_blnJR As Boolean

'13/05/2004 - Daniel
'Vars de tratamento de Percentuais e Totais
'de impostos sobre servi�os
Private m_sngPercentualCSLL   As Single
Private m_sngPercentualCOFINS As Single
Private m_sngPercentualPIS    As Single
Private m_sngPercentualIRRF   As Single
Private m_dblTotalCSLL        As Double
Private m_dblTotalCOFINS      As Double
Private m_dblTotalPIS         As Double
Private m_dblTotalIRRF        As Double
Private m_dblTotalMenosServ   As Double
Private m_dblTotaldeImpostos  As Double

'01/07/2004 - Daniel
'Var para tratamento da limpeza de campos
'conforme perfil do user
'Case: Coneg Campos
Private m_blnClear As Boolean

'26/08/2004 - Daniel
'Criado valida��o para verificar se o usu�rio possui permiss�o
'para enchergar o estoque ou n�o
Private m_blnPermitido As Boolean

'17/09/2004 - Daniel
'Mostrar o valor real em L_Tot_Pagar
'Case: Nilvo Burin
Dim m_blnNilvo   As Boolean
Dim m_blnNotZero As Boolean
Dim m_dblTotal   As Double

'28/10/2004 - Daniel
'Tratamento para o field [Sa�das - Produtos].[Pre�o Final]
'Para o cliente A.S. Wijma (Bel�m - Par�) dever� ser Double
'para os demais clientes continua sendo Single
Dim m_dblPrecoFinalAuxi As Double
Dim m_blnASWijmaBelem   As Boolean

'09/11/2004 - Daniel
'Esta var � utilizada para identificar o cliente
'Teknika que possui tratamento especial para o ICMS
Dim m_blnTeknika As Boolean
'06/05/2005 - Daniel
'
'Implementa��o.: Trabalhar com o c�digo para fornecedor cadastrado na tela de produtos.
'                Impacto: Ao entrar com o c�digo para o fornecedor no campo c�digo do produto
'                o sistema dever� trazer o c�digo do produto que estiver amarrado nele
'Solicita��o...: Cristiano Pavinato - PSI RS
Private m_blnUsaCodFornec As Boolean
'12/05/2005 - Daniel
'
'Solicitante..: Info Social
'
'Finalidade...: Deixamos configur�vel em Par�metros � exibi��o
'               nas telas de Sa�da e Venda R�pida da coluna Fabricante
'               nos dropdowns de pesquisas
Private m_blnExibirColunaFabricante As Boolean

'23/05/2006 - mpdea
'Otimizado a verifica��o do cliente isento em IPI
Private m_blnIsentoIPI As Boolean

'19/10/2007 - Anderson
'Implementa��o do campo Lucro M�nimo Permitido conforme solicita��o da Agrotama
Private m_bolLucroMinimoPermitido As Boolean

'14/12/2007 - Celso
'Utilizada para armazenar o cliente para o qual j� tenha sido solicitada senha
'do gerente qdo o mesmo tiver contas em atraso
Private m_blnSenhaGerJaInformada As Boolean
Private m_strCodigoClienteContas As String
  
'16/10/2009 - mpdea
'Modo de entrada de dados no grid de produtos
Private m_int_modo_grid_produtos As Integer

'10/12/2009 - Andrea
Dim rsSa�da_Cartoes As Recordset

Dim Total_Valor_Desonerado As Double

'Pilatti var
Dim sOrigemToolBarMoveRegistros As String

' Variaveis A3Manager APP - Pilatti Mar�o/18
Public dbA3Manager As New ADODB.Connection

Dim totalNCM_2 As Double    'Total em R$ de imposto pago na movimenta��o

Dim sCodigoProdutoDevolucao As String
Dim sNomeProdutoDevolucao As String
Dim lQuantidadeItensProdutoDevolucao As Long
Dim sValorUnitarioProdutoDevolucao As String

Dim rsEstadosICMS As Recordset
Dim aliquotaICMS_tab_ICMS_PERC_ESTADOS As String
Dim bo_AliquotaICMS_interestadual As Boolean

Dim sTipoOperacaoSaida As String

Private bProdutoSemPrecoNaGrade As Boolean

Dim l_tamanhoOriginal_TAB1 As Long
Dim l_tamanhoOriginal_GRADE1 As Long
Dim l_tamanhoOriginal_GRADE1_Grupo1Produtos As Long
Dim l_txtSeq As Long
Dim l_Label52 As Long
Dim l_txtComanda As Long
Dim l_lblComanda As Long
Dim l_Senha As Long
Dim l_Label26 As Long
Dim l_Nome_Caixa As Long
Dim l_Combo_Caixa As Long
Dim l_Label22 As Long
Dim l_mskValidade As Long
Dim l_lblValidade As Long
Dim l_cmd_tabelaDePrecos As Long
Dim l_txtSubTotal As Long
Dim l_Label35 As Long
Dim l_txtDescSubTotal As Long
Dim l_Label36 As Long
Dim l_Label48 As Long
Dim l_L_Tot_Pagar As Long
Dim l_tamanhoOriginal_Grade_Serv As Long
Dim l_tamanhoOriginal_Grade_Serv_GrupoServicos As Long
Dim l_B_Servi�os_Conc As Long
Dim l_Nome_Cliente_Estica As Long
Dim l_Nome_Digitador_Estica As Long
Dim l_txtRef_Estica As Long
Dim l_cboPresencaComprador_Estica As Long

'Cancela desconto -- PABLO 07/07/2022
Private b_EscondeTelaDesconto As Boolean

Private Sub EmiteCarnesNOVOS()
On Error GoTo Erro:
  Dim Resp As String

  Resp = InputBox("Impriss�o em modelo:" & vbCrLf & vbCrLf & "     1 - TICKET         [40 colunas]" & vbCrLf & vbCrLf & "     2 - RELAT�RIO [110 colunas]", "Qual o modelo de impress�o?", "1")
  If Not IsNumeric(Resp) Then
      DisplayMsg "Op��o de impress�o inv�lida!"
      Exit Sub
  Else
      If Resp <> "1" And Resp <> "2" Then
          DisplayMsg "Op��o de impress�o inv�lida!"
          Exit Sub
      End If
  End If

  Dim strNomeArq As String

  If Resp = "2" Then
      Rel1.Destination = 0
      strNomeArq = gsReportPath & "carne02.rpt"
  Else
      Rel1.WindowShowPrintSetupBtn = True
      Rel1.WindowState = crptMaximized
      Rel1.Destination = IIf(False, crptToWindow, crptToPrinter)
      
      strNomeArq = gsReportPath & "carne02_todasParcelas_46Colunas.rpt"
  End If
  
  If Dir(strNomeArq) = "" Then
    DisplayMsg "Arquivo """ & strNomeArq & """ n�o encontrado."
    Exit Sub
  End If
    
  Rel1.DataFiles(0) = gsQuickDBFileName
  Rel1.ReportFileName = strNomeArq
  Rel1.ParameterFields(0) = "pSequencia;" & rsSaidas("Sequ�ncia") & ";true"
  Rel1.ParameterFields(6) = "pFilial;" & gnCodFilial & ";true"
  
  Dim sEmpresaNome As String
  Dim sEmpresaRuaComNumero As String
  Dim sEmpresaCidadeEstado As String
  Dim sEmpresaFone As String
  Dim sEmpresaCep As String
  
  If Len(gsNomeFilial) > 30 Then
      sEmpresaNome = Mid(gsNomeFilial, 1, 30)
  Else
      sEmpresaNome = gsNomeFilial
  End If
  
  If Len(gsFilialEndereco) > 30 Then
      sEmpresaRuaComNumero = Mid(gsFilialEndereco, 1, 30)
  Else
      sEmpresaRuaComNumero = gsFilialEndereco
  End If

  If Len(gsFilialCidadeEstado) > 30 Then
      sEmpresaCidadeEstado = Mid(gsFilialCidadeEstado, 1, 30)
  Else
      sEmpresaCidadeEstado = gsFilialCidadeEstado
  End If

  If Len(gsFilialFone) > 14 Then
      sEmpresaFone = Mid(gsFilialFone, 1, 14)
  Else
      sEmpresaFone = gsFilialFone
  End If

  If Len(gsFilialCep) > 11 Then
      gsFilialCep = Mid(gsFilialFone, 1, 11)
  Else
      sEmpresaCep = gsFilialCep
  End If
  
  Rel1.ParameterFields(1) = "pEmpresa;" & sEmpresaNome & ";true"
  Rel1.ParameterFields(2) = "pEmpresaEnderecoRua;" & sEmpresaRuaComNumero & ";true"
  Rel1.ParameterFields(3) = "pEmpresaEnderecoCidadeEstado;" & sEmpresaCidadeEstado & ";true"
  Rel1.ParameterFields(4) = "pEmpresaEnderecoFone;" & sEmpresaFone & ";true"
  Rel1.ParameterFields(5) = "pEmpresaEnderecoCep;" & "Cep " & sEmpresaCep & ";true"
  Rel1.WindowState = crptMaximized
  
  Call SetPrinterName("REL", Rel1)

  Rel1.Action = 1
  
  Rel1.Reset

  Exit Sub
Erro:
  MsgBox "Erro tentando gerar Carn�s. Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Function Acha_Grade(Aux As Double) As Double
  rsProdutos2.FindFirst "C�digo = '" & Aux & "'"
  If Not rsProdutos2.NoMatch Then
    Acha_Grade = Aux
  Else
    rsGrade.Index = "C�digo"
    rsGrade.Seek "=", Aux
    If rsGrade.NoMatch Then
      Acha_Grade = Aux
    Else
      Acha_Grade = rsGrade("C�digo Original")
    End If
  End If
End Function

Public Sub Calcula_Linha()
  'Calcula pre�o total e final da linha
  Dim nX As Integer
  Dim Qtde As Double
  Dim Pre�o As Double
  Dim Desconto As Double
  Dim Valor_Desconto As Double
  Dim IPI As Double
  Dim Valor_IPI As Double
  Dim Pre�o_Total As Double
  Dim Pre�o_Final As Double
  Dim Pre�o_Final2 As Double
  Dim Desp_Acessorias As Double
  Dim Valor_Desonerado As Double
  
  With Grade1
    For nX = 1 To 8
      Select Case nX
        Case 1, 4, 6, 7, 8
          If .Columns(nX).Text = "" Then
            .Columns(nX).Text = 0
          End If
      End Select
    Next nX
    
    If .Columns("Desp_Acessorias").Text = "" Then
    .Columns("Desp_Acessorias").Text = 0
    End If
    
    If .Columns("Valor Desonerado").Text = "" Then
      .Columns("Valor Desonerado").Text = 0#
    End If
    
    Desp_Acessorias = Format((.Columns("Desp_Acessorias").Text), "#0.00")
    Valor_Desonerado = Format((.Columns("Valor Desonerado").Text), "#0.00")
    Qtde = .Columns("Qtde").Text
    '04/05/2004 - Daniel
    'Personaliza��o Embalavi
    If g_bln5CasasDecimais Then
      Pre�o = Format((.Columns("Pre�o Unit.").Text), "##,###,##0.00000")
    '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      Pre�o = Format((.Columns("Pre�o Unit.").Text), "##,###,##0.000")
    Else
      'Pre�o = .Columns("Pre�o Unit.").Text
      Pre�o = Format((.Columns("Pre�o Unit.").Text), "##,###,##0.00")
    End If
    Desconto = .Columns("Desconto").Text

    ' ==============================================================
    ' Tratar IPI
    If rsParametros("CodigoRegimeTributario") = 1 Then
        If rsOp_Sa�da.Fields("tipo").Value = "G" Then 'cboFinalidade.ListIndex = 3 Then
            'Finalidade devolu��o
            IPI = .Columns("IPI").Text
        Else
            .Columns("IPI").Text = "0"
            IPI = "0"
        End If
    Else
        IPI = .Columns("IPI").Text
    End If

'''    If Not IsNull(rsProdutos2("IPI_ValidoEntradaSaida").Value) And rsProdutos2("IPI_ValidoEntradaSaida").Value = 1 Then
'''        IPI = .Columns("IPI").Text
'''    ElseIf Not IsNull(rsProdutos2("IPI_ValidoEntradaSaida").Value) And rsProdutos2("IPI_ValidoEntradaSaida").Value <> 1 Then
'''        If cboFinalidade.ListIndex = 3 Then
'''            'Finalidade devolu��o
'''            IPI = .Columns("IPI").Text
'''        Else
'''            .Columns("IPI").Text = "0"
'''        End If
'''    Else
'''        .Columns("IPI").Text = "0"
'''    End If
    ' ==============================================================
    
    Pre�o_Total = Format(Qtde * Pre�o, "#0.00")
    .Columns("Total").Text = Pre�o_Total
    
    Valor_Desconto = Format(Pre�o_Total * Desconto / 100#, "#0.00")
    Pre�o_Final = Format((Pre�o_Total - Valor_Desconto + Desp_Acessorias), "#0.00")
    Valor_IPI = Format(Truncate(Pre�o_Final * IPI / 100, 2), "#0.00")
    
    '23/05/2006 - mpdea
    'Adicionado verifica��o de cliente isento de IPI
    If Not Calcula_IPI Or m_blnIsentoIPI Then
      Valor_IPI = 0
    End If
    
    If Calcula_IPI_TOT Then
      Pre�o_Final2 = Format((Pre�o_Final), "#0.00")
      .Columns("Pre�o Final").Text = Pre�o_Final2
    Else
      Pre�o_Final2 = Format((Pre�o_Final + Valor_IPI), "#0.00")
      .Columns("Pre�o Final").Text = Pre�o_Final2
    End If
    'Calculo do ICM
    If .Columns("Tipo_ICM").Text = "N" Then
      'ICM Normal
      If gbBaseICMSomadoIPI = True Then
        .Columns("Base_ICM").Text = Pre�o_Final2
        .Columns("Valor_ICM").Text = Pre�o_Final2 * CSng(gsHandleNull(.Columns("ICM").Text & "")) / 100
      Else
        .Columns("Base_ICM").Text = Pre�o_Final
        .Columns("Valor_ICM").Text = Pre�o_Final * CSng(gsHandleNull(.Columns("ICM").Text & "")) / 100
      End If
    ElseIf .Columns("Tipo_ICM").Text = "R" Then
      'ICM Retido
      If CDbl(.Columns("Valor_Base_Unit").Text) <> 0 Then
        'Base Fixa
        .Columns("Base_ICM").Text = CDbl(.Columns("Qtde").Text) * CDbl(.Columns("Valor_Base_Unit").Text)
        .Columns("Valor_ICM").Text = CDbl(.Columns("Base_ICM").Text) * CDbl(.Columns("ICM").Text) / 100
      End If
      If CDbl(.Columns("Redu��o_ICM").Text) <> 0 Then
        'Base Reduzida
        .Columns("Base_ICM").Text = Pre�o_Final * CDbl(.Columns("Redu��o_ICM").Text) / 100
        .Columns("Valor_ICM").Text = CDbl(.Columns("Base_ICM").Text) * CDbl(.Columns("ICM").Text) / 100
      End If
    ElseIf .Columns("Tipo_ICM").Text = "Z" Then
      'ICM Reduzido
      If CDbl(.Columns("Valor_Base_Unit").Text) <> 0 Then
        'Base Fixa
        .Columns("Base_ICM").Text = CDbl(.Columns("Qtde").Text) * CDbl(.Columns("Valor_Base_Unit").Text)
        .Columns("Valor_ICM").Text = CDbl(.Columns("Base_ICM").Text) * CDbl(.Columns("ICM").Text) / 100
      End If
      If CDbl(.Columns("Redu��o_ICM").Text) <> 0 Then
        'Base Reduzida
        .Columns("Base_ICM").Text = Pre�o_Final * CDbl(.Columns("Redu��o_ICM").Text) / 100
        .Columns("Valor_ICM").Text = CDbl(.Columns("Base_ICM").Text) * CDbl(.Columns("ICM").Text) / 100
      End If
    End If
  End With
End Sub

Sub ShowRecord()
  Dim i             As Integer
  Dim l             As Integer
  Dim Linha         As Integer
  Dim Aux           As String
  Dim C�d           As String
  Dim Aux_Produto   As String
  Dim Aux_Tamanho   As Integer
  Dim Aux_Cor       As Integer
  Dim Aux_Edi��o    As Long
  Dim Aux_Tipo      As Integer
  Dim Aux_Erro      As Integer
  Dim nSeqSaidas    As Long
  Dim Tool          As ActiveBarLibraryCtl.Tool
  Dim sngQtdeTotal  As Single
  Dim sSQLChaves    As String
  '17/05/2013-Alexandre Afornali
  'Case DiskEmbalagens
  txtComanda.Text = ""
  
  gridChaves.Rows = 1
  
  L_Tot_ICMS_Deson.Text = "0#"
  '22/10/2004 - Daniel
  'Esta var armazenar� o valor real do Total da Sa�da
  'BUG: Ao efetuar o altera total no momento da venda,
  'quando mostr�vamos o valor total depois de passar pelo
  'recalcula estava mostrando o valor do recalcula da grid
  'e n�o o valor real da venda
  Dim dblTotal      As Double
  
  '30/04/2003 - mpdea
  'Zera o total de desconto concedido
  Total_Desconto = 0
  gcDescInTotal = 0
  
  '01/10/2002 - mpdea
  'In�cio da exibi��o do registro
  mblnInShowRecord = True
  
  Err.Clear
  
  '28/10/2002 - mpdea
  'Corrigido declara��o para processamento de erro
  On Error GoTo ErrHandler
  
  'Na mudan�a de registro o Altera Totais � desmarcado
  Set Tool = ActiveBar1.Tools("miComplAlteraTotais")
  If Tool.Checked Then
    Call ActiveBar1_Click(Tool)
  End If
  
  L_Dia.Caption = Format(rsSaidas("Data"), "dd/mm/yyyy")
  
  '20/05/2005 - Daniel
  '
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '
  'O sistema dever� reconhecer se a nota fiscal foi criada
  'automaticamente ou manualmente a partir da opera��o escolhida
  If gbNotaManual(rsSaidas("Opera��o").Value, "SAIDA") Then
    '---[ Manual ]---
    If Not IsNull(rsSaidas("Nota Fiscal")) Then
      txtNF.Text = rsSaidas("Nota Fiscal") & ""
    Else
      txtNF.Text = "0"
    End If
    
    '01/08/2005 - Daniel
    'Tratamento para o Campo Sa�das.DataEmissaoNotaManual
    lblDataEmissaoNotaManual.Visible = True
    mskDataEmissaoNotaManual.Visible = True
    '
    If Not IsNull(rsSaidas("DataEmissaoNotaManual").Value) Then mskDataEmissaoNotaManual.Text = rsSaidas("DataEmissaoNotaManual").Value & ""
    
  Else
    '---[ Nota gerada Autom�tica ]---
    '21/02/2005 - Daniel
    'Tratamento para a exibi��o do campo Nota Impressa
    'estava ocorrendo o erro 94 Invalid of null
    If Not IsNull(rsSaidas("Nota Impressa")) Then
      txtNF.Text = rsSaidas("Nota Impressa") & ""
    Else
      txtNF.Text = "0"
    End If
  
    '01/08/2005 - Daniel
    'Tratamento para o Campo Sa�das.DataEmissaoNotaManual
    lblDataEmissaoNotaManual.Visible = False
    mskDataEmissaoNotaManual.Visible = False
  End If
  
  
  '20/01/2003 - Maikel
  '             Trocadas as linhas comentadas abaixo por uma linha s� que vem a seguir
'  L_Efetivada.Visible = False
'  If rsSaidas("Efetivada") Then
'    L_Efetivada.Visible = True
'  End If
  L_Efetivada.Visible = rsSaidas.Fields("Efetivada").Value
  
  
  '20/01/2003 - Maikel
  '             Trocadas as linhas comentadas abaixo por uma linha s� que vem a seguir
'  Movimenta��o_Desfeita.Visible = False
'  If rsSaidas("Movimenta��o Desfeita") = True Then Movimenta��o_Desfeita.Visible = True
  Movimenta��o_Desfeita.Visible = rsSaidas.Fields("Movimenta��o Desfeita").Value
  
  '20/01/2003 - Maikel
  '             Trocadas as linhas comentadas abaixo por uma linha s� que vem a seguir
'  Nota_Cancelada.Visible = False
'  If rsSaidas("Nota Cancelada") Then
'    Nota_Cancelada.Visible = True
'  End If
  Nota_Cancelada.Visible = rsSaidas.Fields("Nota Cancelada").Value
  
'''  ' Junho 2019
'''  If Movimenta��o_Desfeita.Visible = False And L_Efetivada.Visible = False Then
'''      Grade1.Enabled = True
'''  End If
  
'''  Senha.Text = gSenhaUsuarioLogado
  Senha.Text = ""
  Combo_Operador.Text = rsSaidas("Operador") & ""
'''  Combo_Operador.Text = gnUserCode & ""
  Combo_Operador_LostFocus

  cboOper.Text = rsSaidas("Opera��o")
  cboOper_LostFocus
  Total_Valor_Desonerado = 0#
  L_Tot_ICMS_Deson.Text = Total_Valor_Desonerado
  '19/02/2004 - Daniel
  'Case: PSV
  If m_blnPSV Then
  
    If Not IsNull(rsSaidas.Fields("Data Validade").Value) Then
      mskValidade.Text = rsSaidas.Fields("Data Validade").Value
    End If
      
  End If
  '----------------------------------------------------------
  
  cboDigitador.Text = rsSaidas("Digitador")
  cboDigitador_LostFocus
  
  cboCliente.Text = rsSaidas("Cliente")
  cboCliente_LostFocus
  
  Combo_Pre�o.Text = rsSaidas("Tabela") & ""
  
  txtRef.Text = rsSaidas("Refer�ncia") & ""
  
  txtSeq.Text = rsSaidas("Sequ�ncia")
  
  Obs.Text = rsSaidas("Observa��es") & ""
  
  Total_Pagar = rsSaidas("Total")
  
  Select Case rsSaidas("Consumidor_Final").Value
    Case "1"
      cboConsumidorFinal.Text = "1=Sim"
    Case Else
      cboConsumidorFinal.Text = "0=N�o"
  End Select
  Select Case rsSaidas("Presenca_Comprador").Value
    Case "0"
      cboPresencaComprador.Text = "0=N�o se aplica"
    Case "1"
      cboPresencaComprador.Text = "1 =Opera��o presencial"
    Case "2"
      cboPresencaComprador.Text = "2=Opera��o n�o presencial, pela Internet"
    Case "3"
      cboPresencaComprador.Text = "3=Opera��o n�o presencial, Teleatendimento"
    Case "4"
      cboPresencaComprador.Text = "4=NFC-e em opera��o com entrega em domic�lio"
    Case "9"
      cboPresencaComprador.Text = "9=Opera��o n�o presencial, outros"
    Case Else
      cboPresencaComprador.Text = "1 =Opera��o presencial"
  End Select
  Select Case rsSaidas("FinalidadeNFe").Value
    Case "1"
      cboFinalidade.ListIndex = 0
    Case "2"
      cboFinalidade.ListIndex = 1
    Case "3"
      cboFinalidade.ListIndex = 2
    Case "4"
      cboFinalidade.ListIndex = 3
    Case Else
      cboFinalidade.ListIndex = 0
  End Select
'    Case "1"
'      cboFinalidade.Text = "1=NFe normal"
'    Case "2"
'      cboFinalidade.Text = "2=NF-e complementar"
'    Case "3"
'      cboFinalidade.Text = "3=NF-e de ajuste"
'    Case "4"
'      cboFinalidade.Text = "4=Devolu��o de mercadoria"
'    Case Else
'      cboFinalidade.Text = "1=NFe normal"
'  End Select

  '==========================================================
  Dim rsChaves As Recordset
  Dim nContaChaves As Integer
  sSQLChaves = "Select * from SaidasChaves Where Filial = " & gnCodFilial & " and Sequencia = " & rsSaidas("Sequ�ncia")
  Set rsChaves = db.OpenRecordset(sSQLChaves, dbOpenSnapshot)
  If Not (rsChaves.BOF And rsChaves.EOF) Then
      rsChaves.MoveLast
      rsChaves.MoveFirst
      For nContaChaves = 0 To rsChaves.RecordCount - 1
        gridChaves.AddItem vbTab & rsChaves.Fields("Chave").Value

        rsChaves.MoveNext
      Next
  End If
  rsChaves.Close
  Set rsChaves = Nothing
  '==========================================================
  
  If Len(Trim(rsSaidas.Fields("ComentariosSobreOrcamento"))) > 0 Then
    If MsgBox(rsSaidas.Fields("ComentariosSobreOrcamento") & vbCrLf & vbCrLf & "Deseja apagar o coment�rio ?", vbQuestion + vbYesNo, "Quick Store") = vbYes Then
      If Combo_Operador.Text = gnUserCode Then
        rsSaidas.Edit
        rsSaidas.Fields("ComentariosSobreOrcamento").Value = ""
        rsSaidas.Update
      Else
        MsgBox "Voc� n�o � o propriet�rio do or�amento, portanto n�o tem permiss�o para apagar o coment�rio !", vbInformation, "Quick Store"
      End If
    End If
  End If
  
  '08/11/2002 - mpdea
  'Verifica��o de nulo
  '20/09/2002 - mpdea
  'Desconto no SubTotal
  mcurDescontoSubTotal = 0
  Call IsDataType(dtCurrency, rsSaidas.Fields("DescontoSubTotal").Value, mcurDescontoSubTotal)

  Prometido_Para.Text = rsSaidas("Prometido Para") & ""
  Or�amento_Aprovado.Text = rsSaidas("Or�amento Aprovado") & ""
  
  If IsDate(rsSaidas("Data Acerto Empr�stimo")) Then
    Data_Acerto.Text = Format(rsSaidas("Data Acerto Empr�stimo"), "dd/mm/yyyy")
  End If
  
  nSeqSaidas = rsSaidas("Sequ�ncia")
  
  '14/08/2002 - mpdea
  'Exibi��o do nr. do or�amento
  If rsSaidas.Fields("InfoNrOrcamento").Value & "" <> "" Then
    Me.Caption = "Sa�das - " & rsSaidas.Fields("InfoNrOrcamento").Value
  Else
    Me.Caption = "Sa�das"
  End If
  
  sngQtdeTotal = 0
  
  i = 0
  gnPesoLiquido = 0#
  gnPesoBruto = 0#
  rsSaidas_Prod.Index = "Sequ�ncia"
LP_Prox:
  rsSaidas_Prod.Seek ">", gnCodFilial, rsSaidas("Sequ�ncia"), Linha
  If rsSaidas_Prod.NoMatch Then GoTo Fim_Mostra
  If rsSaidas_Prod("Filial") <> gnCodFilial Then GoTo Fim_Mostra
  If rsSaidas_Prod("Sequ�ncia") <> rsSaidas("Sequ�ncia") Then GoTo Fim_Mostra
  
  Linha = rsSaidas_Prod("Linha")
  
  Tabe(i).C�digo = rsSaidas_Prod("C�digo")
   
  'Acha Produto
   Aux = rsSaidas_Prod("C�digo")
   C�d = Aux
   
   Call Acha_Produto(C�d, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro)
   C�d = Aux_Produto
   
   If Aux_Erro = 0 Then
      rsProdutos2.FindFirst "C�digo = '" & Aux_Produto & "'"
      Tabe(i).Nome = rsProdutos2("Nome")
      Tabe(i).Unidade = rsProdutos2("Unidade Venda") & ""
      Tabe(i).PesoLiquido = gsHandleNull(rsProdutos2("PesoLiquido"))
      Tabe(i).PesoBruto = gsHandleNull(rsProdutos2("PesoBruto"))
      Tabe(i).Tipo_ICM = gsHandleNull(rsProdutos2("Tipo ICM"))
      Tabe(i).Redu��o_ICM = gsHandleNull(rsProdutos2("Redu��o ICM"))
      Tabe(i).Valor_Base_Unit = gsHandleNull(rsProdutos2("Base C�lculo"))
   Else
      Tabe(i).Nome = " ???"
      Tabe(i).Unidade = ""
      Tabe(i).PesoLiquido = 0
      Tabe(i).PesoBruto = 0
      Tabe(i).Tipo_ICM = ""
      Tabe(i).Redu��o_ICM = 0
      Tabe(i).Valor_Base_Unit = 0
   End If
   
  '13-04-2025 pablo
  If rsParametros("EditarNomeProduto").Value Then
    Dim QUERY As String
    QUERY = "SELECT Nome FROM ProdutoNomeNFe WHERE "
    QUERY = QUERY & "Filial = " & gnCodFilial & " AND "
    QUERY = QUERY & "Sequencia = " & rsSaidas("Sequ�ncia") & " AND "
    QUERY = QUERY & "Codigo = '" & Aux_Produto & "';"
    
    Dim rsNomeProd As Recordset
    Set rsNomeProd = db.OpenRecordset(QUERY, dbOpenSnapshot)
    
    If Not (rsNomeProd.BOF And rsNomeProd.EOF) Then
      rsNomeProd.MoveLast
      rsNomeProd.MoveFirst
      If rsNomeProd.RecordCount = 1 Then
        Tabe(i).Nome = Trim(CStr(rsNomeProd.Fields("Nome").Value))
      End If
    End If
    
    rsNomeProd.Close
    Set rsNomeProd = Nothing
  End If
      
   Tabe(i).Pre�o_Total = rsSaidas_Prod("Pre�o") * rsSaidas_Prod("Qtde")
   Tabe(i).Pre�o_Final = rsSaidas_Prod("Pre�o Final")
   Tabe(i).Nada = ""
   Tabe(i).Informa = ""
   Tabe(i).Qtde = rsSaidas_Prod("Qtde")
   'Total_Valor_Desonerado = Total_Valor_Desonerado + Tabe(i).Valor_Desonerado
   
   If IsNull(rsSaidas_Prod("Desp_Acessorias")) Then
      Tabe(i).Desp_Acessorias = 0
   Else
      Tabe(i).Desp_Acessorias = rsSaidas_Prod("Desp_Acessorias")
   End If
   
   sngQtdeTotal = sngQtdeTotal + rsSaidas_Prod("Qtde")
   
   gnPesoLiquido = gnPesoLiquido + Tabe(i).PesoLiquido * Tabe(i).Qtde
   gnPesoBruto = gnPesoBruto + Tabe(i).PesoBruto * Tabe(i).Qtde
   '04/05/2004 - Daniel
   'Personaliza��o Embalavi
   If g_bln5CasasDecimais Then
    Tabe(i).Pre�o = Format((rsSaidas_Prod("Pre�o")), "##,###,##0.00000")
   '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
   ElseIf g_bln3CasasDecimais Then
    Tabe(i).Pre�o = Format((rsSaidas_Prod("Pre�o")), "##,###,##0.000")
   Else
    'Tabe(i).Pre�o = rsSaidas_Prod("Pre�o")
    Tabe(i).Pre�o = Format((rsSaidas_Prod("Pre�o")), "##,###,##0.00")
   End If

   Tabe(i).Desconto = rsSaidas_Prod("Desconto")
   Tabe(i).ICM = rsSaidas_Prod("ICM")
   Tabe(i).IPI = rsSaidas_Prod("IPI")
   Tabe(i).CFOP_Produto = rsSaidas_Prod("CFOP") & "" '20/12/2006 - Anderson - Altera��o para registro de CFOP por produto
   Tabe(i).Etiqueta = rsSaidas_Prod("Etiqueta")
   Tabe(i).Descr_Adicional = rsSaidas_Prod("Descricao Adicional") & ""
   If IsNull(rsSaidas_Prod("ValorICMSDesonerado").Value) Then
    Tabe(i).Valor_Desonerado = 0
   Else
    Tabe(i).Valor_Desonerado = rsSaidas_Prod("ValorICMSDesonerado").Value
   End If
   If IsNull(rsSaidas_Prod("Percentual_Diferimento").Value) Then
    Tabe(i).Percentual_Diferimento = 0
   Else
   Tabe(i).Percentual_Diferimento = rsSaidas_Prod("Percentual_Diferimento").Value
   End If
   lblMovPendencia.Visible = False
   If rsSaidas_Prod("Qtde") > rsSaidas_Prod("QtdeEntregue") Then
     Dim rsOperacao As Recordset
     Set rsOperacao = db.OpenRecordset("SELECT * FROM [Opera��es Sa�da] WHERE C�digo = " & cboOper.Text, dbOpenSnapshot)
     With rsOperacao
       If Not (.BOF And .EOF) Then
         If !ControleEntregas = -1 Then lblMovPendencia.Visible = True
       End If
     End With
     rsOperacao.Close
     Set rsOperacao = Nothing
   End If
  
   i = i + 1
  
  GoTo LP_Prox
  
Fim_Mostra:
  
  Rem apaga resto da linha
  If i <> Linhas_Grade Then
    For l = i To 254
      With Tabe(l)
        .C�digo = 0
        .Nome = ""
        .Unidade = ""
        .Pre�o_Total = 0
        .Pre�o_Final = 0
        .Nada = ""
        .Informa = ""
        .Qtde = 0
        .PesoLiquido = 0
        .PesoBruto = 0
        .Pre�o = 0
        .Desconto = 0
        .ICM = 0
        .IPI = 0
        .Etiqueta = 0
        .Descr_Adicional = ""
        .CFOP_Produto = "" '20/12/2006 - Anderson - Altera��o para registro de CFOP por Produto
        .Valor_Desonerado = 0
      End With
    Next l
  End If
  
  i = 0
  Linha = 0
  rsSaidas_Serv.Index = "Sequ�ncia"
Lp_Serv:
  rsSaidas_Serv.Seek ">", gnCodFilial, rsSaidas("Sequ�ncia"), Linha
  If rsSaidas_Serv.NoMatch Then GoTo Fim_Servi�os
  If rsSaidas_Serv("Filial") <> gnCodFilial Then GoTo Fim_Servi�os
  If rsSaidas_Serv("Sequ�ncia") <> rsSaidas("Sequ�ncia") Then GoTo Fim_Servi�os
  
  Linha = rsSaidas_Serv("Linha")
  rsServi�os.Index = "C�digo"
  
  Tabe_Serv(i).C�digo = rsSaidas_Serv("C�digo")
  Tabe_Serv(i).Descri��o = rsSaidas_Serv("Descri��o")
  Tabe_Serv(i).Tempo = rsSaidas_Serv("Tempo")
  Tabe_Serv(i).Pre�o = rsSaidas_Serv("Pre�o")
  Tabe_Serv(i).Completo = rsSaidas_Serv("Completo")
  Tabe_Serv(i).Total = CDbl(Tabe_Serv(i).Tempo) * CDbl(Tabe_Serv(i).Pre�o)
  Tabe_Serv(i).CFOP_Servico = rsSaidas_Serv("CFOP") & "" '20/12/2006 - Anderson - Altera��o para o registro de CFOP por servi�o
  
  '26/07/2005 - Daniel
  'Personaliza��o para a empresa J.R. Hidroqu�mica
  'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
  'C.S.T. (C�digo de Situa��o Tribut�ria)
  If m_blnJR Then Tabe_Serv(i).CST = rsSaidas_Serv("CST").Value & ""
  '-----------------------------------------------------------------
  rsServi�os.Seek "=", Tabe_Serv(i).C�digo
  If rsServi�os.NoMatch Then Tabe_Serv(i).ISS = 0
  Tabe_Serv(i).ISS = rsServi�os("ISS")
     
  i = i + 1
  
  GoTo Lp_Serv
   
Fim_Servi�os:
  If i <> Linhas_Servi�o Then
   For l = i To 254
    Tabe_Serv(l).C�digo = 0
    Tabe_Serv(l).Descri��o = ""
    Tabe_Serv(l).Tempo = ""
    Tabe_Serv(l).Pre�o = 0
    Tabe_Serv(l).Completo = False
    Tabe_Serv(l).Total = 0
    Tabe_Serv(l).ISS = 0
    Tabe_Serv(l).CFOP_Servico = "" '20/12/2006 - Anderson - Altera��o para Registro de CFOP por servi�o
    '26/07/2005 - Daniel
    'Personaliza��o para a empresa J.R. Hidroqu�mica
    'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
    'C.S.T. (C�digo de Situa��o Tribut�ria)
    If m_blnJR Then Tabe_Serv(l).CST = ""
    '-------------------------------------------------------------
   Next l
  End If
     
  Num_Registro = rsSaidas.Bookmark
     
  Grade1.MoveLast
  Grade1.MoveFirst
  Grade_Serv.MoveLast
  Grade_Serv.MoveFirst
  
  L_Tot_Prod.Text = Format(rsSaidas("Produtos"), "###,###,##0.00")
  L_Tot_Desc.Text = Format(rsSaidas("Desconto"), "###,###,##0.00")
  L_Tot_IPI.Text = Format(rsSaidas("IPI"), "###,###,##0.00")
  L_Frete.Text = Format(rsSaidas("Frete"), "###,###,##0.00")
  
  'Pilatti Novembro 18-11-2017
  If rsSaidas("FreteSomaOuNaoEstimativa") = True Then
    chk_freteNaoSomaPercentual.Value = 1
  Else
    chk_freteNaoSomaPercentual.Value = 0
  End If

  '12/05/2005 - Daniel
  'Adicionado campo Seguro
  If IsNumeric(rsSaidas("Seguro").Value) Then
    txtSeguro.Text = Format(rsSaidas("Seguro"), "###,###,##0.00")
  Else
    txtSeguro.Text = "0,00"
  End If
  
  L_Base_ICM.Text = Format(rsSaidas("Base ICM"), "###,###,##0.00")
  L_Valor_ICM.Text = Format(rsSaidas("Valor ICM"), "###,###,##0.00")
  L_Base_ICM_Subs.Text = Format(rsSaidas("Base ICM Subs"), "###,###,##0.00")
  L_Valor_ICM_Subs.Text = Format(rsSaidas("Valor ICM Subs"), "###,###,##0.00")
  
  '18/05/2005 - Daniel
  'Tratamento para o campo N� da NF
  txtNrSerieNF.Text = rsSaidas("SerieNF").Value & ""
  
  '22/10/2004 - Daniel
  'Armazenar o valor real do Total da Sa�da
  dblTotal = Format(rsSaidas("Total"), "###,###,##0.00")
  L_Tot_Pagar.Text = Format(rsSaidas("Total"), "###,###,##0.00")
  '17/09/2004 - Daniel
  'Case: Nilvo Burin
  'Valida��o para n�o mostrar valor total zerado caso o produto seja exclu�do
  If m_blnNilvo Then
    If rsSaidas("Total").Value <> 0 Then
      m_blnNotZero = True
      m_dblTotal = Format(rsSaidas("Total").Value, FORMAT_VALUE)
    End If
  End If
  '--------------------------------------------------------------------------
  L_Tot_Serv.Text = Format(rsSaidas("Servi�os"), "###,###,##0.00")
  L_Tot_ISS.Text = Format(rsSaidas("Valor ISS"), "###,###,##0.00")
  
  '20/09/2002 - mpdea
  'Exibi��o com o Desconto no SubTotal
  txtSubTotal.Text = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
  txtDescSubTotal.Text = Format(mcurDescontoSubTotal, FORMAT_VALUE)

  Combo_T�cnico.Text = rsSaidas("T�cnico") & ""
  Combo_T�cnico_LostFocus
  
  Combo_Caixa.Text = rsSaidas("Caixa") & ""
  Combo_Caixa_LostFocus
  
  '14/11/2002 - mpdea
  'Calcula Quantidade total de itens no grid
  lblQtdeTotal.Caption = sngQtdeTotal
  
  '-----------------------------------------------------------------------------
  '14/11/2002 - mpdea
  'Removido as fun��es comentadas abaixo por n�o estar apresentando
  'os valores gravados realmente na base de dados
  
'  '23/09/2002 - mpdea
'  'Removido por haver c�lculos dependentes da fun��o - Reavaliar
'  '20/09/2002 - mpdea
'  'Somente recalcula se a movimenta��o n�o foi efetivada
'  If Not rsSaidas("Efetivada").Value Then
'    Call Recalcula
'  End If
  '-----------------------------------------------------------------------------

  
  '-----------------------------------------------------------------------------
  'TESTAR
  '
  '07/04/2004 - mpdea
  'Ref.    : Erro na apresenta��o das informa��es com altera��o de totais
  'Solu��o : N�o efetuar o recalculo dos dados na exibi��o do registro
  '          comentando as linhas no bloco abaixo
  '-----------------------------------------------------------------------------
  Dim X As Integer

  Grade1.Redraw = False
  Grade1.MoveFirst

  For X = 0 To Grade1.Rows - 1
    Calcula_Linha
'    Grade1.Refresh
'    Grade1.Update
    Grade1.MoveNext
  Next X
  
  Grade1.MoveFirst
  Grade1.Redraw = True

  Recalcula
  
  '22/10/2004 - Daniel
  'Mostrar o valor real do total da sa�da
  L_Tot_Pagar.Text = Format(dblTotal, "###,###,##0.00")
  '-----------------------------------------------------------------------------
  
  '30/11/2004 - Daniel
  'Mostrar o valor real do registro ap�s o recalcula...
  'Solicitado por: Medicalway
  L_Tot_Prod.Text = Format(rsSaidas("Produtos"), "###,###,##0.00")
  L_Tot_IPI.Text = Format(rsSaidas("IPI"), "###,###,##0.00")
  L_Tot_Desc.Text = Format(rsSaidas("Desconto"), "###,###,##0.00")
  L_Frete.Text = Format(rsSaidas("Frete"), "###,###,##0.00")
  L_Base_ICM.Text = Format(rsSaidas("Base ICM"), "###,###,##0.00")
  L_Base_ICM_Subs.Text = Format(rsSaidas("Base ICM Subs"), "###,###,##0.00")
  L_Valor_ICM.Text = Format(rsSaidas("Valor ICM"), "###,###,##0.00")
  L_Valor_ICM_Subs.Text = Format(rsSaidas("Valor ICM Subs"), "###,###,##0.00")
  L_Tot_Serv.Text = Format(rsSaidas("Servi�os"), "###,###,##0.00")
  L_Tot_ISS.Text = Format(rsSaidas("Valor ISS"), "###,###,##0.00")
  L_Tot_ICMS_Deson.Text = Format(rsSaidas("TotalDesoneracaoICMS"), "###,###,##0.00")
  
  '17/05/2013-Alexandre Afornali
  'Case DiskEmbalagens
  If txtComanda.Visible = True Then
      Call CarregaComanda
  End If
'''  '29/10/2013 - Jean
'''  'Customiza��o para Disk Embalagens para bloquear a grid quando tiver uma sequencia j� gravada
'''  If CheckSerialCaseMod("QS73520-469") Then
'''    If (txtSeq.Text <> "") Then
'''      Grade1.Enabled = False
'''      DropDown1.Enabled = False
'''    End If
'''  End If
  
  '-----------------------------------------------------------------------------
  
  '01/10/2002 - mpdea
  'Fim da exibi��o do registro
  mblnInShowRecord = False
    
  Exit Sub
  
ErrHandler:
  '01/10/2002 - mpdea
  'Fim da exibi��o do registro
  mblnInShowRecord = False
    
  '12/06/2004 - Daniel
  'Altera��o: N�o h� como mostrar registro deletado "movimenta��o desfeita"
  If Err.Number = 13 And CLng(nSeqSaidas) = 0 Then  'Type mismatch
    rsSaidas.MoveLast
    Exit Sub
  End If

    
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apresentar o registro em Sa�das. Seq��ncia=" & CLng(nSeqSaidas)
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
End Sub

Public Sub Recalcula()
'06/05/2004 - Daniel
'Coment�rios sobre personaliza��es para a Embalavi:
'Diferimento: H� uma fun��o que veridica se h� Diferimento
'em caso afirmativo, faremos uma redu��o na base icm de 33%
'e tiraremos 18% deste valor que ser� o Valor ICM
  Dim nX As Integer
  '14/11/2002 - mpdea
  'Quantidade deve ser single (conforme estrutura da base de dados)
  Dim Qtde As Double
  
  
  Dim Pre�o As Double
  Dim Desconto As Double
  Dim Valor_Desconto As Double
  Dim IPI As Double
'  Dim Valor_IPI As Double
  Dim Pre�o_Total As Double
  Dim Pre�o_Final As Double
  Dim Pre�o_Final2 As Double
  
  Dim Tot_Desoneracao As Double
  Dim Temp_Desoneracao As Double
  
  Dim Tot_Prod As Double
  Dim Tot_Desc As Double
  Dim Tot_IPI As Double
  Dim Tot_Pagar As Double
  Dim Valor_Desc As Double
  Dim Valor_IPI As Double
  Dim temp As Double
  Dim Base_ICM As Double
  Dim Valor_ICM As Double
  Dim Base_ICM_Subs As Double
  Dim Valor_ICM_Subs As Double
  Dim Valor_ISS As Double
  Dim ICM(254, 2) As Double
  Dim nTbValIPI(254) As Currency
  Dim sCodProd As String
  Dim nValorIcmsFrete As Currency
  '14/11/2002 - mpdea
  'Quantidade deve ser single (conforme estrutura da base de dados)
  Dim nQtdeTotal As Single
  '10/11/2004 - Daniel
  Dim strUF As String
  Dim ValorTotalDesoneracao As Double
  
  'Caso esteja com altera totais pressionado
  'n�o executa o rec�lculo dos totais
  If ActiveBar1.Tools("miComplAlteraTotais").Checked Then Exit Sub
  
  Tot_Desc = 0#
  Tot_Prod = 0#
  gnPesoLiquido = 0#
  gnPesoBruto = 0#
  
  Tot_Desoneracao = 0#
  
  nQtdeTotal = 0
  
  ValorTotalDesoneracao = 0#
  
  For nX = 0 To (Linhas_Grade - 1)
    sCodProd = gsHandleNull(Tabe(nX).C�digo)
    If sCodProd <> "0" Then  'Faz somente os preenchidos
      
      Qtde = Tabe(nX).Qtde
      
      
      'Calcula Quantidade total de itens no grid
      nQtdeTotal = nQtdeTotal + Qtde
      
      '04/05/2004 - Daniel
      'Personaliza��o Embalavi
      If g_bln5CasasDecimais Then
        Pre�o = Format((Tabe(nX).Pre�o), "##,###,##0.00000")
      '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Pre�o = Format((Tabe(nX).Pre�o), "##,###,##0.000")
      Else
        'Pre�o = Tabe(nX).Pre�o
        Pre�o = Format((Tabe(nX).Pre�o), "##,###,##0.00")
      End If
      
      Desconto = Tabe(nX).Desconto
      IPI = Tabe(nX).IPI
      
      Pre�o_Total = Format(Qtde * Pre�o, "#0.00")
      Tabe(nX).Pre�o_Total = Pre�o_Total
      
      Valor_Desconto = Format(Pre�o_Total * Desconto / 100#, "#0.00")
      Pre�o_Final = Format((Pre�o_Total - Valor_Desconto), "#0.00")
      
      '------------------------------------------------------
      '23/05/2006 - mpdea
      'Comentado restri��o de isen��o de IPI para a Embalavi
      '� utilizado configura��o do cadastro de clientes
      '
      '06/05/2004 - Daniel
      'Caso seja Embalavi, chamaremos a fun��o IsencaoIPI
      'para verificar se o Cli_For � isento de IPI
'      If m_blnEmbalavi Then
'
'        If Len(cboCliente.Text) = 0 Then Exit Sub
        
        '28/09/2004 - mpdea
        'Otimizado a verifica��o do cliente isento em IPI
        If m_blnIsentoIPI Then
        'If IsencaoIPI(CLng(cboCliente.Text)) Then
          Valor_IPI = 0
        Else
          Valor_IPI = Pre�o_Final * IPI / 100#
          Valor_IPI = Truncate(Valor_IPI, 2)
        End If
        
'      Else 'N�o Embalavi
'        Valor_IPI = Pre�o_Final * IPI / 100#
'      End If
      '------------------------------------------------------
      
      Valor_IPI = Format(Valor_IPI, "#0.00")
      If Not Calcula_IPI Then
        Valor_IPI = 0
      End If
      
      If Calcula_IPI_TOT Then
        Pre�o_Final2 = Format((Pre�o_Final), "#0.00")
        Tabe(nX).Pre�o_Final = Pre�o_Final2
       Else
        Pre�o_Final2 = Format((Pre�o_Final + Valor_IPI), "#0.00")
        Tabe(nX).Pre�o_Final = Pre�o_Final2
      End If
            
'      Pre�o_Final2 = Format((Pre�o_Final + Valor_IPI), "#0.00")
'      Tabe(nX).Pre�o_Final = Pre�o_Final2
  
      
      '16/05/2006 - mpdea
      'Corrigido c�lculos de Base de C�lculo que estavam utilizando
      'tanto a base por valor quanto a base por percentual
      '(ICMS Retido e ICMS com Base Reduzida)
      With Tabe(nX)
        'Calculo do ICM
        If .Tipo_ICM = "N" Then
          'ICM Normal
          If gbBaseICMSomadoIPI = True Then
            .Base_ICM = Pre�o_Final2
            .Valor_ICM = Pre�o_Final2 * CSng(gsHandleNull(.ICM & "")) / 100
          Else
            .Base_ICM = Pre�o_Final
            .Valor_ICM = Pre�o_Final * CSng(gsHandleNull(.ICM & "")) / 100
          End If
        ElseIf .Tipo_ICM = "R" Then
          'ICM Retido
          If CDbl(.Valor_Base_Unit) <> 0 Then
            'Base Fixa
            .Base_ICM = CDbl(.Qtde) * CDbl(.Valor_Base_Unit)
            '''.Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
            
            '             vBCICMSST * Percentual Icm Saida (Tela Cadastro Produtos                  - (Pre�o do produto * Qtde) * percentual (tabela ICMS_PERCENTUAL_ESTADOS)
            .Valor_ICM = (CDbl(.Base_ICM) * CDbl(rsProdutos2.Fields("Percentual Icm Saida")) / 100) - CDbl(Pre�o_Final) * CDbl(.ICM) / 100
          
          ElseIf CDbl(.Redu��o_ICM) <> 0 Then
            'Base Reduzida
            .Base_ICM = Pre�o_Final * CDbl(.Redu��o_ICM) / 100
            .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
          End If
        ElseIf .Tipo_ICM = "Z" Then
          '10/11/2004 - Daniel
          'Personaliza��o para o cliente Teknika
          'Tratamento do ICM
          If Not m_blnTeknika Then
          
              'ICM Reduzido
              If CDbl(.Valor_Base_Unit) <> 0 Then
                'Base Fixa
                .Base_ICM = CDbl(.Qtde) * CDbl(.Valor_Base_Unit)
                .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
              ElseIf CDbl(.Redu��o_ICM) <> 0 Then
                'Base Reduzida
                .Base_ICM = Pre�o_Final * CDbl(.Redu��o_ICM) / 100
                .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
              End If
              
          Else 'Teknika
              
              'ICM Reduzido
              If CDbl(.Valor_Base_Unit) <> 0 Then
                'Base Fixa
                .Base_ICM = CDbl(.Qtde) * CDbl(.Valor_Base_Unit)
                .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
              End If
              
              '10/11/2004 - Daniel
              'Tratamento para base reduzida...
              'Chamamos a Function IE_Isento para verifica��o
              If IE_Isento(strUF) Then 'ISENTO = TRUE
                
                  .Base_ICM = Pre�o_Final
                  .Valor_ICM = Pre�o_Final * CSng(gsHandleNull(.ICM & "")) / 100
                
              Else 'ISENTO = FALSE
                
                If strUF = "PR" Then
                  If CDbl(.Redu��o_ICM) <> 0 Then
                    'Base Reduzida
                    .Base_ICM = Pre�o_Final * CDbl(.Redu��o_ICM) / 100 'CDbl(.Redu��o_ICM) = 66,66
                    .Valor_ICM = CDbl(.Base_ICM) * CDbl(.ICM) / 100
                  End If
                Else
                    .Base_ICM = Pre�o_Final
                    .Valor_ICM = Pre�o_Final * CSng(gsHandleNull(.ICM & "")) / 100
                End If
                
              End If
              
          End If
        End If
      End With
            
      gnPesoLiquido = gnPesoLiquido + Tabe(nX).PesoLiquido * Tabe(nX).Qtde
      gnPesoBruto = gnPesoBruto + Tabe(nX).PesoBruto * Tabe(nX).Qtde
      
      temp = Tabe(nX).Pre�o * Tabe(nX).Qtde
      temp = Format(temp, "#0.00")
      Tot_Prod = Tot_Prod + temp
      Valor_Desc = temp * Tabe(nX).Desconto / 100#
      Valor_Desc = Format(Valor_Desc, "#0.00")
      Tot_Desc = Tot_Desc + Valor_Desc
      temp = temp - Valor_Desc
      
      Temp_Desoneracao = Tabe(nX).Valor_Desonerado
      Tot_Desoneracao = Tot_Desoneracao + Temp_Desoneracao
      
      
      '------------------------------------------------------
      '23/05/2006 - mpdea
      'Comentado restri��o de isen��o de IPI para a Embalavi
      '� utilizado configura��o do cadastro de clientes
      '
      '06/05/2004 - Daniel
      'Caso seja Embalavi, chamaremos a fun��o IsencaoIPI
      'para verificar se o Cli_For � isento de IPI
'      If m_blnEmbalavi Then
'
'        If Len(cboCliente.Text) = 0 Then Exit Sub
        
        '28/09/2004 - mpdea
        'Otimizado a verifica��o do cliente isento em IPI
        If m_blnIsentoIPI Then
        'If IsencaoIPI(CLng(cboCliente.Text)) Then
          Valor_IPI = 0
        Else
          Valor_IPI = temp * Tabe(nX).IPI / 100#
          Valor_IPI = Truncate(Valor_IPI, 2)
        End If
      
'      Else 'N�o Embalavi
'        Valor_IPI = Temp * Tabe(nX).IPI / 100#
'      End If
      '------------------------------------------------------
      
      Valor_IPI = Format(Valor_IPI, "#0.00")
      
      If Calcula_IPI = False Then
        Valor_IPI = 0
      '------------------------------------------------------
      '23/05/2006 - mpdea
      'Comentado c�digo abaixo, pois a verifica��o de IPI
      'j� � realizada acima
      '
'      Else
'      'Else adicionado por Daniel, para fazer novamente a verifica��o
'      'para clientes isentos de IPI, case Embalavi, 06/05/2004
'        If m_blnEmbalavi Then
'          If Len(cboCliente.Text) = 0 Then Exit Sub
'
'          '28/09/2004 - mpdea
'          'Otimizado a verifica��o do cliente isento em IPI
'          If blnIsentoIPI Then
'          'If IsencaoIPI(CLng(cboCliente.Text)) Then
'            Valor_IPI = 0
'          End If
'        End If
      '------------------------------------------------------
      End If
      
      Tot_IPI = Tot_IPI + Valor_IPI
     
      If Calcula_ICM Then
        If Tabe(nX).Valor_Base_Unit <> 0 Or Tabe(nX).Redu��o_ICM <> 0 Then
          If Tabe(nX).Tipo_ICM = "R" Then
            Base_ICM_Subs = Base_ICM_Subs + Tabe(nX).Base_ICM
            Valor_ICM_Subs = Valor_ICM_Subs + Tabe(nX).Valor_ICM
          End If
        End If
      End If
     
      If Tabe(nX).ICM <> 0 Then
        If Calcula_ICM Then
          'ICM(Tabe(nX).ICM) = ICM(Tabe(nX).ICM) + Temp
          'If Tabe(nX).Valor_Base_Unit = 0 And Tabe(nX).Redu��o_ICM = 0 Then
          If Tabe(nX).Tipo_ICM = "Z" Or Tabe(nX).Tipo_ICM = "N" Then
            ICM(Tabe(nX).ICM, 1) = ICM(Tabe(nX).ICM, 1) + Tabe(nX).Base_ICM
            ICM(Tabe(nX).ICM, 2) = ICM(Tabe(nX).ICM, 2) + Tabe(nX).Valor_ICM
          End If
        End If
      End If
     ValorTotalDesoneracao = ValorTotalDesoneracao + Tabe(nX).Valor_Desonerado
     
    End If
  Next nX
  
      
  '14/11/2002 - mpdea
  'Quantidade deve ser single (conforme estrutura da base de dados)
  lblQtdeTotal.Caption = nQtdeTotal 'CDbl(nQtdeTotal)
  
  
  For nX = 1 To 254
    If ICM(nX, 1) <> 0 Then
      If Calcula_ICM Then
        Base_ICM = Base_ICM + ICM(nX, 1)
        Valor_ICM = Valor_ICM + ICM(nX, 2)
      End If
    End If
  Next nX
  
  Total_Servi�os = 0#
  For nX = 0 To 254
   If Tabe_Serv(nX).C�digo <> 0# Then
     Total_Servi�os = Total_Servi�os + Tabe_Serv(nX).Total
     If Tabe_Serv(nX).ISS <> 0 Then
       Valor_ISS = Valor_ISS + (Tabe_Serv(nX).Total * Tabe_Serv(nX).ISS / 100#)
       
       '06/12/2005 - mpdea
       'Corrigido c�lculo da fun��o Int atrav�s da convers�o do resultado
       'anterior para double
       '
       'Como ocorre o erro:
       '
       'Int(95) = 95
       'Int(9.5 * 10) = 95
       'Int(0.95 * 100) = 94 -------! ERRO !-------
       'Int(0.095 * 1000) = 95
    Valor_ISS = Int(CDbl(Valor_ISS * 100#)) / 100#
     End If
   End If
  Next nX
  
  Tot_Pagar = Tot_Prod - Tot_Desc + Tot_IPI + Total_Servi�os
'  Total_Desconto = Tot_Desc
  'Alterado para manter o total de desconto no Total Geral
  Total_Desconto = Tot_Desc + gcDescInTotal
  '- tot_desc
  
  Tot_Pagar = Tot_Pagar - gcDescInTotal
    
  If IsNull(L_Frete.Text) Or L_Frete.Text = "" Then L_Frete.Text = 0
  
  If Calcula_ICM_Frete Then
    If Calcula_ICM_Frete = True And Not IsNull(rsOp_Sa�da("Perc Icms Frete")) Then
       If Estado = "" Then
         PercIcmsFrete = rsOp_Sa�da("Perc Icms Frete")
       ElseIf Estado <> "" Then
         rsEstados.Index = "Estado"
         rsEstados.Seek "=", Estado
         If rsEstados.NoMatch Then
             PercIcmsFrete = rsOp_Sa�da("Perc Icms Frete")
             ElseIf Not rsEstados.NoMatch Then
                If rsEstados("ICM") = -1 Then
                   PercIcmsFrete = rsOp_Sa�da("Perc Icms Frete")
                Else
                   PercIcmsFrete = rsEstados("ICM")
                End If
             End If
        End If
    Else
       PercIcmsFrete = 0
    End If
    
    nValorIcmsFrete = (L_Frete.Text * PercIcmsFrete) / 100
    
    Base_ICM = Base_ICM + L_Frete.Text
    Valor_ICM = Valor_ICM + Format(nValorIcmsFrete, FORMAT_VALUE)
    
  End If
  
  If Soma_Frete Then
    Tot_Pagar = Tot_Pagar + Format(L_Frete.Text, FORMAT_VALUE)
  End If
  
  '12/04/2005 - Daniel
  'Tratamento para soma do seguro ao total a receber
  If Soma_Seguro Then
    Tot_Pagar = Tot_Pagar + Format(txtSeguro.Text, FORMAT_VALUE)
  End If
  
  '13/05/2004 - Daniel
  'Chamada da Fun��o que calcular� os impostos sobre servi�os
  'CSLL, COFINS, PIS, IRRF
  CalculaImpostosSobreServicos (Format(Total_Servi�os, FORMAT_VALUE))
  
  '20/09/2002 - mpdea
  'Adicionado o Desconto no SubTotal
  '14/05/2004 - Daniel
  'Adicionado o desconto de impostos sobre servi�os
  'a soma de CSLL, COFINS, PIS e IRRF (m_dblTotaldeImpostos)
  Tot_Pagar = Format(Tot_Pagar - mcurDescontoSubTotal - m_dblTotaldeImpostos, FORMAT_VALUE)
  
  '11/11/2008 - mpdea
  'Soma o ICMS Retido ao total da nota
  If m_blnSomaIcmsRetidoTotalNota Then
    Tot_Pagar = Format(Tot_Pagar + Valor_ICM_Subs, FORMAT_VALUE)
  End If
  
  Total_Pagar = Round(Tot_Pagar, 2)
  
  '14/05/2004 - Daniel
  'Adicionado o txtImpostosSobreServicos que � total de impostos
  'sobre servi�os (CSLL, COFINS, PIS e IRRF)
  txtImpostosSobreServicos.Text = Format(m_dblTotaldeImpostos, FORMAT_VALUE)
  '05/11/2007 - Anderson
  'Verifica se deve somar os produtos no total da nota
  'L_Tot_Prod.Text = Format(Tot_Prod, FORMAT_VALUE)
  If blnSomarProdutosTotalNota Then
    L_Tot_Prod.Text = Format(Tot_Prod, FORMAT_VALUE)
  Else
    L_Tot_Prod.Text = Format(0, FORMAT_VALUE)
  End If
  L_Tot_Serv.Text = Format(Total_Servi�os, FORMAT_VALUE)
  L_Tot_IPI.Text = Format(Tot_IPI, FORMAT_VALUE)
  L_Tot_ICMS_Deson.Text = Format(ValorTotalDesoneracao, FORMAT_VALUE)
  L_Tot_ICMS_Deson.Text = Format(Tot_Desoneracao, FORMAT_VALUE)
  '05/11/2007 - Anderson
  'Verifica se deve somar os produtos no total da nota
  '24/09/2002 - mpdea
  'Desconto no SubTotal
  'txtSubTotal.Text = Format(Tot_Pagar + mcurDescontoSubTotal, FORMAT_VALUE)
  'txtDescSubTotal.Text = Format(mcurDescontoSubTotal, FORMAT_VALUE)
  '************
  '************ dezembro 2019 COMENTEI ESTE CODIGO AQUI
'''  If blnSomarProdutosTotalNota Then
'''    txtSubTotal.Text = Format(Tot_Pagar + Total_Desconto, FORMAT_VALUE)
'''    txtDescSubTotal.Text = Format(Total_Desconto, FORMAT_VALUE)
'''  Else
'''    '28/12/2007 - Anderson
'''    'Corre��o para c�lculo do IRRF de servi�os.
'''    'txtSubTotal.Text = Format(Total_Servi�os, FORMAT_VALUE)
'''    txtSubTotal.Text = Format(Total_Servi�os - m_dblTotaldeImpostos, FORMAT_VALUE)
'''    txtDescSubTotal.Text = Format(Total_Desconto, FORMAT_VALUE)
'''  End If

  ' e fiz assim
  On Error GoTo ContinuaAbaixo
  If rsOp_Sa�da("SomarProdutosTotalNota") Then
    txtSubTotal.Text = Format(Tot_Pagar + Total_Desconto, FORMAT_VALUE)
    
    If (mcurDescontoSubTotal > 0 And Total_Desconto = 0) Or (Total_Pagar + mcurDescontoSubTotal + Total_Desconto = txtSubTotal.Text) Then
        txtDescSubTotal.Text = Format(mcurDescontoSubTotal + Total_Desconto, FORMAT_VALUE)
        Tot_Pagar = txtSubTotal.Text - (mcurDescontoSubTotal + Total_Desconto)
    Else
        txtDescSubTotal.Text = Format(Total_Desconto, FORMAT_VALUE)
    End If
  Else
    txtSubTotal.Text = Format(0, FORMAT_VALUE)
    txtDescSubTotal.Text = Format(0, FORMAT_VALUE)
  End If
  txtSubTotal.Text = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
  '*****************
  '*****************
ContinuaAbaixo:
  
  '05/11/2007 - Anderson
  'Verifica se deve somar os produtos no total da nota
  '17/09/2004 - Daniel
  'Case: Nilvo Burin
  'Valida��o para n�o mostrar valor total zerado caso o produto seja exclu�do
  'If m_blnNilvo Then
  '  If m_blnNotZero Then
  '    L_Tot_Pagar.Text = Format(m_dblTotal, FORMAT_VALUE)
  '  Else
  '    L_Tot_Pagar.Text = Format(Tot_Pagar, FORMAT_VALUE)
  '  End If
  'Else 'Demais clientes
  '  L_Tot_Pagar.Text = Format(Tot_Pagar, FORMAT_VALUE)
  'End If
  'L_Tot_Desc.Text = Format(Total_Desconto, FORMAT_VALUE)
  If blnSomarProdutosTotalNota Then
    If m_blnNilvo Then
      If m_blnNotZero Then
        L_Tot_Pagar.Text = Format(m_dblTotal, FORMAT_VALUE)
      Else
        L_Tot_Pagar.Text = Format(Tot_Pagar, FORMAT_VALUE)
      End If
    Else 'Demais clientes
      L_Tot_Pagar.Text = Format(Tot_Pagar, FORMAT_VALUE)
    End If
    L_Tot_Desc.Text = Format(Total_Desconto, FORMAT_VALUE)
  Else
    If m_blnNilvo Then
      If m_blnNotZero Then
        L_Tot_Pagar.Text = Format(Total_Servi�os, FORMAT_VALUE)
      Else
        L_Tot_Pagar.Text = Format(Total_Servi�os, FORMAT_VALUE)
      End If
    Else 'Demais clientes
      '28/12/2007 - Anderson
      'Corre��o do c�lculo do imposto
      'L_Tot_Pagar.Text = Format(Total_Servi�os - mcurDescontoSubTotal, FORMAT_VALUE)
      L_Tot_Pagar.Text = Format(Total_Servi�os - m_dblTotaldeImpostos - Total_Desconto, FORMAT_VALUE)
    End If
    L_Tot_Desc.Text = Format(0, FORMAT_VALUE)
  End If
  
  '23/05/2006 - mpdea
  'Centralizado verifica��o do uso de Diferimento
  '
  '06/05/2004 - Daniel
  'Case: Embalavi
  'Verifica��o de Diferimento sobre o ICM
  'quando for Embalavi
  'If m_blnEmbalavi Then
  If g_blnDiferimento Then

    If Len(cboCliente.Text) = 0 Then 'Tudo ser� 0,00 no in�cio...
      L_Base_ICM = Format(Base_ICM, FORMAT_VALUE)
      L_Valor_ICM = Format(Valor_ICM, FORMAT_VALUE)
      L_Base_ICM_Subs = Format(Base_ICM_Subs, FORMAT_VALUE)
      L_Valor_ICM_Subs = Format(Valor_ICM_Subs, FORMAT_VALUE)
    Else
        If Diferimento(CLng(cboCliente.Text)) Then
          Dim dblTotal As Double 'Antiga dblTrintaETres
          Dim dblBase  As Double 'Antiga dblDezoito
          Dim rstDiferimento As Recordset '17/05/2004
          Dim dblTotalTable  As Double
          Dim dblBaseTable   As Double
                  
          Set rstDiferimento = db.OpenRecordset("SELECT Total, Base FROM Diferimento WHERE Filial = " & gnCodFilial, dbOpenDynaset)
          
          With rstDiferimento
            If Not (.BOF And .EOF) Then
              If Not IsNumeric(.Fields("Total").Value) Or Not IsNumeric(.Fields("Base").Value) Then
                dblTotalTable = 0
                dblBaseTable = 0
              Else
                dblTotalTable = .Fields("Total").Value
                dblBaseTable = .Fields("Base").Value
              End If
            Else
              dblTotalTable = 0
              dblBaseTable = 0
            End If
            .Close
          End With
          
          Set rstDiferimento = Nothing
          

          dblTotal = Format(((Base_ICM * dblTotalTable) / 100), "###,##0.00")
          Base_ICM = Base_ICM - dblTotal

          dblBase = Format(((Base_ICM * dblBaseTable) / 100), "###,##0.00")
          Valor_ICM = dblBase

          L_Base_ICM = Format(Base_ICM, FORMAT_VALUE)
          L_Valor_ICM = Format(Valor_ICM, FORMAT_VALUE)

          '-----------[ICM Subs]-----------

          dblTotal = Format(((Base_ICM_Subs * dblTotalTable) / 100), "###,##0.00")
          Base_ICM_Subs = Base_ICM_Subs - dblTotal

          dblBase = Format(((Base_ICM_Subs * dblBaseTable) / 100), "###,##0.00")
          Valor_ICM_Subs = dblBase

          L_Base_ICM_Subs = Format(Base_ICM_Subs, FORMAT_VALUE)
          L_Valor_ICM_Subs = Format(Valor_ICM_Subs, FORMAT_VALUE)


        Else 'Caso n�o haja Diferimento, continua normal...
          L_Base_ICM = Format(Base_ICM, FORMAT_VALUE)
          L_Valor_ICM = Format(Valor_ICM, FORMAT_VALUE)
          L_Base_ICM_Subs = Format(Base_ICM_Subs, FORMAT_VALUE)
          L_Valor_ICM_Subs = Format(Valor_ICM_Subs, FORMAT_VALUE)
        End If

    End If
  Else 'N�o Embalavi continua normal...
    L_Base_ICM = Format(Base_ICM, FORMAT_VALUE)
    L_Valor_ICM = Format(Valor_ICM, FORMAT_VALUE)
    L_Base_ICM_Subs = Format(Base_ICM_Subs, FORMAT_VALUE)
    L_Valor_ICM_Subs = Format(Valor_ICM_Subs, FORMAT_VALUE)
  End If

    
  L_Tot_ISS.Text = Format(Valor_ISS, FORMAT_VALUE)
  
End Sub

'15/02/2007 - Anderson - Filtrar por cliente na tela de vendas - Solicitado por Paulo Ribertec.
Private Sub FiltrarCliente()

  'N�o permite que o bot�o fique em status de checked se as informa��es do filtro n�o estiverem corretas
  If Not IsNumeric(cboCliente) And cboCliente = "" Then
    ActiveBar1.Tools("miComplFiltrarCliente").Checked = False
    Exit Sub
  End If

  'Verifica se � para aplicar o filtro
  If ActiveBar1.Tools("miComplFiltrarCliente").Checked And IsNumeric(cboCliente) And cboCliente <> "" Then
    Set rsSaidas = db.OpenRecordset("SELECT * FROM Sa�das WHERE Filial = " & gnCodFilial & " AND Cliente=" & cboCliente.Text & " ORDER BY Sequ�ncia", dbOpenDynaset)
  Else
    Set rsSaidas = db.OpenRecordset("SELECT * FROM Sa�das WHERE Filial = " & gnCodFilial & " ORDER BY Sequ�ncia", dbOpenDynaset)
  End If

  'Move para o primeiro registro
  Call MoveFirst

End Sub

Private Sub AlteraTotais()
  
  Dim Tool As ActiveBarLibraryCtl.Tool
  Dim bLocked As Boolean
  
  Call StatusMsg("")
  
'  If IsNull(Num_Registro) Then
'    gsTitle = LoadResString(201)
'    gsMsg = "Encontre a movimenta��o de sa�da antes."
'    gnStyle = vbOKOnly + vbExclamation
'    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'    Exit Sub
'  End If
  
  Set Tool = ActiveBar1.Tools("miComplAlteraTotais")
  
  If Not Tool.Checked Then
    If L_Efetivada.Visible Then
'      gsTitle = LoadResString(201)
'      gsMsg = "Esta opera��o j� foi efetivada e n�o pode ser alterada."
'      gnStyle = vbOKOnly + vbExclamation
'      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      frmEfetivada.Show vbModal
      Exit Sub
    Else
      If Not frmGerente.gbSenhaGerente Then
        Exit Sub
      End If
    End If
  End If
  Tool.Checked = Not Tool.Checked
  
  bLocked = Not Tool.Checked
  
  L_Tot_Prod.Locked = bLocked
  L_Tot_Desc.Locked = bLocked
  L_Tot_IPI.Locked = bLocked
'  L_Frete.Locked = bLocked ** comentado na 43, agora o sistema calcula sozinho>> Leandro
  L_Base_ICM.Locked = bLocked
  L_Valor_ICM.Locked = bLocked
  L_Base_ICM_Subs.Locked = bLocked
  L_Valor_ICM_Subs.Locked = bLocked
  L_Tot_Pagar.Locked = bLocked
  L_Tot_Serv.Locked = bLocked
  L_Tot_ISS.Locked = bLocked
  L_Tot_ICMS_Deson.Locked = bLocked
  
'''  If Not bLocked Then
'''    L_Tot_Prod.SetFocus
'''  End If
  
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  sOrigemToolBarMoveRegistros = "MoveFirst"
  
  With rsSaidas
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub PesquisaPorData()
On Error Resume Next

    Dim sData As String
    
    If ActiveBar1.Tools("miOpOrdem").Text <> "Por Data e Seq��ncia" Then
      MsgBox "Deixe selecionada a op��o 'Por Data e Seq��ncia' na combo de pesquisa."
      ActiveBar1.Tools("miOpOrdem").SetFocus
      Exit Sub
    End If
    
    'sData = InputBox("Informe a Data (dd/mm/aaaa)", "Data:")
    sData = frmCalendario.gsDateCalender(Now)
    
    'rsSaidas.Sort = "Data"
    rsSaidas.MoveFirst
    
    While Not rsSaidas.EOF
      If rsSaidas.Fields("Data").Value = sData Then
        Call ShowRecord
        Exit Sub
      End If
      
      rsSaidas.MoveNext
    Wend

End Sub
Private Sub MoveLast()
  On Error Resume Next
  
  sOrigemToolBarMoveRegistros = "MoveLast"
  
  If ActiveBar1.Tools("miComplFiltrarCliente").Checked = False Then
      'Recarrega as SAIDAS se o bot�o de FUNIL n�o estiver selecionado...para n�o prejudicar a pesquisa de saidas por cliente selecionado
      rsSaidas.Close
      Set rsSaidas = db.OpenRecordset("SELECT * FROM Sa�das WHERE Filial = " & gnCodFilial & " ORDER BY Sequ�ncia", dbOpenDynaset)
  End If

  With rsSaidas
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
      
      '18/02/2005 - Daniel
      '
      'Case: Aura Prata
      '
      'Em casos de sa�das originadas na tela de acerto,
      'o primeiro �tem da nota n�o estava sendo exibido.
      'For�amos a mostrar a sa�da
      If CheckSerialCaseMod("QS40898-680") Then Call ShowRecord
      '----------------------------------------------------------------
    End If
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  
  sOrigemToolBarMoveRegistros = "MovePrevious"
  
  With rsSaidas
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
  
  sOrigemToolBarMoveRegistros = "MoveNext"

  With rsSaidas
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
  Dim Sai_Loop As Integer
  Dim Fim As Integer
  Dim Ordem As Long
  Dim Resposta As Integer
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre a movimenta��o de sa�da antes."
    Exit Sub
  End If
  
  If L_Efetivada.Visible = True Then
    DisplayMsg "Esta opera��o j� foi efetivada e n�o pode ser apagada por aqui, veja a ajuda."
    Exit Sub
  End If
 
  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente apagar esta movimenta��o?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    DisplayMsg "Movimenta��o n�o apagada."
    Exit Sub
  End If
 
  '15/05/2013-Alexandre Afornali
  'Case DiskEmbalagens
  Dim rsComandas As Recordset
  Set rsComandas = db.OpenRecordset("SaidasComandas")
  Dim countrs As Long
  countrs = 0
  While Not rsComandas.EOF
    countrs = countrs + 1
    rsComandas.MoveNext
  Wend
  If (countrs > 0) Then
    rsComandas.MoveFirst
  End If
  While Not rsComandas.EOF
    If (rsComandas("CodSaida") = txtSeq.Text) And (rsComandas("Filial").Value = gnCodFilial) Then
      rsComandas.Delete
      rsComandas.MoveLast
    End If
    rsComandas.MoveNext
  Wend
 
  Call StatusMsg("Apagando movimenta��o (produtos).")
  'Apaga produtos
  rsSaidas_Prod.Index = "Sequ�ncia"
  Sai_Loop = False
  
  Do
    rsSaidas_Prod.Seek ">", gnCodFilial, Val(txtSeq.Text)
    
    If rsSaidas_Prod.NoMatch Then Sai_Loop = True
    If Sai_Loop = False Then If rsSaidas_Prod("Filial") <> gnCodFilial Then Sai_Loop = True
    If Sai_Loop = False Then If rsSaidas_Prod("Sequ�ncia") <> Val(txtSeq.Text) Then Sai_Loop = True
    
    If Sai_Loop = False Then
      rsSaidas_Prod.Delete
    End If
  Loop Until Sai_Loop = True

  rsSa�da_Cheques.Index = "Ordem"
  Fim = False
  Ordem = 0
  'apaga cheques existentes
  Do
    rsSa�da_Cheques.Seek ">", gnCodFilial, Val(txtSeq.Text), Ordem
    If rsSa�da_Cheques.NoMatch Then Fim = True
    If Fim = False Then If rsSa�da_Cheques("Sequ�ncia") <> Val(txtSeq.Text) Then Fim = True
    If Fim = False Then If rsSa�da_Cheques("Filial") <> gnCodFilial Then Fim = True
    If Fim = False Then
      rsSa�da_Cheques.Delete
    End If
  Loop Until Fim = True
    
  rsSa�da_Parcelas.Index = "Ordem"
  Fim = False
  'apaga Parcelas existentes
  Do
    rsSa�da_Parcelas.Seek ">", gnCodFilial, Val(txtSeq.Text), Ordem
    If rsSa�da_Parcelas.NoMatch Then Fim = True
    If Fim = False Then If rsSa�da_Parcelas("Sequ�ncia") <> Val(txtSeq.Text) Then Fim = True
    If Fim = False Then If rsSa�da_Parcelas("Filial") <> gnCodFilial Then Fim = True
    If Fim = False Then
      rsSa�da_Parcelas.Delete
    End If
  Loop Until Fim = True
    

  Num_Registro = Null
  L_Efetivada.Visible = False
  rsSaidas.Delete
  
  'LOG *****************
  sSql = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '"
  sSql = sSql & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & Val(txtSeq.Text) & " Vlr:" & L_Tot_Pagar.Text & " Cli:" & cboCliente.Text, 80) & "', 'EXCLUSAO SAIDAS')"
  db.Execute sSql, dbFailOnError
  'fim *******************
  
  txtSeq.Text = ""
  
  Call ClearScreen
  StatusMsg "Opera��o apagada."

End Sub

'30/04/2003 - mpdea
'Dividido rotina em RealizaDescontoSubTotal e RealizaDescontoRateado
Private Sub RealizaDesconto()
  
  If m_blnDescontoRateado Then
    Call RealizaDescontoRateado
  Else
    Call RealizaDescontoSubTotal
  End If
  
End Sub

'22/10/2002 - mpdea
'Atualizado as chamadas para UpdateRecord que retornam valor
'
'20/09/2002 - mpdea
'Totalmente reformulado para suportar o desconto no sub total
'
'O Desconto no SubTotal dever� ser a �ltima opera��o antes do recebimento,
'portanto n�o poder� haver venda de mais itens ap�s o desconto
'
Private Sub RealizaDescontoSubTotal()
  Dim sngMaxDescPerc As Single
  Dim curDesconto As Currency
  Dim curNewTotal As Currency
  Dim blnHasItem As Boolean
  Dim intX As Integer
  Dim strSQL As String
  Dim intRet As Integer
  
  
  '28/11/2002 - mpdea
  'Ajustes da Base de ICM
  Dim curBaseICM As Currency
  Dim curValorICM As Currency
  Dim curValorIPI As Currency
  Dim curValorISS As Currency
  Dim sngDescPerc As Single
  
  
  Call StatusMsg("")
  
  'Atualiza opera��o
  Call cboOper_LostFocus
  
  If Nome_Opera��o.Caption = "" Then
    DisplayMsg "Opera��o n�o selecionada ou incorreta."
    Exit Sub
  End If
  
  If Not rsOp_Sa�da.Fields("Dinheiro").Value Then
    DisplayMsg "Tipo de opera��o n�o movimenta dinheiro para utilizar esta fun��o."
    Exit Sub
  End If
  
  For intX = 0 To (Grade1.Rows - 1)
    If Tabe(intX).C�digo <> "0" And Tabe(intX).C�digo <> "" Then
      blnHasItem = True
    End If
  Next intX

  If Not blnHasItem Then
    DisplayMsg "N�o existe nenhum produto digitado, imposs�vel fornecer desconto."
    Exit Sub
  End If
  
  If Total_Pagar = 0 Then
    DisplayMsg "Total igual a zero, imposs�vel fornecer desconto."
    Exit Sub
  End If
  
  If L_Efetivada.Visible Then
    DisplayMsg "Movimenta��o j� efetivada."
    Exit Sub
  End If
  
  
  '20/11/2002 - mpdea
  'Verifica��o de desconto j� concedido
  If Not IsNull(Num_Registro) Then
    If mcurDescontoSubTotal > 0 Then
      DisplayMsg "Desconto no SubTotal j� concedido."
      Exit Sub
    End If
  End If
  
  
  'Percentual de desconto para o funcion�rio / Filial
  sngMaxDescPerc = IIf(rsFuncionarios("nPercDesconto") = 0, _
    rsParametros("VR Desconto"), rsFuncionarios("nPercDesconto"))
    
  'Exibe o form de desconto
  '05/11/2007 - Anderson
  'Verifica se o total de produtos deve somar a nota
  'If frmDesconto.Start(CCur(Total_Pagar), sngMaxDescPerc, curDesconto, curNewTotal, False) Then
  If Not b_EscondeTelaDesconto Then
  If frmDesconto.Start(IIf(blnSomarProdutosTotalNota, CCur(Total_Pagar), CCur(Total_Servi�os)), sngMaxDescPerc, curDesconto, curNewTotal, False) Then
    
    '03/09/2003 - mpdea
    'Inclu�do IPI e ISS
    '
    '28/11/2002 - mpdea
    'Armazena temporariamente valores de ICM (normal)
    curBaseICM = CCur("0" & L_Base_ICM.Text)
    curValorICM = CCur("0" & L_Valor_ICM.Text)
    curValorIPI = CCur("0" & L_Tot_IPI.Text)
    curValorISS = CCur("0" & L_Tot_ISS.Text)
    
    
    '03/09/2003 - mpdea
    'Removido formata��o do percentual de desconto
    'ocasionava erro de arredondamento
    '
    'Desconto concedido em percentual
    sngDescPerc = CSng(curDesconto / Total_Pagar)
    
    
    'Atualiza valores
    Total_Pagar = CDbl(curNewTotal)
    mcurDescontoSubTotal = curDesconto
    
    'Atualiza exibi��o
    txtSubTotal.Text = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
    txtDescSubTotal.Text = Format(mcurDescontoSubTotal, FORMAT_VALUE)
    L_Tot_Pagar.Text = Format(Total_Pagar, FORMAT_VALUE)
    
    
    '13/12/2002 - mpdea
    'Inclu�do IPI e ISS
    '
    '28/11/2002 - mpdea
    'Atualiza valores de ICM
    L_Base_ICM.Text = Format(curBaseICM * (1 - sngDescPerc), FORMAT_VALUE)
    L_Valor_ICM.Text = Format(curValorICM * (1 - sngDescPerc), FORMAT_VALUE)
    L_Tot_IPI.Text = Format(curValorIPI * (1 - sngDescPerc), FORMAT_VALUE)
    L_Tot_ISS.Text = Format(curValorISS * (1 - sngDescPerc), FORMAT_VALUE)
    
    
    'Atualiza registro
    intRet = UpdateRecord
    
    'Verifica erro
    If intRet <> 0 Or mblnForceUpdate Then
      'Ativa flag para for�ar nova atualiza��o de registro
      mblnForceUpdate = True
      
      'Cancela desconto
      Total_Pagar = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
      mcurDescontoSubTotal = 0
      
      'Atualiza exibi��o
      txtSubTotal.Text = Format(Total_Pagar, FORMAT_VALUE)
      txtDescSubTotal.Text = Format(0, FORMAT_VALUE)
      L_Tot_Pagar.Text = Format(Total_Pagar, FORMAT_VALUE)
      
      
      '13/12/2002 - mpdea
      'Inclu�do IPI e ISS
      '
      '28/11/2002 - mpdea
      'Restaura valores de ICM
      L_Base_ICM.Text = Format(curBaseICM, FORMAT_VALUE)
      L_Valor_ICM.Text = Format(curValorICM, FORMAT_VALUE)
      L_Tot_IPI.Text = Format(curValorIPI, FORMAT_VALUE)
      L_Tot_ISS.Text = Format(curValorISS, FORMAT_VALUE)
      
      
      Exit Sub
    End If
    End If
    
    'Realiza recebimento
    Call RealizaRecebimento
    
    'Verifica confirma��o do recebimento
    'Caso contr�rio restaura valores anteriores ao desconto
    If Not L_Efetivada.Visible Then
      'Ativa flag para for�ar nova atualiza��o de registro
      mblnForceUpdate = True
      
      'Atualiza valores
      Total_Pagar = Format(mcurDescontoSubTotal + Total_Pagar, FORMAT_VALUE)
      mcurDescontoSubTotal = 0
      
      '--------------------------------------------------------------------------
      '13/12/2002 - mpdea
      'Inclu�do IPI e ISS
      '
      '28/11/2002 - mpdea
      'Restaura valores do registro para os campos: Base ICM e Valor ICM
      '
      '07/11/2002 - mpdea
      'Corrigido argumento de valor para a string SQL (RT-3144)
      '
      'Restaura valores no registro gravado
      strSQL = "UPDATE Sa�das SET DescontoSubTotal = 0, Total = " & _
        Replace(Total_Pagar, ",", ".") & _
        ", [Base ICM] = " & Replace(curBaseICM, ",", ".") & _
        ", [Valor ICM] = " & Replace(curValorICM, ",", ".") & _
        ", [Valor ISS] = " & Replace(curValorISS, ",", ".") & _
        ", IPI = " & Replace(curValorIPI, ",", ".") & _
        " WHERE Filial = " & gnCodFilial & " AND Sequ�ncia = " & CLng(txtSeq.Text)
      db.Execute strSQL, dbFailOnError
      '--------------------------------------------------------------------------
      
      
      'Atualiza exibi��o
      txtSubTotal.Text = Format(Total_Pagar, FORMAT_VALUE)
      txtDescSubTotal.Text = Format(0, FORMAT_VALUE)
      L_Tot_Pagar.Text = Format(Total_Pagar, FORMAT_VALUE)
      
      
      '13/12/2002 - mpdea
      'Inclu�do IPI e ISS
      '28/11/2002 - mpdea
      'Restaura valores de ICM
      L_Base_ICM.Text = Format(curBaseICM, FORMAT_VALUE)
      L_Valor_ICM.Text = Format(curValorICM, FORMAT_VALUE)
      L_Tot_IPI.Text = Format(curValorIPI, FORMAT_VALUE)
      L_Tot_ISS.Text = Format(curValorISS, FORMAT_VALUE)
      
      
      Exit Sub
    End If
    
    'Desativa flag, opera��o completada com sucesso
    mblnForceUpdate = False
    
  End If
  
End Sub

Private Sub RealizaDescontoRateado()
  Dim Conta As Integer
  Dim i As Integer
  Dim Desc_Max As Double
  Dim Desc As Double
  Dim Desc_Perc As Double
  Dim Novo_Total As Double
  Dim Tot_IPI As Double
  Dim F As Form
  Dim nValDif As Double
  Dim sPrecoFinal As Double
  Dim sPrecoTotal As Double
  Dim nLastRow As Long
  Dim nVal As Integer
  Dim nQtdeItens As Integer
  Dim nDesconto As Double
  Dim nDescontoUlt As Double
  Dim nTotalDesconto As Double
  Dim nPercMaxDesc As Single
  '23/04/2009 - mpdea
  Dim str_format_casas_decimais As String


  Call StatusMsg("")

  Conta = 0
  For i = 0 To (Grade1.Rows - 1)
   If Tabe(i).C�digo <> "0" And Tabe(i).C�digo <> "" Then Conta = Conta + 1
  Next i

  If Conta = 0 Then
    DisplayMsg "N�o existe nenhum produto digitado, imposs�vel fornecer desconto."
    Exit Sub
  End If

  If L_Efetivada.Visible = True Then
    DisplayMsg "Movimenta��o j� efetivada."
    Exit Sub
  End If

  
  '30/04/2003 - mpdea
  'Adapta��es para o desconto rateado
  '---------------------------------------------------------------------------------
  'Percentual de desconto para o funcion�rio / Filial
'  Desc_Max = Round(CDbl(Total_Pagar * rsParametros("VR Desconto") / 100#), 2)
  nPercMaxDesc = IIf(rsFuncionarios("nPercDesconto") = 0, _
    rsParametros("VR Desconto"), rsFuncionarios("nPercDesconto"))
  
  
  'Exibe o form de desconto
  '05/11/2007 - Anderson
  'Verifica se os produtos devem ser somados a nota
  'If Not frmDesconto.Start(CCur(Total_Pagar), nPercMaxDesc, _
  '                         0, 0, True, Total_Desconto) Then Exit Sub
  If Not b_EscondeTelaDesconto Then _
    If Not frmDesconto.Start(IIf(blnSomarProdutosTotalNota, CCur(Total_Pagar), CCur(Total_Servi�os)), nPercMaxDesc, _
                           0, 0, True, IIf(blnSomarProdutosTotalNota, Total_Desconto, 0)) Then Exit Sub


'  Set F = New frmDesconto
'  F.Desc_Fornecido.Caption = Format(Total_Desconto, "###,###,##0.00")
'  F.Total.Caption = Total_Pagar
'  F.Desconto.Text = ""
'  F.Show vbModal
'  Set F = Nothing
'
'  If gnDesconto = 0# Then Exit Sub
  '---------------------------------------------------------------------------------
  
  
  '23/04/2009 - mpdea
  'Formata��o do pre�o unit�rio
  If g_bln3CasasDecimais Then
    str_format_casas_decimais = "000"
  ElseIf g_bln5CasasDecimais Then
    str_format_casas_decimais = "00000"
  Else
    str_format_casas_decimais = "00"
  End If
  
  
  Desc_Max = (Total_Pagar + Total_Desconto) * nPercMaxDesc / 100
  If (Desc_Max + 0.1) < (Total_Desconto + gnDesconto) Then
    DisplayMsg "Desconto superior ao permitido."
    Exit Sub
  End If

  Total_Desconto = Total_Desconto + gnDesconto
  'Adicionado para manter o total em desconto no Total Geral
  gcDescInTotal = gcDescInTotal + gnDesconto

  Desc_Perc = gnDesconto / Total_Pagar
  Desc_Perc = 1 - Desc_Perc
  Novo_Total = Total_Pagar - gnDesconto
  
  
  For i = 0 To (Grade1.Rows - 1)
    With Tabe(i)
      '23/04/2009 - mpdea
      'Adicionado: And .C�digo <> "0"
      If .C�digo <> "" And .C�digo <> "0" Then
        '23/04/2009 - mpdea
        'Modificado para que o c�lculo do pre�o seja formatado de acordo com as casas decimais de pre�o
        '.Pre�o = Format(.Pre�o * Desc_Perc, FORMAT_VALUE)
        '.Pre�o = Format(.Pre�o * Desc_Perc, "#0." & str_format_casas_decimais)
        .Pre�o_Total = Format(.Qtde * .Pre�o * Desc_Perc, FORMAT_VALUE)
        .Pre�o_Final = Format(.Pre�o_Total * (1 - .Desconto / 100), FORMAT_VALUE)
        .Base_ICM = .Pre�o_Final
        .Valor_ICM = Format(.Pre�o_Total * .ICM / 100, FORMAT_VALUE)
        Tot_IPI = Format(.Pre�o_Final * .IPI / 100, FORMAT_VALUE)
        
        If Calcula_IPI_TOT Then
          .Pre�o_Final = .Pre�o_Final
        Else
          .Pre�o_Final = .Pre�o_Final + Tot_IPI
        End If
      End If
    End With
  Next i
  
  Call Recalcula
  
  If Total_Pagar <> Novo_Total Then
    Desc = Total_Pagar - Novo_Total
    For i = 0 To (Grade1.Rows - 1)
      '23/04/2009 - mpdea
      'Adicionado: And Tabe(i).C�digo <> "0"
      If Tabe(i).C�digo <> "" And Tabe(i).C�digo <> "0" Then
        If Tabe(i).Qtde = 1 Then
          '23/04/2009 - mpdea
          'Modificado para que o c�lculo do pre�o seja formatado de acordo com as casas decimais de pre�o
          'Tabe(i).Pre�o = Format((Tabe(i).Pre�o - Desc), FORMAT_VALUE)
          'Tabe(i).Pre�o = Format((Tabe(i).Pre�o - Desc), "#0." & str_format_casas_decimais)
          Tabe(i).Pre�o_Final = Tabe(i).Qtde * (Tabe(i).Pre�o - Desc)
          Tabe(i).Base_ICM = Tabe(i).Pre�o_Final
          Tabe(i).Valor_ICM = Round(CDbl(Tabe(i).Pre�o_Total * Tabe(i).ICM / 100#), 2)
          Tot_IPI = Format(Tabe(i).Pre�o_Final * Tabe(i).IPI / 100, FORMAT_VALUE)
          If Calcula_IPI_TOT Then
            Tabe(i).Pre�o_Final = Tabe(i).Pre�o_Final
          Else
            Tabe(i).Pre�o_Final = Tabe(i).Pre�o_Final + Tot_IPI
          End If
          Desc = 0
          Exit For
        End If
      End If
    Next i
    Call Recalcula
    L_Tot_Desc.Text = Format(CCur(L_Tot_Desc.Text) - (nValDif / 100#), "###,###,##0.00")
  End If

  '23/04/2009 - mpdea
  'Ajusta desconto caso haja res�duo
  gcDescInTotal = Format(gcDescInTotal - Desc, FORMAT_VALUE)
  Total_Desconto = Format(Total_Desconto - Desc, FORMAT_VALUE)

  Grade1.MoveLast
  Grade1.MoveFirst
  
  
  
'  Desc_Max = (Total_Pagar + Total_Desconto) * nPercMaxDesc / 100
'  If Desc_Max < (Total_Desconto + gnDesconto) Then
'    DisplayMsg "Desconto superior ao permitido."
'    Exit Sub
'  End If
'
'  Total_Desconto = Total_Desconto + gnDesconto
'  'Adicionado para manter o total em desconto no Total Geral
'  gcDescInTotal = gcDescInTotal + gnDesconto
'
'  Desc_Perc = gnDesconto / Total_Pagar
'  Desc_Perc = 1 - Desc_Perc
'  Novo_Total = Total_Pagar - gnDesconto
'  nQtdeItens = 0
'
'  For i = 0 To (Grade1.Rows - 1)
'
'   If gsHandleNull(Tabe(i).C�digo) <> "0" Then
'     Tabe(i).Pre�o = Round(CDbl(Tabe(i).Pre�o * Desc_Perc), 2)
'     Tabe(i).Pre�o_Total = Round(CDbl(Tabe(i).Qtde * Tabe(i).Pre�o), 2)
'     Tabe(i).Pre�o_Final = Round(CDbl(Tabe(i).Pre�o_Total * (1 - Tabe(i).Desconto / 100#)), 2)
'     Tabe(i).Base_ICM = Tabe(i).Pre�o_Final
'     Tabe(i).Valor_ICM = Round(CDbl(Tabe(i).Pre�o_Total * Tabe(i).ICM / 100#), 2)
'     Tot_IPI = Round(CDbl(Tabe(i).Pre�o_Final * Tabe(i).IPI / 100#), 2)
'     If Calcula_IPI_TOT Then
'        Tabe(i).Pre�o_Final = Tabe(i).Pre�o_Final
'     Else
'        Tabe(i).Pre�o_Final = Tabe(i).Pre�o_Final + Tot_IPI
'     End If
'
'     nQtdeItens = nQtdeItens + 1
'     nLastRow = i
'   End If
'
'  Next i
'
'  Call Recalcula
'
'  nValDif = Int((Total_Pagar - Novo_Total) * 100)
'  If nValDif <> 0 Then
'    nDesconto = (nValDif \ nQtdeItens)
'    nDescontoUlt = (nValDif - (nDesconto * (nQtdeItens - 1)))
'    nDesconto = nDesconto / 100#
'    nDescontoUlt = nDescontoUlt / 100#
'    nTotalDesconto = 0
'    For i = 0 To (Grade1.Rows - 1)
'      DoEvents
'      If gsHandleNull(Tabe(i).C�digo) <> "0" Then
'        If i < nQtdeItens - 1 Then
'          Tabe(i).Pre�o_Final = Round(Tabe(i).Pre�o_Total * (1 - Tabe(i).Desconto / 100#) - nDesconto, 2)
'        Else
'          Tabe(i).Pre�o_Final = Round(Tabe(i).Pre�o_Total * (1 - Tabe(i).Desconto / 100#) - nDescontoUlt, 2)
'        End If
'        Tabe(i).Base_ICM = Tabe(i).Pre�o_Final
'        Tabe(i).Valor_ICM = Round(CDbl(Tabe(i).Pre�o_Total * Tabe(i).ICM / 100#), 2)
'        Tot_IPI = Round(CDbl(Tabe(i).Pre�o_Final * Tabe(i).IPI / 100#), 2)
'
'        If Calcula_IPI_TOT Then
'           Tabe(i).Pre�o_Final = Tabe(i).Pre�o_Final
'        Else
'          Tabe(i).Pre�o_Final = Tabe(i).Pre�o_Final + Tot_IPI
'        End If
'
'      End If
'    Next i
'    Call Recalcula
'    L_Tot_Desc.Text = Format(CCur(L_Tot_Desc.Text) - (nValDif / 100#), "###,###,##0.00")
'  End If
'
'  Grade1.MoveLast
'  Grade1.MoveFirst

End Sub

'22/10/2002 - mpdea
'Modificado para function, 0 = Sucesso
Private Function UpdateRecord() As Integer
  Dim nRet As Integer
  Dim nSequencia As Long
  'Vari�veis de Tratamento de Erro
  Dim bSequencia As Boolean
  Dim bSeqChanged As Boolean
  Dim intRepeatUpdate3022 As Integer
  Dim intRepeatUpdateLocked As Integer
  
  Dim i As Integer
  Dim Conta As Integer
  Dim Linha As Integer
  Dim Ordem As Integer
  Dim Tabe1 As Variant
  Dim Aux_Cod_Prod As String
  Dim Aux_Prod As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edi��o As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim Aux_Texto As String
  Dim Limite_Usado As Double
  Dim M�ximo As Double
'  Dim Book_Par As Variant
  Dim sMsg As String
  
  Dim nPercMaxDesc As Single
  Dim cDescMax As Currency
  
  Dim sUnidade As String
  Dim sTributaria As String
  
  '27/05/2004 - Daniel
  'Var para controle do campo Validade quando PSV
  Dim blnValidade As Boolean
  
  '08/08/2002 - mpdea
  'N�mero do terminal para opera��es de or�amento
  Dim bytNrTerminal As Byte
  'N�mero do pr�ximo or�amento
  Dim lngNrOrcamento As Long
  
  '18/11/2002 - mpdea
  'Flag para exibi��o de mensagem avisando
  'a cria��o da nova movimenta��o de or�amento
  Dim blnShowMessageNewBudget As Boolean
  
  Dim rstPrecos             As Recordset
  Dim blnOrcamentoAprovado  As Boolean
  
  '05/05/2004 - mpdea
  'Controle de transa��o
  Dim blnInTransaction As Boolean
  
  '12/05/2004 - Daniel
  'Soma dos impostos sobre Servi�os
  Dim dblSomaImpostosSobreServ As Double
  
  Dim dblTotalIcmsDesonerado As Double
  
  bProdutoSemPrecoNaGrade = False
  
  '22/10/2002 - mpdea
  'Atribui valor inicial de retorno da fun��o
  UpdateRecord = -1
  
  totalNCM_2 = 0#
  
  
  'MsgBox cboOper.Text
  
  Linha = Grade1.Row
  On Error Resume Next
'  Grade1.Row = 2
'  Grade1.Row = 1
'  Grade_Serv.Row = 2
'  Grade_Serv.Row = 1
  Grade1.MoveLast
  Grade1.MoveFirst
  Grade_Serv.MoveLast
  Grade_Serv.MoveFirst
  On Error GoTo 0
  DoEvents
  
  On Error GoTo ErrHandler

  Call cboDigitador_LostFocus

  If IsNull(Num_Registro) And gbDemoVersion Then
    rsSaidas.MoveLast
    rsSaidas.MoveFirst
    If rsSaidas.RecordCount >= NMAXREGDEMO Then
      gsTitle = LoadResString(201)
      gsMsg = LoadResString(13)
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Exit Function
    End If
  End If
 
  If L_Efetivada.Visible = True Then
    frmEfetivada.Show vbModal
    Exit Function
  End If
  
  
  '07/05/2003 - mpdea
  'Verifica se a movimenta��o foi efetivada
  If Not IsNull(Num_Registro) Then
    If rsSaidas.Fields("Efetivada").Value Then
      frmEfetivada.Show vbModal
      Exit Function
    End If
  End If
  
  
  ' ********************
  'APENAS AVISO DE TRATAMENTO DE CONSIST�NCIA DE REGRAS no caso quando for gerado o XML para a SEFAZ
  'Verificar operacao de saida com tipo de finalidade...apenas para aviso
  If sTipoOperacaoSaida = "G" And cboFinalidade.ListIndex <> 3 Then
      Dim retMsgOper As Variant
      retMsgOper = MsgBox("REVISANDO...EST� DE ACORDO?" & vbCrLf & vbCrLf & "          Para OPERA��O DE DEVOLU��O escolha a FINALIDADE = 4 " & vbCrLf & "          Para OPERA��O DE REMESSA      escolha a FINALIDADE = 1 " & vbCrLf & vbCrLf & " CASO OK, PROSSIGA", vbYesNo, "Aten��o")
  
      If retMsgOper = vbNo Then
          Exit Function
      End If
  End If
  
  If cboFinalidade.ListIndex = 3 And gridChaves.Rows <= 1 Then
      Dim retMsgOper2 As Variant
      retMsgOper2 = MsgBox("OPERA��O COM FINALIDADE 4 (Devolu��o de mercadoria) NECESSITA QUE SEJA INFORMADA A 'Chave da Nota Fiscal Original de compra'." & vbCrLf & vbCrLf & "DESEJA PROSSEGUIR MESMO ASSIM? ", vbYesNo, "Aten��o")
  
      If retMsgOper2 = vbNo Then
          txt_chave1.SetFocus
          Exit Function
      End If
  End If
  ' ********************
  
  
  ' TRATAMENTO PARA OPERACOES DE SAIDA ENVOLVENDO EMPRESTIMO...NA QUEST�O ESTOQUE
  If Not IsNull(Num_Registro) Then
    Dim rsOperSaidaAuxiliar As Recordset
    Dim sSqlAux As String
    sSqlAux = "Select Estoque From [Opera��es Sa�da] "
    sSqlAux = sSqlAux + " Where C�digo = " + cboOper.Text
  
    Set rsOperSaidaAuxiliar = db.OpenRecordset(sSqlAux, dbOpenDynaset)
  
    With rsOperSaidaAuxiliar
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        If .Fields("Estoque").Value = True And rsSaidas("Observa��es") = "Venda gerada por Empr�stimo (MovEst)" Then
          ' A opera��o utilizada na realiza��o do emprestimo j� diminuiu o estoque dos produtos
          ' e esta opera��o que esta selecionada nesta venda TAMB�M est� configurada para diminuir estoque.
          ' Ent�o isto n�o pode ocorrer pois ira furar o estoque.
          MsgBox " A opera��o utilizada na realiza��o do empr�stimo j� diminuiu o estoque dos produtos. A opera��o que esta selecionada TAMB�M esta configurada para diminuir estoque. Escolha uma opera��o que n�o baixe o estoque NOVAMENTE."
          .Close
          Set rsOperSaidaAuxiliar = Nothing
          Exit Function
        End If
      End If
      .Close
    End With
    Set rsOperSaidaAuxiliar = Nothing
  End If
  ' FIM TRATAMENTO PARA OPERACOES DE SAIDA ENVOLVENDO EMPRESTIMO...NA QUEST�O ESTOQUE
  
  
  Rem Verifica os dados digitados
'  cboOper_LostFocus
  cboCliente_LostFocus
  
  If Nome_Opera��o.Caption = "" Then
    DisplayMsg "Opera��o incorreta, verifique."
    cboOper.SetFocus
    Exit Function
  End If
  
  If Nome_Digitador.Caption = "" Then
     DisplayMsg "Vendedor incorreto, verifique."
     cboDigitador.SetFocus
     Exit Function
  End If
  
  If Nome_Operador.Caption = "" Then
     DisplayMsg "Operador incorreto, verifique."
     Combo_Operador.SetFocus
     Exit Function
  End If
  
  '-------------------------------------------------------------------------
  '18/09/2002 - mpdea
  'Inclu�do/modificado verifica��o para cliente inativo e bloqueado
  'Alterado mensagem para cliente n�o localizado
  If Nome_Cliente.Caption = "" Then
    DisplayMsg "Cliente inativo, bloqueado ou inexistente."
    If cboCliente.Enabled = True Then cboCliente.SetFocus
    Exit Function
  End If
  
  If rsCliFor("Bloqueado") Then
    DisplayMsg "Este cliente est� bloqueado, imposs�vel gravar."
    If cboCliente.Enabled = True Then cboCliente.SetFocus
    Exit Function
  End If
  
  If rsCliFor("Inativo") Then
    DisplayMsg "Este cliente est� inativo, imposs�vel gravar."
    If cboCliente.Enabled = True Then cboCliente.SetFocus
    Exit Function
  End If
  
  '11/12/07 - Celso
  'Se o cliente tem contas em atraso, exige senha do gerente para continuar com a venda
   If rsParametros.Fields("ExigeSenhaGerVndContaAtraso").Value Then
      If Not m_blnSenhaGerJaInformada Then
         Dim Total_atrasado As Double
         Total_atrasado = Pega_Atrasado_Cliente(cboCliente.Text)
         If Total_atrasado > 0 Then
            DisplayMsg "Cliente [" & rsCliFor.Fields("Nome").Value & "] tem contas em atraso."
            'Senha do gerente
            If Not frmGerente.gbSenhaGerente Then
               Exit Function
            End If
            m_blnSenhaGerJaInformada = True
            m_strCodigoClienteContas = cboCliente.Text
         End If
      End If
   End If
  '------------------------------------------------------
  
  If Not IsNull(Num_Registro) Then
    If rsSaidas("Nota Cancelada") = True Then
      DisplayMsg "A nota fiscal desta movimenta��o j� foi cancelada. A movimenta��o n�o pode ser alterada."
      Exit Function
    End If
  End If
  
  If Not IsNull(Combo_Pre�o.Text) Then
    If Len(Combo_Pre�o.Text) > 15 Then
     DisplayMsg "Tabela de pre�os incorreta, verifique."
     Exit Function
    End If
  End If

  Conta = 0
  For i = 0 To Linhas_Grade
   If Tabe(i).C�digo <> "" Then
     If Tabe(i).Qtde <> 0 Then
       Conta = 1
       Exit For
     End If
   End If
  Next i
  
  If Combo_Pre�o.Text = "" And Conta > 0 Then
    DisplayMsg "Tabela de pre�os incorreta, verifique."
    Combo_Pre�o.SetFocus
    Exit Function
  End If
  
  '---------------------------------------------------------------------------------
  '07/05/2002 - mpdea
  '
  'Alterado verifica��o da exist�ncia da tabela de pre�os para opera��es do tipo
  'WEB (tabela de pre�os din�mica [DB:Pre�os])
  'Somente verifica se WebOrderFormID = 0 (venda n�o WEB)
  '<<-------------------------------------------------------------------------------
  If Not IsNull(Num_Registro) Then
    If Not CLng("0" & rsSaidas.Fields("WebOrderFormID").Value) > 0 Then
      If Len(Combo_Pre�o.Text) > 0 Then
        rsTabelas.Index = "Tabela"
        rsTabelas.Seek "=", Combo_Pre�o.Text
        If rsTabelas.NoMatch Then
           DisplayMsg "Tabela de pre�os n�o existe, verifique."
           Combo_Pre�o.SetFocus
           Exit Function
        End If
      End If
    End If
  End If
  '------------------------------------------------------------------------------->>
  
  If F_Empr�stimo.Visible = True Then
    If Not IsDate(Data_Acerto.Text) Then
      DisplayMsg "Digite a data de acerto para este empr�stimo."
      Data_Acerto.SetFocus
      Exit Function
    End If
  End If
  
  If Nome_Caixa.Caption = "" Then
    If Combo_Caixa.Enabled = False Then
      '20/11/2002 - mpdea
      'Inclu�do mensagem informando erro na grava��o
      DisplayMsg "Caixa n�o informado, assumindo Caixa 1. Execute a opera��o novamente."
      Combo_Caixa.Text = 1
      Combo_Caixa_LostFocus
    Else
      DisplayMsg "Digite o caixa."
      Combo_Caixa.SetFocus
    End If
    Exit Function
  End If
  
  rsOperadores.Index = "C�digo"
  rsOperadores.Seek "=", Combo_Operador.Text
  If rsOperadores.NoMatch Then
    DisplayMsg "Operador incorreto."
    Combo_Operador.SetFocus
    Exit Function
  End If
  
'  If rsOp_Sa�da("Dinheiro") = True Then

'  Mauro - 12/08/2021 Pesquisando como deixar opcionala exig�ncia de senha na tela SA�das

    If rsOperadores("ValorP") <> CriptografaSenha(Senha.Text) Then
      DisplayMsg "Senha incorreta."
      Senha.SetFocus
      Exit Function
    End If
    
'  End If

'  If rsParametros("VR Verifica Limite") = True And rsCliFor("Limite Cr�dito") <> 0 And rsOp_Sa�da("Tipo") = "V" Then
'    Limite_Usado = Pega_Limite_Usado(rsCliFor("C�digo"))
'    If (Limite_Usado + Retorna_Valor(L_Tot_Pagar.Text)) > rsCliFor("Limite Cr�dito") Then
'      M�ximo = rsCliFor("Limite Cr�dito") - Limite_Usado
'      DisplayMsg "Limite de cr�dito excedido. N�o � poss�vel vender. Venda m�xima = " + Format(M�ximo, "###,###,##0.00")
'      Exit Function
'    End If
'  End If

  Conta = 0
  For i = 0 To Linhas_Grade
    If Tabe(i).C�digo <> "" Then
     If Tabe(i).Qtde <> 0 Then
       Conta = Conta + 1
       Exit For
     End If
    End If
  Next i
  
  If Conta = 0 Then
    For i = 0 To Linhas_Servi�o
      If Tabe_Serv(i).C�digo <> 0 Then
        Conta = Conta + 1
        Exit For
      End If
    Next i
  End If
  
  If Conta = 0 Then
    DisplayMsg "Nenhum produto/servi�o digitado ou quantidades zeradas, imposs�vel gravar."
    Grade1.SetFocus
    Exit Function
  End If
    
  If rsOp_Sa�da("Senha") Then
    '14/10/2013 - Jean e Eduardo
    'Customiza��o para cliente Disk Embalagens para s� pedir senha gerente depois da sequencia j� estiver gravada no banco de dados
    If CheckSerialCaseMod("QS73520-469") Then
      If (txtSeq.Text = "") Then
      GoTo Sai_Senha:
      Else
        If Not frmGerente.gbSenhaGerente Then
          Exit Function
        End If
      End If
    End If
    
    If Not frmGerente.gbSenhaGerente Then
      Exit Function
    End If
  End If
  
Sai_Senha:
  
  '05/02/2004 - Maikel
  '---[ Colocado este IF pois o sistema estava deixando gravar venda ou nota de devolu��o a fornecedores com valor total igual ou menor que 0 (zero) ]---'
    If rsOp_Sa�da("Dinheiro") Or rsOp_Sa�da("Nota") Then
      If CDbl(L_Tot_Pagar.Text) <= 0 Then
        MsgBox "O valor total da nota fiscal est� incorreto, verifique !", vbCritical, "Quick Store"
        Exit Function
      End If
    End If
  '---[ Colocado este IF pois o sistema estava deixando gravar venda ou nota de devolu��o a fornecedores com valor total igual ou menor que 0 (zero) ]---'
  
  
  '=======================================================================================
  '07/11/2002 - mpdea
  'Vari�vel mcurDescontoSubTotal n�o estava inclu�da na verifica��o do desconto m�ximo
  
  'Tratamento Jun/2019 para verifiar limite de desconto pelo operador (e n�o pelo VENDEDOR)
  
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", Val(Combo_Operador.Text)
  If rsFuncionarios.NoMatch Then Exit Function
  
  'Verifica a aplica��o do desconto, de acordo com o limite do funcion�rio
  nPercMaxDesc = IIf(rsFuncionarios("nPercDesconto") = 0, _
    rsParametros("VR Desconto"), rsFuncionarios("nPercDesconto"))
  '19/01/2007 - Anderson
  'cDescMax = (Total_Pagar + Total_Desconto + mcurDescontoSubTotal) * nPercMaxDesc / 100
  cDescMax = Format((Total_Pagar + Total_Desconto + mcurDescontoSubTotal) * nPercMaxDesc / 100, "0.00")
  
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", Val(cboDigitador.Text) 'vendedor
  If rsFuncionarios.NoMatch Then Exit Function
  
  '20/09/2002 - mpdea
  'Inclu�do o Desconto no SubTotal
  If Total_Desconto + mcurDescontoSubTotal > cDescMax Then
    DisplayMsg "Desconto superior ao permitido."
    Exit Function
  End If
  '=======================================================================================
  
  
  
  '19/08/2003 - mpdea
  'Modificado nome do campo
  '
  '09/10/2002 - mpdea
  'Verifica estoque conforme configura��es
  If Not rsParametros.Fields("Venda Sem Estoque Saidas").Value And rsOp_Sa�da.Fields("Estoque").Value Then
    If Not mblnCheckStock() Then Exit Function
  End If
  
  
  '21/11/2002 - mpdea
  'Verifica se o registro pode ser alterado (somente para o estado AM)
  If UCase(gstrGetEstadoFilial(gnCodFilial)) = "AM" Then
    If Not IsNull(Num_Registro) Then
      If rsSaidas.Fields("Locked").Value Then
        DisplayMsg "Movimenta��o bloqueada para grava��o."
        Exit Function
      End If
    End If
  End If
  
  
  '08/08/2002 - mpdea
  'Inclus�o da verifica��o para or�amento do nr. do terminal
  bytNrTerminal = 0
  If rsOp_Sa�da.Fields("Tipo").Value = "O" Then
    
'''    If Not IsDataType(dtByte, txtNrTerminal.Text, bytNrTerminal) Then
'''      DisplayMsg "Preencher com o n�mero exclusivo deste Terminal para as Opera��es de Or�amento."
'''      Call SelectAllText(txtNrTerminal, True)
'''      Exit Function
'''    End If
'''    If bytNrTerminal = 0 Or bytNrTerminal > 99 Then
'''      DisplayMsg "N�mero do Terminal deve estar entre 1 e 99."
'''      Call SelectAllText(txtNrTerminal, True)
'''      Exit Function
'''    End If
    
    '-------------------------------------------------------------------------------
    '19/11/2002 - mpdea
    'Inclu�do flag para exibi��o de mensagem quando houver cria��o de um novo
    'registro de or�amento
    '
    '12/11/2002 - mpdea
    'Alterado o tratamento conforme nova solicita��o de cliente (Yanco Norte)
    '(Somente para o Estado de Manaus)
    'A cada atualiza��o de registro, gera novo or�amento
    If UCase(gstrGetEstadoFilial(gnCodFilial)) = "AM" Then
      If Not IsNull(Num_Registro) Then
        blnShowMessageNewBudget = True
      End If
      Num_Registro = Null
    End If
    
    If IsNull(Num_Registro) Then 'Novos or�amentos
      'Obt�m o n�mero do pr�ximo or�amento
      lngNrOrcamento = glngNextNrOrcamento(CByte(gnCodFilial))
      'Valida a fun��o anterior
      If lngNrOrcamento = -1 Then
        Exit Function
      End If
    Else 'Or�amentos gravados anteriores a vers�o 6.0.45
      If rsSaidas.Fields("InfoNrOrcamento").Value & "" = "" Then
        'Obt�m o n�mero do pr�ximo or�amento
        lngNrOrcamento = glngNextNrOrcamento(CByte(gnCodFilial))
        'Valida a fun��o anterior
        If lngNrOrcamento = -1 Then
          Exit Function
        End If
      End If
    End If
    '-------------------------------------------------------------------------------
    
    'Salva o nr. do terminal no registro
    Call SaveSetting("QuickStore", "ConfigSAIDAS", "NrTerminal", bytNrTerminal)
  End If
    
  '27/04/2004 - Daniel
  'Caso for PSV e a fun��o VerificaSeExisteValidade estiver True
  'o campo de validade dever� estar preenchido
  If m_blnPSV Then
    If VerificaSeExisteValidade Then
      blnValidade = True 'Acendemos o flag pois o conte�do do objeto pode estar v�lido...

      If Not ValidaCampoValidade Then Exit Function
    Else
      blnValidade = False
    End If
  End If
  
  '02/05/2005 - Daniel
  '
  'Solicitante..: Jorge Marcos - PSI MT
  '
  'Finalidade...: Verificar o limite de cr�dito do cliente antes da grava��o
  '               Isto � essencial para todas as empresas que trabalham
  '               com pronta entrega
  If Len(Nome_Cliente.Caption) > 0 Then
    If rsParametros("VerificaLimiteCli").Value Then
      Dim dblLimiteCli     As Double
      Dim dblLimiteCredito As Double
      
      Call GetLimiteCliente(cboCliente.Text, dblLimiteCli)
      
      dblLimiteCredito = Format(dblLimiteCli - Pega_Limite_Usado(cboCliente.Text), FORMAT_VALUE)
      
      If ((L_Tot_Pagar.Text) > dblLimiteCredito) Or ((L_Tot_Pagar.Text) > dblLimiteCli) Then
        MsgBox "O cliente ao qual voc� est� fazendo a venda tem R$ " & _
               Format(dblLimiteCredito, FORMAT_VALUE) & " de saldo para novas compras. O recebimento estar� sendo de R$ " & _
               Format(L_Tot_Pagar.Text, FORMAT_VALUE) & ". N�o � possivel continuar !! ", vbCritical, "Quick Store"
        
        Exit Function
      End If
    End If 'If rsParametros("VerificaLimiteCli").Value
  End If
  
  
  '20/05/2005 - Daniel
  '
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '
  'O sistema dever� julgar se a nota fiscal ser� criada
  'automaticamente ou manualmente a partir da opera��o escolhida
  If gbNotaManual(CInt(Trim(cboOper.Text)), "SAIDA") Then
    If Len(txtNF.Text) < 1 Or (txtNF.Text) = "0" Then   'N�o preencheu corretamente...
      'Para a sa�da, o campo nota fiscal tornou-se obrigat�rio
      'devido a atualiza��o que ocorre no Contas a Receber
      MsgBox "Preencha o campo Nota Fiscal.", vbExclamation, "Aten��o"
      txtNF.SetFocus
      Exit Function
    End If

    '01/08/2005 - Daniel
    'Tratamento para o Campo Sa�das.DataEmissaoNotaManual
    'este campo tornou-se obrigat�rio pois estar� envolvido
    'em gera��es de arquivos atrav�s do InfoICMS
    If Not IsDate(mskDataEmissaoNotaManual.Text) Then
      MsgBox "Informe a Data de Emiss�o da Nota Manual.", vbExclamation, "Quick Store"
      mskDataEmissaoNotaManual.SetFocus
      Exit Function
    End If
    
  End If
  
  '10/06/2005 - Daniel
  '
  '----------[ Finalidade da valida��o abaixo ]----------
  'Caso ocorra venda de servi�o, iremos checar se o T�cnico n�o
  'foi mencionado e avisaremos o usu�rio que a comiss�o n�o ser�
  'identificada a qual T�cnico ela pertence
  '
  With Grade_Serv
    .MoveFirst
    If IsNumeric(.Columns("C�digo").Text) And (.Columns("C�digo").Text) <> "0" Then
      If Len(Nome_T�cnico.Caption) <= 0 Then
        Dim strMsg As String

        strMsg = ""
        strMsg = "O campo T�cnico n�o foi preenchido. O valor da(s) comiss�o(�es) sobre" & vbCrLf
        strMsg = strMsg & "servi�o(s) n�o ser�(�o) calculado(s). Deseja continuar ?"

        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Aten��o") = vbNo Then Exit Function
      End If
    End If
  End With
  
  '18/01/2007 - Anderson
  'Solicita��o senha do gerente ao alterar o vendedor relacionado ao cliente
  If rsParametros("VendedorSenhaGerente").Value Then
    If rsCliFor("Vendedor") <> 0 And rsCliFor("Vendedor") <> cboDigitador.Text Then
      If MsgBox("O c�digo do vendedor n�o corresponde ao cliente selecionado. A senha do gerente ser� necess�ria para concluir a grava��o da venda." & Chr(13) & "Deseja continuar assim mesmo?", vbYesNo + vbQuestion, "Aten��o") = vbYes Then
        If Not frmGerente.gbSenhaGerente Then
          Exit Function
        End If
      Else
        Exit Function
      End If
    End If
  End If
  
  '30/04/2008 - mpdea
  'Verifica n�mero de documento do cliente
  Dim str_numero_documento_cliente As String
  If Not IsNull(Num_Registro) Then
    str_numero_documento_cliente = rsSaidas.Fields("NumeroDocumentoCliente").Value & ""
  End If
  str_numero_documento_cliente = g_str_GetNumeroDocumento(CInt(cboOper.Text), CLng(cboCliente.Text), str_numero_documento_cliente)

  
  '--------------------------------------------------------------------------
  'UPDATE RECORD
  '--------------------------------------------------------------------------
  
  
  'Inicia transa��o
  Call ws.BeginTrans: blnInTransaction = True
  
  'Pega n�mero da nova movimenta��o
  If IsNull(Num_Registro) Then
'    Book_Par = rsParametros.Bookmark
'    rsParametros.Bookmark = Book_Par
    nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("�ltima Movimenta��o") + 1
    rsParametros.Edit
    rsParametros("�ltima Movimenta��o") = nSequencia
    rsParametros.Update
  End If
  
  Call StatusMsg("Gravando sa�da ...")
  
  If gbLogError = True Or CDate(Data_Atual) <> CDate(Date) Then  'grava log
    rsLog.AddNew
      rsLog("Tipo") = "DT WIN <> DT QUICK"
      rsLog("Data") = Date
      Aux_Texto = "Filial " + str(gnCodFilial) + " Seq.: " + CStr(nSequencia)
      Aux_Texto = Aux_Texto + " DTQUICK: " + CStr(Data_Atual) + " DTWIN: " + CStr(Date) + " DTMOV: " + L_Dia.Caption
      rsLog("Texto") = Left(Aux_Texto, 80)
    rsLog.Update
  End If
  
  If IsNull(Num_Registro) Then
    rsSaidas.AddNew
    sMsg = "inserida"
    rsSaidas("Filial") = gnCodFilial
    rsSaidas("Sequ�ncia") = nSequencia
    txtSeq.Text = ""
  Else
    rsSaidas.LockEdits = True
    rsSaidas.Edit
    sMsg = "alterada"
    nSequencia = Val(txtSeq.Text)
  End If
  
  ' ======================================================
  ' Gravar Chaves referenciadas
  If nSequencia > 0 Then
    If gridChaves.Rows > 1 Then
        Dim nChaves As Integer
        db.Execute "Delete from SaidasChaves where Filial = " & gnCodFilial & " and Sequencia=" & nSequencia
        
        For nChaves = 1 To gridChaves.Rows - 1
            db.Execute "Insert into SaidasChaves (Filial, Sequencia, Chave) values (" & gnCodFilial & "," & nSequencia & ",'" & gridChaves.TextMatrix(nChaves, 1) & "')"
        Next
    End If
  End If
  ' ======================================================
  
  ' ======================================================
  ' Gravar Comandas
  ' ======================================================
  If nSequencia > 0 Then
    If rsParametros("TrabalharComComanda").Value Then
      frmComanda.Comanda = txtComanda.Text
      If Trim(txtComanda.Text) = "" Then
        frmComanda.Deleta (nSequencia)
      ElseIf frmComanda.Existe(nSequencia) = False Then
        Dim rsSaidasComandas As Recordset
        Set rsSaidasComandas = db.OpenRecordset("SaidasComandas")
        rsSaidasComandas.AddNew
        rsSaidasComandas("CodComanda") = txtComanda.Text
        rsSaidasComandas("CodSaida") = nSequencia
        rsSaidasComandas("Filial") = gnCodFilial
        rsSaidasComandas.Update
        rsSaidasComandas.Close
      End If
    End If
  End If
  ' ======================================================

  rsSaidas("Data") = L_Dia.Caption
  '04/08/2005 - Daniel
  'Tratamento para o Campo Sa�das.DataEmissaoNotaManual
  'Solicitante: Ped�gio Cal�ados e Confec��es
  'Referente ao Projeto Impress�o de Notas Manuais
  If IsDate(mskDataEmissaoNotaManual.Text) Then
    rsSaidas("DataEmissaoNotaManual").Value = Format(mskDataEmissaoNotaManual.Text, "DD/MM/YYYY")
  End If
  '-----------------------------------------------------
  rsSaidas("Tabela") = Combo_Pre�o.Text
  rsSaidas("Refer�ncia") = txtRef.Text
  rsSaidas("Opera��o") = Val(cboOper.Text)
  rsSaidas("Digitador") = Val(cboDigitador.Text)
  rsSaidas("Operador") = Val(Combo_Operador.Text)
  rsSaidas("Cliente") = Val(cboCliente.Text)
  rsSaidas("Observa��es") = Obs.Text
  
  '30/04/2008 - mpdea
  'N�mero de documento do cliente
  rsSaidas.Fields("NumeroDocumentoCliente").Value = str_numero_documento_cliente
  
  rsSaidas("Caixa") = Val(Combo_Caixa.Text)
  '23/07/2004 - Daniel
  'Altera��o: estes campos ser�o utilizados somente para o
  'faturamento da STC, para os demais clientes dever� sempre
  'ser 0 (zero)
  rsSaidas("Num Autorizacao").Value = 0
  rsSaidas("MesX").Value = 0
  '---------------------------------------------------------
  rsSaidas("Produtos") = Retorna_Valor(L_Tot_Prod.Text)
  rsSaidas("Desconto") = Retorna_Valor(L_Tot_Desc.Text)
   
  '19/04/2004 - Daniel
  'Populando o field Data Validade
  'Case: PSV Inform�tica
  If blnValidade Then
    rsSaidas.Fields("Data Validade").Value = Trim(mskValidade.Text)
    blnValidade = False
  End If
  
  '24/05/2004 - Daniel
  'Case: Bic Amaz�nia
  'Populando os campos Sa�das.[Codigo Func Comprador] = 0 e
  'Sa�das.[Status Venda Func] = False pois eles ter�o outro
  'valor somente no crit�rio usado em na tela de venda r�pida
  rsSaidas("Codigo Func Comprador").Value = 0
  rsSaidas("Status Venda Func").Value = False
  '--------------------------------------------------------------
  
  '23/04/2004 - Daniel
  'O campo FaturaSourceReserva sempre ser� False at� o momento
  'que a partir dele seja clonado uma sa�da para venda
  rsSaidas.Fields("FaturaSourceReserva").Value = False
  '--------------------------------------------------------------
  
  '20/09/2002 - mpdea
  'Desconto no SubTotal
  rsSaidas("DescontoSubTotal") = mcurDescontoSubTotal

  rsSaidas("IPI") = Retorna_Valor(L_Tot_IPI.Text)
  rsSaidas("Frete") = Retorna_Valor(L_Frete.Text)
  
  ' Pilatti Novembro 28-11-2017
  rsSaidas("FreteSomaOuNaoEstimativa") = chk_freteNaoSomaPercentual.Value
    
  '12/04/2005 - Daniel
  'Adicionado o campo Seguro
  rsSaidas("Seguro").Value = Retorna_Valor(txtSeguro.Text)
  
  rsSaidas("Base ICM") = Retorna_Valor(L_Base_ICM.Text)
  rsSaidas("Valor ICM") = Retorna_Valor(L_Valor_ICM.Text)
  rsSaidas("Base ICM Subs") = Retorna_Valor(L_Base_ICM_Subs)
  rsSaidas("Valor ICM Subs") = Retorna_Valor(L_Valor_ICM_Subs)
  rsSaidas("Total") = Retorna_Valor(L_Tot_Pagar.Text)
  rsSaidas("Servi�os") = Retorna_Valor(L_Tot_Serv.Text)
  rsSaidas("TotalDesoneracaoICMS") = Retorna_Valor(L_Tot_ICMS_Deson.Text)
  '13/05/2004 - Daniel
  'Campos de tratamento de impostos sobre servi�os
  'percentuais e totais de CSLL, PIS, COFINS, IRRF
  rsSaidas("Percentual CSLL").Value = m_sngPercentualCSLL
  rsSaidas("Percentual COFINS").Value = m_sngPercentualCOFINS
  rsSaidas("Percentual PIS").Value = m_sngPercentualPIS
  rsSaidas("Percentual IRRF").Value = m_sngPercentualIRRF
  rsSaidas("Total CSLL").Value = m_dblTotalCSLL
  rsSaidas("Total COFINS").Value = m_dblTotalCOFINS
  rsSaidas("Total PIS").Value = m_dblTotalPIS
  rsSaidas("Total IRRF").Value = m_dblTotalIRRF
  
  
  rsSaidas("TotalMenosServ").Value = m_dblTotalMenosServ
  
  '18/05/2005 - Daniel
  'Tratamento para o campo N� da NF
  rsSaidas("SerieNF").Value = UCase(CStr(txtNrSerieNF.Text & ""))
  
  '17/09/2009 - mpdea
  'Modelo de documento fiscal
  rsSaidas.Fields("ModeloDocumentoFiscal").Value = gstrGetModeloDocumentoFiscalOperacao(tmSaidas, rsSaidas.Fields("Opera��o").Value)
  
  rsSaidas.Fields("Consumidor_Final").Value = Left(cboConsumidorFinal.Text, 1)
  rsSaidas.Fields("Presenca_Comprador").Value = Left(cboPresencaComprador.Text, 1)
  rsSaidas.Fields("FinalidadeNFe").Value = Left(cboFinalidade.Text, 1)
  
  
  '15/07/2016 Michel
  '4% al�q. inter. p/ prod. importados
  '7% para os Estados de origem do Sul e Sudeste
  '12% para os demais casos.

  'If cmbAliquotaOrigem.Text = "4% al�q. inter. p/ prod. importados" Then
  '  rsSaidas.Fields("aliquota_origem").Value = 4
  'ElseIf cmbAliquotaOrigem.Text = "7% para os Estados de origem do Sul e Sudeste" Then
  '  rsSaidas.Fields("aliquota_origem").Value = 7
  'Else
  '  rsSaidas.Fields("aliquota_origem").Value = 12
  'End If
  
  rsSaidas.Fields("aliquota_origem").Value = aliquotaICMS_tab_ICMS_PERC_ESTADOS
  
  Dim strSQL As String
  strSQL = "SELECT ALIQUOTA FROM ICMS_PERCENTUAL_ESTADOS "
  strSQL = strSQL & "WHERE ESTADO_ORIGEM = ESTADO_DESTINO AND ESTADO_DESTINO = '" & rsCliFor("Estado").Value & "';"
  Dim rsAliqDest As Recordset
  Set rsAliqDest = db.OpenRecordset(strSQL)
  If Not (rsAliqDest.BOF And rsAliqDest.EOF) Then
    rsAliqDest.MoveFirst
    rsSaidas.Fields("aliquota_destino").Value = Retorna_Valor(rsAliqDest("ALIQUOTA").Value)
  End If
  rsAliqDest.Close
  Set rsAliqDest = Nothing
  
  '
  '20/05/2005 - Daniel
  '
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '
  'O sistema dever� julgar se a nota fiscal ser� criada
  'automaticamente ou manualmente a partir da opera��o escolhida
  '
  'Tratamento para o campo [Nota Fiscal]
  If gbNotaManual(CInt(Trim(cboOper.Text)), "SAIDA") Then
    rsSaidas("Nota Fiscal").Value = CLng(Trim("0" & txtNF.Text & ""))
  Else
    rsSaidas("Nota Fiscal").Value = 0
  End If
  '
  '-------------------------------------------------
  
  
  Call ZerarVarsImpostosServi�os
  '-------------------------------------------------
    
  '23/07/2004 - Daniel
  'Havia um bloco de if else...
  'Caso o cliente fosse STC e a Sa�da tivesse origem na tela de
  'Programa��es, for�ar�amos o c�lculo do Valor ISS para STC
  rsSaidas("Valor ISS") = Retorna_Valor(L_Tot_ISS.Text)
  
  rsSaidas("Prometido Para") = Prometido_Para.Text
  rsSaidas("Or�amento Aprovado") = Or�amento_Aprovado.Text
  
  If IsDate(Data_Acerto.Text) Then
    rsSaidas("Data Acerto Empr�stimo") = Data_Acerto.Text
  End If
  
  
  rsSaidas("T�cnico") = 0
  If Nome_T�cnico.Caption <> "" Then
    rsSaidas("T�cnico") = Combo_T�cnico.Text
  End If
  
  
  '08/08/2002 - mpdea
  'Informa o nr. do or�amento e do terminal
  If rsOp_Sa�da.Fields("Tipo").Value = "O" Then
    If rsSaidas.Fields("InfoNrOrcamento").Value & "" = "" Then
      rsSaidas.Fields("InfoNrOrcamento").Value = _
        "Or�amento nr. " & Format(lngNrOrcamento, "000,000") & "/" & Format(bytNrTerminal, "00")
    End If
  End If
  
  bSeqChanged = False
  bSequencia = True
  rsSaidas.Update
  rsSaidas.LockEdits = False
  bSequencia = False
  'Grava novamente a �ltima movimenta��o
  'se a mesma foi alterada
  If bSeqChanged Then
    With rsParametros
      .Edit
      .Fields("�ltima Movimenta��o") = nSequencia
      .Update
    End With
  End If
  
  Num_Registro = rsSaidas.LastModified
  rsSaidas.Bookmark = Num_Registro
  
  'Apaga produtos
  Call EraseTypeMoviment(tmSaidasProdutos, gnCodFilial, nSequencia)
  'Apaga Servi�os
  Call EraseTypeMoviment(tmSaidasServicos, gnCodFilial, nSequencia)
    
  Rem Grava Produtos
  Conta = 1
  blnOrcamentoAprovado = True
  
  For i = 0 To 254
    If Tabe(i).C�digo <> "" Then
      If Tabe(i).Qtde <> 0 Then
        If Tabe(i).Nome <> "" Then
        
'        Para corrigir, precisa mover o c�digo abaixo:
'
'    rsProdutos2.FindFirst ("C�digo = '" & UCase(Tabe(i).C�digo) & "'")
'    If Not IsNull(rsProdutos2("Unidade Venda")) Then
'      sUnidade = rsProdutos2("Unidade Venda")
'    Else
'      sUnidade = " "
'    End If
'    If Not IsNull(rsProdutos2("Situa��o Tribut�ria")) Then
'      sTributaria = rsProdutos2("Situa��o Tribut�ria")
'    Else
'      sTributaria = " "
'    End If
'
'Para depois de (onde � obtido o c�digo sem grade):
'
'    Aux_Cod_Prod = Tabe(i).C�digo
'    Acha_Produto(Aux_Cod_Prod, Aux_Prod, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro)
'
'    rsSaidas_Prod("C�digo Sem Grade") = Aux_Prod
'
'E alterar a pesquisa para o c�digo principal:
'
'    rsProdutos2.FindFirst ("C�digo = '" & Aux_Prod & "'")

          '26/08/2010 - Andrea
          Aux_Cod_Prod = Tabe(i).C�digo
          Acha_Produto Aux_Cod_Prod, Aux_Prod, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro

          'Alterada a pesquisa para o c�digo principal
          'rsProdutos2.FindFirst "C�digo = '" & UCase(Tabe(i).C�digo) & "'"
          rsProdutos2.FindFirst ("C�digo = '" & Aux_Prod & "'")
          If Not IsNull(rsProdutos2("Unidade Venda")) Then
             sUnidade = rsProdutos2("Unidade Venda")
          Else
             sUnidade = " "
          End If
          If Not IsNull(rsProdutos2("Situa��o Tribut�ria")) Then
             sTributaria = rsProdutos2("Situa��o Tribut�ria")
          Else
             sTributaria = " "
          End If
          rsSaidas_Prod.AddNew
            rsSaidas_Prod("Filial") = gnCodFilial
            rsSaidas_Prod("Sequ�ncia") = nSequencia
            rsSaidas_Prod("Linha") = Conta
            rsSaidas_Prod("C�digo") = UCase(Tabe(i).C�digo)
            rsSaidas_Prod("QtdeEntregue") = 0     'Novo campo para entregas
            rsSaidas_Prod("Qtde") = Tabe(i).Qtde
            rsSaidas_Prod("Pre�o") = Tabe(i).Pre�o
            rsSaidas_Prod("Desconto") = Tabe(i).Desconto
            rsSaidas_Prod("ICM") = Tabe(i).ICM
            rsSaidas_Prod("IPI") = Tabe(i).IPI
            rsSaidas_Prod("Etiqueta") = False
            rsSaidas_Prod("Descricao Adicional") = Tabe(i).Descr_Adicional
            rsSaidas_Prod("ValorICMSDesonerado") = Tabe(i).Valor_Desonerado
            rsSaidas_Prod("Percentual_Diferimento") = Tabe(i).Percentual_Diferimento
            
            rsSaidas_Prod("Desp_Acessorias") = Tabe(i).Desp_Acessorias
            
            If rsSaidas_Prod("Pre�o") = 0 Then
                bProdutoSemPrecoNaGrade = True
            End If
            
            
            If Tabe(i).Etiqueta = "-1" Then rsSaidas_Prod("Etiqueta") = True
            
            '28/10/2004 - Daniel
            'Tratamento para o field [Sa�das - Produtos].[Pre�o Final]
            'Para o cliente A.S. Wijma (Bel�m - Par�) dever� ser Double
            'para os demais clientes continua sendo Single
            If m_blnASWijmaBelem Then
              Call IsDataType(dtDouble, Tabe(i).Pre�o_Final, m_dblPrecoFinalAuxi)
              rsSaidas_Prod("Pre�o Final") = m_dblPrecoFinalAuxi
            Else
              rsSaidas_Prod("Pre�o Final") = Tabe(i).Pre�o_Final
            End If
                   
            '26/08/2010 - Andrea
            'Movido este c�digo para cima
            'Aux_Cod_Prod = Tabe(i).C�digo
            'Acha_Produto Aux_Cod_Prod, Aux_Prod, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro
          
            rsSaidas_Prod("C�digo Sem Grade") = Aux_Prod
            
            '27/05/2010 - mpdea
            'Atualiza CFOP ao gravar produto para corrigir o problema de CFOP zerado ou incorreto
            'rsSaidas_Prod("CFOP") = Tabe(i).CFOP_Produto '20/12/2006 - Anderson - Altera��o para o registro de CFOP por produto e servico
            rsSaidas_Prod("CFOP") = GetCfopProduto(Aux_Prod)
            
            If sUnidade = "" Or IsNull(sUnidade) Then
               sUnidade = "  "
               rsSaidas_Prod("Unidade Venda") = sUnidade
            Else
               rsSaidas_Prod("Unidade Venda") = sUnidade
            End If
            
            If sTributaria = "" Or IsNull(sTributaria) Then
               sTributaria = " "
               rsSaidas_Prod("Situa��o Tribut�ria") = sTributaria
            Else
               rsSaidas_Prod("Situa��o Tribut�ria") = sTributaria
            End If
            
            If blnOrcamentoAprovado Then
              Set rstPrecos = db.OpenRecordset("SELECT * FROM Pre�os WHERE Produto = '" & Aux_Prod & "' AND Tabela = '" & Combo_Pre�o.Text & "'", dbOpenDynaset, dbReadOnly)
              With rstPrecos
                If Not (.BOF And .EOF) Then
                  If (Tabe(i).Pre�o <> .Fields("Pre�o")) Then
                    blnOrcamentoAprovado = False
                  End If
                End If
                
                .Close
                Set rstPrecos = Nothing
              End With
            End If
            
            '09/08/2007 - Anderson
            'Altera��o realizada para armazenar o custo do produto no momento da venda
            rsSaidas_Prod("PrecoCusto") = gcGetPrecoProduto(rsSaidas_Prod("C�digo"), "CUSTO")
            
            ' Pilatti Junho/2018
            rsSaidas_Prod("CFOP") = Tabe(i).CFOP_Produto
            
            '************************
            'Trata tributos
            Call UpdateTotalNCM_2(rsSaidas_Prod("C�digo"))
            'Fim trata tributos
            
          rsSaidas_Prod.Update
          
          '13-04-2025 pablo
          If rsParametros("EditarNomeProduto").Value Then
            If UCase(Trim(Tabe(i).Nome)) <> UCase(Trim(CStr(rsProdutos2("Nome").Value))) Then
              Dim apagar As Boolean
              Dim QUERY As String
              QUERY = "SELECT Nome FROM ProdutoNomeNFe WHERE "
              QUERY = QUERY & "Filial = " & gnCodFilial & " AND "
              QUERY = QUERY & "Sequencia = " & nSequencia & " AND "
              QUERY = QUERY & "Codigo = '" & Trim(Tabe(i).C�digo) & "';"
              
              Dim rsNomeProd As Recordset
              Set rsNomeProd = db.OpenRecordset(QUERY, dbOpenSnapshot)
              If Not (rsNomeProd.BOF And rsNomeProd.EOF) Then
                rsNomeProd.MoveLast
                rsNomeProd.MoveFirst
                apagar = (rsNomeProd.RecordCount > 0)
              End If
              rsNomeProd.Close
              Set rsNomeProd = Nothing

              If apagar Then
                QUERY = "DELETE FROM ProdutoNomeNFe WHERE "
                QUERY = QUERY & "Filial = " & gnCodFilial & " AND "
                QUERY = QUERY & "Sequencia = " & nSequencia & " AND "
                QUERY = QUERY & "Codigo = '" & Trim(Tabe(i).C�digo) & "';"
                db.Execute QUERY
              End If
             
              QUERY = "INSERT INTO ProdutoNomeNFe (Filial, Sequencia, Codigo, Nome) VALUES "
              QUERY = QUERY & "(" & gnCodFilial & ", "
              QUERY = QUERY & nSequencia & ", "
              QUERY = QUERY & "'" & Trim(Tabe(i).C�digo) & "', "
              QUERY = QUERY & "'" & Trim(Tabe(i).Nome) & "');"
              db.Execute QUERY
            End If
          End If
          
          
          Conta = Conta + 1
          
        End If
      End If
    End If
  Next i
  
  If bProdutoSemPrecoNaGrade = True Then
      'frm_produtoSemPrecoNaGrade.Left = 4110
      'frm_produtoSemPrecoNaGrade.Top = 5580
      frm_produtoSemPrecoNaGrade.Visible = True
  Else
      frm_produtoSemPrecoNaGrade.Visible = False
  End If
  
  rsSaidas.Edit
      rsSaidas.Fields("OrcamentoAprovado").Value = blnOrcamentoAprovado
      If totalNCM_2 > 0 Then
          rsSaidas("TotalNCM") = totalNCM_2
      End If
  rsSaidas.Update
  
  
  Rem Grava Servi�os
  Conta = 1
  For i = 0 To 254
    If Tabe_Serv(i).C�digo <> 0 Then
      rsSaidas_Serv.AddNew
        rsSaidas_Serv("Filial") = gnCodFilial
        rsSaidas_Serv("Sequ�ncia") = nSequencia
        rsSaidas_Serv("Linha") = Conta
        rsSaidas_Serv("C�digo") = Tabe_Serv(i).C�digo
        rsSaidas_Serv("Descri��o") = Tabe_Serv(i).Descri��o
        rsSaidas_Serv("Tempo") = Tabe_Serv(i).Tempo
        rsSaidas_Serv("Pre�o") = Tabe_Serv(i).Pre�o
        rsSaidas_Serv("Completo") = Tabe_Serv(i).Completo
        rsSaidas_Serv("CFOP") = Tabe_Serv(i).CFOP_Servico '20/12/2006 - Anderson - Altera��o para o registro de CFOP por produto e servico
        '26/07/2005 - Daniel
        'Personaliza��o para a empresa J.R. Hidroqu�mica
        'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
        'C.S.T. (C�digo de Situa��o Tribut�ria)
        If m_blnJR Then
          If Len(Tabe_Serv(i).CST & "") = 1 Then
            rsSaidas_Serv("CST").Value = Tabe_Serv(i).CST & ""
          Else
            rsSaidas_Serv("CST").Value = ""
          End If
        Else
          rsSaidas_Serv("CST").Value = ""
        End If
        '--------------------------------------------------------------
      rsSaidas_Serv.Update
      Conta = Conta + 1
    End If
  Next i
  
  
  nRet = 0
  
  If rsOp_Sa�da("Dinheiro") = False And rsOp_Sa�da("Tipo") <> "O" Then
    Call StatusMsg("Aguarde, efetivando movimenta��o...")
    nRet = Efetiva_Sa�da(gnCodFilial, nSequencia)
    If nRet <> 0 Then
      Select Case nRet
        Case -1
          'A��o cancelada
          Call StatusMsg("A��o cancelada.")
        Case 5
          Call DisplayMsg("Tabela de pre�os inexistente.")
        Case Else
          Call DisplayMsg("Opera��o N�O efetivada. Erro" & str(nRet))
      End Select
      L_Efetivada.Visible = False
      'Cancelamento da transa��o
      ws.Rollback: blnInTransaction = False
    Else
      'Fim da transa��o
      ws.CommitTrans: blnInTransaction = False
      L_Efetivada.Visible = True
      Call StatusMsg("")
    End If
  Else
    'Somente grava a venda
    Call ws.CommitTrans: blnInTransaction = False
  End If
  
  If txtSeq.Text = "" Then
    txtSeq.Text = nSequencia
  End If
  
  '14/08/2002 - mpdea
  'Exibi��o do nr. do or�amento
  If rsSaidas.Fields("InfoNrOrcamento").Value & "" <> "" Then
    Me.Caption = "Sa�das - " & rsSaidas.Fields("InfoNrOrcamento").Value
  Else
    Me.Caption = "Sa�das"
  End If
  
  '20/09/2002 - mpdea
  'Registro atualizado, desativa flag para for�ar atualiza��o
  mblnForceUpdate = False
  
  Call StatusMsg("")
  
'  If L_Efetivada.Visible = True Then
'    DisplayMsg "OPERA��O EFETIVADA. Movimenta��o de Sa�da " & sMsg & " com sucesso."
'  Else
'    If nRet = 0 Then
'      DisplayMsg "Movimenta��o de Sa�da " & sMsg & " com sucesso."
''    Else
''      DisplayMsg "Movimenta��o N�O efetivada. Erro " & CStr(nRet)
'    End If
'  End If
  ActiveBar1.Tools("miComplRecebimento").Enabled = True
  
  '18/11/2002 - mpdea
  'Exibe aviso da cria��o da nova movimenta��o de or�amento
  If blnShowMessageNewBudget Then
    DisplayMsg "Nova movimenta��o de or�amento criada."
  End If
 
'---------------------------------------------------------------------------------------------------------------
' Joga dados da movimenta��o para o banco do GestoPDV por causa do PAF
'---------------------------------------------------------------------------------------------------------------
   Dim GestoBD As Database
   Dim SaidaEstoque As Recordset
   Dim SaidaEstoqueItem As Recordset
   Dim ItemEstoqueAlmox As Recordset
   Dim QuickBD As Database
   Dim produtos As Recordset
   Dim cad_prod As Recordset
   If frmParametros.VerificaPAF = True Then
     Set rsParametros = db.OpenRecordset("Select * from [Par�metros Filial] Where Filial = " & gnCodFilial & ";")
              
     Dim fso As New FileSystemObject
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists(rsParametros("BancoPDV").Value & "\Gesto.mde") Then

     Set GestoBD = OpenDatabase(rsParametros("BancoPDV").Value & "\Gesto.mde", False, False)
     Set SaidaEstoque = GestoBD.OpenRecordset("Select * from SaidaEstoque Where NUMERO = " & txtSeq.Text & ";")
     If SaidaEstoque.EOF Then
       SaidaEstoque.AddNew
       SaidaEstoque!Numero = txtSeq.Text
       SaidaEstoque!CODIGO_CLIENTE = cboCliente.Text
       SaidaEstoque!Cliente = Left(Nome_Cliente.Caption, 40)
       SaidaEstoque!DATA_SAIDA = L_Dia.Caption
       If Obs.Text <> "" Then
         SaidaEstoque!Observacao = Obs.Text
       End If
       If L_Tot_Desc.Text <> "" Then
         SaidaEstoque!VL_DESCONTO = L_Tot_Desc.Text
       End If
       SaidaEstoque!COD_Vendedor = cboDigitador.Text
       SaidaEstoque.Update
     Else
       SaidaEstoque.Edit
       SaidaEstoque!CODIGO_CLIENTE = cboCliente.Text
       SaidaEstoque!Cliente = Left(Nome_Cliente.Caption, 40)
       SaidaEstoque!DATA_SAIDA = L_Dia.Caption
       If Obs.Text <> "" Then
         SaidaEstoque!Observacao = Obs.Text
       End If
       If L_Tot_Desc.Text <> "" Then
         SaidaEstoque!VL_DESCONTO = L_Tot_Desc.Text
       End If
       SaidaEstoque!COD_Vendedor = cboDigitador.Text
       SaidaEstoque.Update
       Set SaidaEstoqueItem = GestoBD.OpenRecordset("Select * from SaidaEstoqueItem Where NUMERO = " & txtSeq.Text & "")
 
         Do Until SaidaEstoqueItem.EOF = True
            If SaidaEstoqueItem.EOF = False Then
              SaidaEstoqueItem.Delete
              SaidaEstoqueItem.MoveNext
            End If
         Loop
 
     End If
 
      'Continuar PAF saidas produtos
     Dim Nome_Prod As Recordset
     Dim Estoque_Prod As Recordset
     Set produtos = db.OpenRecordset("Select * from [Sa�das - Produtos] where Filial = " & gnCodFilial & " and Sequ�ncia = " & txtSeq.Text & ";")
     produtos.MoveFirst
     Do Until produtos.EOF
       Set cad_prod = db.OpenRecordset("Select * from Produtos where C�digo = '" & produtos("C�digo sem Grade") & "';")
       Set Nome_Prod = GestoBD.OpenRecordset("SELECT DESCRICAO From ItemEstoque WHERE CODIGO_FORNECEDOR =  '" & produtos("C�digo sem Grade") & "';")
       Set ItemEstoqueAlmox = GestoBD.OpenRecordset("Select * from ItemEstoqueAlmox Where Codigo_Item =  '" & produtos("C�digo sem Grade") & "';")
       If Nome_Prod.EOF Then
         MsgBox "O produto de c�digo: " & produtos("C�digo sem Grade") & " n�o esta cadastrado no Gesto, para que o erro n�o volte a ocorrer entre no cadastro do produto e mande gravar."
         Exit Function
       End If
       If cad_prod("Tipo") = "N" Then
         Set Estoque_Prod = db.OpenRecordset("Select [Estoque Atual] From [Estoque Final] where Filial = " & gnCodFilial & " and Produto =  '" & produtos("C�digo sem Grade") & "';")
         Set SaidaEstoqueItem = GestoBD.OpenRecordset("Select * from SaidaEstoqueItem Where NUMERO = " & txtSeq.Text & " AND CODIGO_ITEM =  '" & produtos("C�digo sem Grade") & "';")
          'If SaidaEstoqueItem.EOF Then
           SaidaEstoqueItem.AddNew
           SaidaEstoqueItem!Numero = txtSeq.Text
           SaidaEstoqueItem!Item = produtos("Linha")
           SaidaEstoqueItem!Codigo_Item = produtos("C�digo")
           SaidaEstoqueItem!DESCRICAO_ITEM = Nome_Prod("DESCRICAO")
           SaidaEstoqueItem!Quantidade = produtos("Qtde")
           SaidaEstoqueItem!VALOR_UNIT_DESC = produtos("Pre�o") - (produtos("Pre�o") * produtos("Desconto") / 100)
           SaidaEstoqueItem!Valor_Total = produtos("Pre�o Final")
           SaidaEstoqueItem.Update
           If Estoque_Prod.EOF Then
             MsgBox "O produto " & Nome_Prod("DESCRICAO") & " esta com estoque n�o inicializado. Estoque n�o atualizado no Gesto. Favor inicializar estoque do produto na tela de cadastro de produto ou ficar� com estoque errado."
           Else
             If Not ItemEstoqueAlmox.EOF Then
               ItemEstoqueAlmox.Edit
               ItemEstoqueAlmox!Qtde_Disponivel = Estoque_Prod("Estoque Atual")
               ItemEstoqueAlmox.Update
             End If
           End If
          'Produto Grade PAF
         ElseIf cad_prod("Tipo") = "G" Then
           Tamanho = 0
           Cor = 0
           Edicao = 0
           Tipo = 1
           Erro = 0
           Acha_Produto produtos("C�digo"), produtos("C�digo"), Tamanho, Cor, Edicao, Tipo, Erro
           Set Estoque_Prod = db.OpenRecordset("Select [Estoque Atual] From [Estoque Final] where Filial = " & gnCodFilial & " and Produto =  '" & produtos("C�digo sem Grade") & "'  AND Cor = " & Cor & " And Tamanho = " & Tamanho & "")
           Set SaidaEstoqueItem = GestoBD.OpenRecordset("Select * from SaidaEstoqueItem Where NUMERO = " & txtSeq.Text & " AND CODIGO_ITEM =  '" & produtos("C�digo sem Grade") & "';")
            If SaidaEstoqueItem.EOF Then
             SaidaEstoqueItem.AddNew
             SaidaEstoqueItem!Numero = txtSeq.Text
             SaidaEstoqueItem!Item = produtos("Linha")
             SaidaEstoqueItem!Codigo_Item = produtos("C�digo Sem Grade")
             SaidaEstoqueItem!DESCRICAO_ITEM = Nome_Prod("DESCRICAO")
             SaidaEstoqueItem!Quantidade = produtos("Qtde")
             SaidaEstoqueItem!VALOR_UNIT_DESC = produtos("Pre�o") - (produtos("Pre�o") * produtos("Desconto") / 100)
             SaidaEstoqueItem!Valor_Total = produtos("Pre�o Final")
             SaidaEstoqueItem.Update
             If Estoque_Prod.EOF Then
               MsgBox "O produto " & Nome_Prod("DESCRICAO") & " esta com estoque n�o inicializado. Estoque n�o atualizado no Gesto. Favor inicializar estoque do produto na tela de cadastro de produto ou ficar� com estoque errado."
             Else
               If Not ItemEstoqueAlmox.EOF Then
                 ItemEstoqueAlmox.Edit
                 ItemEstoqueAlmox!Qtde_Disponivel = Estoque_Prod("Estoque Atual")
                 ItemEstoqueAlmox.Update
               End If
             End If
         End If
      End If
     produtos.MoveNext
     Loop
     Dim rsLibera As Recordset
     Set rsLibera = GestoBD.OpenRecordset("SELECT id From parametro")
   End If
   End If
'---------------------------------------------------------------------------------------------------------------
  
  '22/10/2002 - mpdea
  'Fun��o executada com sucesso
  UpdateRecord = 0
  
  Exit Function

ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3022 And bSequencia 'Duplicidade de movimenta��o
      If intRepeatUpdate3022 < 1000 Then
        Call StatusMsg("Verificando registro...")
        intRepeatUpdate3022 = intRepeatUpdate3022 + 1
        nSequencia = gnGetNextSequencia(gnCodFilial)
        bSeqChanged = True
        rsSaidas("Sequ�ncia") = nSequencia
        Resume
      End If
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
          'Cancelamento da transa��o
          If blnInTransaction Then ws.Rollback
          Exit Function
        End If
      
'        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
'          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Sa�das - Gravar") = vbOK Then
'          nRepeatUpdateLocked = 0
'          Resume
'        Else
'          On Error Resume Next
'          'Cancelamento da transa��o
'          ws.Rollback
'          Exit Function
'        End If
      End If
    Case Else
      'Cancelamento da transa��o
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      MsgBox "Erro em Sa�das - Gravar: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
      Exit Function
      
'      'Outros Erros
'      Select Case frmErro.gnShowErr(Err.Number, "Sa�das - Gravar")
'        Case 0 'Repetir
'          Resume
'        Case 1 'Prosseguir
'          Resume Next
'        Case 2 'Sair
'          Exit Function
'        Case 3 'Encerrar
'          End
'      End Select
  End Select
  
End Function

Public Sub ClearScreen()
  Dim Linha As Integer
  Dim Tool As ActiveBarLibraryCtl.Tool

  '02/02/2009 - mpdea
  'Corrigido RT-3021
  'Modificado para "limpar" a vari�vel antes de outras verifica��es
  Num_Registro = Null

  frm_produtoSemPrecoNaGrade.Visible = False
  bProdutoSemPrecoNaGrade = False

  '15/02/2007 - Anderson
  'Ferramenta de filtro para c�digo do cliente - Solicitado por: Paulo Ribertec
  If ActiveBar1.Tools("miComplFiltrarCliente").Checked Then
    ActiveBar1.Tools("miComplFiltrarCliente").Checked = False
    Set rsSaidas = db.OpenRecordset("SELECT * FROM Sa�das WHERE Filial = " & gnCodFilial & " ORDER BY Sequ�ncia", dbOpenDynaset)
  End If

  'Na mudan�a de registro o Altera Totais � desmarcado
  Set Tool = ActiveBar1.Tools("miComplAlteraTotais")
  If Tool.Checked Then
    Call ActiveBar1_Click(Tool)
  End If

  '14/08/2002 - mpdea
  'Caption
  Me.Caption = "Sa�das"

  txtNF.Text = ""
  cboConsumidorFinal.Text = "1=Sim"
  cboPresencaComprador.Text = "1 =Opera��o presencial"
  'cboFinalidade.Text = "1=NFe normal"
  cboFinalidade.ListIndex = 0
  
  gridChaves.Rows = 1

  If ActiveBar1.Tools("miOpFreezeOperacao").Checked = False Then
     cboOper.Text = ""
     Nome_Opera��o.Caption = ""
  End If
  If ActiveBar1.Tools("miOpFreezeDigitador").Checked = False Then
    cboDigitador.Text = ""
    Nome_Digitador.Caption = ""
    '01/06/2004 - Daniel
    'Manter ou n�o o Operador
    Combo_Operador.Text = ""
    Nome_Operador.Caption = ""
  End If
  
  If ActiveBar1.Tools("miOpFreezeCliente").Checked = False Then
    cboCliente.Text = ""
    Nome_Cliente.Caption = ""
  End If

  If ActiveBar1.Tools("miOpFreezeTabPrecos").Checked = False Then
    Combo_Pre�o.Text = ""
  End If
  
  '19/01/2004 - Daniel
  'Case.......: PSV Inform�tica
  If m_blnPSV Then
    mskValidade.Mask = ""
    mskValidade.Text = ""
    mskValidade.Mask = "##/##/####"
    mskValidade.Enabled = False
    mskValidade.BackColor = &H808080 'Cinza
  End If
  '-----------------------------------------
    
  txtRef.Text = ""

  Obs.Text = ""
  
  '04/03/2004 - mpdea
  'Limpa a senha
  Senha.Text = ""
  
 
  L_Dia.Caption = Format$(Data_Atual, "dd/mm/yyyy")
  
  Erase Tabe
'  For Linha = 0 To 254
'    With Tabe(Linha)
'      .C�digo = 0
'      .Nome = ""
'      .Unidade = ""
'      .Pre�o_Total = 0
'      .Pre�o_Final = 0
'      .Nada = ""
'      .Informa = ""
'      .Qtde = 0
'      .Pre�o = 0
'      .Desconto = 0
'      .Base_ICM = 0
'      .ICM = 0
'      .Valor_Base_Unit = 0
'      .Redu��o_ICM = 0
'      .Valor_ICM = 0
'      .IPI = 0
'      .Etiqueta = False
'      .Tipo_ICM = ""
'      .Descr_Adicional = ""
'    End With
'  Next Linha
  
  Erase Tabe_Serv
'  For Linha = 0 To 254
'    With Tabe_Serv(Linha)
'      .C�digo = 0
'      .Descri��o = ""
'      .Pre�o = 0
'      .Total = 0
'      .Tempo = ""
'      .Completo = False
'      .ISS = 0
'    End With
'  Next Linha


  Grade1.MoveLast
  Grade1.MoveFirst

  Grade_Serv.MoveLast
  Grade_Serv.MoveFirst

  Tab1.Tab = 0
  
  '08/12/2004 - Daniel
  'Adicionado esta linha com a finalidade de evitar
  'perda de foco
  Grade1.Columns(0).Text = "0"
  '------------------------------------------------
  
  '22/06/2006 - mpdea
  'Corrigido status do grid que alternava entre o modo de edi��o
  Grade1.Update
  
  lblQtdeTotal.Caption = 0
  Total_Desconto = 0
  gcDescInTotal = 0
  
  '20/09/2002 - mpdea
  'Desconto no SubTotal
  mcurDescontoSubTotal = 0
  txtSubTotal.Text = Format("0", FORMAT_VALUE)
  txtDescSubTotal.Text = Format("0", FORMAT_VALUE)
  
  Total_Servi�os = 0
  L_Frete.Text = Format("0", FORMAT_VALUE)
  
  '12/04/2005 - Daniel
  'Adicionado campo Seguro
  txtSeguro.Text = Format("0", FORMAT_VALUE)
  
  '18/05/2005 - Daniel
  'Tratamento para o campo N� de S�rie da NF
  
  If Len(txtNrSerieNF.Text) > 0 Then
    txtNrSerieNF.Text = ""
  End If
  
  '01/08/2005 - Daniel
  '
  'Tratamento para o Campo Sa�das.DataEmissaoNotaManual
  'Solicitante: Ped�gio Cal�ados e Confec��es
  'Projeto    : Impress�o de Notas Manuais
  mskDataEmissaoNotaManual.Mask = ""
  mskDataEmissaoNotaManual.Text = ""
  mskDataEmissaoNotaManual.Mask = "##/##/####"
  
  mskDataEmissaoNotaManual.Visible = False
  lblDataEmissaoNotaManual.Visible = False
  '-----------------------------------------------
  
  Call Recalcula
  
  L_Efetivada.Visible = False
  lblMovPendencia.Visible = False
  
  Movimenta��o_Desfeita.Visible = False
  Nota_Cancelada.Visible = False
    
'  Call F_Pagto_Entrada.Limpa
  
  txtSeq.Text = ""
  Desconto_Cli = 0

  Prometido_Para.Text = ""
  Or�amento_Aprovado.Text = ""
  
  Data_Acerto.Mask = ""
  Data_Acerto.Text = ""
  Data_Acerto.Mask = "##/##/####"
  F_Empr�stimo.Visible = False

  Combo_T�cnico.Text = 0
  Combo_T�cnico_LostFocus
  
  Combo_Caixa.Text = 0
  If gbCaixas = False Then Combo_Caixa.Text = 1
    
  If Not rsSaidas.EOF Then
    On Error Resume Next
    rsSaidas.MoveFirst
    rsSaidas.MovePrevious
    cboOper.SetFocus
    On Error GoTo 0
  End If
  
  '20/09/2002 - mpdea
  'Novo registro, desativa flag para for�ar atualiza��o
  mblnForceUpdate = False
  
    
'  Data4.Refresh 'ATUALIZA OS CLIENTES TODA VEZ QUE LIMPA
'  Data1.Refresh 'ATUALIZARIA OS PRODUTOS MAS FICA MUITO LENTO
  Combo_Caixa_LostFocus
  'elefante
  txtComanda.Text = ""
  btnComandaVendas.Visible = False
  txtComanda.Width = txtSeq.Width
  L_Tot_ICMS_Deson.Text = Format("0", FORMAT_VALUE)
    '29/10/2013 - Jean
'''  'Customiza��o para Disk Embalagens para bloquear a grid quando tiver uma sequencia j� gravada
'''  If CheckSerialCaseMod("QS73520-469") Then
'''    If (txtSeq.Text = "") Then
'''      Grade1.Enabled = True
'''      DropDown1.Enabled = True
'''    End If
'''  End If

End Sub
Public Sub ReplicaMov()

  Dim Tool As ActiveBarLibraryCtl.Tool

  If IsNull(Num_Registro) Then
      DisplayMsg "Encontre uma movimenta��o antes."
      Exit Sub
  End If
 
  Set Tool = ActiveBar1.Tools("miComplAlteraTotais")
  If Tool.Checked Then
    Call ActiveBar1_Click(Tool)
  End If
  
  gridChaves.Rows = 1
  
  txtNF.Text = ""
  txtRef.Text = ""
  txtSeq.Text = ""
  
  L_Dia.Caption = Format$(Data_Atual, "dd/mm/yyyy")
  
  L_Efetivada.Visible = False
  Movimenta��o_Desfeita.Visible = False
  Nota_Cancelada.Visible = False
  Combo_Caixa_LostFocus
  
  Num_Registro = Null
    
  DisplayMsg "Movimenta��o Replicada. Revise os valores e Grave."
End Sub

'04/05/2004 - mpdea
'Corrigido e otimizado o c�digo em geral
'
'08/04/2003 - mpdea
'Implementado tratamento de erro
Private Sub PrintNota()
  Dim frmX As Form
  
  Dim rsTempOpSaidas As Recordset
  Dim strSQL As String
  Dim blnExit As Boolean
  Dim blnShowObs As Boolean
  Dim intX As Integer
  
  Dim strFileNF As String
  Dim intRet As Integer
  Dim lngNotaFiscal As Long
  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked As Integer
  
  '18/12/2007 - Anderson
  'Implementa��o do NSU para SC
  Dim blnGerarNSU As Boolean
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")
  
  '18/12/2007 - Anderson
  'Implementa��o do NSU para SC
  blnGerarNSU = True
  
  If txtSeq.Text = "" Then
    DisplayMsg "Ache ou grave uma venda antes."
    Exit Sub
  End If
  
  If rsSaidas.Fields("Nota Cancelada").Value Then
    DisplayMsg "Esta nota est� cancelada e n�o pode ser reimpressa."
    Exit Sub
  End If
  
  '04/12/2007 - Anderson
  'Verifica permiss�o para imprimir nota somente em movimenta��es efetivadas
  If rsParametros.Fields("ImprimeNotaMovEfetivada").Value Then
    If Not rsSaidas.Fields("Efetivada").Value Then
      DisplayMsg "Movimenta��o n�o efetivada. N�o � poss�vel imprimir nota fiscal."
      Exit Sub
    End If
  End If
  
  'Verifica��es referente a opera��o de Sa�da
  strSQL = "SELECT * FROM [Opera��es Sa�da] WHERE C�digo = " & rsSaidas.Fields("Opera��o").Value
  Set rsTempOpSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rsTempOpSaidas
    If .RecordCount > 0 Then
      If Not .Fields("Nota").Value Then
        DisplayMsg "Opera��o n�o permite Nota Fiscal."
        blnExit = True
      End If
      blnShowObs = .Fields("InTelaObsTransp").Value
    Else
      DisplayMsg "Opera��o de Sa�da n�o encontrada."
      blnExit = True
    End If
    .Close
  End With
  Set rsTempOpSaidas = Nothing
  If blnExit Then Exit Sub
  
'  Call RecalculaPesos
  
  If blnShowObs Then
    Set frmX = New frmObsNota
    frmX.gsCliente = rsCliFor.Fields("Transportadora").Value
    frmX.lngSequencia = rsSaidas.Fields("Sequ�ncia").Value
    frmX.bytTipoTabela = 1
    frmX.Show vbModal
    Set frmX = Nothing
    If gsRetornoDoc <> "OK" Then
      StatusMsg "Nota n�o impressa."
      Exit Sub
    End If
  Else
    For intX = 0 To 7
      gsObsDoc(intX) = ""
    Next intX
    gsPlaca = ""
    gsUfrmPlaca = ""
    gsQtdeTrans = ""
    gsMarcaTrans = ""
    gsEspecieTrans = ""
    gsPesoBruto = ""
    gsPesoLiquido = ""
    gsTransportadora = ""
  End If
  
  '20/05/2005 - Daniel
  '
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '             Tratamento para nota gerada manualmente
  If Not (gbNotaManual(rsSaidas.Fields("Opera��o").Value, "SAIDA")) Then
    '-----------------------------------------------------------------------
    'Impress�o de Nota autom�tica (com incrementa��o no contador do Quick...
    'Isto sempre ocorreu no Quick
    '-----------------------------------------------------------------------
    Call IsDataType(dtLong, rsSaidas.Fields("Nota Impressa").Value, lngNotaFiscal)
    If lngNotaFiscal <> 0 Then
      If MsgBox("A Nota fiscal j� foi impressa, deseja imprimir novamente?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Aten��o") = vbNo Then
        Exit Sub
      End If
      
      '18/12/2007 - Anderson
      'Implementa��o do NSU para SC
      blnGerarNSU = False

    End If
  
  End If
  
  '--------------------------------------------------------------------------
  'Grava nova NF
  '--------------------------------------------------------------------------
  If lngNotaFiscal = 0 Then
    'Modificado leitura e grava��o do n�mero da �ltima nota fiscal
    'Inclu�do transa��o durante grava��o
    'lngNotaFiscal = rsParametros.Fields("�ltima Nota").Value + 1
    '
    ws.BeginTrans
    blnInTransaction = True
    
    '20/05/2005 - Daniel
    'lngNotaFiscal = g_lngNextNotaFiscal(rsSaidas.Fields("Filial").Value) Mudamos a linha da chamada para n�o invocar gbNotaManual 2x
    
    'With rsParametros
    '  .Edit
    '  .Fields("�ltima Nota").Value = lngNotaFiscal
    '  .Update
    'End With
    '
    With rsSaidas
      .LockEdits = True
      .Edit
            
      '20/05/2005 - Daniel
      '
      'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
      '             para todos usu�rios do Quick Store
      '
      'O sistema dever� julgar se a nota fiscal ser� criada
      'automaticamente ou manualmente a partir da opera��o escolhida
      'Nota: Caso seja manualmente (notas de bloquinho), o sistema n�o
      'dever� incrementar o contador pois o sistema estava fora do ar
      If Not (gbNotaManual(rsSaidas.Fields("Opera��o").Value, "SAIDA")) Then
        '20/05/2005 - Daniel
        'Adicionamos a linha de comando abaixo da busca da pr�xima nota
        lngNotaFiscal = g_lngNextNotaFiscal(rsSaidas.Fields("Filial").Value)
        .Fields("Nota Impressa").Value = lngNotaFiscal
      End If
      
      
      'Grava��o dos campos de observa��es na tela de sa�das
      'For intX = 0 To 7
      '  .Fields("obs_Obs" & intX + 1).Value = gsObsDoc(intX)
      'Next intX
      For intX = 0 To 1
        .Fields("obs_infCpl" & intX + 1).Value = gsObsDoc(intX)
      Next intX
      .Fields("obs_Transportadora") = gsTransportadora
      .Fields("obs_Placa") = gsPlaca
      .Fields("obs_Uf") = gsUfrmPlaca
      .Fields("obs_Especie") = gsEspecieTrans
      .Fields("obs_Qtde") = gsQtdeTrans
      .Fields("obs_Marca") = gsMarcaTrans
      .Fields("obs_PesoBruto") = IIf(IsNumeric(gsPesoBruto), gsPesoBruto, 0)
      .Fields("obs_PesoLiquido") = IIf(IsNumeric(gsPesoLiquido), gsPesoLiquido, 0)
      .Fields("obs_FretePago") = IIf(IsNumeric(gsFretePago), gsFretePago, 0)
      .Update
      .LockEdits = False
    End With
    '
    '20/05/2005 - Daniel
    If Not (gbNotaManual(rsSaidas.Fields("Opera��o").Value, "SAIDA")) Then
      txtNF.Text = lngNotaFiscal
    
      '05/05/2005 - mpdea
      'Atualiza a Nota Fiscal e Fatura do Contas a Receber
      Call StatusMsg("Verificando e atualizando contas a receber...")
      '
      strSQL = "UPDATE [Contas a Receber] SET Nota = " & lngNotaFiscal
      strSQL = strSQL & ", Fatura = '" & lngNotaFiscal & "/ ' & Parcela"
      strSQL = strSQL & " WHERE Tipo = 'R'"
      strSQL = strSQL & " AND Filial = " & rsSaidas.Fields("Filial").Value
      strSQL = strSQL & " AND Sequ�ncia = " & rsSaidas.Fields("Sequ�ncia").Value
      '
      db.Execute strSQL, dbFailOnError
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
        strSQL, _
        "frmSaidas_PrintNota", _
        "Contas a Receber", g_strArquivoSystemLog
      End If

    
    Else
    
      '20/05/2005 - Daniel
      'Tratamento para a nota manual...
      '
      'Atualiza a Nota Fiscal e Fatura do Contas a Receber
      Call StatusMsg("Verificando e atualizando contas a receber...")
      '
      strSQL = "UPDATE [Contas a Receber] SET Nota = " & CLng("0" & txtNF.Text)
      strSQL = strSQL & ", Fatura = '" & CLng("0" & txtNF.Text) & "/ ' & Parcela"
      strSQL = strSQL & " WHERE Tipo = 'R'"
      strSQL = strSQL & " AND Filial = " & rsSaidas.Fields("Filial").Value
      strSQL = strSQL & " AND Sequ�ncia = " & rsSaidas.Fields("Sequ�ncia").Value
      '
      db.Execute strSQL, dbFailOnError
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
        strSQL, _
        "frmSaidas_PrintNota", _
        "Contas a Receber", g_strArquivoSystemLog
      End If
    End If
    
''    Aux_Data = CDate("01/01/1980")
'    Aux_Int = 1
'    Aux_Conta = 0
'    rsContas_Receber.Index = "Cliente"
''    Erro = False
'Lp1_Receber:
'    rsContas_Receber.Seek ">", "R", rsSaidas("Cliente"), Aux_Conta
'    If rsContas_Receber.NoMatch Then GoTo Fim_Receber
'    If rsContas_Receber("Tipo") <> "R" Then GoTo Fim_Receber
'    If rsContas_Receber("Cliente") <> rsSaidas("Cliente") Then GoTo Fim_Receber
'    Aux_Conta = rsContas_Receber("Contador")
'    If rsContas_Receber("Sequ�ncia") <> rsSaidas("Sequ�ncia") Then GoTo Lp1_Receber
'    rsContas_Receber.Edit
'      rsContas_Receber("Nota") = rsSaidas("Nota Impressa")
'      rsContas_Receber("Fatura") = str(rsSaidas("Nota Impressa")) + "/" + str(Aux_Int)
'      Aux_Int = Aux_Int + 1
'    rsContas_Receber.Update
'    GoTo Lp1_Receber
'
'Fim_Receber:


    Call StatusMsg("")
    
    'Finaliza transa��o
    ws.CommitTrans
    blnInTransaction = False
  End If
  '--------------------------------------------------------------------------
  
  '18/12/2007 - Anderson
  'Implementa��o do NSU
  If blnGerarNSU Then
    Call GerarNSU(rsSaidas, "Sa�das")
  End If
  
  '--------------------------------------------------------------------------
  'Imprime NF
  '--------------------------------------------------------------------------
  strFileNF = gsConfigPath + rsParametros.Fields("Nota Sa�da").Value + ".CNF"
  intRet = Imprime_Nota(strFileNF, rsSaidas.Fields("Filial").Value, rsSaidas.Fields("Sequ�ncia").Value)
  If intRet = 0 Then
    '14/04/2003 - mpdea
    'Atualiza a data da impress�o da nota fiscal
    strSQL = "UPDATE Sa�das SET DataEmissaoNota = #"
    strSQL = strSQL & Format(Date, "mm/dd/yyyy") & "# "
    strSQL = strSQL & "WHERE Filial = " & rsSaidas.Fields("Filial").Value
    strSQL = strSQL & " AND Sequ�ncia = " & rsSaidas.Fields("Sequ�ncia").Value
    db.Execute strSQL, dbFailOnError
       
    '
    '20/05/2005 - Daniel
    'Tratamento para nota manual
    If Not (gbNotaManual(rsSaidas.Fields("Opera��o").Value, "SAIDA")) Then
      DisplayMsg "Nota [" & lngNotaFiscal & "] impressa com sucesso."
    Else
      DisplayMsg "Nota [" & rsSaidas.Fields("Nota Fiscal").Value & "] impressa com sucesso."
    End If
  Else
    DisplayMsg "Houve o erro " & intRet & " durante a impress�o da Nota."
  End If
  '--------------------------------------------------------------------------
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3197, 3187, 3218, 3260 'Registro bloqueado
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
          'Cancelamento da transa��o
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
      
'        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
'          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Sa�das - Imprimir Nota Fiscal") = vbOK Then
'          intRepeatUpdateLocked = 0
'          Resume
'        Else
'          'Cancelamento da transa��o
'          If blnInTransaction Then ws.Rollback
'          Exit Sub
'        End If
      End If
    Case Else
      'Cancelamento da transa��o
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  End Select
  
End Sub


'Private Sub MoveNext()
' Dim Atual As Variant
' Dim Atual2 As Variant
' Dim Atual3 As Variant
'
'
' Atual = N�mero.Text
' If IsNull(Atual) Then Atual = 0
' If Not IsNumeric(Atual) Then Atual = 0
' If Atual <> 0 And IsNull(Num_Registro) Then Atual = Atual - 1
'
' If O_Sequ�ncia.Value = True Then
'   rsSaidas.Index = "Sequ�ncia"
'
'   rsSaidas.Seek ">", gnCodFilial, Atual
'   If rsSaidas.NoMatch Then
'     Beep
'     If Not IsNull(Num_Registro) Then
'      rsSaidas.Bookmark = Num_Registro
'     End If
'     Exit Sub
'   End If
'
'   If rsSaidas("Filial") <> gnCodFilial Then
'     Beep
'     If Not IsNull(Num_Registro) Then
'      rsSaidas.Bookmark = Num_Registro
'     End If
'     Exit Sub
'   End If
'
' End If
'
'
'
' If O_Data.Value = True Then
'   Atual2 = CDate(L_Dia.Caption)
'
'   rsSaidas.Index = "Data"
'
'   rsSaidas.Seek ">", gnCodFilial, Atual2, Atual
'   If rsSaidas.NoMatch Then
'     Beep
'     If Not IsNull(Num_Registro) Then
'      rsSaidas.Bookmark = Num_Registro
'     End If
'     Exit Sub
'   End If
'
'
'   If rsSaidas("Filial") <> gnCodFilial Then
'     Beep
'     If Not IsNull(Num_Registro) Then
'      rsSaidas.Bookmark = Num_Registro
'     End If
'     Exit Sub
'   End If
' End If
'
' If O_Fornecedor.Value = True Then
'   Atual2 = CDate(L_Dia.Caption)
'   Atual3 = cboCliente.Text
'
'   If IsNull(Atual3) Then Atual3 = 0
'   If Atual3 = "" Then Atual3 = 0
'
'   rsSaidas.Index = "Fornecedor"
'
'   rsSaidas.Seek ">", Atual3, Atual2, Atual
'   If rsSaidas.NoMatch Then
'     Beep
'     If Not IsNull(Num_Registro) Then
'      rsSaidas.Bookmark = Num_Registro
'     End If
'     Exit Sub
'   End If
'
'   If rsSaidas("Filial") <> gnCodFilial Then
'     Beep
'     If Not IsNull(Num_Registro) Then
'      rsSaidas.Bookmark = Num_Registro
'     End If
'     Exit Sub
'   End If
'
' End If
'
'
'
' If O_Refer�ncia.Value = True Then
'   Atual2 = txtRef.Text
'   rsSaidas.Index = "Refer�ncia"
'
'   rsSaidas.Seek ">", gnCodFilial, Atual2, Atual
'   If rsSaidas.NoMatch Then
'     Beep
'     If Not IsNull(Num_Registro) Then
'      rsSaidas.Bookmark = Num_Registro
'     End If
'     Exit Sub
'   End If
'
'   If rsSaidas("Filial") <> gnCodFilial Then
'     Beep
'     If Not IsNull(Num_Registro) Then
'      rsSaidas.Bookmark = Num_Registro
'     End If
'     Exit Sub
'   End If
'End If
'
'Call ShowRecord
'
'End Sub
'
Private Sub RealizaRecebimento()
  Dim nRet As Integer
  Dim intRepeatUpdateLocked As Integer
  
  Dim Ordem As Integer
  Dim Fim As Integer
  Dim Resposta As Integer
  Dim R_Banco As Integer
  Dim R_Cheque As String
  Dim R_Bom As Date
  Dim R_Valor As Double
  Dim Conta As Integer
  Dim Resp As Integer
  Dim Parcelas As Integer
  
  Dim blnInTransaction As Boolean
  
  On Error GoTo ProcessErr

  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre ou grave uma venda antes."
    Exit Sub
  End If
  
  '20/09/2002 - mpdea
  'For�a a atualiza��o do registro
  If mblnForceUpdate Then
    DisplayMsg "Valores alterados, grave a venda antes."
    Exit Sub
  End If
  
  '12/08/2002 - mpdea
  'Caso seja uma movimenta��o WEB, verifica se � necess�rio atualizar o registro
  If CLng("0" & rsSaidas.Fields("WebOrderFormID").Value) > 0 Then
    If CInt("0" & rsSaidas.Fields("Digitador").Value) = 0 Then
      '22/10/2002 - mpdea
      'Modificado para fun��o
      If UpdateRecord <> 0 Then
        Exit Sub
      End If
    End If
  End If
  
  rsOperadores.Index = "C�digo"
  rsOperadores.Seek "=", Combo_Operador.Text
  
  If rsOperadores.NoMatch Then
    DisplayMsg "Operador incorreto."
    Exit Sub
  End If
  
  If rsOperadores("Recebimento Sa�das") = False Then
    Beep
    DisplayMsg "Este usu�rio n�o tem permiss�o para usar a tela de recebimento."
    Exit Sub
  End If
  If rsOperadores("ValorP") <> CriptografaSenha(Senha.Text) Then
    DisplayMsg "Senha incorreta."
    Senha.SetFocus
    Exit Sub
  End If
  
   
  If rsOp_Sa�da("Dinheiro") = False Then
    Beep
    DisplayMsg "Esta opera��o n�o necessita que seja usada a tela de recebimento."
    Exit Sub
  End If
  
  
  If IsNumeric(rsParametros("DiasBloqueioVenda").Value) Then
    If rsParametros.Fields("DiasBloqueioVenda") > 0 Then
      If IsDate(rsCliFor.Fields("�ltima Compra")) Then
        If (CDate(Data_Atual) - CDate(rsCliFor.Fields("�ltima Compra"))) > CInt(rsParametros.Fields("DiasBloqueioVenda")) Then
          If MsgBox("O cliente que voc� escolheu n�o compra h� " & (CDate(Data_Atual) - CDate(rsCliFor.Fields("�ltima Compra"))) & " dias, deseja continuar ? ", vbQuestion + vbYesNo, "Quick Store") = vbNo Then
            Exit Sub
          End If
        End If
      End If
    End If
  End If
  
  If rsSaidas("Recebimento") = True Then
    Resp = MsgBox("Esta opera��o j� foi efetivada. Os dados de recebimento est�o dispon�veis apenas para visualiza��o. Caso queira alterar os dados do recebimento, use a op��o DESFAZ movimenta��o no menu Op��es antes.", vbInformation, "Aten��o")
    
    frmRecebimento.Limpa_Tela (0)
    frmRecebimento.Receber.Caption = Total_Pagar
    frmRecebimento.L_Sequ�ncia = rsSaidas("Sequ�ncia")
    frmRecebimento.S�_Leitura.Value = 1
    
    frmRecebimento.Show vbModal
    Exit Sub
    
  End If
  
  If rsSaidas("Opera��o") <> cboOper.Text Then
     DisplayMsg "Grave a movimenta��o para atualizar a opera��o alterada."
     Exit Sub
  End If
  
  
  '19/08/2003 - mpdea
  'Modificado nome do campo
  '
  '09/10/2002 - mpdea
  'Verifica estoque conforme configura��es
  If Not rsParametros.Fields("Venda Sem Estoque Saidas").Value And rsOp_Sa�da.Fields("Estoque").Value Then
    If Not mblnCheckStock Then Exit Sub
  End If
  
  
  Total_Pagar = rsSaidas("Total")
  Call StatusMsg("")
  
  frmRecebimento.Limpa_Tela (0)
  frmRecebimento.Receber.Caption = Total_Pagar
  frmRecebimento.L_Sequ�ncia = rsSaidas("Sequ�ncia")
  frmRecebimento.S�_Leitura = 0
  frmRecebimento.Acerta_Tela
  frmRecebimento.Combo_Banco.Text = rsCliFor("Conta Cobran�a")
  frmRecebimento.lngCodigoCliente = CLng(cboCliente.Text)
  frmRecebimento.bytTelaChamada = 2  'Sa�das
  
  frmRecebimento.Intervalo_Parc.Caption = rsParametros("Sa�da Intervalo Parc")
    
  frmRecebimento.Show vbModal
  
  If frmRecebimento.Retorno.Caption <> "OK" Then
'    DisplayMsg "Recebimento n�o efetivado."
    Exit Sub
  End If
  
'  Call WaitSeconds(1, True) 'Aguarda um segundo para o refresh
  DoEvents
  Me.Refresh
  
  Screen.MousePointer = vbHourglass
  
  Call StatusMsg("Gravando recebimento...")
  
  'In�cio da transa��o
  ws.BeginTrans
  blnInTransaction = True
  
  rsSaidas.Edit
   rsSaidas("Recebe - Conta") = False
   If frmRecebimento.Conta.Value = 1 Then rsSaidas("Recebe - Conta") = True
   rsSaidas("Recebe - Dinheiro") = CDbl(frmRecebimento.Dinheiro.Text)
   rsSaidas("Recebe - Emp Cart�o") = Val(frmRecebimento.Combo_Empresa.Text)
   rsSaidas("Recebe - Num Cart�o") = frmRecebimento.strNumeroCartao
   rsSaidas("Recebe - Cart�o") = CDbl(frmRecebimento.Cart�o.Text)
   rsSaidas("Recebe - Vale") = CDbl(frmRecebimento.Vale.Text)
   rsSaidas("Recebimento") = True
   rsSaidas("TotalCartaoDebito") = frmRecebimento.TxtDebito.Text
   rsSaidas("TotalCartaoCredito") = frmRecebimento.txtCredito.Text
  
   If frmRecebimento.Conta.Value = 1 Then
      rsSaidas("Total Prazo") = rsSaidas("Total")
   Else
      rsSaidas("Total Prazo") = frmRecebimento.Pega_Total_Parcelas
   End If
     
   If frmRecebimento.O_Banco.Value = True Then
      rsSaidas("Tipo Parcela") = "B"
      If rsSaidas("Total Prazo") <> 0 Then rsSaidas("Conta") = frmRecebimento.Combo_Banco.Text
   End If
   
   If frmRecebimento.O_Carteira.Value = True Then rsSaidas("Tipo Parcela") = "C"
   If frmRecebimento.O_Carnet.Value = True Then rsSaidas("Tipo Parcela") = "T"
   
   '10/12/2009 - Andrea
   'Na tela de Recebimentos, n�o ser� mais atualizado este campo Label_Cart�o2,
   'O recebimento em cart�es ficarao no grid de cart�es
   'mas o campo Cart�o.text tem o valor total recebido em cart�o
   'If Len(Trim(frmRecebimento.Label_Cart�o2.Caption)) > 0 Then
     'rsSaidas("Parcela Cart�o") = "S"
     'rsSaidas("Qtde Parcelas") = CInt(gsHandleNull(frmRecebimento.Label_Cart�o2.Caption & ""))
     'rsSaidas("Valor Parcela") = CDbl(gsHandleNull(frmRecebimento.Label_Cart�o4.Caption) & "")
   'End If
  
   '07/01/2004 - Daniel
   'Realiza Recebimento ent�o:
   'Alimentamos os campos Valor Recebido e Troco
   'da tabela Sa�das
   rsSaidas("Valor Recebido") = frmRecebimento.g_dblValorRecebidoFrmRec
   rsSaidas("Troco") = frmRecebimento.g_dblTrocoFrmRec
  
  rsSaidas.Update
    
  '10/12/2009 - Andrea
  'Apaga Cartoes
  Call EraseTypeMoviment(tmMovimentoCartoes, gnCodFilial, Val(txtSeq.Text))
  'Grava Cartoes
  Dim lng_row As Long
  Dim var_book As Variant
  Dim str_administradora As String
  Dim dbl_valor As Double
  Dim int_qtde_parcelas As Double
  Dim dbl_valor_parcela As Double
  Dim str_numero As String
  Dim bln_credito As Boolean

  'Valor em cart�o
  With frmRecebimento.Grade_Cartoes
    'Verifica ocorr�ncia
    If .Rows > 0 Then
      
      For lng_row = 0 To .Rows - 1
          
        var_book = .AddItemBookmark(lng_row)
              
        'Verifica registro informado
        Call IsDataType(dtString, .Columns("Administradora").CellText(var_book), str_administradora)
        If str_administradora <> "" Then
          'Valores
          Call IsDataType(dtDouble, .Columns("Valor").CellText(var_book), dbl_valor)
          Call IsDataType(dtInteger, .Columns("Qtde Parcelas").CellText(var_book), int_qtde_parcelas)
          If int_qtde_parcelas = 0 Then int_qtde_parcelas = 1
          Call IsDataType(dtDouble, .Columns("Valor Parcelas").CellText(var_book), dbl_valor_parcela)
          Call IsDataType(dtString, .Columns("Numero").CellText(var_book), str_numero)
          Call IsDataType(dtBoolean, .Columns("Credito").CellValue(var_book), bln_credito)
          
          rsSa�da_Cartoes.AddNew
            rsSa�da_Cartoes("Filial") = gnCodFilial
            rsSa�da_Cartoes("Sequ�ncia") = Val(txtSeq.Text)
            rsSa�da_Cartoes("Ordem") = (lng_row + 1)
            rsSa�da_Cartoes("Administradora") = str_administradora
            rsSa�da_Cartoes("Valor") = dbl_valor
            rsSa�da_Cartoes("Parcelas") = int_qtde_parcelas
            rsSa�da_Cartoes("ValorParcelas") = dbl_valor_parcela
            '15/12/2009 - Andrea
            'Maikel e Marcelo pediram para n�o gravar o n�mero do cart�o
            rsSa�da_Cartoes("NumeroCartao") = str_numero
            
            rsSa�da_Cartoes("Credito") = bln_credito
            
          rsSa�da_Cartoes.Update
          
        End If
      Next lng_row
    End If
  End With

  'Apaga Cheques
  Call EraseTypeMoviment(tmMovimentoCheques, gnCodFilial, Val(txtSeq.Text))
  'Grava Cheques
  If rsParametros("VR Permite Cheques") Then
    Ordem = 1
    Do
      Resposta = frmRecebimento.Pega_Banco(Ordem, R_Banco, R_Cheque, R_Bom, R_Valor)
      If Resposta = 1 Then
        rsSa�da_Cheques.AddNew
          rsSa�da_Cheques("Filial") = gnCodFilial
          rsSa�da_Cheques("Sequ�ncia") = Val(txtSeq.Text)
          rsSa�da_Cheques("Ordem") = Ordem
          rsSa�da_Cheques("Banco") = R_Banco
          rsSa�da_Cheques("Cheque") = R_Cheque
          rsSa�da_Cheques("Bom") = R_Bom
          rsSa�da_Cheques("Valor") = R_Valor
        rsSa�da_Cheques.Update
      End If
      Ordem = Ordem + 1
    ' altera��o parametro cheque (Pablo)
    'Loop Until Ordem > 50
    Loop Until Ordem > rsParametros("VR Qtde Cheques")
  End If
    
  'Apaga Parcelas
  Call EraseTypeMoviment(tmMovimentoParcelas, gnCodFilial, Val(txtSeq.Text))
  'Grava Parcelas
  If rsParametros("VR Permite Parcela") Then
    Ordem = 1
    Do
      Resposta = frmRecebimento.Pega_Parcela(Ordem, R_Bom, R_Valor, Parcelas)
      If Resposta = 1 Then
        rsSa�da_Parcelas.AddNew
        rsSa�da_Parcelas("Filial") = gnCodFilial
        rsSa�da_Parcelas("Sequ�ncia") = Val(txtSeq.Text)
        rsSa�da_Parcelas("Ordem") = Ordem
        rsSa�da_Parcelas("Bom") = R_Bom
        rsSa�da_Parcelas("Valor") = R_Valor
        rsSa�da_Parcelas("Parcelas") = Parcelas
        rsSa�da_Parcelas.Update
      End If
      Ordem = Ordem + 1
    ' altera��o parametro parcela (Pablo)
    'Loop Until Ordem > 50
    Loop Until Ordem > rsParametros("VR Qtde Parcela")
  End If
   
  Call StatusMsg("Aguarde, efetivando venda...")
  
  nRet = Efetiva_Sa�da(gnCodFilial, Val(txtSeq.Text))
  
  If nRet <> 0 Then
    Select Case nRet
      Case -1
        'A��o cancelada
        Call StatusMsg("A��o cancelada.")
      Case 5
        Call DisplayMsg("Tabela de pre�os inexistente.")
      Case Else
        Call DisplayMsg("Opera��o N�O efetivada. Erro" & str(nRet))
    End Select
    L_Efetivada.Visible = False
    'Cancelamento da transa��o
    ws.Rollback
    blnInTransaction = False
  Else
    'Fim da transa��o
    ws.CommitTrans
    blnInTransaction = False
    L_Efetivada.Visible = True
    m_blnSenhaGerJaInformada = False
    
    ' *************
    ' Pilatti Abril/2018
    Dim sSql As String
    Dim sNumSequenciaERP_APP As String
    Dim sOperacaoERP_APP As String
    
    sOperacaoERP_APP = cboOper.Text
    
    If gINTEGRACAO_APP_ERR_QUICK = True And (sOperacaoERP_APP = sOPERACAO_APPQuick01 Or sOperacaoERP_APP = sOPERACAO_APPQuick02) Then
          
      sNumSequenciaERP_APP = txtSeq.Text

      ' Atualizar A3Manager para pedido efetivado
      ' Conectar com o DB A3Manager (sql server)
      ConectaDB_A3Manager
          
      ' Atualizar registro (status_processamento = 5   PEDIDO EFETIVADO NO ERP
      sSql = "update top (1) [dbo].[pedido] "
      sSql = sSql + " set status_processamento = 5 "
      sSql = sSql + " where [numSequenciaERP]='" + sNumSequenciaERP_APP + "'"
      sSql = sSql + " and [tpERP] = 1 and codClienteA3CadResult = '" + gCodClienteA3CadResult + "'"

      dbA3Manager.Execute sSql

      dbA3Manager.Close
      Set dbA3Manager = Nothing
    End If
    ' Fim Pilatti
    
    Call StatusMsg("")
  End If
  
  Screen.MousePointer = vbDefault
    
  Exit Sub
  
ProcessErr:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3197, 3187, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Resume
        Else
          'Cancelamento da transa��o
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
'        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
'          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Sa�das - Recebimento") = vbOK Then
'          nRepeatUpdateLocked = 0
'          Resume
'        Else
'          Call StatusMsg("")
'          On Error Resume Next
'          'Cancelamento da transa��o
'          ws.Rollback
'          Exit Sub
'        End If
      End If
    Case Else
      'Cancelamento da transa��o
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      MsgBox "Erro em Sa�das - Recebimento: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
      Exit Sub
      
'      'Outros Erros
'      Select Case frmErro.gnShowErr(Err.Number, "Sa�das - Recebimento")
'        Case 0 'Repetir
'          Resume
'        Case 1 'Prosseguir
'          Resume Next
'        Case 2 'Sair
'          Call StatusMsg("")
'          Exit Sub
'        Case 3 'Encerrar
'          End
'      End Select
  End Select

End Sub

' **************************************************
' M�todo A3MANAGER APP Pilatti/Mar�o/18
' ConectaDB_A3Manager
Private Function ConectaDB_A3Manager()
    Dim sConnectionString As String

    '-- Build the connection string
    'sConnectionString = "PROVIDER = MSDASQL;driver={SQL Server};database=A3Manager;server=LUIZ-PILATTI\SQL2012;uid=sa;pwd=@dmin123;"
    sConnectionString = "PROVIDER = MSDASQL;driver={SQL Server};database=A3Manager;server=AMAZONA-F74E4RM\SQLEXPRESS;uid=sa;pwd=admin@A3;"

    dbA3Manager.ConnectionString = sConnectionString
    dbA3Manager.Open
End Function

' **************************************************
' M�todo A3MANAGER APP Pilatti/Mar�o/18
' ContarCaracteres
Private Function ContarCaracteres(Texto As String, separador As String) As Integer
    Dim c As Integer
    Dim Caracter As String
    
    ContarCaracteres = 0
    
    For c = 1 To Len(Texto)
      Caracter = Mid$(Texto, c, 1)
      
      If Caracter = separador Then
        ContarCaracteres = ContarCaracteres + 1
      End If
    Next c

End Function

Private Sub B_Servi�os_Conc_Click()
  Dim i As Integer
  
  For i = 0 To 254
    If Tabe_Serv(i).C�digo <> 0 Then
      Tabe_Serv(i).Completo = True
    End If
  Next i
  Grade_Serv.MoveLast
  Grade_Serv.MoveFirst
  
End Sub

'30/01/2009 - mpdea
'Implementado op��o para email
Private Sub ImprimirTicket(ByVal blnEmail As Boolean)
  ' Dim Str1 As String
  ' Dim Str_Rel As String
  Dim Aux As Variant
  Dim Nome_Ticket As String
  Dim F As Form

  Dim rsTempOpSaidas As Recordset
  Dim sSql As String
  Dim bExit As Boolean
  Dim bShowObs As Boolean
  Dim nX As Integer

  
  On Error GoTo ErrHandler


  Aux = txtSeq.Text
  If IsNull(Aux) Or Aux = "" Then
    DisplayMsg "Ache ou grave uma venda antes."
    Exit Sub
  End If


  '03/07/2006 - mpdea
  'Verifia permiss�o para imprimir ticket somente em movimenta��es efetivadas
  'Solicitante: Bem me quer
  If Not blnEmail Then
    If rsParametros.Fields("ImprimeTicketMovEfetivada").Value Then
      If Not rsSaidas.Fields("Efetivada").Value Then
        DisplayMsg "Movimenta��o n�o efetivada. N�o � poss�vel imprimir o Ticket."
        Exit Sub
      End If
    End If
  End If

  'Verifica��es referente a opera��o de Sa�da
  sSql = "SELECT * FROM [Opera��es Sa�da] WHERE C�digo = " & rsSaidas("Opera��o")
  Set rsTempOpSaidas = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsTempOpSaidas
    If .RecordCount > 0 Then
      bShowObs = .Fields("InTelaObsTransp")
    Else
      DisplayMsg "Opera��o de Sa�da n�o encontrada."
      bExit = True
    End If
    .Close
  End With
  Set rsTempOpSaidas = Nothing
  If bExit Then Exit Sub


 If rsOp_Sa�da("Ticket Imprimir") <> "" Then
   Nome_Ticket = gsConfigPath + rsOp_Sa�da("Ticket Imprimir")
 End If
 
 If rsOp_Sa�da("Ticket Imprimir") = "" Then
   Set F = New frmEscolheTicket
   F.Show vbModal
   Set F = Nothing
   If gsRetornoDoc = "CANCELADO" Then
     'StatusMsg "Ticket n�o impresso."
     Exit Sub
   End If
   Nome_Ticket = gsConfigPath + gsRetornoDoc
   If Dir(Nome_Ticket) = "" Then
     DisplayMsg "Arquivo """ & Nome_Ticket & """ n�o encontrado."
     Exit Sub
   End If
 End If
 
  If bShowObs Then
    Set F = New frmObsNota
    F.gsCliente = rsCliFor("Transportadora")
    F.lngSequencia = rsSaidas.Fields("Sequ�ncia").Value
    F.bytTipoTabela = 1
    F.Show vbModal
    Set F = Nothing
    If gsRetornoDoc <> "OK" Then
      StatusMsg "Opera��o cancelada."
      Exit Sub
    End If
  Else
    For nX = 0 To 7
      gsObsDoc(nX) = ""
    Next nX
    gsPlaca = ""
    gsUfrmPlaca = ""
    gsQtdeTrans = ""
    gsMarcaTrans = ""
    gsEspecieTrans = ""
    gsPesoBruto = ""
    gsPesoLiquido = ""
  End If
  
  '11/08/2003 - maikel
  '             Grava��o dos campos de observa��es na tela de sa�das
  '----------------------------------------------------------------'
    rsSaidas.Edit
    
    'For nX = 0 To 7
    '  rsSaidas.Fields("obs_Obs" & nX + 1).Value = gsObsDoc(nX)
    'Next nX
    For nX = 0 To 1
      rsSaidas.Fields("obs_infCpl" & nX + 1).Value = gsObsDoc(nX)
    Next nX
    
    rsSaidas.Fields("obs_Transportadora") = gsTransportadora
    rsSaidas.Fields("obs_Placa") = gsPlaca
    rsSaidas.Fields("obs_Uf") = gsUfrmPlaca
    rsSaidas.Fields("obs_Especie") = gsEspecieTrans
    rsSaidas.Fields("obs_Qtde") = gsQtdeTrans
    rsSaidas.Fields("obs_Marca") = gsMarcaTrans
    rsSaidas.Fields("obs_PesoBruto") = IIf(IsNumeric(gsPesoBruto), gsPesoBruto, 0)
    rsSaidas.Fields("obs_PesoLiquido") = IIf(IsNumeric(gsPesoLiquido), gsPesoLiquido, 0)
    
    rsSaidas.Fields("obs_FretePago") = IIf(IsNumeric(gsFretePago), gsFretePago, 0)
    rsSaidas.Update
  '----------------------------------------------------------------'
  
  If blnEmail Then
    'Prepara para enviar por email
    Call EnviarEmailModeloTicket(Nome_Ticket, gnCodFilial, rsSaidas.Fields("Sequ�ncia").Value, rsSaidas.Fields("Cliente").Value)
  Else
    'Imprime o ticket
    Call Imprime_Ticket(Nome_Ticket, gnCodFilial, rsSaidas.Fields("Sequ�ncia").Value)
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub btnComandaVendas_Click()
  If frmComanda.Total > 1 Then frmComanda.Show vbModal
End Sub

Private Sub cboCliente_Click()
  cboCliente.Text = cboCliente.Columns(1).Text
End Sub

Private Sub cboDigitador_Click()
  cboDigitador.Text = cboDigitador.Columns(2).Text
End Sub

Private Sub cboOper_Click()
  cboOper.Text = cboOper.Columns(1).Text
End Sub

Private Sub cmd_acataUsuarioLogadoComoOperador_Click()
  Senha.Text = gSenhaUsuarioLogado
  Combo_Operador.Text = gnUserCode & ""
  Combo_Operador_LostFocus
End Sub


Private Sub cmd_ajudaDespesasAcessorias_Click()
    MsgBox "Se necess�rio, digite o valor em R$ na coluna da grade Despesas Acess�rias que ser� adicionado ao produto, este valor poder� ser calculado pelo sistema e adicionado ao valor total da venda quando for gerada a NF-e.", vbInformation, "Aviso Fiscal"
End Sub

Private Sub cmd_ajudaIPI_Click()
    MsgBox "Caso houver % de IPI destacado na coluna da grade acima em alguns produtos, este % poder� ser calculado pelo sistema e adicionado ao valor total da venda quando for gerada a NF-e.", vbInformation, "Aviso Fiscal"
End Sub

Private Sub cmd_devolucaoProdutos_Click()

    If txtSeq.Text = "" Then
        MsgBox "Selecione uma venda.", vbInformation, "Aten��o"
        Exit Sub
    End If

    Dim objFormDevolucaoProduto As frmSaidasDevolucaoProdutos
    Set objFormDevolucaoProduto = New frmSaidasDevolucaoProdutos
    objFormDevolucaoProduto.lsSequenciaVenda = txtSeq.Text
    
    'Pegar o produto selecionado na grade...
    If IsNull(sCodigoProdutoDevolucao) Or sCodigoProdutoDevolucao = "" Then
        MsgBox "Selecione um produto na grade.", vbInformation, "Aten��o"
    Else
        objFormDevolucaoProduto.sCodigoProdutoDevolucao = sCodigoProdutoDevolucao
        objFormDevolucaoProduto.sNomeProdutoDevolucao = sNomeProdutoDevolucao
        objFormDevolucaoProduto.lsQuantidade = lQuantidadeItensProdutoDevolucao
        objFormDevolucaoProduto.sDescontoVenda = txtDescSubTotal.Text
        objFormDevolucaoProduto.sEmpresaFilial = Nome_Filial.Caption
        objFormDevolucaoProduto.sCliente = cboCliente.Text & "-" & Nome_Cliente.Caption
        objFormDevolucaoProduto.sDataDaVenda = L_Dia.Caption
        objFormDevolucaoProduto.sValorUnitarioProdutoDevolucao = sValorUnitarioProdutoDevolucao
    End If
    objFormDevolucaoProduto.Show
End Sub

Private Sub cmd_fecharFrameProdutoSemPrecoNaGrade_Click()
    frm_produtoSemPrecoNaGrade.Visible = False
End Sub

Private Sub cmd_gerarNFe_Click()
    ''##############################################################
    '' PABLO - 14/10/2022
    ''##############################################################
    ' envia a movimenta��o para tela de gerenciamento da NFe
    If Trim(txtSeq.Text) <> "" Then
        Dim p_seq As Long
        p_seq = CLng(Trim(txtSeq.Text))
        frmNFe.SetParametros (p_seq)
    End If
    ''##############################################################
    
    
    frmNFe.Show
End Sub

Private Sub cmd_tabelaDePrecos_Click()
  
  Dim bm As Variant
  Dim obj_formPreco As Form
  
  If Grade1.Columns("C�digo").CellValue(bm) = "" Or Grade1.Columns("C�digo").CellValue(bm) = "0" Then
      MsgBox "Selecione um produto na grade", vbInformation, "Aten��o"
  Else
      Set obj_formPreco = New frmTabelasDePrecosProdutos
      
      obj_formPreco.valorProdutoAcatado = ""
      obj_formPreco.CodigoProduto = Grade1.Columns("C�digo").CellValue(bm)
      obj_formPreco.nomeProduto = Grade1.Columns("Nome").CellValue(bm)
      
      obj_formPreco.Show vbModal
      
      
      'MsgBox "Produto " & Grade1.Columns("C�digo").CellValue(bm)
      
      If obj_formPreco.valorProdutoAcatado <> "" Then
          Grade1.Columns("Pre�o Unit�rio").Value = obj_formPreco.valorProdutoAcatado
          Call Calcula_Linha
      End If
      
      Set obj_formPreco = Nothing
  End If
End Sub

Private Sub cmdInsertItens_Click()
  Dim nX As Integer
  Call ClearScreen
  Grade1.SetFocus
  SendKeys "^{HOME}", True
  For nX = 1 To 255
    SendKeys "1{DOWN}", True
  Next nX
  SendKeys "1{UP}", True
End Sub

Private Sub Combo_Caixa_Click()
  Combo_Caixa.Text = Combo_Caixa.Columns(1).Text
End Sub

Private Sub Combo_Caixa_CloseUp()
  
  Combo_Caixa.Text = Combo_Caixa.Columns(1).Text
  Combo_Caixa_LostFocus

End Sub

Private Sub Combo_Caixa_LostFocus()

  Nome_Caixa.Caption = ""
  If IsNull(Combo_Caixa.Text) Then Exit Sub
  If Combo_Caixa.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Caixa.Text) Then Exit Sub
  If Val(Combo_Caixa.Text) > 99 Then Exit Sub
  If Val(Combo_Caixa.Text) < 1 Then Exit Sub
  
  rsCaixas.Index = "Caixa"
  rsCaixas.Seek "=", Val(Combo_Caixa.Text)
  If rsCaixas.NoMatch Then Exit Sub
  
  Nome_Caixa.Caption = rsCaixas("Descri��o") & ""
  
  
End Sub

Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(1).Text
  cboCliente_LostFocus
End Sub

Public Sub cboCliente_LostFocus() '16/04/2004 - Daniel - Mudado para Public
  Dim Aux As Variant
  
  Call StatusMsg("")
  
  'Indica que ainda n�o foi informada Senha Gerente para este cliente
  If cboCliente.Text <> m_strCodigoClienteContas Then
     m_blnSenhaGerJaInformada = False
  End If
 
  Nome_Cliente.Caption = ""
  Desconto_Cli = 0
  
  Aux = cboCliente.Text
  If IsNull(Aux) Then Exit Sub
  If Aux = "" Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) < 1 Then Exit Sub
  If Val(Aux) > 99999999 Then Exit Sub
  
  rsCliFor.Index = "C�digo"
  rsCliFor.Seek "=", Val(Aux)
  If rsCliFor.NoMatch Then
    '28/10/2002 - mpdea
    'Somente exibe o aviso se n�o estiver em navega��o dos registros
    If Not mblnInShowRecord Then
      DisplayMsg "Cliente incorreto."
      cboCliente.SetFocus
    End If
    Exit Sub
  End If
  
  '01/10/2002 - mpdea
  'Somente exibe os avisos se n�o estiver em navega��o dos registros
  If Not mblnInShowRecord Then
  
    '18/09/2002 - mpdea
    'Verifica se o cliente est� bloqueado ou inativo
    If rsCliFor("Bloqueado") Then
      DisplayMsg "Cliente [" & rsCliFor.Fields("Nome").Value & "] est� bloqueado."
      Call SelectAllText(cboCliente, True)
      Exit Sub
    End If
    
    If rsCliFor("Inativo") Then
      DisplayMsg "Cliente [" & rsCliFor.Fields("Nome").Value & "] est� inativo."
      Call SelectAllText(cboCliente, True)
      Exit Sub
    End If
  End If
  
  Nome_Cliente.Caption = rsCliFor("Nome") & ""
    
  If Nome_Digitador.Caption = "" Then
    If rsCliFor("Vendedor") <> 0 Then
      cboDigitador.Text = rsCliFor("Vendedor") & ""
      cboDigitador_LostFocus
    End If
  End If
  
  '28/10/2005 - mpdea
  'Movido c�digo, pois dependia de possuir vendedor para o cliente
  'para exibir a tabela de pre�os
  If Len(Trim(rsCliFor.Fields("TabelaPrecoPadrao"))) > 0 Then
    Combo_Pre�o.Text = rsCliFor.Fields("TabelaPrecoPadrao") & ""
    '06/04/2004 - mpdea
    'Movido execu��o para o final do c�digo onde sempre � realizado
'      Combo_Pre�o_LostFocus
  End If
  
  Desconto_Cli = gsHandleNull(rsCliFor("Desconto") & "")
 
  
  '23/05/2006 - mpdea
  'Cliente isento em IPI
  m_blnIsentoIPI = rsCliFor.Fields("IsentoIPI").Value


  Estado = ""
  rsEstados.Index = "Estado"
  If IsNull(rsCliFor("Estado")) Then Exit Sub
  If rsCliFor("Estado") <> "" Then
    rsEstados.Seek "=", rsCliFor("Estado")
    If Not rsEstados.NoMatch Then
      Estado = rsEstados("Estado")
    End If
  End If

  '--------------------------------------------------------------------------------------------------
  ' TRATAMENTO PARA DESTAQUE DE ICMS PARA EMPRESAS <LUCRO REAL>....'N�O SIMPLES'
  '--------------------------------------------------------------------------------------------------
  If gblnSimplesNacional = False Then
      aliquotaICMS_tab_ICMS_PERC_ESTADOS = ""

      If Not (rsEstadosICMS.EOF And rsEstadosICMS.BOF) Then
          rsEstadosICMS.MoveFirst
          While Not rsEstadosICMS.EOF
              If UCase(rsEstadosICMS.Fields("ESTADO_ORIGEM").Value) = UCase(gsEstadoOrigemEmpresaLogado) And _
                UCase(rsEstadosICMS.Fields("ESTADO_DESTINO").Value) = UCase(rsCliFor("Estado")) Then
                  aliquotaICMS_tab_ICMS_PERC_ESTADOS = rsEstadosICMS.Fields("ALIQUOTA").Value
                  rsEstadosICMS.MoveLast
              End If
              rsEstadosICMS.MoveNext
          Wend
      End If

      If UCase(gsEstadoOrigemEmpresaLogado) = UCase(rsCliFor("Estado")) Then
          bo_AliquotaICMS_interestadual = False
      Else
          bo_AliquotaICMS_interestadual = True
      End If
  End If
  '--------------------------------------------------------------------------------------------------
  '--------------------------------------------------------------------------------------------------

  
  '06/04/2004 - mpdea
  'Realiza sempre o recalculo dos pre�os devido a poss�veis
  'modifica��es de desconto
'  Call Combo_Pre�o_LostFocus
End Sub

Private Sub cboDigitador_CloseUp()
  cboDigitador.Text = cboDigitador.Columns(2).Text
  cboDigitador_LostFocus
End Sub

Public Sub cboDigitador_LostFocus() '16/04/2004 - Daniel - Mudado para Public
 Dim Aux As Variant
 
 Call StatusMsg("")
  
 'ActiveBar1.Tools("miComplDesconto").Enabled = False
  
 Nome_Digitador.Caption = ""
 
 Aux = cboDigitador.Text
 If IsNull(Aux) Then Exit Sub
 If Aux = "" Then Exit Sub
 If Not IsNumeric(Aux) Then Exit Sub
 If Val(Aux) < 1 Then Exit Sub
 If Val(Aux) > 9999 Then Exit Sub
 
 rsFuncionarios.Index = "C�digo"
 rsFuncionarios.Seek "=", Val(Aux)
 
 If rsFuncionarios.NoMatch Then
   DisplayMsg "Funcion�rio incorreto."
   Exit Sub
 End If
 
 Nome_Digitador.Caption = rsFuncionarios("Nome")
  
 'ActiveBar1.Tools("miComplDesconto").Enabled = rsFuncionarios("bPermiteDesconto")

End Sub

Private Sub cboOper_CloseUp()
  cboOper.Text = cboOper.Columns(1).Text
  cboOper_LostFocus
End Sub

Public Sub cboOper_LostFocus() '16/04/2004 - Daniel - Mudado para Public
 Dim Aux As Variant
 
 Call StatusMsg("")
 
 Nome_Opera��o.Caption = ""
 
 '04/12/2007 - Anderson
 'Vari�vel criada para verificar se a opera��o soma o total de produtos na nota
 blnSomarProdutosTotalNota = False
 Calcula_ICM = False
 Calcula_IPI = False
 gbBaseICMSomadoIPI = False
 Calcula_IPI_TOT = False
 
 '11/11/2008 - mpdea
 m_blnSomaIcmsRetidoTotalNota = False
 
 Calcula_ICM_Frete = False
 Soma_Frete = False
 '12/04/2005 - Daniel
 'Tratamento para soma de seguro ao
 'total a receber
 Soma_Seguro = False
 
 '14/08/2002 - mpdea
 'Nr. Terminal
 txtNrTerminal.Visible = False
 lblNrTerminal.Visible = False
 
 '09/10/2002 - mpdea
 'Posiciona o recordset como NoMatch
 rsOp_Sa�da.Index = "C�digo"
 rsOp_Sa�da.Seek "=", 0
 
 Aux = cboOper.Text
 If IsNull(Aux) Then
    Exit Sub
 End If
 
 If Aux = "" Then
    Exit Sub
 End If
 
 If Not IsNumeric(Aux) Then
    Exit Sub
 End If
 
 If Val(Aux) < 1 Then
    Exit Sub
 End If
 
 If Val(Aux) > 999 Then
    Exit Sub
 End If
 rsOp_Sa�da.Seek "=", Val(Aux)
 
 If rsOp_Sa�da.NoMatch Then
   DisplayMsg "Opera��o incorreta."
   cboOper.SetFocus
   cboOper.SelStart = 0
   cboOper.SelLength = Len(cboOper.Text)
   Exit Sub
 End If
 
 Nome_Opera��o = rsOp_Sa�da("Nome")
 
 '04/12/2007 - Anderson
 'Vari�vel criada para verificar se a opera��o soma o total de produtos na nota
 blnSomarProdutosTotalNota = rsOp_Sa�da("SomarProdutosTotalNota")
 Calcula_ICM = rsOp_Sa�da("ICM")
 Calcula_IPI = rsOp_Sa�da("IPI")
 gbBaseICMSomadoIPI = rsOp_Sa�da("Base ICM com IPI")
 Calcula_IPI_TOT = rsOp_Sa�da("IPI TOT")
 Calcula_ICM_Frete = rsOp_Sa�da("Calcula Icm Frete")
 Soma_Frete = rsOp_Sa�da("Soma Frete")
 '12/04/2005 - Daniel
 'Tratamento para soma de seguro ao
 'total a receber
 Soma_Seguro = rsOp_Sa�da("SomarSeguro").Value
 
  '11/11/2008 - mpdea
  m_blnSomaIcmsRetidoTotalNota = rsOp_Sa�da.Fields("SomaIcmsRetidoTotalNota").Value

 '08/08/2002 - mpdea
 'Verifica a tabula��o em caso de or�amento
 txtNrTerminal.Visible = rsOp_Sa�da.Fields("Tipo").Value = "O"
 lblNrTerminal.Visible = txtNrTerminal.Visible
 
 sTipoOperacaoSaida = rsOp_Sa�da.Fields("Tipo").Value
 
 '01/08/2005 - Daniel
 '
 'Tratamento para o Campo Sa�das.DataEmissaoNota
 'Solicitante: Ped�gio Cal�ados e Confec��es
 'Projeto    : Impress�o de Notas Manuais
 If rsOp_Sa�da("EmitirNFManualmente").Value Then
    lblDataEmissaoNotaManual.Visible = True
    mskDataEmissaoNotaManual.Visible = True
 Else
    lblDataEmissaoNotaManual.Visible = False
    mskDataEmissaoNotaManual.Visible = False
 End If
 
 
' If Calcula_ICM = True And Not IsNull(rsOp_Sa�da("Perc Icms Frete")) Then
'    PercIcmsFrete = rsOp_Sa�da("Perc Icms Frete")
' Else
'    PercIcmsFrete = 0
' End If
 
 F_Empr�stimo.Visible = False
 If rsOp_Sa�da("Tipo") = "E" Then F_Empr�stimo.Visible = True
  
  'For�a a atualiza��o dos valores de impostos
  Dim nRow As Integer
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Recalculando...")
  'Refaz o pre�o ao alterar a tabela de pre�os
  For nRow = 0 To Linhas_Grade - 1
   Call Calcula_Linha_Tabe(nRow)
  Next nRow
  
  On Error Resume Next
  If rsSaidas.Fields("Efetivada").Value = False Then
      'Recalcula valores
      Call Recalcula
  End If

  
  '21/12/2006 - Anderson
  'Linhas retiradas para evitar erro de uso do Quick Store na tela de vendas
  'Descri��o do Erro: AO digitar um c�digo inv�lido, o quick exibe uma mensagem de erro e coloca o foco na
  'coluna do c�digo do produto. O usu�rio usava as teclas de movimenta��o para a direita e depois para a
  'esquerda e abria a combo para selecionar um produto. Assim que escolhia o produto correto, o Quick n�o
  'estava atualizando os valores de impostos como por exemplo ICMS.
  'With Grade1
  '  .MoveLast
  '  .MoveFirst
  'End With
  
  '19/01/2004 - Daniel
  'Case.......: PSV Inform�tica
  'Finalidade.: Comp�r o field Validade em Opera��es Sa�da
  If m_blnPSV Then
     If Not VerificaSeExisteValidade Then
      mskValidade.Enabled = False
      mskValidade.BackColor = &H808080
      mskValidade.Mask = ""
      mskValidade.Text = ""
      mskValidade.Mask = "##/##/####"
      'Foi recomendado n�o mostrar...
      lblValidade.Visible = False
      mskValidade.Visible = False
     Else
      mskValidade.Enabled = True
      mskValidade.BackColor = &H80000005
      lblValidade.Visible = True
      mskValidade.Visible = True
     End If
     
  End If
  '----------------------------------------------------------
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
End Sub


Private Sub Combo_Operador_Click()
  Combo_Operador.Text = Combo_Operador.Columns(2).Text
End Sub

Private Sub Combo_Operador_CloseUp()

  Combo_Operador.Text = Combo_Operador.Columns(2).Text
  Combo_Operador_LostFocus

End Sub

Public Sub Combo_Operador_LostFocus() '16/04/2004 - Daniel - Mudado para Public
 Dim Aux As Variant
 
 Call StatusMsg("")
 Nome_Operador.Caption = ""
 
 Aux = Combo_Operador.Text
 If IsNull(Aux) Then Exit Sub
 If Aux = "" Then Exit Sub
 If Not IsNumeric(Aux) Then Exit Sub
 If Val(Aux) < 1 Then Exit Sub
 If Val(Aux) > 9999 Then Exit Sub
 
 rsOperadores.Index = "C�digo"
 rsOperadores.Seek "=", Val(Aux)
 
 If rsOperadores.NoMatch Then
   DisplayMsg "Operador incorreto."
   Exit Sub
 End If
 
 Nome_Operador.Caption = rsOperadores("Nome")
 
 

End Sub

Private Sub Combo_Pre�o_Click()
  Combo_Pre�o.Text = Combo_Pre�o.Columns(0).Text
End Sub

Private Sub Combo_Pre�o_CloseUp()
  Dim nRow As Integer
  Dim bm As Variant
  Dim strFullCode As String
  Dim strCodProd As String
  Dim intErro As Integer
  Dim Aux_Pre�o As Double
  
  
  If Trim(Combo_Pre�o.Text) <> "" Then
    Screen.MousePointer = vbHourglass
    Call StatusMsg("Refazendo tabela...")
    'Refaz o pre�o ao alterar a tabela de pre�os
    rsPre�os.Index = "Tabela"
    For nRow = 0 To Linhas_Grade - 1
      strFullCode = gsHandleNull(Tabe(nRow).C�digo)
      '-------------------------------------------------------------------------------
      '07/05/2002 - mpdea
      '
      'Alterado para que localize o pre�o para os produtos do tipo Grade e Edi��o
      '(Procura c�digo principal)
      '-------------------------------------------------------------------------------
      Call Acha_Produto(strFullCode, strCodProd, 0, 0, 0, 0, intErro)
      If intErro = 0 Then
        If strFullCode <> "0" Then
'          rsPre�os.Seek "=", Combo_Pre�o.Text, strCodProd
'          If rsPre�os.NoMatch Then
'            Tabe(nRow).Pre�o = 0
'          Else
'            Tabe(nRow).Pre�o = rsPre�os("Pre�o")
'          End If
          
          '---------------------------------------------------------------------------
          '06/04/2004 - mpdea
          '
          'Alterado para que inclua o desconto do produto, cliente e cota��o no
          'c�lculo do pre�o
          '---------------------------------------------------------------------------
          'Posiciona recordset
          rsProdutos2.FindFirst "C�digo = '" & strCodProd & "'"
          'Acha pre�o
          If Combo_Pre�o.Text = "" Then
            Tabe(nRow).Pre�o = 0
          Else
            rsPre�os.Index = "Tabela"
            rsPre�os.Seek "=", Combo_Pre�o.Text, strCodProd
            If rsPre�os.NoMatch Then
              Tabe(nRow).Pre�o = 0
            Else
               Aux_Pre�o = rsPre�os("Pre�o") * ((100 - (rsProdutos2("Desconto") + Desconto_Cli)) / 100)
               If rsProdutos2("Moeda") <> 1 Then
                 rsCota��es.Index = "Moeda"
                 rsCota��es.Seek "<=", rsProdutos2("Moeda"), Data_Atual
                 If Not rsCota��es.NoMatch Then
                   If rsCota��es("Moeda") = rsProdutos2("Moeda") Then
                     Aux_Pre�o = Aux_Pre�o * rsCota��es("Cota��o")
                   End If
                 End If
               End If
               
               '04/05/2004 - Daniel
               'Personaliza��o Embalavi
               If g_bln5CasasDecimais Then
                Tabe(nRow).Pre�o = Format(Aux_Pre�o, "#0.00000")
              '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
               ElseIf g_bln3CasasDecimais Then
                Tabe(nRow).Pre�o = Format(Aux_Pre�o, "#0.000")
               Else
                Tabe(nRow).Pre�o = Format(Aux_Pre�o, "#0.00")
               End If
               
            End If
          End If
          '---------------------------------------------------------------------------
          
        Else
          Tabe(nRow).Pre�o = 0
        End If
      Else
        Tabe(nRow).Pre�o = 0
      End If
      Call Calcula_Linha_Tabe(nRow)
    Next nRow

    On Error Resume Next
    If rsSaidas.Fields("Efetivada").Value = False Then
      'Recalcula valores
      Call Recalcula
    End If
    
    With Grade1
      .MoveLast
      .MoveFirst
    End With
    Screen.MousePointer = vbDefault
    Call StatusMsg("")
  End If


  
'  Call RecalculaPrecos
'  With Grade1
'    .MoveFirst
'    For nRow = 0 To Linhas_Grade - 1
'      If Tabe(nRow).C�digo <> "0" Then
'        .Columns("Pre�o Unit.").Text = Tabe(nRow).Pre�o
'        .Columns("Total").Text = Tabe(nRow).Pre�o_Total
'        .Columns("Pre�o Final").Text = Tabe(nRow).Pre�o_Final
'        Call Calcula_Linha
'      End If
'      .MoveNext
'    Next nRow
'    Call Recalcula
'    .MoveLast
'    .MoveFirst
'  End With
  
End Sub

Private Sub Combo_Pre�o_LostFocus()
  If IsNull(Combo_Pre�o.Text) Then
    Exit Sub
  ElseIf Combo_Pre�o.Text = "" Then
    Exit Sub
  Else
    Combo_Pre�o.Text = UCase(Combo_Pre�o.Text)
  End If
  Call Combo_Pre�o_CloseUp
End Sub

Private Sub Combo_T�cnico_CloseUp()
 Combo_T�cnico.Text = Combo_T�cnico.Columns(2).Text
 Combo_T�cnico_LostFocus
End Sub

Private Sub Combo_T�cnico_LostFocus()

  Nome_T�cnico.Caption = ""
  If IsNull(Combo_T�cnico.Text) Then Exit Sub
  If Combo_T�cnico.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_T�cnico.Text) Then Exit Sub
  If Val(Combo_T�cnico.Text) < 1 Then Exit Sub
  If Val(Combo_T�cnico.Text) > 9999 Then Exit Sub
  
  
  
  rsT�cnicos.Index = "C�digo"
  rsT�cnicos.Seek "=", Val(Combo_T�cnico.Text)
  If rsT�cnicos.NoMatch Then Exit Sub
  Nome_T�cnico.Caption = rsT�cnicos("Nome") & ""
  
End Sub


Private Sub Command1_Click()
  frmFundoCombateAPobreza.Show


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

Private Sub DropDown1_Click()
Rem Acha pre�o e mostra
'' rsPre�os.Index = "Tabela"
'' rsPre�os.Seek "=", Combo_Pre�o.Text, DropDown1.Columns(1).Text
''
'' If rsPre�os.NoMatch Then
''    Grade1.Columns(4).Text = ""
''  Else
''    Grade1.Columns(4).Text = Format$(rsPre�os("Pre�o"), "###,###,##0.00")
'' End If
'' Grade1.Columns(0).Text = DropDown1.Columns(1).Text
'' Grade1.Columns(2).Text = DropDown1.Columns(0).Text
'' Call RecalculaPrecos
'
' Call Grade1_BeforeColUpdate(0, 0, 0)
' Call Calcula_Linha
End Sub

Private Sub DropDown1_CloseUp()
  
  With DropDown1
'    .DataFieldToDisplay = "C�digo"
    rsPre�os.Index = "Tabela"
    rsPre�os.Seek "=", Combo_Pre�o.Text, .Columns("C�digo").Text
    If rsPre�os.NoMatch Then
      Grade1.Columns("Pre�o Unit.").Text = "0.00"
    Else
      '04/05/2004 - Daniel
      'Personaliza��o Embalavi
      If g_bln5CasasDecimais Then
        Grade1.Columns("Pre�o Unit.").Text = Format$(rsPre�os("Pre�o"), "###,###,##0.00000")
      '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Grade1.Columns("Pre�o Unit.").Text = Format$(rsPre�os("Pre�o"), "###,###,##0.000")
      Else
        Grade1.Columns("Pre�o Unit.").Text = Format$(rsPre�os("Pre�o"), "###,###,##0.00")
      End If
    End If
'    Grade1.Columns("Pre�o Unit.").Text = Format$(gsHandleNull(.Columns("Pre�o").Text), "###,###,##0.00")
    Grade1.Columns("C�digo").Text = .Columns("C�digo").Text
    Grade1.Columns("Nome").Text = .Columns("Nome").Text
    
'    Call RecalculaPrecos
  End With
  
  '21/12/2006 - Anderson
  'for�a a execu��o do evento para evitar erro de uso do Quick Store na tela de vendas
  'Descri��o do Erro: AO digitar um c�digo inv�lido, o quick exibe uma mensagem de erro e coloca o foco na
  'coluna do c�digo do produto. O usu�rio usava as teclas de movimenta��o para a direita e depois para a
  'esquerda e abria a combo para selecionar um produto. Assim que escolhia o produto correto, o Quick n�o
  'estava atualizando os valores de impostos como por exemplo ICMS.
  Call Grade1_BeforeColUpdate(0, 0, 0)
  'Call Calcula_Linha
'' Call RecalculaPrecos
  
  
'''  DoEvents
'''  Grade1.Columns(0).Text = DropDown1.Columns(1).Text
'''  Grade1.Columns(1).Text = "1"
'''  Grade1.Columns(2).Text = Trim(DropDown1.Columns(0).Text)
'''  Call StatusMsg("")

'  Grade1.Columns(Grade1.Col).Text = DropDown1.Columns(1).Text
'  If Len(Trim(Grade1.Columns(1).Text)) > 0 Then
'    If Val(Grade1.Columns(1).Text) = 0 Then
'      Grade1.Columns(1).Text = "1"
'    End If
'  End If
'  Grade1.Columns(0).Text = DropDown1.Columns(1).Text
'  Grade1.Columns(2).Text = DropDown1.Columns(0).Text
'  Call StatusMsg("")
'  Lista_Aberta = False
'  Call Calcula_Linha
End Sub

Private Sub DropDown1_DropDown()
'  Lista_Aberta = True
'  DropDown1.DataFieldToDisplay = "C�digo"
  'Data1.Recordset.FindFirst "C�digo = '" & Grade1.Columns(0).Text & "'"
  'Grade1.Columns(0).Text = Grade1.Columns(2).Text
'  With DropDown1
'    rsProdutos.Index = "C�digo"
'    rsProdutos.Seek "=", .Text
'    If rsProdutos.NoMatch Then
'      .DataFieldList = "Nome"
'    Else
'      .DataFieldList = "C�digo"
'    End If
'  End With
  
  Dim rsTemp As Recordset
  Set rsTemp = db.OpenRecordset("SELECT C�digo FROM Produtos WHERE C�digo = '" & Grade1.Columns("C�digo").Text & "'", dbOpenSnapshot)
  If rsTemp.EOF Then
    DropDown1.DataFieldList = "Nome"
  Else
    DropDown1.DataFieldList = "C�digo"
    'Data1.Recordset.FindFirst "C�digo = '" & DropDown1.Columns("C�digo").Text & "'"
  End If
  rsTemp.Close
  Set rsTemp = Nothing
End Sub


Private Sub Dropdown1_RowLoaded(ByVal Bookmark As Variant)
  Dim nEstoque As Double
  Dim sMsgEstoque As String
  Dim nErro As Integer
  
  With DropDown1
    'Estoque
    nEstoque = Acha_Estoque(gnCodFilial, .Columns("C�digo").Text, 0, 0, 0, nErro)
    Select Case nErro
      Case 0
        '26/08/2004 - Daniel
        'Criado valida��o para verificar se o usu�rio possui permiss�o
        'para enchergar o estoque ou n�o
        If m_blnPermitido Then
          sMsgEstoque = nEstoque
        Else
          sMsgEstoque = "Usu�rio n�o permitido"
        End If
      Case 1
        sMsgEstoque = "Estoque n�o iniciado"
      Case 2
        sMsgEstoque = "Depende da grade"
      Case 3
        sMsgEstoque = "Depende da edi��o"
      Case 4
        sMsgEstoque = "Produto n�o existe"
    End Select
    .Columns("Estoque").Text = sMsgEstoque
    'Pre�o
    If Combo_Pre�o.Text = "" Then
      .Columns("Pre�o").Text = "Pre�o n�o encontrado"
    Else
      rsPre�os.Index = "Tabela"
      rsPre�os.Seek "=", Combo_Pre�o.Text, .Columns("C�digo").Text
      If rsPre�os.NoMatch Then
        .Columns("Pre�o").Text = "Pre�o n�o encontrado"
      Else
        '04/05/2004 - Daniel
        'Personaliza��o Embalavi
        If g_bln5CasasDecimais Then
          .Columns("Pre�o").Text = Format(rsPre�os("Pre�o"), "##,###,##0.00000")
        '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
          .Columns("Pre�o").Text = Format(rsPre�os("Pre�o"), "##,###,##0.000")
        Else
          '.Columns("Pre�o").Text = Format(rsPre�os("Pre�o"), Formato_Pre�o)
          .Columns("Pre�o").Text = Format(rsPre�os("Pre�o"), "##,###,##0.00")
        End If
      End If
    End If
    'Fim pre�o
    
    '---[Coluna Fabricante]---
    '29/03/2005 - Daniel
    'Case: El�trica Leal - Mostrar o fabricante
    '
    '12/05/2005 - Daniel
    'Rotina parametrizada conforme solicita��o da Info Social
    If m_blnExibirColunaFabricante Then
      If Len(.Columns("C�digo").Text) > 0 Then .Columns("Fabricante").Text = GetFabricante(.Columns("C�digo").Text) & ""
    End If
    '---[Fim Fabricante]---
    
  End With
End Sub

Private Sub DropDown2_CloseUp()
  With Grade_Serv
    .Columns("C�digo").Text = DropDown2.Columns("C�digo").Text
    .Columns("Descri��o").Text = Trim(DropDown2.Columns("Descri��o").Text)
    .Columns("Pre�o").Text = DropDown2.Columns("Pre�o").Text
  End With
  Call StatusMsg("")
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case 0
          gsOrder = " ORDER BY Sequ�ncia"
        Case 1
          gsOrder = " ORDER BY Data, Sequ�ncia"
        Case 2
          gsOrder = " ORDER BY Cliente, Sequ�ncia"
        Case 3
          gsOrder = " ORDER BY Refer�ncia"
        Case 4
          gsOrder = " ORDER BY [Nota Impressa]"
      End Select
  End Select
  Set rsSaidas = db.OpenRecordset(gsSql & gsWhere & gsOrder, dbOpenDynaset)
End Sub

Public Sub limparVariaveisDevolucaoProduto()
  sCodigoProdutoDevolucao = ""
  lQuantidadeItensProdutoDevolucao = 0
  sNomeProdutoDevolucao = ""
  sValorUnitarioProdutoDevolucao = ""
End Sub

Private Function FormataValorTextoComVirgula(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTextoComVirgula = Format(dblValor, "#0." & String(lngCasasDecimais, "0"))
End Function

Private Sub GeraXML_ticket()
  On Error GoTo Erro
  
  Dim rsVenda           As Recordset
  Dim rsVendaProdutos   As Recordset
  Dim sSql              As String
  Dim sXML_Ticket       As String
  Dim sDataVenda        As String
  Dim sHoraVenda        As String
  Dim sTotalProdutos    As String
  Dim dTotalProdutos    As Double
  Dim sTotalPago        As String
  Dim dTotalPago        As Double
  Dim sAuxValor1        As String
  Dim sAuxValor2        As String
  Dim sTotalDesconto    As String
  Dim sCodCliente       As String
  Dim sNomCliente       As String
  Dim sNomVendedor      As String
  Dim sNumeroLinhasProd As String
  Dim sCaixa            As String
  
  Dim rsParametros      As Recordset
  Dim sBancoPDV         As String
  
  sSql = "SELECT * FROM [Par�metros Filial] Where Filial = " & gnCodFilial
  Set rsParametros = db.OpenRecordset(sSql, dbOpenDynaset)
  sBancoPDV = rsParametros("BancoPDV").Value
  rsParametros.Close
  Set rsParametros = Nothing
  
  sSql = "SELECT S.Data, S.NSU_Hora, S.Produtos, S.Total, S.Digitador, S.Cliente, C.Nome, F.Nome as NomeVendedor, S.Caixa "
  sSql = sSql & " FROM Sa�das S, Cli_For C, Funcion�rios F "
  sSql = sSql & " Where S.Filial = " & gnCodFilial & " and "
  sSql = sSql & " S.Sequ�ncia = " & txtSeq.Text & " and "
  sSql = sSql & " S.Cliente = C.C�digo and "
  sSql = sSql & " S.Digitador = F.C�digo "
  Set rsVenda = db.OpenRecordset(sSql, dbOpenDynaset)
  
  If Not (rsVenda.EOF And rsVenda.BOF) Then
      sDataVenda = rsVenda.Fields("Data").Value
      
      If Not IsNull(rsVenda.Fields("NSU_Hora").Value) Then
          sHoraVenda = rsVenda.Fields("NSU_Hora").Value
      Else
          sHoraVenda = "00:00"
      End If
      
      dTotalProdutos = rsVenda.Fields("Produtos").Value
      sAuxValor1 = FormataValorTexto(dTotalProdutos, 2)
      sAuxValor1 = Replace(sAuxValor1, ".", ",")
      sTotalProdutos = sAuxValor1
  
      dTotalPago = rsVenda.Fields("Total").Value
      sAuxValor2 = FormataValorTexto(dTotalPago, 2)
      sAuxValor2 = Replace(sAuxValor2, ".", ",")
      sTotalPago = sAuxValor2
      
      sTotalDesconto = CDbl(sAuxValor1) - CDbl(sAuxValor2)
      sTotalDesconto = FormataValorTexto(sTotalDesconto, 2)
      sTotalDesconto = Replace(sTotalDesconto, ".", ",")
      
      sCodCliente = rsVenda.Fields("Cliente").Value
      sNomCliente = rsVenda.Fields("Nome").Value
      sNomVendedor = rsVenda.Fields("NomeVendedor").Value
      sCaixa = rsVenda.Fields("Caixa").Value
  End If
  rsVenda.Close
  Set rsVenda = Nothing
  
  sXML_Ticket = ""
  sXML_Ticket = "<TicketQS>"
  
  ' =========================================================
  ' Dados empresa/filial
  sXML_Ticket = sXML_Ticket & "<DadosCabecalho>"
  sXML_Ticket = sXML_Ticket & "<Nome>" & gsNomeFilial & "</Nome>"
  sXML_Ticket = sXML_Ticket & "<Endereco>" & gsFilialEndereco & "</Endereco>"
  sXML_Ticket = sXML_Ticket & "<Bairro>" & gsFilialBairro & "</Bairro>"
  sXML_Ticket = sXML_Ticket & "<CidadeEstado>" & gsFilialCidadeEstado & "</CidadeEstado>"
  sXML_Ticket = sXML_Ticket & "<Cep>" & gsFilialCep & "</Cep>"
  sXML_Ticket = sXML_Ticket & "<Fone>" & gsFilialFone & "</Fone>"
  
  sXML_Ticket = sXML_Ticket & "<CodCliente>" & sCodCliente & "</CodCliente>"
  sXML_Ticket = sXML_Ticket & "<NomCliente>" & sNomCliente & "</NomCliente>"
  sXML_Ticket = sXML_Ticket & "<NomVendedor>" & sNomVendedor & "</NomVendedor>"
  sXML_Ticket = sXML_Ticket & "<Sequencia>" & txtSeq.Text & "</Sequencia>"
  sXML_Ticket = sXML_Ticket & "<DataVenda>" & sDataVenda & "</DataVenda>"
  sXML_Ticket = sXML_Ticket & "<HoraVenda>" & sHoraVenda & "</HoraVenda>"
  sXML_Ticket = sXML_Ticket & "<Caixa>" & sCaixa & "</Caixa>"
  
  sXML_Ticket = sXML_Ticket & "</DadosCabecalho>"
  ' =========================================================
  
  ' =========================================================
  ' Dados Produtos
  
  sSql = "SELECT S.C�digo, P.Nome, S.Qtde, S.Pre�o, S.Desconto, S.[Pre�o Final] as PrecoFinal, S.Linha "
  sSql = sSql & " FROM [Sa�das - Produtos] S, Produtos P "
  sSql = sSql & " Where S.Filial = " & gnCodFilial & " and "
  sSql = sSql & " S.Sequ�ncia = " & txtSeq.Text & " and "
  sSql = sSql & " S.[C�digo sem Grade] = P.C�digo "
  sSql = sSql & " Order by S.Linha "
  Set rsVendaProdutos = db.OpenRecordset(sSql, dbOpenDynaset)
  
  If Not (rsVendaProdutos.EOF And rsVendaProdutos.BOF) Then
    rsVendaProdutos.MoveLast
    sNumeroLinhasProd = rsVendaProdutos.RecordCount
    rsVendaProdutos.MoveFirst
    
    sXML_Ticket = sXML_Ticket & "<TotalLinhasProduto>" & sNumeroLinhasProd & "</TotalLinhasProduto>"
    sXML_Ticket = sXML_Ticket & "<Produtos>"
  
    While Not rsVendaProdutos.EOF
    
        If rsVendaProdutos.Fields("C�digo").Value <> "" Then
        
          sXML_Ticket = sXML_Ticket & "<LinhaProduto" & rsVendaProdutos.Fields("Linha").Value & ">"
      
          sXML_Ticket = sXML_Ticket & "<CodProduto>" & rsVendaProdutos.Fields("C�digo").Value & "</CodProduto>"
          sXML_Ticket = sXML_Ticket & "<NomProduto>" & rsVendaProdutos.Fields("Nome").Value & "</NomProduto>"
          sXML_Ticket = sXML_Ticket & "<QtdeProduto>" & rsVendaProdutos.Fields("Qtde").Value & "</QtdeProduto>"
          sXML_Ticket = sXML_Ticket & "<PrecoProduto>" & FormataValorTextoComVirgula(rsVendaProdutos.Fields("Pre�o").Value, 2) & "</PrecoProduto>"
          sXML_Ticket = sXML_Ticket & "<DescProduto>" & FormataValorTextoComVirgula(rsVendaProdutos.Fields("Desconto").Value, 2) & "</DescProduto>"
          sXML_Ticket = sXML_Ticket & "<PrecoFinalProduto>" & FormataValorTextoComVirgula(rsVendaProdutos.Fields("PrecoFinal").Value, 2) & "</PrecoFinalProduto>"
      
          sXML_Ticket = sXML_Ticket & "</LinhaProduto" & rsVendaProdutos.Fields("Linha").Value & ">"
        End If
        
    
        rsVendaProdutos.MoveNext
    Wend
  End If
  rsVendaProdutos.Close
  Set rsVendaProdutos = Nothing
  
  sXML_Ticket = sXML_Ticket & "</Produtos>"
  ' =========================================================
  
  ' =========================================================
  ' Dados Totais
  sXML_Ticket = sXML_Ticket & "<Totais>"
  sXML_Ticket = sXML_Ticket & "<SubTotal>" & sTotalProdutos & "</SubTotal>"
  sXML_Ticket = sXML_Ticket & "<TotalDesconto>" & sTotalDesconto & "</TotalDesconto>"
  sXML_Ticket = sXML_Ticket & "<Total>" & sTotalPago & "</Total>"
  sXML_Ticket = sXML_Ticket & "</Totais>"
  ' =========================================================
  
  sXML_Ticket = sXML_Ticket & "</TicketQS>"
  
  ' =========================================================
  ' Gravar no banco de dados local do IMPRESSOR
  Dim rsTesteConexao As Recordset
  On Error GoTo SegueFluxo
    
  Set rsTesteConexao = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = ''")
SegueFluxo:

  If Err.Number <> 0 Then
      Set BancoPDV = OpenDatabase(sBancoPDV & "\QuickStore.mdb", False, False, ";PWD='" & gsGetPValue() & "'")
  End If
  
  Dim rsNFCe_Job As Recordset
  Set rsNFCe_Job = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = 'XXXXXXX' And cnpj = '999999'")

  Dim iTipo As Integer
' iTipo = 1     ' N�o esta em contingencia
' iTipo = 2     ' Esta em contingencia
  iTipo = 9     ' XML Impress�o de Ticket do QuickStore

  If rsNFCe_Job.EOF Then
      rsNFCe_Job.AddNew
      rsNFCe_Job!CNPJ = "TICKET_QS"
      rsNFCe_Job!xml = sXML_Ticket
      rsNFCe_Job!Tipo = iTipo
      rsNFCe_Job!Serie = 1
      rsNFCe_Job!N_NF = txtSeq.Text
      rsNFCe_Job!Chave = txtSeq.Text
      rsNFCe_Job!CPF = "TICKET_QS"
      rsNFCe_Job!Nome_Consumidor = "TICKET_QS"
      rsNFCe_Job!Data_Emissao = ""
      rsNFCe_Job!Total_Tributos = ""
      rsNFCe_Job!Nome_Emitente = ""
      rsNFCe_Job!Endereco_Emitente = ""
      rsNFCe_Job!IE_Emitente = ""
      rsNFCe_Job!retFazenda = "SEM RET"
      rsNFCe_Job.Update
  End If

  rsNFCe_Job.Close
  Set rsNFCe_Job = Nothing
  ' =========================================================

  Exit Sub
Erro:
  MsgBox "Inconsist�ncia em rotina GeraXML_ticket " & Err.Number & " " & Err.Description, vbInformation, "Aten��o"
End Sub


Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
Call cmd_acataUsuarioLogadoComoOperador_Click
  Select Case Tool.Name
    Case "miComplPesquisaVendasHj"
        gOrigemTelaSaidasChamadorDaTelaAcharVendaHoje = True
        frmVendasHoje.Show vbModal
    Case "miComplPesquisaCliente"
      Dim objFrmPesquisaCliFor As frmPesquisaCliFor
      Set objFrmPesquisaCliFor = New frmPesquisaCliFor
      objFrmPesquisaCliFor.iOrigemSaidas = True
      objFrmPesquisaCliFor.Show
    Case "miOpFirst"
      limparVariaveisDevolucaoProduto
      Call MoveFirst
    Case "miOpPrevious"
      limparVariaveisDevolucaoProduto
      Call MovePrevious
    Case "miOpNext"
      limparVariaveisDevolucaoProduto
      Call MoveNext
    Case "miOpLast"
      limparVariaveisDevolucaoProduto
      Call MoveLast
    Case "miPesqData"
      Call PesquisaPorData
    Case "miOpClear"
      '25/06/2004 - Daniel
      'Adicionada a Fun��o PossuiPermissao que verificar� se o usu�rio tem
      'ou n�o permiss�o para usar o bot�o de limpar
      'Case: Coneg Campos - Aproveitado para os demais
''''      If PossuiPermissao Then
        Call ClearScreen
''''      Else
''''        If m_blnClear Then
''''          Call ClearScreen
''''          m_blnClear = False
''''        Else
''''          If Not frmGerente.gbSenhaGerente Then
''''            Exit Sub
''''          Else
''''            Call ClearScreen
''''          End If
''''        End If
''''      End If
      '20/03/2013-Alexandre Afornali
      'Ao limpar a tela, o rsSaidas perde o where
      Set rsSaidas = db.OpenRecordset(gsSql & gsOrder, dbOpenDynaset)
    Case "miOpUpdate"
      '01/07/2004 - Daniel
      'Gravou ent�o habilitamos a op��o de limpeza tamb�m
      m_blnClear = True
      Call UpdateRecord
      If (Senha.Text <> "" And txtComanda.Text <> "") Then 'Valida se a senha esta digitada
        'Call UpdateComanda
      End If
      If (txtSeq.Text <> "" And Senha.Text <> "") Then
        Call UpdateTotalNCM
      End If
    Case "miOpDelete"
      '21/10/2013 - Jean
      'Customiza��o para Disk Embalagens que faz com que a senha do gerente seja obrigat�rio no caso de exclus�o de registro de sa�da
'''      If CheckSerialCaseMod("QS73520-469") Then
        If Not frmGerente.gbSenhaGerente Then
            Exit Sub
'''          Else
'''            Call ClearScreen
          End If
'''      End If
      Call DeleteRecord
      Call ClearScreen
      
    Case "miOpSearch"
      Call SearchRecord
    Case "miComplInfo"
      Call GetInformation
    Case "miComplRecebimento"
      Call RealizaRecebimento
    Case "miComplDesconto"
      'grava venda se ainda n�o tiver sido gravada -- PABLO 07/07/2022
      If Trim(txtSeq.Text) = "" Then Call ActiveBar1_Click(ActiveBar1.Tools("miOpUpdate"))
      If Trim(txtSeq.Text) <> "" Then
        Call UpdateTotalNCM
        Call RealizaDesconto
      End If
    Case "miComplFindNextOrcam"
      Call FindNextOrcam
    Case "miComplTransformOrcamVenda"
      Call TransformaOrcamEmVenda
    Case "miComplUndoMovim"
      Call UndoMovimSaida
    Case "miComplPrintNotaFiscal"
      Call PrintNota
    
    Case "miComplPrintTicket"
        If giQuick_viaRDP_ticket = 1 Then
          'Ser� impresso pelo IMPRESSOR EXE c#
          GeraXML_ticket
        Else
          'Impresso padr�o antigo
          Call ImprimirTicket(False)
        End If
    Case "miComplAlteraTotais"
      Call AlteraTotais
    Case "miComplCancelNota"
      Call CancelNota
    Case "miComplConsultaProdutos"
      nChamaConsulta = 2
      Call SearchProduto
    Case "miComplFiltrarCliente" '15/02/2007 - Anderson - Filtrar por cliente na tela de vendas - Solicitado por Paulo Ribertec.
      Tool.Checked = Not Tool.Checked
      Call FiltrarCliente
    Case "miVerificaPedido"
      frmVerificaPedido.Show
    Case "miComplLeitorOtico"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigSAIDAS", "Scanner", Tool.Checked)
    Case "miOpFreezeOperacao"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigSAIDAS", "Mantem Operacao", Tool.Checked)
    Case "miOpFreezeDigitador"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigSAIDAS", "Mantem Digitador", Tool.Checked)
    Case "miOpFreezeCliente"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigSAIDAS", "Mantem Cliente", Tool.Checked)
    Case "miOpFreezeTabPrecos"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigSAIDAS", "Mantem TabPrecos", Tool.Checked)
    Case "miOpEtiquetas"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("ConfigSAIDAS", "Etiqueta Balanca", Tool.Checked)
    Case "miOpReplica"
      'Cancela desconto -- PABLO 07/07/2022
      mcurDescontoSubTotal = 0
      Total_Desconto = 0
      b_EscondeTelaDesconto = True
      Call RealizaDesconto
      b_EscondeTelaDesconto = False
      
      Call ReplicaMov
    Case "miEtiquetaEnderecamento"
      frmEtiquetaEnderecamento.Show
    Case "miEmissaoCertificados"
      frmEmissaoCertificado.Show
    '01/04/2005 - Daniel
    'Adicionado refresh para corrigir problemas de exibi��o
    'das informa��es na tela de sa�das.
    'Exemplo de bug: Ao importar o pedido da web a 1� linha
    'da grid estava ficando invis�vel.
    'Case "miRefreshTela"
    '  Call RefreshTela
    Case "miEmissaoDuplicatas"
      '17/06/2005 - Daniel
      '
      'Solicitante: TI Brasil (Pavinato)
      '
      'Inserido rotina para emiss�o de Duplicatas (Faturas)
      'rotina baseada na rotina j� existente na tela de VR
      Call EmitirDuplicatas
      
    '05/03/2007 - Anderson
    'Impress�o customizada de Or�amentos
    Case "miImprimirOrcamento", "miImprimirOrcamentoVenda"
      'QS38785-386 - NewTech Inform�tica
      If CheckSerialCaseMod("QS38785-386") Then
        Call Imprimir_Orcamento(IIf(Tool.Name = "miImprimirOrcamento", "Orcam1.rpt", "Orcam2.rpt"))
      Else
        'Para os outros clientes
        Call Imprimir_Orcamento2(IIf(Tool.Name = "miImprimirOrcamento", "Orcam1.rpt", "Orcam2.rpt"))
      End If
    
    Case "miImprimirOrcamentoVenda_Servicos"
        Call Imprimir_Orcamento2("Orcam4_de_servicos.rpt")
      
    'Eduardo - 17/10/2013
    Case "miImprimirOrcamentoSemFrete" 'Sem frete no subtotal.
    'QS73520-469 - Disk Embalagens Ltda - ME
    Call Imprimir_Orcamento3(IIf(Tool.Name = "miImprimirOrcamento", "Orcam1.rpt", "Orcam3.rpt"))
    
    'Eduardo - 19/11/2013
    Case "LimpaComandas"
      Call ChamaLimpaComandas
    
    '22/06/2007 - Anderson
    'Exportar Sa�das para Excel
    Case "miOpExportarExcel"
      Call ExportarExcel
      
    '27/09/2007 - Anderson
    'Implementa��o da impress�o de carn� com c�digo de barras
    'Solicitado por: Naativa
    Case "miOpImprimirCarne"
      Call ImprimirCarne
      
    '17/10/2007 - Anderson
    'Customiza��o de pesquisa da combo de produtos
    Case "miOpPesquisarCodigo"
      Tool.Checked = True
      ActiveBar1.Tools("miOpPesquisarDescricao").Checked = False
      Call UpdateArqConfig("ConfigSAIDAS", "PesquisarCodigo", Tool.Checked)
      Call UpdateArqConfig("ConfigSAIDAS", "PesquisarDescricao", False)

    '17/10/2007 - Anderson
    'Customiza��o de pesquisa da combo de produtos
    Case "miOpPesquisarDescricao"
      Tool.Checked = True
      ActiveBar1.Tools("miOpPesquisarCodigo").Checked = False
      Call UpdateArqConfig("ConfigSAIDAS", "PesquisarDescricao", Tool.Checked)
      Call UpdateArqConfig("ConfigSAIDAS", "PesquisarCodigo", False)
    
    '30/01/2009 - mpdea
    Case "miEnviarEmail"
      ImprimirTicket True
      
    Case "miComplRetNFCe"
'''      If txtSeq.Text = "" Then
'''        Exit Sub
'''      End If
'''      Dim VerificaRetorno As New clsNFCe
'''      VerificaRetorno.VerificaRetorno (txtSeq.Text)
      
    Case "miComplNFC-e"
      If txtSeq.Text = "" Then
        DisplayMsg "NFC-e s� pode ser emitido a partir de uma venda efetivada. Encontre uma venda efetivada."
        Exit Sub
      End If
      
      If L_Tot_Desc.Text <> "0,00" And gcDescInTotal = 0 Then
        gcDescInTotal = CCur(L_Tot_Desc.Text)
      End If
      
      Dim EnviaNFCe As New clsNFCe
      EnviaNFCe.EnviaNFCe txtSeq.Text, gcDescInTotal
    Case "miOrcamentoExpresso"
      'Esta opcao � igual a Imprime Orcamento...apenas aqui tem um bot�o de atalho que imprime direto o orcamento
'      sMENSAGEM_LOG_TESTE_GERAL = "STEP 1: " & Now & vbCrLf
      ImprimeOrcamentoExpresso
      
    Case "miComplNFe"
      If txtSeq.Text = "" Or L_Efetivada.Visible = False Then
        DisplayMsg "NFe s� pode ser emitido a partir de uma venda efetivada. Encontre uma venda efetivada."
        Exit Sub
      End If
      
      
      origemTelaSaidasParaTelaNFe = txtSeq.Text
      frmNFe.Show
      
    Case "mnImprimeCarneTipo1"
        Call EmiteCarnesNOVOS
  End Select
End Sub

Private Sub ImprimeOrcamentoExpressoToyoBens()
  On Error GoTo ErrHandler
  
  With frmRelatorioTicket
    .Caption = "Sequ�ncia: " & CStr(rsSaidas.Fields("Sequ�ncia").Value)
    .Filial = rsSaidas.Fields("Filial").Value
    .Sequencia = rsSaidas.Fields("Sequ�ncia").Value
    .Show vbModal
  End With
  
  Exit Sub
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub ImprimeOrcamentoExpresso()
On Error GoTo Erro

  '*******************
  'Esta opcao � igual a Imprime Orcamento...apenas aqui tem um bot�o de atalho que imprime direto o orcamento

  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Encontre um registro antes."
    Exit Sub
  End If
  
  If rsParametros("CGC") = "41.070.699/0001-43" Or rsParametros("CGC") = "39.305.002/0001-24" Then
    Call ImprimeOrcamentoExpressoToyoBens
    Exit Sub
  End If

  'Status
  Call StatusMsg("Aguarde...")
  MousePointer = vbHourglass

  With Rel1
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsQuickDBFileName
    .DataFiles(3) = gsQuickDBFileName
    .DataFiles(4) = gsQuickDBFileName
    .DataFiles(5) = gsQuickDBFileName
    .DataFiles(6) = gsQuickDBFileName
    .DataFiles(7) = gsQuickDBFileName
    
'    sMENSAGEM_LOG_TESTE_GERAL = sMENSAGEM_LOG_TESTE_GERAL & "STEP 1.1: " & Now & vbCrLf
    
    .Destination = IIf(False, crptToWindow, crptToPrinter)
    .ReportFileName = gsReportPath & "Orcam1.rpt"
    
    .SelectionFormula = "{Sa�das.Filial} = " & rsSaidas.Fields("Filial").Value & " AND {Sa�das.Sequ�ncia} = " & rsSaidas.Fields("Sequ�ncia").Value

'    sMENSAGEM_LOG_TESTE_GERAL = sMENSAGEM_LOG_TESTE_GERAL & "STEP 1.2: " & Now & vbCrLf

    .Formulas(0) = "mensagem = '" & GetSetting("QuickStore", "RelOrcamento", "Mensagem", "") & "'"
    
'    sMENSAGEM_LOG_TESTE_GERAL = sMENSAGEM_LOG_TESTE_GERAL & "STEP 1.3: " & Now & vbCrLf

    'Seta a impressora para relat�rio
    Call SetPrinterName("REL", Rel1)
    
'    sMENSAGEM_LOG_TESTE_GERAL = sMENSAGEM_LOG_TESTE_GERAL & "STEP 4: " & Now & vbCrLf
    .Action = 1
  End With
  
  Call StatusMsg("")
  MousePointer = vbDefault
  
  Rel1.Reset
  
'  sMENSAGEM_LOG_TESTE_GERAL = sMENSAGEM_LOG_TESTE_GERAL & "STEP 5: " & Now & vbCrLf
  
'  MsgBox sMENSAGEM_LOG_TESTE_GERAL, vbInformation, "LOG"

  Exit Sub
  
Erro:
  MsgBox "Erro ao imprimir: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub ChamaLimpaComandas()
On Error GoTo Erro

    'ws.BeginTrans
    'db.Execute "DELETE FROM SaidasComandas Where Filial = " & gnCodFilial
    ''db.Execute "DELETE FROM SaidasComandas "
    'ws.CommitTrans

    'MsgBox "Tabela de comandas zerada com sucesso", vbInformation, "Sucesso"
    
    Exit Sub
Erro:
    ws.Rollback
    MsgBox "Erro ao tentar zerar a tabela de Comandas " & Err.Number & " " & Err.Description, vbInformation, "Aten��o"
    
    'Shell App.Path & "\LimpaComandas.exe", vbHide 'vbNormalFocus
End Sub

'Eduardo - 17/10/2013
Private Sub Imprimir_Orcamento3(ByVal strRelatorio As String)
   
  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Encontre um registro antes."
    Exit Sub
  End If
  
  With frmRelSaidasMov
    .Filial = rsSaidas.Fields("Filial").Value
    .Sequencia = rsSaidas.Fields("Sequ�ncia").Value
    .Relatorio = strRelatorio
    .Show vbModal
  End With
 
End Sub

'18/07/2012 - mpdea
Private Sub Imprimir_Orcamento2(ByVal strRelatorio As String)
   
  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Encontre um registro antes."
    Exit Sub
  End If
  
  With frmRelSaidasMov
    .Filial = rsSaidas.Fields("Filial").Value
    .Sequencia = rsSaidas.Fields("Sequ�ncia").Value
    .Relatorio = strRelatorio
    .Show vbModal
  End With
 
End Sub

Public Sub Imprimir_Orcamento(strRelatorio As String)
  
  Dim Str1 As String
  
  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Encontre um registro antes."
    Exit Sub
  End If
  
  Call StatusMsg("Aguarde ...")
  DoEvents
  
 Rem  Nome do BD
   With Rel1
     .DataFiles(0) = gsQuickDBFileName
     .DataFiles(1) = gsQuickDBFileName
     .DataFiles(2) = gsQuickDBFileName
     .DataFiles(3) = gsQuickDBFileName
     .DataFiles(4) = gsQuickDBFileName
     .DataFiles(5) = gsQuickDBFileName
     .DataFiles(6) = gsQuickDBFileName
   End With

 
 '18/07/2012 - mpdea
 'Corrige sele��o
 ' Rel1.GroupSelectionFormula = "{Sa�das.Sequ�ncia} = " + txtSeq.Text
 Rel1.SelectionFormula = "{Sa�das.Filial} = " & rsSaidas.Fields("Filial").Value & " AND {Sa�das.Sequ�ncia} = " & rsSaidas.Fields("Sequ�ncia").Value

 Rel1.Destination = 0
 
 Rem Nome do arquivo .rpt
 Str1 = gsReportPath & strRelatorio
 
 Rel1.ReportFileName = Str1
 
 
 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relat�rio
  Call SetPrinterName("REL", Rel1)
  

 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault
End Sub

Private Sub DropDown2_DropDown()
  Data5.Refresh
End Sub

Public Sub CheckMovimentacao()
  If Erro_Data Then
    Erro_Data = False
    If frmErroMov.gbContinue Then
      If Not frmGerente.gbSenhaGerente Then
        Unload Me
      End If
    Else
      Unload Me
    End If
  End If
End Sub

'23/10/2002 - mpdea
'Retornado evento de KeyUp para KeyDown, devido ao form de recebimento estar
'enviando keycodes (Alt+C) para o form de Sa�das no evento KeyUp
'
'01/10/2002 - mpdea
'Inclu�do KeyCode = 0: Shift = 0 quando o KeyCode for atalho
'Alterado evento de KeyDown para KeyUp

'22/01/2003 - mpdea
'Verifica Quick em modo limitado
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim Tool As ActiveBarLibraryCtl.Tool

  Set Tool = New ActiveBarLibraryCtl.Tool
'
  If Shift = vbCtrlMask Then
    Select Case KeyCode
      Case vbKeyR
        Tool.Name = "miComplRecebimento"
        Call Screen.ActiveForm.ActiveBar1_Click(Tool)
        KeyCode = 0: Shift = 0
        Exit Sub
      Case vbKeyT
        If gblnQuickFull Then
          Tool.Name = "miComplPrintTicket"
          Call Screen.ActiveForm.ActiveBar1_Click(Tool)
        End If
        KeyCode = 0: Shift = 0
        Exit Sub
      Case vbKeyF
        If gblnQuickFull Then
          Tool.Name = "miComplPrintNotaFiscal"
          Call Screen.ActiveForm.ActiveBar1_Click(Tool)
        End If
        KeyCode = 0: Shift = 0
        Exit Sub
      Case vbKeyD
        Tool.Name = "miComplDesconto"
        '09/02/2007 - Anderson
        'Corre��o do BUG quando o usu�rio n�o pode dar desconto e pressiona as teclas CTRL+D
        If ActiveBar1.Tools("miComplDesconto").Enabled Then
          Call Screen.ActiveForm.ActiveBar1_Click(Tool)
        End If
        KeyCode = 0: Shift = 0
        Exit Sub
    End Select
  End If
  
  Call HandleKeyDown(KeyCode, Shift)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub cmd_excluirChave_Click()
On Error GoTo Erro

  If L_Efetivada.Visible = True Then
      MsgBox "Opera��o n�o pode mais ser alterada", vbInformation, "Aten��o"
      Exit Sub
  End If

  If gridChaves.RowSel > 0 Then
      gridChaves.RemoveItem gridChaves.RowSel
  Else
      MsgBox "Selecione uma chave na grade", vbInformation, "Aten��o"
  End If
    
  Exit Sub
Erro:
    MsgBox "Erro na exclus�o da Chave na grade " & Err.Number & " " & Err.Description, vbInformation, "Aten��o"

End Sub

Private Sub cmd_incluirChave_Click()
On Error GoTo Erro

    If L_Efetivada.Visible = True Then
        MsgBox "Opera��o n�o pode mais ser alterada", vbInformation, "Aten��o"
        Exit Sub
    End If

    If Len(Trim(txt_chave1.Text)) < 4 Or Len(Trim(txt_chave2.Text)) < 4 Or _
        Len(Trim(txt_chave3.Text)) < 4 Or Len(Trim(txt_chave4.Text)) < 4 Or _
        Len(Trim(txt_chave5.Text)) < 4 Or Len(Trim(txt_chave6.Text)) < 4 Or _
        Len(Trim(txt_chave7.Text)) < 4 Or Len(Trim(txt_chave8.Text)) < 4 Or _
        Len(Trim(txt_chave9.Text)) < 4 Or Len(Trim(txt_chave10.Text)) < 4 Or _
        Len(Trim(txt_chave11.Text)) < 4 Then
    
        MsgBox "Informe corretamente a Chave.", vbInformation, "Aten��o"
        Exit Sub
    End If
    
    gridChaves.AddItem vbTab & txt_chave1.Text & txt_chave2.Text & txt_chave3.Text & _
                              txt_chave4.Text & txt_chave5.Text & txt_chave6.Text & _
                              txt_chave7.Text & txt_chave8.Text & txt_chave9.Text & _
                              txt_chave10.Text & txt_chave11.Text
                              
    txt_chave1.Text = ""
    txt_chave2.Text = ""
    txt_chave3.Text = ""
    txt_chave4.Text = ""
    txt_chave5.Text = ""
    txt_chave6.Text = ""
    txt_chave7.Text = ""
    txt_chave8.Text = ""
    txt_chave9.Text = ""
    txt_chave10.Text = ""
    txt_chave11.Text = ""
    
    txt_chave1.SetFocus

    Exit Sub
Erro:
    MsgBox "Erro na inclus�o da Chave na grade " & Err.Number & " " & Err.Description, vbInformation, "Aten��o"

End Sub

'09/03/2023 - Pablo
Private Sub lblChave_DblClick(Index As Integer)
  Dim tmp_chave As String
  tmp_chave = InputBox("CHAVE", "CHAVE")
  If Len(tmp_chave) = 44 And IsNumeric(tmp_chave) Then
    txt_chave1.Text = Mid(tmp_chave, 0 * 4 + 1, 4)
    txt_chave2.Text = Mid(tmp_chave, 1 * 4 + 1, 4)
    txt_chave3.Text = Mid(tmp_chave, 2 * 4 + 1, 4)
    txt_chave4.Text = Mid(tmp_chave, 3 * 4 + 1, 4)
    txt_chave5.Text = Mid(tmp_chave, 4 * 4 + 1, 4)
    txt_chave6.Text = Mid(tmp_chave, 5 * 4 + 1, 4)
    txt_chave7.Text = Mid(tmp_chave, 6 * 4 + 1, 4)
    txt_chave8.Text = Mid(tmp_chave, 7 * 4 + 1, 4)
    txt_chave9.Text = Mid(tmp_chave, 8 * 4 + 1, 4)
    txt_chave10.Text = Mid(tmp_chave, 9 * 4 + 1, 4)
    txt_chave11.Text = Mid(tmp_chave, 10 * 4 + 1, 4)
    Call cmd_incluirChave_Click
  End If
End Sub

Private Sub txt_chave1_Change()
    If Len(txt_chave1) = 4 Then
      txt_chave2.SetFocus
    End If
End Sub

Private Sub txt_chave1_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If
    
      If Len(txt_chave1) = 3 Or Len(txt_chave1) = 4 Then
        txt_chave1.Text = txt_chave1.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave10_Change()
    If Len(txt_chave10) = 4 Then
      txt_chave11.SetFocus
    ElseIf Len(txt_chave10) = 0 Then
      txt_chave9.SetFocus
    End If
End Sub

Private Sub txt_chave10_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave10) = 3 Or Len(txt_chave10) = 4 Then
        txt_chave10.Text = txt_chave10.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave11_Change()
    If Len(txt_chave11) = 0 Then
      txt_chave10.SetFocus
    End If
End Sub

Private Sub txt_chave11_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave11) = 3 Or Len(txt_chave11) = 4 Then
        txt_chave11.Text = txt_chave11.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave2_Change()
    If Len(txt_chave2) = 4 Then
      txt_chave3.SetFocus
    ElseIf Len(txt_chave2) = 0 Then
      txt_chave1.SetFocus
    End If
End Sub

Private Sub txt_chave2_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave2) = 3 Or Len(txt_chave2) = 4 Then
        txt_chave2.Text = txt_chave2.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave3_Change()
    If Len(txt_chave3) = 4 Then
      txt_chave4.SetFocus
    ElseIf Len(txt_chave3) = 0 Then
      txt_chave2.SetFocus
    End If
End Sub

Private Sub txt_chave3_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave3) = 3 Or Len(txt_chave3) = 4 Then
        txt_chave3.Text = txt_chave3.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave4_Change()
    If Len(txt_chave4) = 4 Then
      txt_chave5.SetFocus
    ElseIf Len(txt_chave4) = 0 Then
      txt_chave3.SetFocus
    End If
End Sub

Private Sub txt_chave4_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave4) = 3 Or Len(txt_chave4) = 4 Then
        txt_chave4.Text = txt_chave4.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave5_Change()
    If Len(txt_chave5) = 4 Then
      txt_chave6.SetFocus
    ElseIf Len(txt_chave5) = 0 Then
      txt_chave4.SetFocus
    End If
End Sub

Private Sub txt_chave5_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave5) = 3 Or Len(txt_chave5) = 4 Then
        txt_chave5.Text = txt_chave5.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave6_Change()
    If Len(txt_chave6) = 4 Then
      txt_chave7.SetFocus
    ElseIf Len(txt_chave6) = 0 Then
      txt_chave5.SetFocus
    End If
End Sub

Private Sub txt_chave6_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave6) = 3 Or Len(txt_chave6) = 4 Then
        txt_chave6.Text = txt_chave6.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave7_Change()
    If Len(txt_chave7) = 4 Then
      txt_chave8.SetFocus
    ElseIf Len(txt_chave7) = 0 Then
      txt_chave6.SetFocus
    End If
End Sub

Private Sub txt_chave7_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave7) = 3 Or Len(txt_chave7) = 4 Then
        txt_chave7.Text = txt_chave7.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave8_Change()
    If Len(txt_chave8) = 4 Then
      txt_chave9.SetFocus
    ElseIf Len(txt_chave8) = 0 Then
      txt_chave7.SetFocus
    End If
End Sub

Private Sub txt_chave8_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave8) = 3 Or Len(txt_chave8) = 4 Then
        txt_chave8.Text = txt_chave8.Text
        SendKeys "{End}", True
      End If
  End If
End Sub

Private Sub txt_chave9_Change()
    If Len(txt_chave9) = 4 Then
      txt_chave10.SetFocus
    ElseIf Len(txt_chave9) = 0 Then
      txt_chave8.SetFocus
    End If

End Sub

Private Sub txt_chave9_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
      Dim strValid As String
      strValid = "0123456789"
    
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If

      If Len(txt_chave9) = 3 Or Len(txt_chave9) = 4 Then
        txt_chave9.Text = txt_chave9.Text
        SendKeys "{End}", True
      End If
  End If
End Sub


'04/03/2004 - mpdea
'Implementado tratamento de erro
Private Sub Form_Load()
  Dim Resp As Integer
  Dim Aux As String
  Dim strRet As String
  Dim strSQL As String
  Dim rstCheckSaidas As Recordset
  
  On Error GoTo ErrHandler
  
  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  btnComandaVendas.Visible = False
  txtComanda.Width = txtSeq.Width   '  24,007
  
  l_tamanhoOriginal_TAB1 = Tab1.Width
  l_tamanhoOriginal_GRADE1 = Grade1.Width
  l_tamanhoOriginal_GRADE1_Grupo1Produtos = Grade1.Groups(0).Width
  l_tamanhoOriginal_Grade_Serv = Grade_Serv.Width
  l_tamanhoOriginal_Grade_Serv_GrupoServicos = Grade_Serv.Groups(0).Width
  l_txtSeq = txtSeq.Left
  l_Label52 = Label52.Left
  l_txtComanda = txtComanda.Left
  l_lblComanda = lblComanda.Left
  l_Senha = Senha.Left
  l_Label26 = Label26.Left
  l_Nome_Caixa = Nome_Caixa.Left
  l_Combo_Caixa = Combo_Caixa.Left
  l_Label22 = Label22.Left
  l_mskValidade = mskValidade.Left
  l_lblValidade = lblValidade.Left
  l_cmd_tabelaDePrecos = cmd_tabelaDePrecos.Left
  l_txtSubTotal = txtSubTotal.Left
  l_Label35 = Label35.Left
  l_txtDescSubTotal = txtDescSubTotal.Left
  l_Label36 = Label36.Left
  l_Label48 = Label48.Left
  l_L_Tot_Pagar = L_Tot_Pagar.Left
  l_B_Servi�os_Conc = B_Servi�os_Conc.Left
  
  l_Nome_Cliente_Estica = Nome_Cliente.Width
  l_Nome_Digitador_Estica = Nome_Digitador.Width
  l_txtRef_Estica = txtRef.Width
  l_cboPresencaComprador_Estica = cboPresencaComprador.Width
  
  'Combo_Pre�o.BackColor = F7ED03
  
  '16/10/2009 - mpdea
  'Modo de entrada de dados no grid de produtos
  strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "SAIDAS", "ModoGradeProdutos")
  Call IsDataType(dtInteger, strRet, m_int_modo_grid_produtos)
  
  
  KeyPreview = True
  
  '18/07/2012 - mpdea
  'Liberado para todos os usu�rios e personalizado
  '05/03/2007 - Anderson
  'Libera as customiza��es de impress�o de or�amentos.
  'QS38785-386 - NewTech Inform�tica
  'bolImprimirOrcamento = CheckSerialCaseMod("QS38785-386")
  ActiveBar1.Tools("miImprimirOrcamento").Visible = True 'bolImprimirOrcamento
  ActiveBar1.Tools("miImprimirOrcamentoVenda").Visible = True ' bolImprimirOrcamento
  ActiveBar1.Tools("miImprimirOrcamentoVendaSemFrete").Visible = CheckSerialCaseMod("QS73520-469") 'Eduardo - 17/10/2013 - QS73520-469 Disk Embalagens
  
  '22/11/2013 - Eduardo
  ActiveBar1.Tools("LimpaComandas").Visible = True
  
  '22/06/2007 - Anderson
  'Exportar dados para excel. Customiza��o Candy-Clean
  ActiveBar1.Tools("miOpExportarExcel").Visible = CheckSerialCaseMod("QS37957-281", "QS73206-768")
  
  '27/09/2007 - Anderson
  'Implementa��o da impress�o de carn� com c�digo de barras
  'Solicitado por: Naativa
  ActiveBar1.Tools("miOpImprimirCarne").Visible = g_bolCarneCodigoBarras
  
  With ActiveBar1.Tools("miOpOrdem").CBList
    .Clear
    .AddItem "Por Seq��ncia"
    .AddItem "Por Data e Seq��ncia"
    .AddItem "Por Cliente e Seq��ncia"
    .AddItem "Por Refer�ncia Interna"
    .AddItem "Por Nota Fiscal"
  End With
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  ActiveBar1.RecalcLayout
  ActiveBar1.Refresh

  Desconto_Cli = 0

  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  Data6.DatabaseName = gsQuickDBFileName
  Data7.DatabaseName = gsQuickDBFileName
  Data8.DatabaseName = gsQuickDBFileName
  Data9.DatabaseName = gsQuickDBFileName
  
  Set rsProdutos2 = rsProdutos.Clone
  
  'Set rsServi�os = db.OpenRecordset("Servi�os", , dbReadOnly)
  'Set rsOp_Sa�da = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
  'Set rsFuncionarios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  'Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  'Set rsGrade = db.OpenRecordset("C�digos da Grade", , dbReadOnly)
  'Set rsMovi_Parcelas = db.OpenRecordset("Movimento - Parcelas")
  'Set rsParametros = db.OpenRecordset("Par�metros Filial")
  'Set rsPre�os = db.OpenRecordset("Pre�os")
  
  Set rsServi�os = db.OpenRecordset("Servi�os", , dbReadOnly)
  Set rsOp_Sa�da = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("C�digos da Grade", , dbReadOnly)
  Set rsMovi_Parcelas = db.OpenRecordset("Movimento - Parcelas")
  Set rsParametros = db.OpenRecordset("Par�metros Filial")
  Set rsPre�os = db.OpenRecordset("Pre�os")
  
  If gblnSimplesNacional = False Then
      'Set rsEstadosICMS = db.OpenRecordset("ICMS_Percentual_Estados", , dbReadOnly)
      Set rsEstadosICMS = db.OpenRecordset("ICMS_Percentual_Estados", , dbReadOnly)
  End If
  
  gsSql = "SELECT * FROM Sa�das WHERE Filial = " & gnCodFilial
  gsWhere = ""
  gsOrder = " ORDER BY Sequ�ncia"
  
  ' Pilatti Novembro 2017
  Dim sAcessoCustoProdutos As Variant
  sAcessoCustoProdutos = rsFuncionarios("Custo Produtos").Value
  
  If sAcessoCustoProdutos = False Then
    gsWhere = " and UCASE(Tabela) Not Like '*CUSTO*' and UCASE(Tabela) Not Like '*ORIGEM*' "
  End If
  ' Pilatti fim
  
  Set rsSaidas = db.OpenRecordset(gsSql & gsWhere & gsOrder, dbOpenDynaset)
  gsWhere = ""
  
  'Set rsSaidas_Prod = db.OpenRecordset("Sa�das - Produtos")
  'Set rsSaidas_Serv = db.OpenRecordset("Sa�das - Servi�os")
  'Set rsSa�da_Cheques = db.OpenRecordset("Movimento - Cheques")
  'Set rsSa�da_Parcelas = db.OpenRecordset("Movimento - Parcelas")
  'Set rsUsu�rios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  'Set rsTabelas = db.OpenRecordset("Tabela de Pre�os", , dbReadOnly)
  'Set rsCota��es = db.OpenRecordset("Cota��es", , dbReadOnly)
  'Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
  'Set rsEstados = db.OpenRecordset("Estados", , dbReadOnly)
  'Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  'Set rsT�cnicos = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  'Set rsOperadores = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  'Set rsLog = db.OpenRecordset("ZZZLog")
  
  Set rsSaidas_Prod = db.OpenRecordset("Sa�das - Produtos")
  Set rsSaidas_Serv = db.OpenRecordset("Sa�das - Servi�os")
  Set rsSa�da_Cheques = db.OpenRecordset("Movimento - Cheques")
  Set rsSa�da_Parcelas = db.OpenRecordset("Movimento - Parcelas")
  Set rsUsu�rios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsTabelas = db.OpenRecordset("Tabela de Pre�os", , dbReadOnly)
  Set rsCota��es = db.OpenRecordset("Cota��es", , dbReadOnly)
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
  Set rsEstados = db.OpenRecordset("Estados", , dbReadOnly)
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  Set rsT�cnicos = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsOperadores = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsLog = db.OpenRecordset("ZZZLog")
  
  '10/12/2009 - Andrea
  'Set rsSa�da_Cartoes = db.OpenRecordset("Movimento - Cartoes")
  Set rsSa�da_Cartoes = db.OpenRecordset("Movimento - Cartoes")
  
  '20/12/2006 - Anderson - Registro de CFOP por produto e servi�o
  'Set rsProdutoCFOP = db.OpenRecordset("ProdutoCFOP", , dbReadOnly)
  'Set rsServicoCFOP = db.OpenRecordset("ServicoCFOP", , dbReadOnly)
  
  Set rsProdutoCFOP = db.OpenRecordset("ProdutoCFOP", , dbReadOnly)
  Set rsServicoCFOP = db.OpenRecordset("ServicoCFOP", , dbReadOnly)
  
  '17/10/2007 - Anderson
  'Implementa��o de tipo de pesquisa na combo de produtos
  strRet = GetSetting("QuickStore", "ConfigSAIDAS", "PesquisarDescricao", "")
  If strRet = "" Then strRet = True
  ActiveBar1.Tools("miOpPesquisarDescricao").Checked = CBool(strRet)
  
  '17/10/2007 - Anderson
  'Implementa��o de tipo de pesquisa na combo de produtos
  strRet = GetSetting("QuickStore", "ConfigSAIDAS", "PesquisarCodigo", False)
  ActiveBar1.Tools("miOpPesquisarCodigo").Checked = CBool(strRet)
  
  strRet = GetSetting("QuickStore", "ConfigSAIDAS", "Scanner", False)
  ActiveBar1.Tools("miComplLeitorOtico").Checked = CBool(strRet)
 
  strRet = GetSetting("QuickStore", "ConfigSAIDAS", "Mantem Operacao", False)
  ActiveBar1.Tools("miOpFreezeOperacao").Checked = CBool(strRet)
 
  strRet = GetSetting("QuickStore", "ConfigSAIDAS", "Mantem Digitador", False)
  ActiveBar1.Tools("miOpFreezeDigitador").Checked = CBool(strRet)
 
  strRet = GetSetting("QuickStore", "ConfigSAIDAS", "Mantem Cliente", False)
  ActiveBar1.Tools("miOpFreezeCliente").Checked = CBool(strRet)
 
  strRet = GetSetting("QuickStore", "ConfigSAIDAS", "Mantem TabPrecos", False)
  ActiveBar1.Tools("miOpFreezeTabPrecos").Checked = CBool(strRet)
 
  strRet = GetSetting("QuickStore", "ConfigSAIDAS", "Etiqueta Balanca", False)
  ActiveBar1.Tools("miOpEtiquetas").Checked = CBool(strRet)
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then
    MsgBox "Filial n�o encontrada", vbCritical, "Erro"
    Exit Sub
  End If
  Nome_Filial.Caption = rsParametros("Nome")
  
  '13-04-2025 pablo
  If rsParametros("EditarNomeProduto").Value Then Grade1.Columns("Nome").Locked = False
  
  'Tratamento NrSerie pelo arquivo .txt
  If gNrSerieNF > 0 Then
      Dim xConta As Integer
      Dim sCGC As String
      sCGC = rsParametros("CGC")
      sCGC = Replace(sCGC, " ", "")
      sCGC = Replace(sCGC, "/", "")
      sCGC = Replace(sCGC, "\", "")
      sCGC = Replace(sCGC, ".", "")
      sCGC = Replace(sCGC, ",", "")
      sCGC = Replace(sCGC, "-", "")
      For xConta = 0 To gNrSerieNF - 1
          If gArrayNrSerieNF(xConta, 0) = sCGC Then
              'NrCnpj1 , SerieNFe1, SerieNFCe1
              txtNrSerieNF.Text = gArrayNrSerieNF(xConta, 1)
              Exit For
          End If
      Next
  End If
  'Fim tratamento NrSerie
  
  
  '06/05/2003 - mpdea
  'Desconto rateado
  m_blnDescontoRateado = rsParametros.Fields("DescSubTotalRateado").Value
  
  
  '07/05/2003 - mpdea
  'Objetos para Desconto rateado
  lblSubTotal.Visible = Not m_blnDescontoRateado
  txtSubTotal.Visible = Not m_blnDescontoRateado
  lblDescSubTotal.Visible = Not m_blnDescontoRateado
  txtDescSubTotal.Visible = Not m_blnDescontoRateado
  
  
  '---------------------------------------------------------------------------------
  '28/10/2002 - mpdea
  'Modificado a atribui��o do objeto de 'Set objeto = ...' para a propriedade
  'RecordSource, corrigindo o erro de navega��o dos registro com as teclas
  'para cima e para baixo

  'strSQL = "SELECT Nome, C�digo, Tipo, Cidade, Estado From Cli_For Where Inativo = False AND Tipo = 'C' ORDER BY Nome"
  strSQL = "SELECT Nome, C�digo, Tipo, Cidade, Estado From Cli_For Where Inativo = False ORDER BY Nome"
  Data4.RecordSource = strSQL
  Data4.Refresh
'  Set Data4.Recordset = db.OpenRecordset(strSQL, dbOpenDynaset)
  '---------------------------------------------------------------------------------
  
  strSQL = "SELECT Descri��o, C�digo, Pre�o From Servi�os Where ((C�digo) <> 0) ORDER BY Descri��o"
  Set Data5.Recordset = db.OpenRecordset(strSQL, dbOpenDynaset)
  Data5.Refresh
   
  Set Data1.Recordset = db.OpenRecordset("SELECT Nome, C�digo FROM Produtos WHERE C�digo <> '0' AND Desativado = False ORDER BY Nome", dbOpenDynaset)
  Data1.Refresh
     
  '---------------------------------------------------------------------------------
  '24/10/2002 - mpdea
  'Modificado a atribui��o do objeto de 'Set objeto = ...' para a propriedade
  'RecordSource, corrigindo o erro de navega��o dos registro com as teclas
  'para cima e para baixo
  '
  '07/05/2002 - mpdea
  '
  'Atualizado SQL para exibi��o das tabelas de pre�os
  '>>-------------------------------------------------------------------------------
'  Set Data9.Recordset = db.OpenRecordset(SQL_CONS_TAB_PRECO_SHOW, dbOpenDynaset)
  
  
  ' Pilatti Novembro de 2017
  'Dim sAcessoCustoProdutos As Variant
  'sAcessoCustoProdutos = rsFuncionarios("Custo Produtos").Value
  
  Dim sSql As String
  
  If sAcessoCustoProdutos = True Then
'''    Data9.RecordSource = SQL_CONS_TAB_PRECO_SHOW

      sSql = "SELECT DISTINCT [Tabela de Pre�os].Tabela "
      sSql = sSql & " FROM AcessoTabelasDePrecosProdutos, [Tabela de Pre�os] "
      sSql = sSql & " INNER JOIN Pre�os ON [Tabela de Pre�os].Tabela = Pre�os.Tabela WHERE "
      sSql = sSql & " [Tabela de Pre�os].Tabela <> 'CUSTO' "
      sSql = sSql & " AND AcessoTabelasDePrecosProdutos.Tabela = [Tabela de Pre�os].Tabela "
      sSql = sSql & " AND AcessoTabelasDePrecosProdutos.Usuario = " & gnUserCode
      sSql = sSql & " ORDER BY [Tabela de Pre�os].Tabela "
  
      Data9.RecordSource = sSql
  Else
    ' Inclui o tratamento para n�o buscar as TABELAS de PRE�OS que tenham o nome CUSTO como parte do nome
    Dim sSqlTabelaPreco As String
'''    sSqlTabelaPreco = "SELECT DISTINCT [Tabela de Pre�os].Tabela FROM [Tabela de Pre�os] " & _
'''    "INNER JOIN Pre�os ON [Tabela de Pre�os].Tabela = Pre�os.Tabela WHERE " & _
'''    "UCASE([Tabela de Pre�os].Tabela) Not Like '*CUSTO*' ORDER BY [Tabela de Pre�os].Tabela"
  
    sSqlTabelaPreco = "SELECT DISTINCT [Tabela de Pre�os].Tabela FROM AcessoTabelasDePrecosProdutos, [Tabela de Pre�os] "
    sSqlTabelaPreco = sSqlTabelaPreco & " INNER JOIN Pre�os ON [Tabela de Pre�os].Tabela = Pre�os.Tabela WHERE "
    sSqlTabelaPreco = sSqlTabelaPreco & " UCASE([Tabela de Pre�os].Tabela) Not Like '*CUSTO*' "
    sSqlTabelaPreco = sSqlTabelaPreco & " AND AcessoTabelasDePrecosProdutos.Tabela = [Tabela de Pre�os].Tabela "
    sSqlTabelaPreco = sSqlTabelaPreco & " AND AcessoTabelasDePrecosProdutos.Usuario = " & gnUserCode
    sSqlTabelaPreco = sSqlTabelaPreco & " ORDER BY [Tabela de Pre�os].Tabela"
  
    Data9.RecordSource = sSqlTabelaPreco
  End If
' Pilatti fim
  
  '-------------------------------------------------------------------------------<<
   
  ' =========================================================
  ' Grade Chaves
  gridChaves.ColWidth(0) = 0
  gridChaves.ColWidth(1) = 4700
  
  gridChaves.Row = 0
  gridChaves.TextMatrix(0, 1) = ""
  ' =========================================================
  
   
'  Grade1.StyleSets("Total").Font.Size = 12
'  Grade1.StyleSets("Total").Font.Bold = True
'  Grade1.StyleSets("Normal").Font.Size = 10
'  Grade1.StyleSets("Normal").Font.Bold = True
  
'  Grade1.RowHeight = 345.2599
  
  
  '17/09/2003 - mpdea
  'Valida��o para o estado de SC
  'verifica se pode alterar ou n�o o pre�o do produto
'  If rsParametros("Saida Altera Preco") Then
'     Grade1.Columns(4).Locked = False
'  Else
'     Grade1.Columns(4).Locked = True
'  End If
  'If UCase(gstrGetEstadoFilial(gnCodFilial)) = "SC" Then
    'Grade1.Columns(4).Locked = True
  'Else
    Grade1.Columns(4).Locked = Not rsParametros.Fields("Saida Altera Preco").Value
  'End If
  
  '09-07-2015 - Jean Ricardo Zanella
  'Valida��o sobre usuario poder alterar pre�os
  If blnPermissaoAlterarPrecos(gnUserCode) = False Then
    Grade1.Columns(4).Locked = True
  End If
  
  Grade1.Rows = rsParametros("Linhas Digita��o")
  Linhas_Grade = rsParametros("Linhas Digita��o")
   
  Grade_Serv.Rows = rsParametros("Linhas Servi�o")
  Linhas_Servi�o = rsParametros("Linhas Servi�o")
  
  
  '23/09/2002 - mpdea
  'Alterado o tratamento para a utiliza��o de Servi�os
  '(objetos vis�vel ou n�o)
  Tab1.TabVisible(1) = gbServico
'''  L_Tot_Serv.Visible = gbServico
'''  L_Tot_ISS.Visible = gbServico
'''  lblTotServ.Visible = gbServico
'''  lblTotISS.Visible = gbServico
  
'  If rsParametros("Usar Servi�os") = False Then
'    Tab1.TabEnabled(1) = False
'  End If
  Alterar_Servi�os = rsParametros("Alterar Servi�os")
  
  If rsParametros("Saida Descr Adicional") = True Then
     Grade1.Columns("Descri��o Adicional").Visible = True
  Else
     Grade1.Columns("Descri��o Adicional").Visible = False
  End If
  
  '19/12/2006 - Anderson
  'Verifica se a coluna CFOP deve ser exibida na grade
  If rsParametros("ExibeCFOP") = True Then
     Grade1.Columns("CFOP").Visible = True
     Grade_Serv.Columns("CFOP").Visible = True
  Else
     Grade1.Columns("CFOP").Visible = False
     Grade_Serv.Columns("CFOP").Visible = False
  End If
  
  If gbCaixas = False Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
    Combo_Caixa.Enabled = False
  End If
  
''  Lista_Aberta = False
   
  rsUsu�rios.Index = "C�digo"
  rsUsu�rios.Seek "=", gnUserCode
  If rsUsu�rios.NoMatch Then Exit Sub
   
  L_Dia.Caption = Format$(Data_Atual, "dd/mm/yyyy")
   
  Call ActiveBarLoadToolTips(Me)
   
  Grade_Serv.StyleSets("Total").Font.Size = 12
  Grade_Serv.StyleSets("Total").Font.Bold = True
  Grade_Serv.StyleSets("Normal").Font.Size = 10
  Grade_Serv.StyleSets("Normal").Font.Bold = True
  
  '08/08/2002 - mpdea
  'Obt�m o nr. do terminal do registro
  txtNrTerminal.Text = GetSetting("QuickStore", "ConfigSAIDAS", "NrTerminal", "")
  
  
  '22/01/2003 - mpdea
  'Quick em modo limitado
  If Not gblnQuickFull Then
    With ActiveBar1
      With .Bands("tbrComplem")
        .Tools("miComplPrintTicket").Visible = False
        .Tools("miComplPrintNotaFiscal").Visible = False
        .Tools("miComplCancelNota").Visible = False
        .Tools("miComplRetNFCe").Visible = False
        .Tools("miComplNFC-e").Visible = False
        .Tools("miComplNFe").Visible = False
      End With
      .RecalcLayout
      .Refresh
    End With
  End If
  
  ' Pilatti Outubro/17
  ActiveBar1.Tools("miComplCancelNota").Visible = False
  ActiveBar1.Tools("miComplPrintNotaFiscal").Visible = False
  ActiveBar1.Tools("miComplNFe").Visible = False

  
  
  '04/03/2004 - mpdea
  'Otimizado verifica��o
  gbLogError = False
  strSQL = "SELECT Data FROM Sa�das WHERE Data > #" & Format(Data_Atual, "MM/dd/yyyy") & "#"
  Set rstCheckSaidas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  If rstCheckSaidas.RecordCount > 0 Then
    Erro_Data = True
    gbLogError = True
  End If
  rstCheckSaidas.Close
  Set rstCheckSaidas = Nothing
'  rsSaidas.FindLast "Data > #" & Format(Data_Atual, "mm/dd/yyyy") & "#"
'  If Not rsSaidas.NoMatch Then
'    Erro_Data = True
'    gbLogError = True
'  End If

  
  '19/02/2004 - Daniel
  'Case.......: PSV Inform�tica
  'Finalidade.: Comp�r ou n�o o field Data Validade em Sa�das
  lblValidade.Visible = False
  mskValidade.Visible = False 'Ser� habilitado somente se a opera��o sa�da tiver o campo validade como true
  
  m_blnPSV = CheckSerialCaseMod("QS35552-811", "QS37705-639", "QS37825-830", "QS38933-772", "QS39369-521")
  '---------------------------------------------------------------------------------------------------------
  
  '17/06/2004 - Daniel
  'Valida��o para o cliente Nilvo Burin
  'Estava aparecendo o Total da Sa�da zerado ao excluir o produto do cadastro
  m_blnNilvo = CheckSerialCaseMod("QS33398-647")
  
  '20/10/2004 - Daniel
  'Case.......: A.S. Wijman
  'Finalidade.: Tratamento para o campo [Sa�das - Produtos].[Pre�o Final]
  m_blnASWijmaBelem = CheckSerialCaseMod("QS39881-068", "QS40377-377")
  
  '09/11/2004 - Daniel
  'Case: Cliente Teknika
  m_blnTeknika = CheckSerialCaseMod("QS40966-243")
  
  '26/07/2005 - Daniel
  'Personaliza��o para a empresa J.R. Hidroqu�mica
  'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
  'C.S.T. (C�digo de Situa��o Tribut�ria)
  m_blnJR = CheckSerialCaseMod("QS31135-807")
  '
  If m_blnJR Then
    With Grade_Serv
      '20/12/2006 - Anderson - Altera��o para o registro de CFOP por servi�o
      '.Columns("Descri��o").Width = 4980
      .Columns("Descri��o").Width = 3660.189
      .Columns("CST").Visible = True
      .Columns("CST").Width = 750.0473
      .Columns("CST").Locked = False
    End With
  Else
    With Grade_Serv
      '20/12/2006 - Anderson - Altera��o para o registro de CFOP por servi�o
      '.Columns("Descri��o").Width = 5699.906
      .Columns("Descri��o").Width = 4380.095
      .Columns("CST").Visible = False
    End With
  End If
  '---------------------------------------------------------------------------------------------------------
  
  '---------------------------------------------------------------------------------------------------------
  '04/05/2004 - Daniel
  'Case.......: Embalavi
  'Finalidade.: Monitorar personaliza��es para a Embalavi
  '
  '01/09/2005 - Daniel
  'Foi aberto o tratamento com 5 casas decimais para a empresa
  'Actel Ltda [QS36688-609, QS36664-089, QS38091-093, QS38186-428]
  m_blnEmbalavi = CheckSerialCaseMod("QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162", "QS36688-609", "QS36664-089", "QS38091-093", "QS38186-428")
  
  With Grade1
    If g_bln5CasasDecimais Then
      .Columns("Pre�o Unit�rio").NumberFormat = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Columns("Pre�o Unit�rio").NumberFormat = "##,###,##0.000"
    Else
        .Columns("Pre�o Unit�rio").NumberFormat = "##,###,##0.00"
    End If
  End With
  
  With DropDown1
    If g_bln5CasasDecimais Then
      .Columns("Pre�o").NumberFormat = "##,###,##0.00000"
    '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Columns("Pre�o").NumberFormat = "##,###,##0.000"
    Else
      .Columns("Pre�o").NumberFormat = "##,###,##0.00"
    End If
  End With
  '---------------------------------------------------------------------------------------------------------
    
  '26/08/2004 - Daniel
  'Criado valida��o para verificar se o usu�rio possui permiss�o
  'para enchergar o estoque ou n�o
  Call EnchergarEstoque
  
  '06/05/2005 - Daniel
  '
  'Implementa��o.: Trabalhar com o c�digo para fornecedor cadastrado na tela de produtos.
  '                Impacto: Ao entrar com o c�digo para o fornecedor no campo c�digo do produto
  '                o sistema dever� trazer o c�digo do produto que estiver amarrado nele
  'Solicita��o...: Cristiano Pavinato - PSI RS
  m_blnUsaCodFornec = g_blnVerificarUsoCodFornece
  '-------------------------------------------------------------
  
  '12/05/2005 - Daniel
  '
  'Solicitante..: Info Social
  '
  'Finalidade...: Deixamos configur�vel em Par�metros � exibi��o
  '               nas telas de Sa�da e Venda R�pida da coluna Fabricante
  '               nos dropdowns de pesquisas
  If rsParametros("ExibirFabricante").Value Then
    m_blnExibirColunaFabricante = True
    DropDown1.Columns("Fabricante").Visible = True
    DropDown1.Columns("Nome").Width = 4665.26
  Else
    m_blnExibirColunaFabricante = False
    DropDown1.Columns("Fabricante").Visible = False
    DropDown1.Columns("Nome").Width = 6420.26
  End If
  '----------------------------------------------------------------------
  '17/05/2013-Alexandre Afornali
  'Mostra ou n�o a o campo de Comandas de acordo com os parametros
  If (rsParametros("TrabalharComComanda").Value = 0) Then
    lblComanda.Visible = False
    txtComanda.Visible = False
  End If
  '----------------------------------------------------------------------
  
  '19/10/2007 - Anderson
  'Implementa��o do campo Lucro M�nimo Permitido conforme solicita��o da Agrotama
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", gnUserCode
  
  If Not rsFuncionarios.NoMatch Then
    m_bolLucroMinimoPermitido = rsFuncionarios("LucroMinimoPermitido")
  Else
    m_bolLucroMinimoPermitido = False
  End If
  
  'Verifica permiss�o para achar venda
  ActiveBar1.Tools("miComplPesquisaVendasHj").Visible = rsFuncionarios.Fields("PermiteAcharVenda").Value Or rsFuncionarios.Fields("Superusu�rio").Value
  
  Senha.Text = gSenhaUsuarioLogado
  
  'Teste
  cmdInsertItens.Visible = gbTeste
  
  ' Pilatti APP QUICK
  Dim iContaRegApp As Integer
  iContaRegApp = 0
  sSql = "SELECT C�digo FROM [Opera��es Sa�da] WHERE EmitirNFManualmente = -1 or InformanteProprio = -1 "
  Set rsVerificaOperacaoERP_APP = db.OpenRecordset(sSql, dbOpenSnapshot)
  While Not rsVerificaOperacaoERP_APP.EOF
    If iContaRegApp = 0 Then
      sOPERACAO_APPQuick01 = rsVerificaOperacaoERP_APP.Fields(0).Value
    Else
      sOPERACAO_APPQuick02 = rsVerificaOperacaoERP_APP.Fields(0).Value
    End If
    iContaRegApp = iContaRegApp + 1
    rsVerificaOperacaoERP_APP.MoveNext
  Wend
  rsVerificaOperacaoERP_APP.Close
  Set rsVerificaOperacaoERP_APP = Nothing
  
  Me.Show
  DoEvents

  Call ClearScreen
  
  '''Combo_Pre�o.Text = rsParametros("VR Tab Pre�o")
  If rsParametros("VR Tab Pre�o") <> "" Then
    Dim rsAcessosTabPrecoUsu As Recordset
    Dim iTemTabelasPreco As Integer
  
    iTemTabelasPreco = 0
  
    sSql = "Select Tabela From AcessoTabelasDePrecosProdutos Where Usuario=" & gnUserCode & " And Tabela='" & rsParametros("VR Tab Pre�o") & "' "
  
    Set rsAcessosTabPrecoUsu = db.OpenRecordset(sSql, dbOpenDynaset)
  
    If Not (rsAcessosTabPrecoUsu.EOF And rsAcessosTabPrecoUsu.BOF) Then
        iTemTabelasPreco = 1
        Combo_Pre�o.Text = rsParametros("VR Tab Pre�o") & ""
    Else
        iTemTabelasPreco = 0
    End If
    rsAcessosTabPrecoUsu.Close
    Set rsAcessosTabPrecoUsu = Nothing
  End If
  
  ActiveBar1.Tools("mnImprimeCarneTipo1").Visible = True
  
  '
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", Val(gnUserCode)
 
  Nome_Operador.Caption = rsFuncionarios("Nome")
  ActiveBar1.Tools("miComplDesconto").Enabled = rsFuncionarios("bPermiteDesconto")
  Combo_Operador.Text = gnUserCode
  Senha.Text = gSenhaUsuarioLogado
  '
  
  'Grade1.MoveFirst
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
 
ErrHandler:
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Resize()
  ' 0 - Normal
  ' 1 - Minimizado
  ' 2 - Maximizado
  If Me.WindowState = 0 Then
    Tab1.Width = l_tamanhoOriginal_TAB1
    Grade1.Width = l_tamanhoOriginal_GRADE1
    Grade1.Groups(0).Width = l_tamanhoOriginal_GRADE1_Grupo1Produtos
    Grade_Serv.Groups(0).Width = l_tamanhoOriginal_Grade_Serv_GrupoServicos
    Grade_Serv.Width = l_tamanhoOriginal_Grade_Serv

    txtSeq.Left = l_txtSeq
    Label52.Left = l_Label52
    txtComanda.Left = l_txtComanda
    lblComanda.Left = l_lblComanda
    Senha.Left = l_Senha
    Label26.Left = l_Label26
    Nome_Caixa.Left = l_Nome_Caixa
    Combo_Caixa.Left = l_Combo_Caixa
    Label22.Left = l_Label22
    mskValidade.Left = l_mskValidade
    lblValidade.Left = l_lblValidade
    cmd_tabelaDePrecos.Left = l_cmd_tabelaDePrecos
    txtSubTotal.Left = l_txtSubTotal
    Label35.Left = l_Label35
    txtDescSubTotal.Left = l_txtDescSubTotal
    Label36.Left = l_Label36
    Label48.Left = l_Label48
    L_Tot_Pagar.Left = l_L_Tot_Pagar
    B_Servi�os_Conc.Left = l_B_Servi�os_Conc
    
    Nome_Cliente.Width = l_Nome_Cliente_Estica
    Nome_Digitador.Width = l_Nome_Digitador_Estica
    txtRef.Width = l_txtRef_Estica
    cboPresencaComprador.Width = l_cboPresencaComprador_Estica
    
 
  ElseIf Me.WindowState = 1 Then
    'aqui nada por hora
  Else
    If Grade1.Width < Screen.Width Then
      Tab1.Width = (Screen.Width - 200) / (l_tamanhoOriginal_GRADE1 / l_tamanhoOriginal_TAB1) '
      Grade1.Width = Screen.Width - 200
      Grade1.Groups(0).Width = Screen.Width - 500
      Grade_Serv.Width = Screen.Width - 200
      Grade_Serv.Groups(0).Width = Screen.Width - 500
      
      txtSeq.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_txtSeq)
      Label52.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Label52)
      txtComanda.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_txtComanda)
      lblComanda.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_lblComanda)
      Senha.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Senha)
      Label26.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Label26)
      Nome_Caixa.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Nome_Caixa)
      Combo_Caixa.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Combo_Caixa)
      Label22.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Label22)
      mskValidade.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_mskValidade)
      lblValidade.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_lblValidade)
      cmd_tabelaDePrecos.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_cmd_tabelaDePrecos)
      txtSubTotal.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_txtSubTotal)
      Label35.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Label35)
      txtDescSubTotal.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_txtDescSubTotal)
      Label36.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Label36)
      Label48.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_Label48)
      L_Tot_Pagar.Left = Tab1.Width - (l_tamanhoOriginal_TAB1 - l_L_Tot_Pagar)
      
      B_Servi�os_Conc.Left = Screen.Width - B_Servi�os_Conc.Width - 200
      
      Nome_Cliente.Width = l_Nome_Cliente_Estica + (txtSeq.Left - l_txtSeq)
      Nome_Digitador.Width = l_Nome_Digitador_Estica + (txtSeq.Left - l_txtSeq)
      txtRef.Width = l_txtRef_Estica + (txtSeq.Left - l_txtSeq)
      cboPresencaComprador.Width = l_cboPresencaComprador_Estica + (txtSeq.Left - l_txtSeq)
    
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  
  If gblnSimplesNacional = False Then
      rsEstadosICMS.Close
      Set rsEstadosICMS = Nothing
  End If
  
  rsProdutos2.Close
  rsServi�os.Close
  rsOp_Sa�da.Close
  rsFuncionarios.Close
  rsCliFor.Close
  rsGrade.Close
  rsMovi_Parcelas.Close
  rsParametros.Close
  rsPre�os.Close
  rsSaidas.Close
  rsSaidas_Prod.Close
  rsSaidas_Serv.Close
  rsSa�da_Cheques.Close
  rsSa�da_Parcelas.Close
  rsUsu�rios.Close
  rsTabelas.Close
  rsCota��es.Close
  rsContas_Receber.Close
  rsEstados.Close
  rsCaixas.Close
  rsT�cnicos.Close
  rsOperadores.Close
  rsLog.Close
  '20/12/2006 - Anderson - Altera��o para o registro de CFOP por produto e servi�o
  rsProdutoCFOP.Close
  rsServicoCFOP.Close
 
  '10/12/2009 - Andrea
  rsSa�da_Cheques.Close
 
  Set rsProdutos2 = Nothing
  Set rsServi�os = Nothing
  Set rsOp_Sa�da = Nothing
  Set rsFuncionarios = Nothing
  Set rsCliFor = Nothing
  Set rsGrade = Nothing
  Set rsMovi_Parcelas = Nothing
  Set rsParametros = Nothing
  Set rsPre�os = Nothing
  Set rsSaidas = Nothing
  Set rsSaidas_Prod = Nothing
  Set rsSaidas_Serv = Nothing
  Set rsSa�da_Cheques = Nothing
  Set rsSa�da_Parcelas = Nothing
  Set rsUsu�rios = Nothing
  Set rsTabelas = Nothing
  Set rsCota��es = Nothing
  Set rsContas_Receber = Nothing
  Set rsEstados = Nothing
  Set rsCaixas = Nothing
  Set rsT�cnicos = Nothing
  Set rsOperadores = Nothing
  Set rsLog = Nothing
  '20/12/2006 - Anderson - Altera��o para o registro de CFOP por produto e servi�o
  Set rsProdutoCFOP = Nothing
  Set rsServicoCFOP = Nothing

 
 Set frmSaidas = Nothing
 Unload frmRecebimento
 Set frmRecebimento = Nothing
End Sub

Private Sub Grade_Serv_AfterColUpdate(ByVal ColIndex As Integer)
  Dim nValor As Double
'  Dim bm As Variant
'  bm = Grade_Serv.GetBookmark(0)
'  Select Case Grade_Serv.Columns(ColIndex).Name
'    Case "Pre�o", "Qtde"
      nValor = CDbl(gsHandleNull(Grade_Serv.Columns("Qtde").Text)) * CDbl(gsHandleNull(Grade_Serv.Columns("Pre�o").Text))
      Grade_Serv.Columns("Total").Text = Format(nValor, "#0.00")
'  End Select

End Sub

Private Sub Grade_Serv_AfterUpdate(RtnDispErrMsg As Integer)
  Call Recalcula
End Sub

Private Sub Grade_Serv_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  Dim C�d As Double
  Dim Valor As Single
  Dim Valor_Int As Long
  Dim Aux_Pre�o As Double
  
  Call StatusMsg("")
  
  Aux = Grade_Serv.Columns(ColIndex).Text
  
  '20/12/2006 - Anderson
  'Altera��o para o registro do CFOp por Servi�o
  If cboOper.Text <> "" Then
    rsServicoCFOP.Index = "PrimaryKey"
    rsServicoCFOP.Seek "=", Grade_Serv.Columns(0).Text, cboOper.Text
    If rsServicoCFOP.NoMatch Then
      rsOp_Sa�da.Index = "C�digo"
      rsOp_Sa�da.Seek "=", cboOper.Text
      If Not rsOp_Sa�da.NoMatch Then
        Grade_Serv.Columns("CFOP").Text = rsOp_Sa�da("C�digo Fiscal")
      End If
    Else
      Grade_Serv.Columns("CFOP").Text = "" & rsServicoCFOP("CFOP")
    End If
  End If

  If Grade_Serv.Columns(ColIndex).Name = "C�digo" Then
    With Grade_Serv
      If IsNull(Aux) Or Aux = "" Or Aux = "0" Then
        .Columns("C�digo").Text = 0
        .Columns("Descri��o").Text = ""
        .Columns("Qtde").Text = 0
        .Columns("CFOP").Text = "" '20/12/2006 - Anderson - Altera��o para o registro do CFOp por servi�o
        .Columns("Pre�o").Text = 0
        .Columns("Total").Text = 0
        .Columns("Completo").Text = vbUnchecked
        .Columns("iss").Text = 0
        '26/07/2005 - Daniel
        'Personaliza��o para a empresa J.R. Hidroqu�mica
        'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
        'C.S.T. (C�digo de Situa��o Tribut�ria)
        If m_blnJR Then .Columns("CST").Text = ""
        
        Exit Sub
      ElseIf Not IsNumeric(Aux) Or Val(Aux) < 0 Then
        DisplayMsg "Servi�o incorreto."
        .Columns(ColIndex).Text = ""
        Cancel = True
        Exit Sub
      ElseIf Val(Aux) > 9999# Then
        DisplayMsg "Servi�o incorreto."
        Cancel = True
        Exit Sub
      End If
      C�d = Aux
      rsServi�os.Index = "C�digo"
      rsServi�os.Seek "=", C�d
      If rsServi�os.NoMatch Then
        DisplayMsg "Servi�o n�o encontrado."
        Cancel = True
        Exit Sub
      Else
        .Columns("Descri��o").Text = Trim(rsServi�os("Descri��o") & "")
        .Columns("Pre�o").Text = rsServi�os("Pre�o") & ""
        .Columns("iss").Text = rsServi�os("ISS") & ""
        If gsHandleNull(.Columns("Qtde").Text) = "0" Then
          .Columns("Qtde").Text = "1"
        End If
        '26/07/2005 - Daniel
        'Personaliza��o para a empresa J.R. Hidroqu�mica
        'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
        'C.S.T. (C�digo de Situa��o Tribut�ria)
        If m_blnJR Then .Columns("CST").Text = "0"
        
      End If
    End With
'    If IsNull(Aux) Or Aux = "" Or Val(Aux) = 0 Then
'      Grade_Serv.Columns("Descri��o").Text = ""
'      Grade_Serv.Columns("Qtde").Text = "0"
'      Grade_Serv.Columns("Pre�o").Text = 0
'      Grade_Serv.Columns("Total").Text = 0
'      Grade_Serv.Columns("iss").Text = 0
'      Exit Sub
'    End If
  End If
 
  If Grade_Serv.Columns(ColIndex).Name = "Qtde" Then
    If IsNull(Aux) Then
      Grade_Serv.Columns("Qtde").Text = 0
      Grade_Serv.Columns("Total").Text = 0
      Exit Sub
    ElseIf Aux = "" Then
      Grade_Serv.Columns("Qtde").Text = 0
      Grade_Serv.Columns("Total").Text = 0
      Exit Sub
    ElseIf Not IsVarGoodNumber(Aux) Then
      DisplayMsg "Quantidade incorreta."
      Cancel = True
      Exit Sub
    ElseIf Not IsNumeric(Aux) Then '03/06/2008 - mpdea - Corrigido RT-13
      DisplayMsg "Quantidade incorreta."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) < 0 Then
      DisplayMsg "Quantidade n�o pode ser menor que 0."
      Cancel = True
      Exit Sub
    End If
    
    ' 12/09/2022 - PABLO VER�OSA SILVA
    ' CORRIGE ERRO DE N�MERO COME�ADO COM O CARACTER "," (V�RGULA)
    If Left(Trim(Aux), 1) = "," Then
      Grade_Serv.Columns("Qtde").Text = 0 & Grade_Serv.Columns("Qtde").Text
    End If
  End If
    
  If Grade_Serv.Columns(ColIndex).Name = "Pre�o" Then
    If IsNull(Aux) Then
      Grade_Serv.Columns("Qtde").Text = 0
      Grade_Serv.Columns("Total").Text = 0
      Exit Sub
    ElseIf Aux = "" Then
      Grade_Serv.Columns("Qtde").Text = 0
      Grade_Serv.Columns("Total").Text = 0
      Exit Sub
    ElseIf Not IsNumeric(Aux) Then
      DisplayMsg "Pre�o incorreto."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) < 0 Then
      DisplayMsg "Pre�o n�o pode ser menor que 0."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) > 9999999 Then
      DisplayMsg "Pre�o incorreto, m�ximo � 9.999.999"
      Cancel = True
      Exit Sub
    End If
  End If

End Sub

Private Sub Grade_Serv_GotFocus()
  Grade_Serv.Col = 0
End Sub

'Private Sub Grade_Serv_ComboDropDown()
'  Data5.Refresh
'End Sub
'
Private Sub Grade_Serv_InitColumnProps()
  Grade_Serv.Columns("C�digo").DropDownHwnd = DropDown2.hwnd
End Sub

Private Sub Grade_Serv_KeyPress(KeyAscii As Integer)
  Dim Linha As Integer
  Dim Texto As Variant
  Dim Tamanho As Integer
  
  With Grade_Serv
    If .Col = 0 Then
      If KeyAscii = vbKeyReturn Then
        If ActiveBar1.Tools("miComplLeitorOtico").Checked And Not DropDown2.DroppedDown Then
          If .Columns("C�digo").Text <> "0" Then
            .Columns("Qtde").Text = 1
            SendKeys "{DOWN}{HOME}", True
            KeyAscii = 0
            Exit Sub
          End If
        End If
      End If
    End If
  End With
  
  If KeyAscii = 8 Then
    Exit Sub
  End If
  
  Texto = Grade_Serv.ActiveCell.Text
  Tamanho = 0
  If Not IsNull(Texto) Then
    Tamanho = Len(Texto)
  End If
   
  If Grade_Serv.Col = 0 Then
    Exit Sub
  End If
   
  If Grade_Serv.Col = 1 Then
    If Grade_Serv.Columns("C�digo").Text <> 9999 Then
      If Alterar_Servi�os = 0 Then
        KeyAscii = 0
        Exit Sub
      End If
    End If
'    If Tamanho = 49 Then
'      If Grade_Serv.Row = Grade_Serv.Rows Then
'        KeyAscii = 0
'        Exit Sub
'      End If
'      Linha = Grade_Serv.Row
'      If Tabe_Serv(Linha + 1).Descri��o <> "" Then
'         KeyAscii = 0
'         Exit Sub
'      End If
'      SendKeys ("{Down}")
'      SendKeys ("{Left}")
'      SendKeys ("9999")
'      SendKeys ("{Right}")
'
'      'KeyAscii = 0
'      Exit Sub
'    End If
  End If
    
  If Grade_Serv.Col = 2 Then
    If Tamanho = 10 Then
      KeyAscii = 0
    End If
    Exit Sub
  End If
  
  If Grade_Serv.Columns("C�digo").Text = "" Or Grade_Serv.Columns("C�digo").Text = "0" Then
    KeyAscii = 0
  End If
 
End Sub

Private Sub Grade_Serv_LostFocus()
  If Grade_Serv.RowChanged Then
    Grade_Serv.Update
  End If
End Sub

Private Sub Grade_Serv_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  With Grade_Serv
    .SetFocus
    .Col = .ColContaining(X, y)
    If .Col = 0 Then  'C�digo
      If .ActiveCell.Text = "0" Then
        .ActiveCell.SelStart = 0
        .ActiveCell.SelLength = Len(.ActiveCell.Text)
      End If
    End If
  End With
 
'  If Grade_Serv.ColContaining(X, Y) = 4 Then
'    Exit Sub
'  End If
'
'  Grade_Serv.SetFocus
'  Grade_Serv.Col = Grade_Serv.ColContaining(X, Y)
'
'  If Grade_Serv.Col = 0 Then
'    If Grade_Serv.Columns("C�digo").Text = "0" Then
'      Grade_Serv.ActiveCell.SelStart = 0
'      Grade_Serv.ActiveCell.SelLength = 100
'    End If
'  End If

End Sub

Private Sub Grade_Serv_RowLoaded(ByVal Bookmark As Variant)
'  Dim nCol As Integer
'  For nCol = 0 To Grade_Serv.Cols - 1
'    If Grade_Serv.Columns(nCol).Name = "Total" Then
'      Grade_Serv.Columns("Total").CellStyleSet "Total", Grade_Serv.Row
'    Else
'      Grade_Serv.Columns(nCol).CellStyleSet "Normal", Grade_Serv.Row
'    End If
'  Next nCol
End Sub

Private Sub Grade_Serv_UnboundAddData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, NewRowBookmark As Variant)
  Dim nLinha As Integer
  
  nLinha = Grade1.Row
  
  With Tabe_Serv(nLinha)
    .C�digo = Grade_Serv.Columns("C�digo").Text
    .Descri��o = Grade_Serv.Columns("Descri��o").Text
    .Tempo = Grade_Serv.Columns("Qtde").Text
    .Pre�o = CDbl(Grade_Serv.Columns("Pre�o").Text)
    .Total = CDbl(Grade_Serv.Columns("Total").Text)
    .Completo = gsHandleNull(Grade_Serv.Columns("Completo").Value & "")
    .ISS = Grade_Serv.Columns("iss").Text
    .CFOP_Servico = Grade_Serv.Columns("CFOP").Text '20/12/2006 - Anderson - Altera��o para o registro de CFOP
    '26/07/2005 - Daniel
    'Personaliza��o para a empresa J.R. Hidroqu�mica
    'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
    'C.S.T. (C�digo de Situa��o Tribut�ria)
    If m_blnJR Then .CST = Grade_Serv.Columns("CST").Text & ""
  End With
End Sub


Private Sub Grade_Serv_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim nX As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      nX = Grade_Serv.Rows
    Else
      nX = 0
    End If
  Else
    nX = StartLocation
  End If
  NewLocation = nX + NumberOfRowsToMove
End Sub


Private Sub Grade_Serv_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Dim r, i, J, p As Integer
  Dim nTempo As Single
  Dim nPreco As Double
  Dim sTempo As String
  Dim sPreco As String

  Dim nPos As Integer
  Dim nX As Integer
  Dim nCount As Integer
  
  '27/06/2005 - Daniel
  'Adicionado rotina para tratamento de erro
  On Error GoTo TratarErro
  
  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      nPos = Grade_Serv.Rows
    Else
      nPos = 0
    End If
  Else
    nPos = StartLocation
    If ReadPriorRows Then
      nPos = nPos - 1
    Else
      nPos = nPos + 1
    End If
  End If
  
  With RowBuf
    For nX = 0 To .RowCount - 1
      If nPos < 0 Or nPos >= Grade1.Rows Then
        Exit For
      Else
        .Value(nX, 0) = Tabe_Serv(nPos).C�digo
        .Value(nX, 1) = Tabe_Serv(nPos).Descri��o
        .Value(nX, 2) = Tabe_Serv(nPos).Tempo
        '20/12/2006 - Anderson - Altera��o para o registro de CFOP
        '.Value(nX, 3) = Tabe_Serv(nPos).Pre�o
        '.Value(nX, 4) = Tabe_Serv(nPos).Total
        .Value(nX, 4) = Tabe_Serv(nPos).Pre�o
        .Value(nX, 5) = Tabe_Serv(nPos).Total
        '26/07/2005 - Daniel
        'Personaliza��o para a empresa J.R. Hidroqu�mica
        'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
        'C.S.T. (C�digo de Situa��o Tribut�ria)
        'Nota: (nX, 7) pois o 5 e 6 j� est�o sendo usados
        '20/12/2006 - Anderson - Altera��o para o registro de CFOP
        'If m_blnJR Then .Value(nX, 7) = Tabe_Serv(nPos).CST
        If m_blnJR Then .Value(nX, 8) = Tabe_Serv(nPos).CST
      
        sTempo = gsHandleNull(Tabe_Serv(nPos).Tempo)
        sPreco = gsHandleNull(Tabe_Serv(nPos).Pre�o)
        If IsNumeric(sTempo) And IsNumeric(sPreco) Then
          '20/12/2006 - Anderson - Altera��o para o registro de CFOP
          '.Value(nX, 4) = CSng(sTempo) * CDbl(sPreco)  'Tabe_Serv(p).Total
          .Value(nX, 5) = CSng(sTempo) * CDbl(sPreco)  'Tabe_Serv(p).Total
        Else
          '20/12/2006 - Anderson - Altera��o para o registro de CFOP
          '.Value(nX, 4) = "0"
          .Value(nX, 5) = "0"
        End If
        '20/12/2006 - Anderson - Altera��o para o registro de CFOP
        '.Value(nX, 5) = Tabe_Serv(nPos).Completo
        '.Value(nX, 6) = Tabe_Serv(nPos).ISS
        .Value(nX, 6) = Tabe_Serv(nPos).Completo
        .Value(nX, 7) = Tabe_Serv(nPos).ISS
        .Value(nX, 3) = Tabe_Serv(nPos).CFOP_Servico '20/12/2006 - Anderson - Altera��o para o registro de CFOP por servi�o

        .Bookmark(nX) = nPos
        If ReadPriorRows Then
          nPos = nPos - 1
        Else
          nPos = nPos + 1
        End If
        nCount = nCount + 1
      End If
    Next nX
    .RowCount = nCount
  End With
      
  Exit Sub

TratarErro:
  MsgBox "Ocorr�ncia de erro em Private <Grade_Serv_UnboundReadData>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Err.Clear
      
End Sub

Private Sub Grade_Serv_UnboundWriteData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, WriteLocation As Variant)
  Dim nLinha As Integer
 
  If IsNull(WriteLocation) Then
    Exit Sub
  End If
  nLinha = WriteLocation

  With Tabe_Serv(nLinha)
    .C�digo = gsHandleNull(Grade_Serv.Columns("C�digo").Text)
    .Descri��o = Grade_Serv.Columns("Descri��o").Text
    .Tempo = Grade_Serv.Columns("Qtde").Text
    .Pre�o = CDbl(gsHandleNull(Grade_Serv.Columns("Pre�o").Text))
    .Total = CDbl(gsHandleNull(Grade_Serv.Columns("Total").Text))
    .Completo = gsHandleNull(Grade_Serv.Columns("Completo").Value & "")
    .CFOP_Servico = Grade_Serv.Columns("CFOP").Text '20/12/2006 - Anderson - Altera��o para o registro de CFOP por servi�o
    If IsNull(Grade_Serv.Columns("iss").Text) Then
      Grade_Serv.Columns("iss").Text = 0
    End If
    If Grade_Serv.Columns("iss").Text = "" Then
      Grade_Serv.Columns("iss").Text = 0
    End If
    .ISS = Grade_Serv.Columns("iss").Text
    '26/07/2005 - Daniel
    'Personaliza��o para a empresa J.R. Hidroqu�mica
    'Visualiza��o e tratamento para o Campo [Sa�das - Servi�o].CST
    'C.S.T. (C�digo de Situa��o Tribut�ria)
    If m_blnJR Then .CST = Grade_Serv.Columns("CST").Text & ""
  End With
End Sub

Private Sub Grade1_AfterColUpdate(ByVal ColIndex As Integer)
  Call Calcula_Linha
End Sub

Private Sub Grade1_AfterUpdate(RtnDispErrMsg As Integer)
  Call Recalcula
End Sub

Public Sub Grade1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  Dim C�d As String
  Dim Valor As Single
  Dim Valor_Int As Long
  Dim Aux_Pre�o As Double
  Dim Aux_Produto As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edi��o As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim Aux_Str As String
  Dim Aux_Peso As String
  
  '19/10/2007 - Anderson
  'Implementa��o do campo Lucro M�nimo Permitido conforme solicita��o da Agrotama
  Dim bolLucroMinimoPermitido As Boolean
  
  '09/10/2002 - mpdea
  'Verifica��o de estoque
  Dim dblEstoque As Double
  Dim blnCheckStock As Boolean
  
  '21/10/2002 - mpdea
  'Quantidade
  Dim dblQtde As Double
  
  Dim Balan�a As Integer
  Dim Comp_Prod As Integer
  Dim Pre�o_Balan�a As Double
  Dim In�cio_Pre�o As Integer
  Dim Tam_Pre�o As Integer
  
  '10/11/2004 - Daniel
  Dim strUF As String
  
  ' 25/06/2004 - Maikel Cordeiro
  '---[ Variaveis criadas para a fun��o que elimina o ENTER e TAB da vari�vel AUX na emiss�o da nota fiscal ]---'
    Dim intI       As Integer
    Dim bytAsc     As Byte
    Dim strConcat  As String
  '---[ Variaveis criadas para a fun��o que elimina o ENTER e TAB da vari�vel AUX na emiss�o da nota fiscal ]---'
  
  '08/03/2007 - Anderson
  'Inclus�o de c�digo para resolver problema ao digitar um c�digo do fornecedor igual ao c�digo do produto
  Dim rstProdutos As Recordset
  Dim strSQL      As String

  '10/02/2010 - mpdea
  'Flag para quantidade inicializada (padr�o 1)
  Dim bln_qtde_inicializada As Boolean
  'Flag para quantidade zerada
  Dim bln_qtde_zerada As Boolean

  Call StatusMsg("")
  
  Aux = Trim(Grade1.Columns(ColIndex).Text)
  
  ' 25/06/2004 - Maikel Cordeiro
  '---[ Loop criado para retirar o ENTER e TAB da variavel AUX na emiss�o de nota fiscal ]---'
    For intI = 1 To Len(Aux)
      bytAsc = Asc(Mid(Aux, intI, 1))
      
      If Not (bytAsc = 13 Or bytAsc = 10) Then
        strConcat = strConcat & Chr(bytAsc)
      End If
    Next intI
  '---[ Loop criado para retirar o ENTER e TAB da variavel AUX na emiss�o de nota fiscal ]---'
  
  Aux = strConcat
  
  If ColIndex = 0 Then 'C�digo
    If IsNull(Aux) Or Aux = "" Or Aux = "0" Then
      With Grade1
        .Columns("C�digo").Text = 0
        .Columns("Qtde").Text = 0
        .Columns("Nome").Text = ""
        .Columns("UN").Text = ""
        .Columns("Pre�o Unit.").Text = 0
        .Columns("Total").Text = 0
        .Columns("Desconto").Text = 0
        .Columns("ICM").Text = 0
        .Columns("IPI").Text = 0
        .Columns("CFOP").Text = "" '19/12/2006 - Anderson - Registro do CFOP por produto
        .Columns("Etiq").Text = 0
        .Columns("Pre�o Final").Text = 0
        .Columns("Descri��o Adicional").Text = ""
      End With
      Exit Sub
    End If
    
    '---------------------------------------------------------------------------------------------
    '06/05/2005 - Daniel
    '
    'Implementa��o.: Trabalhar com o c�digo para fornecedor cadastrado na tela de produtos.
    '                Impacto: Ao entrar com o c�digo para o fornecedor no campo c�digo do produto
    '                o sistema dever� trazer o c�digo do produto que estiver amarrado nele
    'Solicita��o...: Cristiano Pavinato - PSI RS
    If m_blnUsaCodFornec Then
      Dim strCodParaFornec As String
      
      '-----------------------
      '08/03/2007 - Anderson
      'Inclus�o de c�digo para resolver problema ao digitar um c�digo do fornecedor igual ao c�digo do produto
      strSQL = "SELECT C�digo, [C�digo do Fornecedor] FROM Produtos WHERE C�digo = '" & Aux & "'"
      
      Set rstProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
      
      If rstProdutos.RecordCount = 0 Then
        strCodParaFornec = Aux
      Else
        '29/05/2008 - mpdea
        'Corrigido RT-94
        strCodParaFornec = rstProdutos("C�digo do Fornecedor") & ""
      End If
      
      Set rstProdutos = Nothing
      
      If Not (strCodParaFornec = "0" Or strCodParaFornec = "") Then
        'strCodParaFornec = Aux
        Aux = g_strBuscarCodProd(strCodParaFornec)
        Grade1.Columns("C�digo").Text = Aux
        
        '07/12/2006 - Anderson
        'Alterado pois causando problemas quando o c�digo do produto fornecedor era igual ao c�digo do produto
        'Se n�o existir nenhum produto amarrado
        'If Aux = strCodParaFornec Then Exit Sub
        If Aux = "" Then
          Aux = strCodParaFornec
          Exit Sub
        End If
      End If
    End If
    '---------------------------------------------------------------------------------------------
  
    Aux_Str = CStr(Aux)
    '26/05/2004 - Daniel
    'Tratamento para 0 'zero' a esquerda
    If Not gbZeroEsquerda Then
      Aux_Str = Retira_Zeros(Aux_Str)
    End If
    Grade1.Columns("C�digo").Text = Aux_Str
    
    '-------------------------------------
    'Rotina para verificar se � de balan�a
    '-------------------------------------
    Balan�a = False
    If ActiveBar1.Tools("miOpEtiquetas").Checked = True Then
      Aux_Str = Aux
      If Len(Aux_Str) >= 12 Then
        Aux_Peso = Left$(Aux_Str, 1)
        If Aux_Peso = "2" Then '� produto pes�vel
          Balan�a = True
          Comp_Prod = rsParametros("Qtde Balan�a")
          If Comp_Prod = 3 Then In�cio_Pre�o = 5
          If Comp_Prod = 3 Then Tam_Pre�o = 8
          If Comp_Prod = 4 Then In�cio_Pre�o = 6
          If Comp_Prod = 4 Then Tam_Pre�o = 7
          If Comp_Prod = 5 Then In�cio_Pre�o = 7
          If Comp_Prod = 5 Then Tam_Pre�o = 6
          If Comp_Prod = 6 Then In�cio_Pre�o = 8
          If Comp_Prod = 6 Then Tam_Pre�o = 5
          Aux_Str = Aux
          Aux = Mid(Aux, 2, Comp_Prod)
          '26/05/2004 - Daniel
          'Tratamento para 0 'zero' a esquerda
          If Not gbZeroEsquerda Then
            Aux = Retira_Zeros(Trim(str(Aux)))
          End If
          C�d = Aux
          Grade1.Columns(0).Text = Aux
          Pre�o_Balan�a = Val(Mid(Aux_Str, In�cio_Pre�o, Tam_Pre�o))
          Pre�o_Balan�a = Pre�o_Balan�a / 100
          'Exit Sub
        End If
      End If
    End If
    
    C�d = Trim(CStr(Aux_Str))
    Aux_Tamanho = 0
    Aux_Cor = 0
    Aux_Edi��o = 0
    
'    Aux_Str = Trim(C�d)
    
    '15/06/2005 - Daniel
    'Solicitante...: On Site - O problema foi encontrado na Agliardi - RS;
    '                N�o estava sendo encontrado o produto
    'Corre��o......: Deixamos como est� na busca do produto na tela de VR
    If Balan�a Then
      Dim strProduto As String
      
      strProduto = CStr(Aux)
      
      Call Acha_Produto(strProduto, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro)
    Else
      Call Acha_Produto(Aux_Str, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro)
    End If
    '
    'Call Acha_Produto... colocado acima... (15/06/2005 - Daniel)
    If Aux_Erro <> 0 Then
      Cancel = True
      If Aux_Erro = 1 Then
        DisplayMsg "Produto n�o existe."
      ElseIf Aux_Erro = 2 Then
        DisplayMsg "Produto com grade, digite tamanho e cor."
      ElseIf Aux_Erro = 3 Then
        DisplayMsg "Produto com edi��o, digite a edi��o tamb�m."
      End If
      Exit Sub
    End If
    
    '16/10/2009 - mpdea
    'Corrigido busca do CFOP para produtos com grade
    '20/12/2006 - Anderson
    'Altera��o para o registro do CFOp por produto
    If cboOper.Text <> "" Then
      rsProdutoCFOP.Index = "PrimaryKey"
      rsProdutoCFOP.Seek "=", Aux_Produto, cboOper.Text
      If rsProdutoCFOP.NoMatch Then
        rsOp_Sa�da.Index = "C�digo"
        rsOp_Sa�da.Seek "=", cboOper.Text
        If Not rsOp_Sa�da.NoMatch Then
          '15/03/2008 - mpdea
          'Corrigido RT-13 ao ler o c�digo fiscal como nulo
          Grade1.Columns("CFOP").Text = rsOp_Sa�da("C�digo Fiscal") & ""
        End If
      Else
        Grade1.Columns("CFOP").Text = rsProdutoCFOP("CFOP") & ""
      End If
    End If
     
    '07/10/2002 - mpdea
    'Posiciona recordset
    rsProdutos2.FindFirst "C�digo = '" & Aux_Produto & "'"
    
    If rsProdutos2.NoMatch Then
        MsgBox "Produto " & Aux_Produto & " ainda N�O est� dispon�vel na tela de SA�DAS." & vbCrLf & vbCrLf & "Ent�o FECHE a tela de SA�DAS e ABRA novamente caso queira lan�ar este produto.", vbInformation, "Aten��o"
        Exit Sub
    End If
    
    Aux_Produto = Trim(Aux_Produto)
    
    If Balan�a = False Then
      If Grade1.Columns(1).Text = "0" Then
        Grade1.Columns(1).Text = "1"
        '10/02/2010 - mpdea
        'Campo quantidade inicializado com valor padr�o
        bln_qtde_inicializada = True
      End If
    End If
    
    '------------------------------------------------------------------------------
    '19/08/2003 - mpdea
    'Modificado nome do campo

    '21/10/2002 - mpdea
    'Corrigido verifica��o da quantidade atrav�s da vari�vel dblQtde

    '07/10/2002 - mpdea
    'Verifica��o de estoque
    
    If ((Not rsParametros.Fields("Venda Sem Estoque Saidas").Value) And _
       rsProdutos2.Fields("Estoque").Value) Then

      blnCheckStock = False

      'Opera��o de sa�da
      Call cboOper_LostFocus
      If Not rsOp_Sa�da.NoMatch Then
        'Ativa flag se a opera��o movimenta estoque
        If rsOp_Sa�da.Fields("Estoque").Value Then
          blnCheckStock = True
        End If
      End If

      If blnCheckStock Then
        dblQtde = CDbl(Grade1.Columns("Qtde").Text)  '"0" &
        dblEstoque = -999999
        dblEstoque = Acha_Estoque(gnCodFilial, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Erro)
        If Aux_Erro = 0 Then
          If CDbl(dblQtde) > dblEstoque Then
            If dblEstoque <> -999999 Then
              '10/02/2010 - mpdea
              'Zera quantidade quando o produto for fracionado, a quantidade for inicializada automaticamente (padr�o 1),
              'possuir estoque maior do que 0 e inferior a 1
              'Resolve quest�es para vendas de produtos fracionados que possuem estoque como 0,8
              If gbIsFrac(C�d, 0) And bln_qtde_inicializada And dblEstoque > 0 And dblEstoque < 1 Then
                Grade1.Columns(1).Text = "0"
                bln_qtde_zerada = True
              Else
                DisplayMsg "Quantidade superior ao estoque. Estoque atual: " & dblEstoque
                If CDbl(dblQtde) <> 0 Then Cancel = True
                '13/08/2007 - Anderson
                'Linha retirada para evitar que ocorra a mensagem run-time error 5
                'Grade1.SetFocus
                Exit Sub
              End If
            End If
          End If
        Else
          '08/11/2002 - mpdea
          'Adicionado descri��o do erro 1
          If Aux_Erro = 1 Then
            DisplayMsg "Produto com estoque n�o inicializado."
          Else
            DisplayMsg "Erro [" & Aux_Erro & "] ao encontrar estoque do produto."
          End If
          Cancel = True
          Exit Sub
        End If
      End If
    End If
    '------------------------------------------------------------------------------
    
    If rsProdutos2.Fields("Desativado") Then
      MsgBox "Produto Inativo, verifique !", vbCritical, "Quick Store"
      Grade1.Columns(0).Text = "0"
      Grade1.Columns(1).Text = "0"
      Exit Sub
    End If
    
    C�d = Trim(rsProdutos2("C�digo"))
    
    With Grade1
      .Columns("Nome").Text = rsProdutos2("Nome") & ""
      .Columns("UN").Text = rsProdutos2("Unidade Venda") & ""
      
      '''.Columns("IPI").Text = rsProdutos2("Percentual IPI")
      
      ' ==============================================================
      ' Tratar IPI
      ' Se a Opera��o de Sa�da estiver classificada como G = Devolu��o/Remessas/GratisSaida utiliza-se o Percentual de Entrada
      ' e n�o importa se � uma Devolu��o com finalidade=4 ou Remessa com finalidade=1
      If Trim(cboOper.Text) <> "" Then
          If rsParametros("CodigoRegimeTributario") = 1 Then
              ' Empresa SIMPLES NACIONAL
              If rsOp_Sa�da.Fields("tipo").Value = "G" Then    'or cboFinalidade.ListIndex = 3
                  'Finalidade devolu��o
                  If Not IsNull(rsProdutos2("Percentual_IPI_Entrada")) Then
                      .Columns("IPI").Text = rsProdutos2("Percentual_IPI_Entrada")
                  Else
                      .Columns("IPI").Text = "0"
                  End If
              Else
                  .Columns("IPI").Text = "0"
              End If
          Else
              ' Empresa LUCRO REAL
              If rsOp_Sa�da.Fields("tipo").Value = "G" Then    'or cboFinalidade.ListIndex = 3
                  'Finalidade devolu��o
                  If Not IsNull(rsProdutos2("Percentual_IPI_Entrada")) Then
                      .Columns("IPI").Text = rsProdutos2("Percentual_IPI_Entrada")
                  Else
                      .Columns("IPI").Text = "0"
                  End If
              Else
                  If Not IsNull(rsProdutos2("Percentual IPI")) Then
                      .Columns("IPI").Text = rsProdutos2("Percentual IPI") 'saida
                  Else
                      .Columns("IPI").Text = "0"
                  End If
              End If
          End If
      Else
          MsgBox "ESCOLHA UMA OPERA��O ANTES DE LAN�AR PRODUTOS", vbInformation, "ATEN��O"
          .Columns("IPI").Text = "0"
      End If
      
'''      If Not IsNull(rsProdutos2("IPI_ValidoEntradaSaida").Value) And rsProdutos2("IPI_ValidoEntradaSaida").Value = 1 Then
'''          .Columns("IPI").Text = rsProdutos2("Percentual IPI")
'''      ElseIf Not IsNull(rsProdutos2("IPI_ValidoEntradaSaida").Value) And rsProdutos2("IPI_ValidoEntradaSaida").Value <> 1 Then
'''          If cboFinalidade.ListIndex = 3 Then
'''              'Finalidade devolu��o
'''              .Columns("IPI").Text = rsProdutos2("Percentual IPI")
'''          Else
'''              .Columns("IPI").Text = "0"
'''          End If
'''      Else
'''          .Columns("IPI").Text = "0"
'''      End If
      ' ==============================================================
      
      ' ***********************************************************************
      ' Tratamento NOVO para visualizar o ICMS
      If cboFinalidade.ListIndex = 3 Then  'Nota de Devolu��o
          If Estado = "" Then
              .Columns("ICM").Text = rsProdutos2("Percentual ICM Entrada") & ""
          Else
              ' Tratar "mais pra frente" se for Devolu��o Interestadual com a tabela de ICMS_PERCENTUAL_ESTADOS
              ' Por enquanto fazer o tratamento igual a Devolu��o Estadual
              .Columns("ICM").Text = rsProdutos2("Percentual ICM Entrada") & ""
          End If

      Else 'Nota de sa�da normal
          If gblnSimplesNacional = False Then
              If Estado = "" Then
                  .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida") & ""
              Else
                  aliquotaICMS_tab_ICMS_PERC_ESTADOS = ""
          
                  If Not (rsEstadosICMS.EOF And rsEstadosICMS.BOF) Then
                      rsEstadosICMS.MoveFirst
                      While Not rsEstadosICMS.EOF
                          If UCase(rsEstadosICMS.Fields("ESTADO_ORIGEM").Value) = UCase(gsEstadoOrigemEmpresaLogado) And _
                            UCase(rsEstadosICMS.Fields("ESTADO_DESTINO").Value) = UCase(rsCliFor("Estado")) Then
                              aliquotaICMS_tab_ICMS_PERC_ESTADOS = rsEstadosICMS.Fields("ALIQUOTA").Value
                              rsEstadosICMS.MoveLast
                          End If
                          rsEstadosICMS.MoveNext
                      Wend
                  End If

                  If UCase(gsEstadoOrigemEmpresaLogado) = UCase(rsCliFor("Estado")) Then
                      bo_AliquotaICMS_interestadual = False
                      .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida") & ""
                  Else
                      bo_AliquotaICMS_interestadual = True
                      .Columns("ICM").Text = aliquotaICMS_tab_ICMS_PERC_ESTADOS
                  End If
              End If
          End If
      End If
      ' ***********************************************************************

'''      'Mostra ICM do Estado
'''      If Estado = "" Then
'''        .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida") & ""
'''      ElseIf Estado <> "" Then
'''        rsEstados.Index = "Estado"
'''        rsEstados.Seek "=", Estado
'''        If rsEstados.NoMatch Then
'''          .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida") & ""
'''        ElseIf Not rsEstados.NoMatch Then
'''          If rsEstados("ICM") = -1 Then 'Estado Local
'''             .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida") & ""
'''          Else
'''              '10/11/2004 - Daniel
'''              'Tratamento do ICM solicitado pela Teknika
'''              If Not m_blnTeknika Then 'Demais clientes
'''
'''                '10/01/2005 - Daniel
'''                'Adicionado tratamento especial para a Embalavi
'''                'Sempre que o cliente seja pessoa 'F�sica' independente
'''                'do Estado a taxa ser� '18%' o valor de [rsProdutos2("Percentual ICM Saida").Value]
'''                If m_blnEmbalavi Then
'''
'''                  If Len(cboCliente.Text) > 0 Then 'Est� preenchido
'''                    If PessoaFisica(cboCliente.Text) Then
'''                      .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida").Value & ""
'''                    Else
'''                      .Columns("ICM").Text = rsEstados("ICM")
'''                    End If
'''
'''                  Else 'N�o ter� como verificar sem saber o cliente
'''                    .Columns("ICM").Text = rsEstados("ICM")
'''                  End If
'''
'''                Else 'Demais clientes
'''                  .Columns("ICM").Text = rsEstados("ICM")
'''                End If
'''
'''              Else
'''
'''                If IE_Isento(strUF) Then 'ISENTO = TRUE
'''                  If strUF = "PR" Then
'''                    .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida") & ""
'''                  Else
'''                    .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida") & ""
'''                  End If
'''                Else 'ISENTO = FALSE
'''                  If strUF = "PR" Then
'''                    .Columns("ICM").Text = rsProdutos2("Percentual ICM Saida") & ""
'''                  Else
'''                    .Columns("ICM").Text = rsEstados("ICM")
'''                  End If
'''
'''                End If
'''
'''              End If
'''
'''          End If
'''        End If
'''      End If
      
      .Columns("Base_ICM").Text = 0
      .Columns("Valor_ICM").Text = 0
      .Columns("Valor_Base_Unit").Text = 0
      .Columns("Redu��o_ICM").Text = 0
      .Columns("Tipo_ICM").Text = rsProdutos2("Tipo ICM") & ""
      Select Case rsProdutos2("Tipo ICM")
        Case "I"
          .Columns("ICM").Text = "0"
        Case "R" 'ICM Retido
          If rsProdutos2("Base C�lculo") <> 0 Then
            .Columns("Valor_Base_Unit").Text = rsProdutos2("Base C�lculo")
          End If
          If rsProdutos2("Redu��o ICM") <> 0 Then
            .Columns("Redu��o_ICM").Text = rsProdutos2("Redu��o ICM")
          End If
        Case "Z" 'ICM Reduzido
          If rsProdutos2("Base C�lculo") <> 0 Then
            .Columns("Valor_Base_Unit").Text = rsProdutos2("Base C�lculo")
          End If
          If rsProdutos2("Redu��o ICM") <> 0 Then
            .Columns("Redu��o_ICM").Text = rsProdutos2("Redu��o ICM")
          End If
      End Select
      'Acha pre�o
      rsPre�os.Index = "Tabela"
      If Combo_Pre�o.Text = "" Then
        .Columns("Pre�o Unit.").Text = 0
      Else
      
          ' *********************************************
          ' AJUSTE ABRIL/19 PARA TRATAMENTO DE VALOR ACATADO NA TELA DE PESQUISA DE PRODUTO
          If gTabelaPrecoAcatadaTelaPesquisaProduto <> "" Then
              rsPre�os.Seek "=", gTabelaPrecoAcatadaTelaPesquisaProduto, C�d
          Else
              rsPre�os.Seek "=", Combo_Pre�o.Text, C�d
          End If
          
          gTabelaPrecoAcatadaTelaPesquisaProduto = ""
        
'''          rsPre�os.Seek "=", Combo_Pre�o.Text, C�d
          If rsPre�os.NoMatch Then
              .Columns("Pre�o Unit.").Text = 0
          Else
              Aux_Pre�o = rsPre�os("Pre�o") * ((100 - (rsProdutos2("Desconto") + Desconto_Cli)) / 100)
              If rsProdutos2("Moeda") <> 1 Then
                 rsCota��es.Index = "Moeda"
                 rsCota��es.Seek "<=", rsProdutos2("Moeda"), Data_Atual
                 If Not rsCota��es.NoMatch Then
                     If rsCota��es("Moeda") = rsProdutos2("Moeda") Then
                         Aux_Pre�o = Aux_Pre�o * rsCota��es("Cota��o")
                     End If
                 End If
              End If
               
              '04/05/2004 - Daniel
              'Personaliza��o Embalavi
              If g_bln5CasasDecimais Then
                .Columns("Pre�o Unit.").Text = Format(Aux_Pre�o, "#0.00000")
                '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
              ElseIf g_bln3CasasDecimais Then
                .Columns("Pre�o Unit.").Text = Format(Aux_Pre�o, "#0.000")
              Else
                .Columns("Pre�o Unit.").Text = Format(Aux_Pre�o, "#0.00")
              End If
          End If
      End If
      If Balan�a = True Then
        '04/05/2004 - Daniel
        'Personaliza��o Embalavi
        If g_bln5CasasDecimais Then
          Grade1.Columns(1).Text = Format(Pre�o_Balan�a / rsPre�os("Pre�o"), "######0.00000")
        '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
          Grade1.Columns(1).Text = Format(Pre�o_Balan�a / rsPre�os("Pre�o"), "######0.000")
        Else
          Grade1.Columns(1).Text = Format(Pre�o_Balan�a / rsPre�os("Pre�o"), "######0.00#")
        End If
      End If
      If gsHandleNull(.Columns("Qtde").Text) = "0" And Not bln_qtde_zerada Then
        .Columns("Qtde").Text = "1"
      End If
    End With
  End If
 ' ***********
'  If Calcula_ICM = True And Not IsNull(rsOp_Sa�da("Perc Icms Frete")) Then
'     If Estado = "" Then
'         PercIcmsFrete = rsOp_Sa�da("Perc Icms Frete")
'     ElseIf Estado <> "" Then
'         rsEstados.Index = "Estado"
'         rsEstados.Seek "=", Estado
'         If rsEstados.NoMatch Then
'              PercIcmsFrete = rsOp_Sa�da("Perc Icms Frete")
'         ElseIf Not rsEstados.NoMatch Then
'             If rsEstados("ICM") = -1 Then
'                 PercIcmsFrete = rsOp_Sa�da("Perc Icms Frete")
'              Else
'                 PercIcmsFrete = rsEstados("ICM")
'              End If
'         End If
'    End If
'  Else
'    PercIcmsFrete = 0
'  End If
 '**************
 
  If ColIndex = 1 Then 'Qtde
  
    If Grade1.Columns(0).Text = "" Or Grade1.Columns(0).Text = "0" Then
      Grade1.Columns(1).Text = "0"
      Exit Sub
    End If
    
    
    '07/10/2002 - mpdea
    'Verifica se o produto existe e obt�m dados para consulta de estoque
    Aux_Str = Trim(Grade1.Columns(0).Text)
    Call Acha_Produto(Aux_Str, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro)
    If Aux_Erro <> 0 Then
      Cancel = True
      If Aux_Erro = 1 Then
        DisplayMsg "Produto n�o existe."
      ElseIf Aux_Erro = 2 Then
        DisplayMsg "Produto com grade, digite tamanho e cor."
      ElseIf Aux_Erro = 3 Then
        DisplayMsg "Produto com edi��o, digite a edi��o tamb�m."
      End If
      Exit Sub
    End If
    
    '07/10/2002 - mpdea
    'Posiciona recordset
    rsProdutos2.FindFirst "C�digo = '" & Aux_Produto & "'"
    
    
    If IsNull(Aux) Then
      Grade1.Columns("Qtde").Text = "0"
      Calcula_Linha
      Exit Sub
    ElseIf Aux = "" Then
      Grade1.Columns("Qtde").Text = "0"
      Calcula_Linha
      Exit Sub
    ElseIf Not IsVarGoodNumber(Aux) Then
      DisplayMsg "Quantidade incorreta."
      Cancel = True
      Exit Sub
   ElseIf CDbl(Aux) <= 0 Then
      DisplayMsg "Quantidade n�o pode ser zero."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) > 9999999 Then
      DisplayMsg "Quantidade incorreta, m�xima � 9999999."
      Cancel = True
      Exit Sub
    End If
    
    
    '------------------------------------------------------------------------------
    '19/08/2003 - mpdea
    'Modificado nome do campo
    '
    '07/10/2002 - mpdea
    'Verifica��o de estoque
    If (Not (rsParametros.Fields("Venda Sem Estoque Saidas").Value) And _
       rsProdutos2.Fields("Estoque").Value) Then

      blnCheckStock = False

      'Opera��o de sa�da
      Call cboOper_LostFocus
      If Not rsOp_Sa�da.NoMatch Then
        'Ativa flag se a opera��o movimenta estoque
        If rsOp_Sa�da.Fields("Estoque").Value Then
          blnCheckStock = True
        End If
      End If

      If blnCheckStock Then
        dblEstoque = -999999
        dblEstoque = Acha_Estoque(gnCodFilial, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Erro)
        If Aux_Erro = 0 Then
          If CDbl(Aux) > dblEstoque Then
            If dblEstoque <> -999999 Then
              DisplayMsg "Quantidade superior ao estoque. Estoque atual: " & dblEstoque
              If CDbl(Aux) <> 0 Then Cancel = True
              '13/08/2007 - Anderson
              'Linha retirada para evitar que ocorra a mensagem run-time error 5
              'Grade1.SetFocus
              Exit Sub
            End If
          End If
        Else
          '08/11/2002 - mpdea
          'Adicionado descri��o do erro 1
          If Aux_Erro = 1 Then
            DisplayMsg "Produto com estoque n�o inicializado."
          Else
            DisplayMsg "Erro [" & Aux_Erro & "] ao encontrar estoque do produto."
          End If
          Cancel = True
          Exit Sub
        End If
      End If
    End If
    '------------------------------------------------------------------------------
    
    
    'Verifica se Qtde � decimal
    Valor = Aux
    Valor_Int = Aux
    If Valor = Valor_Int Then
      Calcula_Linha
      Exit Sub
    End If
    
    Aux = Grade1.Columns("C�digo").Text
    'Acha produto
    If IsNull(Aux) Or Aux = "" Or Val(Aux) = 0 Then
      Exit Sub
    End If
    Aux_Str = Trim(Aux)
    Call Acha_Produto(Aux_Str, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro)
    If Aux_Erro <> 0 Then
      Cancel = True
      Exit Sub
    End If
    
    rsProdutos2.FindFirst "C�digo = '" & Aux_Produto & "'"
    If rsProdutos2.NoMatch Then
      Exit Sub
    ElseIf Not rsProdutos2("Fracionado") Then
      DisplayMsg "Produto n�o aceita quantidade fracionada."
      Cancel = True
      Exit Sub
    
    
    '14/11/2002 - mpdea
    'Inclu�do formata��o de produtos fracionados
    ElseIf rsProdutos2("Fracionado").Value Then
      Grade1.Columns("Qtde").Text = Format(Valor, "#0." & String(rsProdutos2("QtdeCasasDecimais").Value, "0"))
    End If


  End If
    
  If ColIndex = 4 Then  'Pre�o
    If IsNull(Aux) Then
       Grade1.Columns("Pre�o Unit.").Text = 0
       Calcula_Linha
       Exit Sub
    ElseIf Aux = "" Then
       Grade1.Columns("Pre�o Unit.").Text = 0
       Calcula_Linha
       Exit Sub
    ElseIf Not IsNumeric(Aux) Then
      DisplayMsg "Pre�o incorreto."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) < 0 Then
      DisplayMsg "Pre�o n�o pode ser menor que 0."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) > 9999999 Then
      DisplayMsg "Pre�o incorreto, m�ximo � 9.999.999"
      Cancel = True
      Exit Sub
    '29/10/2007 - Anderson
    'Implementa��o do campo Lucro M�nimo Permitido conforme solicita��o da Agrotama
    ElseIf g_bolLucroMinimoClasse Then
       If Not PermiteDescontoMargemLucro(Grade1.Columns("C�digo").Text, Grade1.Columns("Desconto").Text, Grade1.Columns("Qtde").Text, Grade1.Columns("Pre�o Unit.").Text) And Not m_bolLucroMinimoPermitido Then
        DisplayMsg "Pre�o unit�rio n�o permitido para este produto. Esta opera��o � permitada apenas com a senha do gerente."
        Cancel = True
        Exit Sub
      End If
    End If
  End If

  If ColIndex = 6 Then  'Desconto
    If IsNull(Aux) Then
      Grade1.Columns("Desconto").Text = 0
      Calcula_Linha
      Exit Sub
    ElseIf Aux = "" Then
      Grade1.Columns("Desconto").Text = 0
      Calcula_Linha
      Exit Sub
    ElseIf Not IsNumeric(Aux) Then
      DisplayMsg "Desconto incorreto."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) < 0 Or CDbl(Aux) > 100 Then
      DisplayMsg "Desconto n�o pode ser menor que 0 ou maior que 100."
      Cancel = True
      Exit Sub
    '19/10/2007 - Anderson
    'Implementa��o do campo Lucro M�nimo Permitido conforme solicita��o da Agrotama
    ElseIf g_bolLucroMinimoClasse Then
      If Not PermiteDescontoMargemLucro(Grade1.Columns("C�digo").Text, Aux, Grade1.Columns("Qtde").Text, Grade1.Columns("Pre�o Unit.").Text) And Not m_bolLucroMinimoPermitido Then
        DisplayMsg "Desconto n�o permitido para este produto. Esta opera��o � permitada apenas com a senha do gerente."
        Cancel = True
        Exit Sub
      End If
    End If
  End If

  If ColIndex = 7 Then  'ICM
    If IsNull(Aux) Then
      Grade1.Columns("ICM").Text = 0
      Exit Sub
    ElseIf Aux = "" Then
      Grade1.Columns("ICM").Text = 0
      Exit Sub
    ElseIf Not IsNumeric(Aux) Then
      DisplayMsg "ICM incorreto."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) < 0 Or CDbl(Aux) > 999 Then
      DisplayMsg "ICM incorreto, deve ser entre 0 e 999."
      Cancel = True
      Exit Sub
    End If
  End If

  If ColIndex = 8 Then  'IPI
    If IsNull(Aux) Then
      Grade1.Columns("IPI").Text = 0
      Calcula_Linha
      Exit Sub
    ElseIf Aux = "" Then
      Grade1.Columns("IPI").Text = 0
      Calcula_Linha
      Exit Sub
    ElseIf Not IsNumeric(Aux) Then
      DisplayMsg "IPI incorreto."
      Cancel = True
      Exit Sub
    ElseIf CDbl(Aux) < 0 Or CDbl(Aux) > 999 Then
      DisplayMsg "IPI incorreto, deve ser entre 0 e 999."
      Cancel = True
      Exit Sub
    End If
  End If

'  Call Calcula_Linha
'  Call RecalculaPrecos

End Sub

Private Sub Grade1_Click()

  sCodigoProdutoDevolucao = Grade1.Columns(0).Value
  lQuantidadeItensProdutoDevolucao = Grade1.Columns(1).Value
  sNomeProdutoDevolucao = Grade1.Columns(2).Value
  sValorUnitarioProdutoDevolucao = Grade1.Columns(4).Value
 
End Sub

Private Sub Grade1_GotFocus()

  '30/06/2005 - Daniel
  'Adicionado tratamento de erros
  On Error GoTo TratarErro
  
  Grade1.Col = 0
  SendKeys "{Home}+{End}"
  
  Exit Sub
  
TratarErro:
  MsgBox "Erro: " & Err.Number & Err.Description, vbExclamation, "Quick Store"
  Err.Clear
  Exit Sub
  
End Sub

Private Sub Grade1_InitColumnProps()
  With Grade1
    .Columns("C�digo").DropDownHwnd = DropDown1.hwnd
    .Columns("Etiq").Style = ssStyleCheckBox
  End With
End Sub

Private Sub Grade1_KeyPress(KeyAscii As Integer)
  
  If Len(Grade1.Columns("C�digo").Text) > 0 Then
    If Asc(Grade1.Columns("C�digo").Text) = 13 Then Grade1.Columns("C�digo").Text = ""
  End If
  
  If Grade1.Col = 0 Then
    If DropDown1.DroppedDown Then
      '17/10/2007 - Anderson
      'Implementa��o do tipo de pesquisa na combo de c�digo do produto
      If ActiveBar1.Tools("miOpPesquisarDescricao").Checked = True Then
        DropDown1.DataFieldList = "Nome"
      Else
        DropDown1.DataFieldList = "C�digo"
      End If
    End If
    
    If KeyAscii = vbKeyReturn Then
      If ActiveBar1.Tools("miComplLeitorOtico").Checked And Not DropDown1.DroppedDown Then
        
        If Grade1.Columns("C�digo").Text <> "" And Grade1.Columns("C�digo").Text <> "0" Then
              Grade1.Columns("Qtde").Text = 1
              'Grade1.SetFocus '07/10/2004 - Daniel - Case: Sucupira
              SendKeys "{DOWN}{HOME}", True
              
              '07/10/2004 - Daniel - Case: Sucupira
              If Grade1.Columns("C�digo").Text = "" Then
                Grade1.Col = 0
                Grade1.SetFocus '21/10/2004 - For�ar o foco
                Grade1.Col = 0  'ADICIONADA ESTA LINHA....FOCO SEM O ZERO 2019 AGOSTO
              Else
                SendKeys "{HOME}", True
                'Grade1.SetFocus '21/10/2004 - For�ar o foco {comentada esta linha em 29/11/2004 - Daniel}
              End If
              
              '27/07/2004 - mpdea
              'Comentado devido a perda de performance da busca
              'pela lista de produtos (permanece como 0 - zero)
              '.Text = "" 'Replace(Grade1.Columns("C�digo").Text, Chr(13), "")
              
              KeyAscii = 0
            
            End If
      End If
    End If
    
  ElseIf Grade1.Col = 1 Then
    
    If KeyAscii = vbKeyReturn Then
      If ActiveBar1.Tools("miComplLeitorOtico").Checked Then
        If Grade1.Columns("C�digo").Text <> "0" Then
          '16/10/2009 - mpdea
          'Modo de entrada de dados no grid de produtos
          Select Case m_int_modo_grid_produtos
            Case 1
              SendKeys "{DOWN}{HOME}", True
          
          End Select
          
          KeyAscii = 0
        End If
      End If
    End If
    
  End If

End Sub

Private Sub Grade1_LostFocus()
  If Grade1.RowChanged Then
    Grade1.Update
  End If
End Sub

Private Sub Grade1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  With Grade1
    .SetFocus
    .Col = .ColContaining(X, y)
    If .Col = 0 Then  'C�digo
      If .ActiveCell.Text = "0" Then
        .ActiveCell.SelStart = 0
        .ActiveCell.SelLength = Len(.ActiveCell.Text)
      End If
    End If
  End With
End Sub

Private Sub Grade1_UnboundAddData(ByVal RowBuf As ssRowBuffer, NewRowBookmark As Variant)
  Dim nLinha As Integer

  nLinha = Grade1.Row
  
  With Tabe(nLinha)
    .C�digo = Grade1.Columns("C�digo").Text
    .Qtde = CDbl(Grade1.Columns("Qtde").Text)
    .Nome = Grade1.Columns("Nome").Text
    .Unidade = Grade1.Columns("UN").Text
    '04/05/2004 - Embalavi
    'Personaliza��o Embalavi
    If g_bln5CasasDecimais Then
      .Pre�o = CDbl(Format((Grade1.Columns("Pre�o Unit.").Text), "##,###,##0.00000"))
    '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Pre�o = CDbl(Format((Grade1.Columns("Pre�o Unit.").Text), "##,###,##0.000"))
    Else
      '.Pre�o = CDbl(Grade1.Columns("Pre�o Unit.").Text)
      .Pre�o = CDbl(Format((Grade1.Columns("Pre�o Unit.").Text), "##,###,##0.00"))
    End If
    .Pre�o_Total = CDbl(Grade1.Columns("Total").Text)
    .Desconto = CDbl(Grade1.Columns("Desconto").Text)
    .ICM = CDbl(Grade1.Columns("ICM").Text)
    .IPI = CDbl(Grade1.Columns("IPI").Text)
    .Etiqueta = Grade1.Columns("Etiq").Text
    .Pre�o_Final = CDbl(Grade1.Columns("Pre�o Final").Text)
    .Base_ICM = CDbl(Grade1.Columns("Base_ICM").Text)
    .Valor_ICM = CDbl(Grade1.Columns("Valor_ICM").Text)
    .Valor_Base_Unit = CDbl(Grade1.Columns("Valor_Base_Unit").Text)
     '19/12/2006 - Anderson - Altera��o realizada para evitar o erro de 13 - Type Mismacth
     .Redu��o_ICM = CDbl("0" & Grade1.Columns("Redu��o_ICM").Text)
    '.Redu��o_ICM = CDbl(Grade1.Columns("Redu��o_ICM").Text)
    .Tipo_ICM = Grade1.Columns("Tipo_ICM").Text
    .CFOP_Produto = Grade1.Columns("CFOP").Text '20/12/2006 - Anderson - Altera��o para o Registro de SCFOP por produto
    .Desp_Acessorias = Grade1.Columns("Desp_Acessorias").Text
  End With
End Sub

Private Sub Grade1_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim nX As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      nX = Grade1.Rows
    Else
      nX = 0
    End If
  Else
    nX = StartLocation
  End If
  NewLocation = nX + NumberOfRowsToMove
End Sub

Private Sub Grade1_UnboundReadData(ByVal RowBuf As ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Dim nPos As Integer
  Dim nX As Integer
  Dim nCount As Integer
  
  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      nPos = Grade1.Rows
    Else
      nPos = 0
    End If
  Else
    nPos = StartLocation
    If ReadPriorRows Then
      nPos = nPos - 1
    Else
      nPos = nPos + 1
    End If
  End If
  
  With RowBuf
    For nX = 0 To .RowCount - 1
      If nPos < 0 Or nPos >= Grade1.Rows Then
        Exit For
      Else
        .Value(nX, 0) = Tabe(nPos).C�digo
        .Value(nX, 1) = Tabe(nPos).Qtde
        .Value(nX, 2) = Tabe(nPos).Nome
        .Value(nX, 3) = Tabe(nPos).Unidade
        .Value(nX, 4) = Tabe(nPos).Pre�o
        .Value(nX, 5) = Tabe(nPos).Pre�o_Total
        .Value(nX, 6) = Tabe(nPos).Desconto
        .Value(nX, 7) = Tabe(nPos).ICM
        .Value(nX, 8) = Tabe(nPos).IPI
'15/10/2007 - Anderson
'Implementa��o do CFOP por produto
'        .Value(nX, 9) = Tabe(nPos).Etiqueta
'        .Value(nX, 10) = Tabe(nPos).Pre�o_Final
'        .Value(nX, 11) = Tabe(nPos).Base_ICM
'        .Value(nX, 12) = Tabe(nPos).Valor_ICM
'        .Value(nX, 13) = Tabe(nPos).Valor_Base_Unit
'        .Value(nX, 14) = Tabe(nPos).Redu��o_ICM
'        .Value(nX, 15) = Tabe(nPos).Tipo_ICM
'        .Value(nX, 16) = Tabe(nPos).Descr_Adicional
        .Value(nX, 9) = Tabe(nPos).CFOP_Produto
        .Value(nX, 10) = Tabe(nPos).Etiqueta
        .Value(nX, 11) = Tabe(nPos).Pre�o_Final
        .Value(nX, 12) = Tabe(nPos).Base_ICM
        .Value(nX, 13) = Tabe(nPos).Valor_ICM
        .Value(nX, 14) = Tabe(nPos).Valor_Base_Unit
        .Value(nX, 15) = Tabe(nPos).Redu��o_ICM
        .Value(nX, 16) = Tabe(nPos).Tipo_ICM
        
        '''''''''''''
        .Value(nX, 17) = Tabe(nPos).Desp_Acessorias
        .Value(nX, 18) = Tabe(nPos).Valor_Desonerado
        .Value(nX, 19) = Tabe(nPos).Percentual_Diferimento
        .Value(nX, 20) = Tabe(nPos).Descr_Adicional
        '''''''''''''
        
        
'''''''''''        .Value(nX, 17) = Tabe(nPos).Descr_Adicional
'''''''''''        '.Value(nX, 18) = Tabe(nPos).[Motivo Desoneramento]
'''''''''''        .Value(nX, 18) = Tabe(nPos).Valor_Desonerado
'''''''''''        .Value(nX, 19) = Tabe(nPos).Percentual_Diferimento

        .Bookmark(nX) = nPos
        If ReadPriorRows Then
          nPos = nPos - 1
        Else
          nPos = nPos + 1
        End If
        nCount = nCount + 1
      End If
    Next nX
    .RowCount = nCount
  End With
  
End Sub

Private Sub Grade1_UnboundWriteData(ByVal RowBuf As ssRowBuffer, WriteLocation As Variant)
On Error GoTo Erro

  Dim nLinha As Integer
  
  nLinha = WriteLocation
  
  If Grade1.Columns("Qtde").Text = "" Then
    Grade1.Columns("Qtde").Text = "0"
  End If
  
  If Grade1.Columns("Valor Desonerado").Text = "" Then
    Grade1.Columns("Valor Desonerado").Text = "0"
  End If
  
  With Tabe(nLinha)
    .C�digo = Grade1.Columns("C�digo").Text
    .Qtde = CDbl(Grade1.Columns("Qtde").Text)
    .Nome = Grade1.Columns("Nome").Text
    .Unidade = Grade1.Columns("UN").Text
    '04/05/2004 - Daniel
    'Personaliza��o Embalavi
    If g_bln5CasasDecimais Then
      .Pre�o = CDbl(Format((Grade1.Columns("Pre�o Unit.").Text), "##,###,##0.00000"))
    '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
    ElseIf g_bln3CasasDecimais Then
      .Pre�o = CDbl(Format((Grade1.Columns("Pre�o Unit.").Text), "##,###,##0.000"))
    Else
      '.Pre�o = CDbl(Grade1.Columns("Pre�o Unit.").Text)
      .Pre�o = CDbl(Format((Grade1.Columns("Pre�o Unit.").Text), "##,###,##0.00"))
    End If
    .Pre�o_Total = CDbl(Grade1.Columns("Total").Text)
    .Desconto = CDbl(Grade1.Columns("Desconto").Text)
    .ICM = CDbl(Grade1.Columns("ICM").Text)
    .IPI = CDbl(Grade1.Columns("IPI").Text)
    .Etiqueta = Grade1.Columns("Etiq").Text
    .Pre�o_Final = CDbl(Grade1.Columns("Pre�o Final").Text)
    .Base_ICM = CDbl(Grade1.Columns("Base_ICM").Text)
    .Valor_ICM = CDbl(Grade1.Columns("Valor_ICM").Text)
    .Valor_Base_Unit = CDbl(Grade1.Columns("Valor_Base_Unit").Text)
     '19/12/2006 - Anderson - Altera��o realizada para evitar o erro de 13 - Type Mismacth
    '.Redu��o_ICM = CDbl(Grade1.Columns("Redu��o_ICM").Text)
     .Redu��o_ICM = CDbl("0" & Grade1.Columns("Redu��o_ICM").Text)
    .Tipo_ICM = Grade1.Columns("Tipo_ICM").Text
    .Descr_Adicional = Grade1.Columns("Descri��o Adicional").Text
    .CFOP_Produto = Grade1.Columns("CFOP").Text 'Altera��o para o Registro de CFOP por Produto
    .Valor_Desonerado = Grade1.Columns("Valor Desonerado").Text
    .Percentual_Diferimento = Grade1.Columns("% Diferimento").Text
    
    If IsNull(Grade1.Columns("Desp_Acessorias").Text) Or Grade1.Columns("Desp_Acessorias").Text = "" Then
      .Desp_Acessorias = "0"
    Else
      .Desp_Acessorias = Grade1.Columns("Desp_Acessorias").Text
    End If
  End With
  
  Exit Sub
Erro:
    MsgBox "Aten��o...ocorreu uma simples inconsist�ncia na sele��o do produto, tente novamente", vbInformation

End Sub

Private Sub L_Base_ICM_GotFocus()
  Call SelectAllText(L_Base_ICM)
End Sub

Private Sub L_Base_ICM_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub L_Base_ICM_Subs_GotFocus()
  Call SelectAllText(L_Base_ICM_Subs)
End Sub

Private Sub L_Base_ICM_Subs_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub L_Base_ICM_Subs_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Base_ICM_Subs)
End Sub

Private Sub L_Base_ICM_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Base_ICM)
End Sub


Private Sub L_Frete_GotFocus()
  Call SelectAllText(L_Frete)
End Sub

Private Sub L_Frete_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub L_Frete_LostFocus()
 Call Recalcula
End Sub

Private Sub L_Frete_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Frete)
End Sub

Private Sub L_Tot_Desc_GotFocus()
  Call SelectAllText(L_Tot_Desc)
End Sub

Private Sub L_Tot_Desc_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub L_Tot_Desc_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Tot_Desc)
End Sub

Private Sub L_Tot_IPI_GotFocus()
  Call SelectAllText(L_Tot_IPI)
End Sub

Private Sub L_Tot_IPI_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub L_Tot_IPI_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Tot_IPI)
End Sub

Private Sub L_Tot_ISS_GotFocus()
  Call SelectAllText(L_Tot_ISS)
End Sub

Private Sub L_Tot_ISS_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub


Private Sub L_Tot_ISS_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Tot_ISS)
End Sub

Private Sub L_Tot_Pagar_GotFocus()
  Call SelectAllText(L_Tot_Pagar)
End Sub

Private Sub L_Tot_Pagar_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub L_Tot_Pagar_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Tot_Pagar)
End Sub


Private Sub L_Tot_Prod_GotFocus()
  Call SelectAllText(L_Tot_Prod)
End Sub

Private Sub L_Tot_Prod_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub L_Tot_Prod_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Tot_Prod)
End Sub


Private Sub L_Tot_Serv_GotFocus()
  Call SelectAllText(L_Tot_Serv)
End Sub

Private Sub L_Tot_Serv_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub


Private Sub L_Tot_Serv_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Tot_Serv)
End Sub

Private Sub L_Valor_ICM_GotFocus()
  Call SelectAllText(L_Valor_ICM)
End Sub

Private Sub L_Valor_ICM_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub L_Valor_ICM_Subs_GotFocus()
  Call SelectAllText(L_Valor_ICM_Subs)
End Sub

Private Sub L_Valor_ICM_Subs_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub FindNextOrcam()
  Dim Seq As Variant
  Dim Cli As Long
  
  Call StatusMsg("")
  If Nome_Cliente.Caption = "" Then
    Beep
    DisplayMsg "Selecione um cliente antes."
    cboCliente.SetFocus
    Exit Sub
  End If
  
  Seq = gsHandleNull(txtSeq.Text & "")
  If Not IsNumeric(Seq) Then Seq = 0
  
Lp1:
  rsSaidas.FindFirst "Cliente = " & gsHandleNull(cboCliente.Text) & " And Sequ�ncia > " & Seq
  If Not rsSaidas.NoMatch Then
    Seq = rsSaidas("Sequ�ncia")
    rsOp_Sa�da.Index = "C�digo"
    rsOp_Sa�da.Seek "=", rsSaidas("Opera��o")
    If rsOp_Sa�da.NoMatch Then GoTo Lp1
    If rsOp_Sa�da("Tipo") <> "O" Then GoTo Lp1
    Call ShowRecord
  Else
    DisplayMsg "N�o existem outros or�amentos para este cliente."
  End If
  
End Sub

Private Sub AddNovoCliente()
  Dim F As Form
  Set F = New frmCliFor
  F.Show
End Sub

Private Sub CancelNota()
 Dim Resposta As Integer
 Dim Texto As String

 Call StatusMsg("")
 If IsNull(Num_Registro) Then
   DisplayMsg "Encontre uma movimenta��o antes."
   Exit Sub
 End If
 
 If rsSaidas("NFCe") > 0 And rsSaidas("Nota Cancelada") = False Then
    Dim CancelaNFCe As New clsNFCe

    Call StatusMsg("Aguarde, cancelando nota...")

    CancelaNFCe.CancelaNFCe (txtSeq.Text)

    If gsRetornoDoc <> "OK" Then
        Exit Sub
    End If
    
    rsSaidas.Edit
    rsSaidas("Nota Cancelada") = True
    rsSaidas.Update
    
    db.Execute "Delete * From Comiss�o WHERE Sequ�ncia = " & txtSeq.Text & " AND Filial = " & gnCodFilial & ""
    
    
    DisplayMsg "Pedido de Cancelamento de NFCe feito com sucesso"
    Exit Sub
 End If
 
'''''' If rsSaidas("Nota Impressa") = 0 Then
''''''   DisplayMsg "N�o foi emitida nota fiscal para esta movimenta��o. Imposs�vel cancelar. "
''''''   Exit Sub
'''''' End If
''''''
'''''' If rsSaidas("Nota Cancelada") = True Then
''''''   DisplayMsg "Esta nota j� foi cancelada."
''''''   Exit Sub
'''''' End If
 
'''''' Texto = "Nota fiscal: " + str(rsSaidas("Nota Impressa")) + Chr(13)
'''''' Texto = Texto + "Ap�s cancelar uma nota voc� N�O poder� mais desfazer ou gravar qualquer altera��o na movimenta��o. Deseja cancelar esta nota ?"
'''''' Resposta = MsgBox(Texto, vbYesNo + vbQuestion, "Aten��o")
'''''' If Resposta = vbNo Then
''''''   StatusMsg "Nota n�o cancelada."
''''''   Exit Sub
'''''' End If
''''''
'''''' Call StatusMsg("Aguarde, cancelando nota...")
 
 db.Execute "Delete * From Comiss�o WHERE Sequ�ncia = " & txtSeq.Text & " AND Filial = " & gnCodFilial & ""
 
'''''' rsSaidas.Edit
'''''' rsSaidas("Nota Cancelada") = True
'''''' rsSaidas.Update
''''''
'''''' DisplayMsg "Nota cancelada."
   
End Sub

Private Sub SearchProduto()
  Dim F As Form
  Call StatusMsg("Aguarde....")
  '31/08/2006 - Anderson
  'Implementa��o de pesquisa avan�ada na tela de consulta do produto
  'Set F = New frmConsultaProd
  'F.Show vbModal
  
  If gbMostrarTelaPesquisaProdutoTipoFoto = True Then
      frmConsultaProd.Show
  Else
      frmPesquisaProduto.Show
  End If
  
  'Set F = Nothing
  'Call StatusMsg("")
End Sub

'27/03/2006 - mpdea
'Implementado tratamento de erro
Public Sub SearchRecord()
  Dim lngSequencia As Long

  
  On Error GoTo ErrHandler
  
  
  If Not IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Apague todos os campos da tela com o bot�o NOVO."
    gsMsg = gsMsg & vbCrLf & "Preencha para a pesquisa uma ou mais das seguintes informa��es:"
    '15/05/2013-Alexandre Afornali
    gsMsg = gsMsg & vbCrLf & "Opera��o, Digitador, Cliente, Seq��ncia, Nota Fiscal, Refer�ncia, Comanda"
    gsMsg = gsMsg & vbCrLf & "E pressione novamente este bot�o PROCURAR."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  gsWhere = ""
  If Len(cboOper.Text) > 0 Then
    gsWhere = gsWhere & " AND Opera��o = " & cboOper.Text
  End If
  
  If Len(cboDigitador.Text) > 0 Then
    gsWhere = gsWhere & " AND Digitador = " & cboDigitador.Text
  End If
  
  If Len(cboCliente.Text) > 0 Then
    gsWhere = gsWhere & " AND Cliente = " & cboCliente.Text
  End If
  
  If Len(txtSeq.Text) > 0 Then
    '27/03/2006 - mpdea
    'Implementado valida��o de dados
    If Not IsDataType(dtLong, txtSeq.Text, lngSequencia) Then
      DisplayMsg "N�mero de sequ�ncia para pesquisa inv�lida."
      Exit Sub
    End If
    gsWhere = gsWhere & " AND Sequ�ncia = " & lngSequencia
  End If
  
  If Len(txtNF.Text) > 0 Then
    '20/05/2005 - Daniel
    'Tratamento para nota criadas manualmente
    If MsgBox("A nota foi impressa automaticamente (padr�o do Quick) ?", vbQuestion + vbYesNo) = vbYes Then
      gsWhere = gsWhere & " AND [Nota Impressa] Like '" & txtNF.Text & "*'"
    Else
      gsWhere = gsWhere & " AND [Nota Fiscal] Like '" & txtNF.Text & "*'"
    End If
  End If
  
  If Len(txtRef.Text) > 0 Then
    gsWhere = gsWhere & " AND Refer�ncia Like '" & txtRef.Text & "*'"
  End If
  
  If (txtComanda.Text <> "") Then
    Dim rsComanda As Recordset
    Set rsComanda = db.OpenRecordset("SaidasComandas")
    rsComanda.MoveFirst
    While Not rsComanda.EOF
      If (rsComanda("CodComanda") = txtComanda.Text And rsComanda("Filial") = gnCodFilial) Then
        gsWhere = gsWhere & " AND [Sequ�ncia] = " & rsComanda("CodSaida")
        rsComanda.MoveLast
      End If
      rsComanda.MoveNext
    Wend
  End If
  '20/03/2013-Alexandre Afornali
  'Filtra apenas parametros informados
  Set rsSaidas = db.OpenRecordset(gsSql & gsWhere & gsOrder, dbOpenDynaset)
  'Set rsSaidas = db.OpenRecordset(gsSql & gsOrder, dbOpenDynaset)
  
  '24/11/2006 - Anderson
  'Essa altera��o permite que o registro seja localizado e que a navega��o seja permitida tanto para o registro anterior, como para o posterior
  If Len(gsWhere) <> 0 Then
    rsSaidas.FindFirst Mid(gsWhere, 5)
  End If
  
  If Not rsSaidas.NoMatch Then
    If Not rsSaidas.EOF Then
      Call ShowRecord
    End If
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em fun��o dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    '20/03/2013-Alexandre Afornali
    Set rsSaidas = db.OpenRecordset(gsSql & gsOrder, dbOpenDynaset)
    '18/02/2005 - Daniel
    Exit Sub
    '-------------------
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Public Sub SearchRecord_peloNumSeq(ByVal Num As Long)
  Dim lngSequencia As Long

  On Error GoTo ErrHandler

  gsWhere = ""

  If Num > 0 Then
      gsWhere = gsWhere & " AND Sequ�ncia = " & Num
  Else
      DisplayMsg "N�mero de sequ�ncia para pesquisa inv�lida."
      Exit Sub
  End If

  Set rsSaidas = db.OpenRecordset(gsSql & gsWhere & gsOrder, dbOpenDynaset)

  'Essa altera��o permite que o registro seja localizado e que a navega��o seja permitida tanto para o registro anterior, como para o posterior
  If Len(gsWhere) <> 0 Then
    rsSaidas.FindFirst Mid(gsWhere, 5)
  End If

  If Not rsSaidas.NoMatch Then
    If Not rsSaidas.EOF Then
      Call ShowRecord
    End If
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em fun��o dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Set rsSaidas = db.OpenRecordset(gsSql & gsOrder, dbOpenDynaset)
    Exit Sub
  End If
  
  Exit Sub

ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Public Sub SearchRecord_peloNumComanda(ByVal Num As String)
  Dim lngSequencia As Long
  Dim sSQlComanda As String
  Dim rsSaidasComandas As Recordset

  On Error GoTo ErrHandler
  
'  If Not IsNumeric(Num) Then
'      DisplayMsg "N�mero de comanda para pesquisa inv�lida."
'      Exit Sub
'  End If

  'If Num > 0 Then
      sSQlComanda = "SELECT CodSaida FROM SaidasComandas WHERE CodComanda = '" & Num & "'"
      sSQlComanda = sSQlComanda & " And Filial = " & gnCodFilial & ""
  'Else
  '    DisplayMsg "N�mero de sequ�ncia para pesquisa inv�lida."
  '    Exit Sub
  'End If

  Set rsSaidasComandas = db.OpenRecordset(sSQlComanda, dbOpenDynaset)
  
  If Not (rsSaidasComandas.EOF And rsSaidasComandas.BOF) Then
      rsSaidasComandas.MoveFirst
      Num = rsSaidasComandas.Fields("CodSaida").Value
      rsSaidasComandas.Close
      Set rsSaidasComandas = Nothing
  Else
      rsSaidasComandas.Close
      Set rsSaidasComandas = Nothing
      DisplayMsg "N�o existe uma venda relacionada a este n�mero de comanda."
      Exit Sub
  End If

  SearchRecord_peloNumSeq Num
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub



'20/04/2006 - mpdea
'Implementado tratamento de erro e transa��o
Private Sub UndoMovimSaida()
  Dim nRet As Integer
  Dim nMoviment As Long
  
  Dim blnInTransaction As Boolean
  Dim lngSequenciaPai As Long
  Dim strSQL As String
  Dim rstSaidaProdutosPai As Recordset
  Dim rstSaidaProdutos As Recordset
  
  
  On Error GoTo ErrHandler
  
  If rsSaidas("Data") < CDate(Data_Atual) Then
    If MsgBox("Esta movimenta��o N�O foi realizada hoje. Pode inclusive estar fora do prazo legal de cancelamento." & _
        " Mesmo assim deseja desfazer a movimenta��o e cancelar a Nota Fiscal?", vbExclamation + vbYesNo, "Aten��o") = vbNo Then
        Exit Sub
    End If
  End If
  
  Call StatusMsg("")
  If IsNull(Num_Registro) Then
      DisplayMsg "Encontre uma sa�da antes."
  ElseIf Not rsSaidas("Efetivada") Then
      DisplayMsg "Esta opera��o n�o foi efetivada."
  ElseIf rsSaidas("Movimenta��o Desfeita") Then
      DisplayMsg "Esta movimenta��o j� foi desfeita."
  ElseIf rsSaidas("Nota Cancelada") And rsSaidas("Movimenta��o Desfeita") Then
      DisplayMsg "Esta Movimenta��o j� foi desfeita e a Nota Fiscal j� foi cancelada."
'  ElseIf rsSaidas("Data") < CDate(Data_Atual) Then
'     DisplayMsg "ATEN��O" & Chr(13) & Chr(13) & "Esta movimenta��o N�O foi feita hoje e " & _
'       "por isso N�O PODE SER DESFEITA." & Chr(13) & Chr(13) & "Se desejar ajuste o " & _
'       "estoque dos produtos e contas a receber manualmente."
  Else
    If rsSaidas("Nota Impressa") <> 0 Then
        If MsgBox("Esta movimenta��o j� teve a Nota Fiscal impressa." & _
            " Deseja desfazer a movimenta��o e cancelar a Nota Fiscal?" _
            , vbExclamation + vbYesNo, "Aten��o") = vbNo Then
            Exit Sub
        End If
    End If
    'If rsSaidas("Cupom Fiscal Impresso") = True Then
      'MsgBox ("Esta movimenta��o j� teve o cupom fical impresso, por isso n�o pode ser desfeita")
      'Exit Sub
    'End If
    'Senha do gerente
    If Not frmGerente.gbSenhaGerente Then
        Exit Sub
    End If
    nMoviment = Val(txtSeq.Text)
    
    ws.BeginTrans: blnInTransaction = True
    nRet = Desefetiva_Sa�da(gnCodFilial, nMoviment)
    If nRet <> 0 Then
        ws.Rollback: blnInTransaction = False
        DisplayMsg "Erro n�" & CStr(nRet) & " ao desfazer movimenta��o de sa�da."
        Exit Sub
    End If
    
    
    '--------------------------------------------------------------------------
    '20/04/2006 - mpdea
    'Verifica se a movimenta��o pertence a movimenta��o de entrega
    'e restaura a quantidade
    If IsDataType(dtLong, rsSaidas.Fields("Sequ�nciaPai").Value, lngSequenciaPai) Then
      If lngSequenciaPai > 0 Then
        
        'Sele��o dos produtos da movimenta��o
        strSQL = "SELECT Qtde, [C�digo sem Grade] "
        strSQL = strSQL & "FROM [Sa�das - Produtos] "
        strSQL = strSQL & "WHERE Filial = " & gnCodFilial
        strSQL = strSQL & " AND Sequ�ncia = " & nMoviment
        
        Set rstSaidaProdutos = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      
        'Verifica os produtos da movimenta��o
        With rstSaidaProdutos
          If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
              'Conforme modelo de entregas, pesquisa origem
              strSQL = "SELECT QtdeEntregue "
              strSQL = strSQL & "FROM [Sa�das - Produtos] "
              strSQL = strSQL & "WHERE Filial = " & gnCodFilial
              strSQL = strSQL & " AND Sequ�ncia = " & lngSequenciaPai
              strSQL = strSQL & " AND [C�digo sem Grade] = '" & .Fields("C�digo sem Grade").Value & "'"
              
              Set rstSaidaProdutosPai = db.OpenRecordset(strSQL, dbOpenDynaset)
              With rstSaidaProdutosPai
                If Not (.BOF And .EOF) Then
                  .Edit
                  'Diminui da qtde entregue a qtde de produtos da movimenta��o
                  .Fields("QtdeEntregue").Value = .Fields("QtdeEntregue").Value - rstSaidaProdutos.Fields("Qtde").Value
                  .Update
                End If
                .Close
              End With
              .MoveNext
            Loop
          End If
          .Close
        End With
        
        Set rstSaidaProdutos = Nothing
        Set rstSaidaProdutosPai = Nothing
      End If
    End If
    '--------------------------------------------------------------------------
    
    
    If rsSaidas("Nota Impressa") = 0 Then
      Call StatusMsg("Aguarde...")
      '30/01/2017 Jean
      'Comentei a parte do c�digo responsav�l pela exclus�o da venda e de suas dependencias ao desfazer a movimenta��o a pedido do Mauro
      'Apaga movimenta��o de Sa�das
'      Call EraseTypeMoviment(tmSaidas, gnCodFilial, nMoviment)
'      Call EraseTypeMoviment(tmSaidasProdutos, gnCodFilial, nMoviment)
'      Call EraseTypeMoviment(tmSaidasServicos, gnCodFilial, nMoviment)
'      Call EraseTypeMoviment(tmMovimentoCheques, gnCodFilial, nMoviment)
'      Call EraseTypeMoviment(tmMovimentoParcelas, gnCodFilial, nMoviment)
'      txtSeq.Text = ""
'      Num_Registro = Null
'      L_Efetivada.Visible = False
      rsSaidas.Edit
      rsSaidas("Movimenta��o Desfeita") = True
      rsSaidas.Update
    Else
      rsSaidas.Edit
      rsSaidas("Nota Cancelada") = True
      rsSaidas("Movimenta��o Desfeita") = True
      rsSaidas.Update
    End If
    
    '20/04/2006 - mpdea
    'Somente agora finaliza transa��o e n�o como antes, quando havia opera��es de bd pendentes
    ws.CommitTrans: blnInTransaction = False
    
    ' Pilatti Outubro/17
    ' Caso seja NFCe
    CancelNota
    
    Call StatusMsg("")
    DisplayMsg "Opera��o desfeita."
  End If
  
  Exit Sub
  
ErrHandler:
  If blnInTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub GetInformation()
 
  Call StatusMsg("")
  
  If IsNull(Nome_Cliente.Caption) Then Exit Sub
  If Nome_Cliente.Caption = "" Then Exit Sub
  If IsNull(cboCliente.Text) Then Exit Sub
  If cboCliente.Text = "" Then Exit Sub
  
  If Not IsNumeric(cboCliente.Text) Then Exit Sub
  If Val(cboCliente.Text) = 0 Then Exit Sub
  
  gsCodCliente = Val(cboCliente.Text)
  
  '20/04/2006 - mpdea
  'Modificado exibi��o do form de informa��es do cliente
  'para poder ser acess�vel de diversas maneiras (Ex.: VR CheckOut)
  frmInformacoes.Show , frmMain
 
End Sub


'20/11/2002 - mpdea
'Geral - Realizado diversas modifica��es e comentado alguns c�digos
Private Sub TransformaOrcamEmVenda()
'  Dim Sai_Loop As Integer
  
  Dim Num_Registro_temp As Variant
  Dim intCodOper_temp As Integer
  Dim strPWD_temp As String
  Dim intRet As Integer
  Dim blnInTransaction As Boolean
  Dim strAux As String
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")

  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Encontre um or�amento antes."
    Exit Sub
  End If
  
  If rsOp_Sa�da("Tipo") <> "O" Then
    Beep
    DisplayMsg "N�o � um or�amento."
    Exit Sub
  End If
  
  If rsOp_Sa�da("ExigeAprovacaoOrcamento") And (Not (rsSaidas.Fields("OrcamentoAprovado"))) Then
    If Not frmGerente.gbSenhaGerente Then
      Exit Sub
    End If
  End If
  
  '28/11/2002 - mpdea
  'Verifica se foi configurado a op. de sa�das em Par�metros
  If CInt("0" & rsParametros.Fields("OpSaidaOrcVenda").Value) = 0 Then
    DisplayMsg "Opera��o de Sa�da a ser utilizada na transforma��o de Or�amento em Venda deve ser configurada em Par�metros, pasta Sa�das."
    Exit Sub
  End If
  
  
  If MsgBox("Esta opera��o n�o poder� ser desfeita, deseja realmente transformar este or�amento em uma venda? ", vbQuestion + vbYesNo, "Aten��o") = vbNo Then
    Exit Sub
  End If
  
  Call StatusMsg("Aguarde, alterando opera��o...")
  
  
'  Rem Apaga produtos
'  rsSaidas_Prod.Index = "Sequ�ncia"
'  Sai_Loop = False
'
'  Do
'   rsSaidas_Prod.Seek ">", gnCodFilial, Val(txtSeq.Text)
'
'   If rsSaidas_Prod.NoMatch Then Sai_Loop = True
'   If Sai_Loop = False Then If rsSaidas_Prod("Filial") <> gnCodFilial Then Sai_Loop = True
'   If Sai_Loop = False Then If rsSaidas_Prod("Sequ�ncia") <> Val(txtSeq.Text) Then Sai_Loop = True
'
'   If Sai_Loop = False Then
'     rsSaidas_Prod.Delete
'   End If
'  Loop Until Sai_Loop = True
'
'  Rem Apaga Sa�das
'  rsSaidas.Delete
'  Num_Registro = Null
  
'  txtSeq.Text = ""
  
  
  
  '20/11/2002 - mpdea
  'Atualiza informa��es da tela com os dados da base
  If UCase(gstrGetEstadoFilial(gnCodFilial)) = "AM" Then
    strPWD_temp = Senha.Text
    Call ShowRecord
    Senha.Text = strPWD_temp
  End If
  
  'Atualiza data
  L_Dia.Caption = Format$(Data_Atual, "dd/mm/yyyy")
  L_Efetivada.Visible = False
  
  
  '----------------------------------------------------------------------------
  '20/11/2002 - mpdea
  'Obt�m c�digo da opera��o atual
  '
  '19/11/2002 - mpdea
  'Obt�m a opera��o cadastrada em Par�metros da Filial
  
'  cboOper.Text = 500
'  cboOper_LostFocus
  intCodOper_temp = rsSaidas.Fields("Opera��o").Value
  cboOper.Text = CInt("0" & rsParametros.Fields("OpSaidaOrcVenda").Value)
  cboOper_LostFocus
  '----------------------------------------------------------------------------
  
  
  '----------------------------------------------------------------------------
  '19/11/2002 - mpdea
  'Verifica��o para o estado do Amazonas que exige tratamento diferenciado
  'para o or�amento
  
  'Criar novo registro para o estado do Amazonas
  If UCase(gstrGetEstadoFilial(gnCodFilial)) = "AM" Then
    'Cria nova movimenta��o preservando a anterior
    
    If rsSaidas.Fields("Locked").Value Then
      'Retorna valores originais
      Call ShowRecord
      Senha.Text = strPWD_temp
      DisplayMsg "Or�amento bloqueado - Venda j� foi criada com este or�amento."
      Exit Sub
    End If
    
    'Informa��es tempor�rias
    Num_Registro_temp = Num_Registro
    Num_Registro = Null
    
    ws.BeginTrans
    blnInTransaction = True
    
    strAux = rsSaidas.Fields("InfoNrOrcamento").Value
        
    'Bloqueia or�amento atual
    With rsSaidas
      .Edit
      .Fields("Locked").Value = True
      '21/11/2002 - mpdea
      'Marca or�amento como efetivado
      .Fields("Efetivada").Value = True
      .Update
    End With
    
    'Grava registro
    intRet = UpdateRecord
    
    If intRet = 0 Then
      'Adiciona a informa��o de nr. do or�amento e bloqueia novo registro
      With rsSaidas
        .Edit
        .Fields("InfoNrOrcamento").Value = strAux
        .Fields("Locked").Value = True
        .Update
      End With
      ws.CommitTrans
      blnInTransaction = False
      
      'Atualiza informa��es da tela
      Call ShowRecord
      Senha.Text = strPWD_temp
      
      DisplayMsg "Venda criada. Realizar recebimento."
    Else
      ws.Rollback
      blnInTransaction = False
      'Retorna valores originais
      Num_Registro = Num_Registro_temp
      'Posiciona registro
      rsSaidas.Bookmark = Num_Registro
      'Atualiza exibi��o dos dados
      Call ShowRecord
      Exit Sub
    End If
  '----------------------------------------------------------------------------
    
  Else
    DisplayMsg "Venda criada. Verifique o c�digo da opera��o, os produtos, quantidades e pre�o. Grave a opera��o e n�o se esque�a do recebimento."
  End If
  
  Exit Sub
  
ErrHandler:
  If blnInTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub L_Valor_ICM_Subs_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Valor_ICM_Subs)
End Sub

Private Sub L_Valor_ICM_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(L_Valor_ICM)
End Sub

Private Sub mskDataEmissaoNotaManual_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataEmissaoNotaManual.Text = frmCalendario.gsDateCalender(mskDataEmissaoNotaManual.Text)
  End If
End Sub

Private Sub mskDataEmissaoNotaManual_LostFocus()
  mskDataEmissaoNotaManual.Text = Ajusta_Data(mskDataEmissaoNotaManual.Text)
End Sub

Private Sub mskValidade_KeyDown(KeyCode As Integer, Shift As Integer)
'A tecla est� pressionada para baixo
  If KeyCode = vbKeyF2 Then
    mskValidade.Text = frmCalendario.gsDateCalender(mskValidade.Text)
  End If
End Sub

Private Sub mskValidade_LostFocus()
  mskValidade.Text = Ajusta_Data(mskValidade.Text)
End Sub

Private Sub Obs_GotFocus()
  Call SelectAllText(Obs)
End Sub


Private Sub Senha_GotFocus()
  Call SelectAllText(Senha)
End Sub

Private Sub Tab1_GotFocus()
  Call Tab1_Click(Tab1.Tab)
End Sub

Private Sub txtComanda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 9 Or KeyAscii = 10 Then
    btnComandaVendas.Visible = False
    txtComanda.Width = txtSeq.Width

    If Trim(txtComanda.Text) <> "" Then
      frmComanda.Comanda = Trim(txtComanda.Text) 'BBB123 e AA01
      If frmComanda.ComandaOK Then
        If frmComanda.Total > 0 Then
          If frmComanda.Sequencia > 0 Then
            SearchRecord_peloNumSeq frmComanda.Sequencia
          Else
            txtComanda.Width = 24.007
            btnComandaVendas.Visible = True
          End If
        End If
      End If
    End If
    'SearchRecord_peloNumComanda (txtComanda.Text)
  ElseIf ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  ElseIf KeyAscii <> 8 And KeyAscii <> 127 Then
    KeyAscii = 0
  End If
  Refresh
End Sub

Private Sub txtComanda_LostFocus()
  txtComanda_KeyPress (13)
End Sub

Private Sub txtDescSubTotal_GotFocus()
  Call SelectAllText(txtDescSubTotal)
End Sub

Private Sub txtNrSerieNF_LostFocus()
  '20/05/2005 - Daniel
  txtNrSerieNF.Text = UCase(txtNrSerieNF.Text & "")
End Sub

Private Sub txtRef_GotFocus()
  Call SelectAllText(txtRef)
End Sub

Private Sub txtRef_LostFocus()
  If IsNull(txtRef.Text) Then Exit Sub
  If txtRef.Text = "" Then Exit Sub
  txtRef = UCase(txtRef)
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
  Dim grdAux As SSDBGrid

  If PreviousTab = 0 Then
    Set grdAux = Grade1
  Else
    Set grdAux = Grade_Serv
  End If

  With grdAux
    .MoveLast
    .MoveFirst
    .Scroll -99, -99
    '-------------------------------------------------------------'
    ' 01/10/2002 - Maikel Cordeiro                                '
    ' O set focus estava causando erro no execut�vel...           '
    ' - OBS: apenas no execut�vel, em projeto o erro n�o ocorre   '
    '-------------------------------------------------------------'
    '    .SetFocus                                                '
    '-------------------------------------------------------------'
  End With
  Set grdAux = Nothing
'  DoEvents
'  SendKeys "{Home}+{End}"
End Sub

Private Sub txtNF_GotFocus()
'  If Not IsNull(Num_Registro) Then
'    SendKeys "{tab}"
'  End If
  Call SelectAllText(txtNF)
End Sub

Private Sub RecalculaPrecos()
'  Dim nRow As Long
'  Dim sCodProd As String
'  Dim bm As Variant
  
'  Screen.MousePointer = vbHourglass
'  Call StatusMsg("Refazendo tabela..."
'
'  rsPre�os.Index = "Tabela"
'  For nRow = 0 To Linhas_Grade - 1
'    sCodProd = gsHandleNull(Tabe(nRow).C�digo)
'    If sCodProd <> "0" Then
'      rsPre�os.Seek "=", Combo_Pre�o.Text, sCodProd
'      If rsPre�os.NoMatch Then
'        Tabe(nRow).Pre�o = 0#
'      Else
'        Tabe(nRow).Pre�o = rsPre�os("Pre�o")
'      End If
'    Else
'      Tabe(nRow).Pre�o = 0#
'    End If
'    Call Calcula_Linha_Tabe(nRow)
'  Next nRow
'
''  Grade1.MoveLast
''  Grade1.MoveFirst
  
  Call Recalcula
  
'  Screen.MousePointer = vbDefault
'  Call StatusMsg("")

End Sub

Private Sub Calcula_Linha_Tabe(ByVal nRow As Long)
  Dim Qtde            As Double
  Dim Pre�o           As Double
  Dim Desconto        As Double
  Dim Valor_Desconto  As Double
  Dim IPI             As Double
  Dim Valor_IPI       As Double
  Dim Pre�o_Total     As Double
  Dim Pre�o_Final     As Double
  Dim Desp_Acessorias As Double
  
  Qtde = Tabe(nRow).Qtde
  '04/05/2004 - Daniel
  'Personaliza��o Embalavi
  If g_bln5CasasDecimais Then
    Pre�o = Format((Tabe(nRow).Pre�o), "##,###,##0.00000")
  '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    Pre�o = Format((Tabe(nRow).Pre�o), "##,###,##0.000")
  Else
    'Pre�o = Tabe(nRow).Pre�o
    Pre�o = Format((Tabe(nRow).Pre�o), "##,###,##0.00")
  End If
  
  'If Tabe(nRow).Desp_Acessorias = "" Then
  ' Tabe(nRow).Desp_Acessorias = 0
  'End If
  
  Desp_Acessorias = Tabe(nRow).Desp_Acessorias
  
  'Desp_Acessorias = Format((Tabe(nRow).Desp_Acessorias"), "#0.00")
  
  Desconto = Tabe(nRow).Desconto
  IPI = Tabe(nRow).IPI
  
  Pre�o_Total = Format(Qtde * Pre�o, "#0.00")
  Valor_Desconto = Pre�o_Total * Desconto / 100
  Pre�o_Final = (Pre�o_Total - Valor_Desconto)
  Valor_IPI = Pre�o_Final * IPI / 100
  
  If Not Calcula_IPI Then
    Valor_IPI = 0
  End If
  
  If Not Calcula_IPI_TOT Then
    Pre�o_Final = Pre�o_Final + Valor_IPI
  End If
  
  Tabe(nRow).Pre�o_Total = Pre�o_Total
  Tabe(nRow).Pre�o_Final = Pre�o_Final
  Tabe(nRow).Total_Valor_Desonerado = Total_Valor_Desonerado
End Sub

Private Sub RecalculaPesos()
  Dim sCodProd As String
  Dim sCod As String
  Dim nRow As Long
'  Dim bm As Variant
  Dim Aux_Produto As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edi��o As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim nQtde As Single
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Refazendo tabela...")
  
  Grade1.Update
  Grade1.MoveFirst
  gnPesoLiquido = 0#
  gnPesoBruto = 0#
  
'  For nRow = 0 To (Linhas_Grade - 1)
  For nRow = LBound(Tabe) To UBound(Tabe)
'    bm = Grade1.GetBookmark(nRow)
'    Grade1.Bookmark = bm
    sCodProd = gsHandleNull(Tabe(nRow).C�digo)
    If sCodProd <> "0" Then  'Faz somente os preenchidos
'      nQtde = Grade1.Columns(1).CellText(bm)
      nQtde = Tabe(nRow).Qtde
      sCod = sCodProd
      Call Acha_Produto(sCod, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_Tipo, Aux_Erro)
      sCod = Aux_Produto
      If Aux_Erro = 0 Then
        rsProdutos2.FindFirst "C�digo = '" & Aux_Produto & "'"
        gnPesoLiquido = gnPesoLiquido + nQtde * gsHandleNull(rsProdutos2("PesoLiquido"))
        gnPesoBruto = gnPesoBruto + nQtde * gsHandleNull(rsProdutos2("PesoBruto"))
      End If
    End If
  Next nRow
  
  Grade1.MoveFirst
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
End Sub

Private Sub txtSeguro_GotFocus()
  Call SelectAllText(txtSeguro)
End Sub

Private Sub txtSeguro_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub txtSeguro_LostFocus()
  Call Recalcula
End Sub

Private Sub txtSeguro_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(txtSeguro)
End Sub

Private Sub txtSeq_GotFocus()
  Call SelectAllText(txtSeq)
End Sub

Private Sub txtSeq_KeyPress(KeyAscii As Integer)
    If IsNumeric(txtSeq.Text) = True Then
        If KeyAscii = 13 Then
            SearchRecord_peloNumSeq (txtSeq.Text)
        End If
    End If
End Sub

Private Sub txtSubTotal_GotFocus()
  Call SelectAllText(txtSubTotal)
End Sub

'24/10/2002 - mpdea
'Corrigido verifica��o de estoque para produtos com grade/edi��o
'
'09/10/2002 - mpdea
'Fun��o que verifica o estoque dos produtos
Private Function mblnCheckStock() As Boolean
  Dim intRow As Integer
  Dim intX As Integer
  Dim strCodProdPrincipal As String
  Dim intTamanho As Integer
  Dim intCor As Integer
  Dim lngEdicao As Long
  Dim intErr As Integer
  Dim intCountItem As Integer
  Dim blnNewItem As Boolean
  Dim dblEstoque As Double
  Dim blnStockInsufficient As Boolean

  'Tabela para controle dos produtos a serem verificados o estoque
  intCountItem = 0
  ReDim typCheckStock(intCountItem) As CheckStock

  For intRow = LBound(Tabe) To UBound(Tabe)
    With Tabe(intRow)

      If .C�digo <> "" And .Qtde <> 0 And .Nome <> "" Then

        Call Acha_Produto(.C�digo, strCodProdPrincipal, 0, 0, 0, 0, intErr)

        Select Case intErr
          Case 1
            DisplayMsg "Produto [" & .C�digo & "] n�o existe."
          Case 2
            DisplayMsg "Produto [" & .C�digo & "] com grade sem tamanho e cor informados."
          Case 3
            DisplayMsg "Produto [" & .C�digo & "] com edi��o sem edi��o informada."
        End Select
        If intErr <> 0 Then Exit Function

        rsProdutos2.FindFirst "C�digo = '" & strCodProdPrincipal & "'"
        If rsProdutos2.NoMatch Then
          DisplayMsg "Produto [" & .C�digo & "] n�o existe."
          Exit Function
        End If

        'Verifica se o produto possui estoque controlado
        If rsProdutos2.Fields("Estoque").Value Then

          blnNewItem = True
          For intX = LBound(typCheckStock) To UBound(typCheckStock)
            'Item j� listado
            If typCheckStock(intX).strCode = .C�digo Then
              typCheckStock(intX).dblQuantity = _
              typCheckStock(intX).dblQuantity + .Qtde
              blnNewItem = False
              Exit For
            End If
          Next intX

          'Novo item (agrupa)
          If blnNewItem Then
            ReDim Preserve typCheckStock(intCountItem) As CheckStock
            typCheckStock(intCountItem).strCode = .C�digo
            typCheckStock(intCountItem).dblQuantity = .Qtde
            intCountItem = intCountItem + 1
            intX = intCountItem
          End If

        End If

      End If

    End With
  Next intRow

  'Fim da cria��o da lista

  'In�cio da verifica��o do estoque
  For intX = LBound(typCheckStock) To UBound(typCheckStock)

    With typCheckStock(intX)
      
      '12/11/2002 - mpdea
      'Adicionado valida��o dos itens no array
      If .strCode <> "" And .dblQuantity <> 0 Then
      
        Call Acha_Produto(.strCode, strCodProdPrincipal, intTamanho, intCor, _
                          lngEdicao, 0, 0)
  
        'Estoque atual
        dblEstoque = -999999
        dblEstoque = Acha_Estoque(gnCodFilial, strCodProdPrincipal, intTamanho, intCor, lngEdicao, intErr)
        
        .dblStock = dblEstoque
        
        If intErr = 0 And dblEstoque <> -999999 Then
          'Se o estoque atual for superior a quantidade a ser movimentada
          'ativa flag de estoque insuficiente
          If .dblQuantity > dblEstoque Then
            .blnStockInsufficient = True
            blnStockInsufficient = True
          End If
        Else
          '20/11/2002 - mpdea
          'Adicionado descri��o do erro 1
          If intErr = 1 Then
            DisplayMsg "Produto [" & .strCode & "] com estoque n�o inicializado."
          Else
            DisplayMsg "Erro [" & intErr & "] ao encontrar estoque do produto."
          End If
          Exit Function
        End If
      
      End If
      
    End With

  Next intX
  
  'Exibe os produtos com estoque insuficiente
  If blnStockInsufficient Then
    frmCheckStock.ShowStockInsufficient
    Exit Function
  End If
  
  'Todos os produtos possuem estoque para a movimenta��o
  mblnCheckStock = True

End Function

Private Function ValidaCampoValidade() As Boolean
'26/02/2004 - Daniel
'Case: PSV

  If Not IsDate(mskValidade.Text) Then
    MsgBox "Data da Validade da Reserva � inv�lida.", vbExclamation, "Quick Store"
    mskValidade.SetFocus
    ValidaCampoValidade = False
    Exit Function
  End If
  
  If CDate(mskValidade.Text) < Data_Atual Then
    MsgBox "Data da Validade da Reserva deve ser posterior a de hoje.", vbExclamation, "Quick Store"
    mskValidade.SetFocus
    ValidaCampoValidade = False
    Exit Function
  End If
  
  ValidaCampoValidade = True

End Function

Private Function VerificaSeExisteValidade() As Boolean
'Checaremos se a opera��o carregada na combo opera��es
'l� na tabela de Opera��es Sa�da o campo Validade est�
'como true
'Case: PSV Inform�tica
  Dim rstOperacoesSaida As Recordset
  Dim strSQL            As String
  
  strSQL = " SELECT C�digo, Validade "
  strSQL = strSQL & " FROM [Opera��es Sa�da] "
  strSQL = strSQL & " WHERE [Opera��es Sa�da].C�digo =" & CInt(Trim(cboOper.Text))
  
  Set rstOperacoesSaida = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstOperacoesSaida
    If Not (.BOF And .EOF) Then
      VerificaSeExisteValidade = .Fields("Validade").Value
    End If
    
    If .RecordCount = 0 Then VerificaSeExisteValidade = False
    
  End With
  
  rstOperacoesSaida.Close
  Set rstOperacoesSaida = Nothing

End Function

Private Function AtualizarTableProgramacao(ByVal lngNumAutorizacao As Long, intMesX As Integer) As String
  Dim rstProgramacao      As Recordset
  Dim strSQLProgramacao   As String
  
  strSQLProgramacao = "SELECT Faturado FROM Programacao "
  strSQLProgramacao = strSQLProgramacao & " WHERE [Num Autorizacao] = " & lngNumAutorizacao
  strSQLProgramacao = strSQLProgramacao & " AND MesX = " & intMesX
  
  Set rstProgramacao = db.OpenRecordset(strSQLProgramacao, dbOpenDynaset)

  With rstProgramacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Edit
      .Fields("Faturado").Value = True
      .Update
    End If
    .Close
  End With

  Set rstProgramacao = Nothing

End Function

'23/05/2006 - mpdea
'Comentado fun��o abaixo devido otimizado na verifica��o de cliente isento de IPI
'
'Private Function IsencaoIPI(ByVal CodCliente As Long) As Boolean
'  '06/05/2004 - Daniel
'  'Case: Embalavi
'  'Esta fun��o tem a finalidade de verificar na tabela Cli_For se o
'  'Cli_For � Isento de IPI
'  Dim rstCliFor As Recordset
'
'  Set rstCliFor = db.OpenRecordset("SELECT IsentoIPI FROM Cli_For WHERE C�digo = " & CodCliente, dbOpenDynaset)
'
'  With rstCliFor
'    If Not (.BOF And .EOF) Then
'      IsencaoIPI = .Fields("IsentoIPI").Value
'    End If
'    .Close
'  End With
'
'  Set rstCliFor = Nothing
'
'End Function

Private Function Diferimento(ByVal CodCliente As Long) As Boolean
  '06/05/2004 - Daniel
  'Case: Embalavi
  'Esta fun��o tem a finalidade de verificar na tabela Cli_For se o
  'estado do Cli_For � PR e se � pessoa jur�dica
  Dim rstCliFor As Recordset
  
  Set rstCliFor = db.OpenRecordset("SELECT Estado, F�sica_Jur�dica FROM Cli_For WHERE C�digo = " & CodCliente, dbOpenDynaset)

  With rstCliFor
    If Not (.BOF And .EOF) Then
      If .Fields("Estado").Value = "PR" And .Fields("F�sica_Jur�dica").Value = "J" Then
        Diferimento = True
      Else
        Diferimento = False
      End If
    End If
    .Close
  End With
  
  Set rstCliFor = Nothing

End Function

'03/06/2008 - mpdea
'Modificado de function para sub, pois n�o retorna valor
'Otimizado c�digo
Private Sub CalculaImpostosSobreServicos(ByVal TotaldeServicos As Double)
  '13/05/2004 - Daniel
  'Esta fun��o tem a finalidade de calcular os percentuais e totais
  'de impostos sobre servi�os: CSLL, COFINS, PIS, IRRF
  Dim rstParametros As Recordset
  Dim rstOpSaida As Recordset
  Dim dbl_base_calculo As Double
  Dim dbl_valor_isencao As Double
  
  '12/06/2008 - mpdea
  'Verifica se a movimenta��o foi efetivada
  If Not IsNull(Num_Registro) Then
    If rsSaidas.Fields("Efetivada").Value Then
      Exit Sub
    End If
  End If
 
  '03/06/2008 - mpdea
  'Sempre zera as vari�veis antes do c�lculo
  Call ZerarVarsImpostosServi�os
  
  If TotaldeServicos > 0 Then 'Primeiro passo verificamos se � Maior que zero
    '03/06/2008 - mpdea
    'Realiza a valida��o antes de prosseguir
    If Len(Nome_Opera��o.Caption) <= 0 Then
      DisplayMsg "Informe a Opera��o para o c�lculo correto de impostos sobre os servi�os."
      Exit Sub
    ElseIf Len(Nome_Cliente.Caption) <= 0 Then
      DisplayMsg "Informe o Cliente para o c�lculo correto de impostos sobre os servi�os."
      Exit Sub
    End If
    
    '03/06/2008 - mpdea
    'Inclu�do atributo somente leitura
    Set rstParametros = db.OpenRecordset("SELECT CSLL, COFINS, PIS, IRRF, ValorIsencaoPisCofinsCsll FROM [Par�metros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset, dbReadOnly)
    With rstParametros
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        'Verificamos o conte�do de CSLL, COFINS, PIS, IRRF em Par�metros
        If Not IsNumeric(.Fields("CSLL").Value) Or Not IsNumeric(.Fields("COFINS").Value) Or _
           Not IsNumeric(.Fields("PIS").Value) Or Not IsNumeric(.Fields("IRRF").Value) Then
          .Close 'rstParametros
          Set rstParametros = Nothing
          '03/06/2008 - mpdea
          'Exibe mensagem de alerta caso os valores estejam incorretos
          DisplayMsg "Percentuais de impostos sobre servi�os inv�lidos. Favor configurar em 'Par�metros da Empresa/Filial', aba 'Outros'."
          Exit Sub
        End If
            
        '03/06/2008 - mpdea
        'Inclu�do atributo somente leitura
        Set rstOpSaida = db.OpenRecordset("SELECT ComissaoServicos FROM [Opera��es Sa�da] WHERE C�digo = " & CInt(cboOper.Text), dbOpenDynaset, dbReadOnly)
        If Not rstOpSaida.Fields("ComissaoServicos").Value Then  'Calcula CSLL, COFINS, PIS
          '12/06/2008 - mpdea
          'Inclu�do regra para c�lculo conforme lei 10.925/2004
          Call IsDataType(dtDouble, .Fields("ValorIsencaoPisCofinsCsll").Value, dbl_valor_isencao)
          dbl_base_calculo = g_dbl_ValorBaseCalculoImpostosServicos(gnCodFilial, CLng(cboCliente.Text), dbl_valor_isencao, TotaldeServicos)
          
          'CSLL
          m_sngPercentualCSLL = Format(.Fields("CSLL").Value, FORMAT_VALUE)
          m_dblTotalCSLL = Format(m_sngPercentualCSLL * dbl_base_calculo / 100, FORMAT_VALUE)
          'COFINS
          m_sngPercentualCOFINS = Format(.Fields("COFINS").Value, FORMAT_VALUE)
          m_dblTotalCOFINS = Format(m_sngPercentualCOFINS * dbl_base_calculo / 100, FORMAT_VALUE)
          'PIS
          m_sngPercentualPIS = Format(.Fields("PIS").Value, FORMAT_VALUE)
          m_dblTotalPIS = Format(m_sngPercentualPIS * dbl_base_calculo / 100, FORMAT_VALUE)
        End If
        
        '27/12/2007 - Anderson
        'O valor m�nimo do c�lculo para o IRRF � de R$ 10,00.
        'esta altera��o � para considerar este valor.
        'IRRF
        'If TotaldeServicos > 666 Then
        If CDbl((.Fields("IRRF").Value * TotaldeServicos) / 100) >= 10 Then
          m_sngPercentualIRRF = Format(.Fields("IRRF").Value, FORMAT_VALUE)
          m_dblTotalIRRF = Format(m_sngPercentualIRRF * TotaldeServicos / 100, FORMAT_VALUE)
        Else 'Caso seja menor suspendemos o IRRF
          m_sngPercentualIRRF = 0
          m_dblTotalIRRF = 0
        End If
      End If
    End With
    
    rstParametros.Close
    rstOpSaida.Close
    Set rstParametros = Nothing
    Set rstOpSaida = Nothing
  
    m_dblTotaldeImpostos = (m_dblTotalCSLL + m_dblTotalCOFINS + m_dblTotalPIS + m_dblTotalIRRF)
    m_dblTotaldeImpostos = Format(m_dblTotaldeImpostos, FORMAT_VALUE)
    
    m_dblTotalMenosServ = TotaldeServicos - (m_dblTotalCSLL + m_dblTotalCOFINS + m_dblTotalPIS + m_dblTotalIRRF)
    m_dblTotalMenosServ = Format(m_dblTotalMenosServ, FORMAT_VALUE)
  End If

End Sub

'13/05/2004 - Daniel
'Zera as Vars de tratamento de percentuais e totais sobre servi�os
Private Sub ZerarVarsImpostosServi�os()
  m_sngPercentualCSLL = 0
  m_sngPercentualCOFINS = 0
  m_sngPercentualPIS = 0
  m_sngPercentualIRRF = 0
  
  m_dblTotalCSLL = 0
  m_dblTotalCOFINS = 0
  m_dblTotalPIS = 0
  m_dblTotalIRRF = 0
  
  m_dblTotalMenosServ = 0
  m_dblTotaldeImpostos = 0
End Sub

Private Function PossuiPermissao() As Boolean
  '25/06/2004 - Daniel
  'Criado rotina de valida��o para checar se o user tem permiss�o ou
  'n�o de limpar os campos. Solicitado pelo cliente Coneg Campos e
  'aproveitado para os demais
  Dim rstFuncionarios As Recordset
  Dim strQuery        As String
  
  PossuiPermissao = True
  
  strQuery = "SELECT C�digo, Nome, SenhaClear "
  strQuery = strQuery & " FROM Funcion�rios "
  strQuery = strQuery & " WHERE Funcion�rios.C�digo = " & gnUserCode
  
  Set rstFuncionarios = db.OpenRecordset(strQuery, dbOpenDynaset)

  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If Not .Fields("SenhaClear").Value Then
        PossuiPermissao = False
      End If
      
    End If
    .Close
  End With
  
  Set rstFuncionarios = Nothing
 
End Function

Private Sub EnchergarEstoque()
  '26/08/2004 - Daniel
  'Criado valida��o para verificar se o usu�rio possui permiss�o
  'para enchergar o estoque ou n�o
  Dim rstFuncionarios As Recordset
  
  Set rstFuncionarios = db.OpenRecordset("SELECT C�digo, VRVisualizarEstoque FROM Funcion�rios WHERE C�digo = " & gnUserCode, dbOpenDynaset)
  
  With rstFuncionarios
   If Not (.BOF And .EOF) Then
     .MoveFirst
     
     m_blnPermitido = .Fields("VRVisualizarEstoque").Value
   End If
   .Close
  End With
  
  Set rstFuncionarios = Nothing

End Sub

Private Function IE_Isento(ByRef Estado As String) As Boolean
  '09/11/2004 - Daniel
  'Verifica��o da I.E. retorna o estado do Cliente
  Dim rstCliFor As Recordset
  Dim strSQL    As String
  Dim strIE     As String
  
  If Len(Nome_Cliente.Caption) <= 0 Then Exit Function
  
  strSQL = "SELECT Inscri��o, Estado FROM Cli_For "
  strSQL = strSQL & " WHERE C�digo = " & CLng(cboCliente.Text)

  Set rstCliFor = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstCliFor
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      strIE = .Fields("Inscri��o").Value & ""
      strIE = UCase(strIE)
      
      If strIE = "ISENTO" Or strIE = "" Then IE_Isento = True
      
      Estado = Trim(.Fields("Estado").Value)
      
    End If
    .Close
  End With
  
  Set rstCliFor = Nothing

End Function

Private Function PessoaFisica(ByVal CodCliente As Long) As Boolean
  'Function criada em 07/12/2004 por Daniel
  'Finalidade: Atender as necessidades da Embalavi
  Dim rstCliFor As Recordset
  Dim strSQL    As String
  
  strSQL = "SELECT F�sica_Jur�dica FROM Cli_For "
  strSQL = strSQL & " WHERE C�digo = " & CodCliente
    
  Set rstCliFor = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstCliFor
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .Fields("F�sica_Jur�dica").Value = "F" Then PessoaFisica = True
      
    End If
    .Close
  End With
  
  Set rstCliFor = Nothing
  
End Function

Private Function GetFabricante(ByVal CodProdu As String) As String
  '29/03/2005 - Daniel
  'Case: El�trica Leal
  'Exibi��o da coluna Fabricante
  Dim rstProdutos As Recordset
  
  GetFabricante = ""
  
  Set rstProdutos = db.OpenRecordset("SELECT Fabricante FROM Produtos WHERE C�digo = '" & CodProdu & "'", dbOpenDynaset)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      GetFabricante = .Fields("Fabricante").Value & ""
    End If
    .Close
  End With
  
  Set rstProdutos = Nothing

End Function

Private Sub GetLimiteCliente(ByVal lngCliente As Long, ByRef dblLimite As Double)
  '02/05/2005 - Daniel
  Dim rstCliente As Recordset
  
  dblLimite = 0
  
  Set rstCliente = db.OpenRecordset("SELECT [Limite Cr�dito] FROM Cli_For WHERE C�digo = " & lngCliente, dbOpenDynaset)
  
  With rstCliente
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      dblLimite = Format(CDbl("0" & .Fields("Limite Cr�dito").Value), FORMAT_VALUE)
    End If
    .Close
  End With

  Set rstCliente = Nothing

End Sub

'Private Sub RefreshTela()
'  '01/04/2005 - Daniel
'
'End Sub

Private Sub EmitirDuplicatas()
  '17/06/2005 - Daniel
  'Rotina para Emiss�o de Duplicatas a partir da tela de Sa�das
  Dim Resp As String
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Grave ou encontre uma venda antes."
    Exit Sub
  End If
  
  With frmEmiteFatura
    .Tipo.Caption = "F"
    .L_Encontrar.Caption = "N�O"
    .L_Nota.Caption = rsSaidas("Nota Impressa") & ""
    .L_Fatura.Caption = ""
    .L_Valor.Caption = rsSaidas("Total")
    .L_Vencimento.Caption = Resp
    .L_Cliente.Caption = rsSaidas("Cliente")
    .lblDataEmissao.Caption = Data_Atual
    .Caption = "Impress�o de Fatura"
    .optTotalParcela.Value = True
    'Foi solicitado para imprimir �s parcelas
    .optTotalNota.Enabled = False
    'Precisamos saber a Sequ�ncia para buscarmos �s parcelas
    .lblSequencia.Caption = rsSaidas("Sequ�ncia") & ""
    '-------------------------------------------------------
    .Show vbModal
  End With
  
End Sub

'18/06/2007 - Anderson
'Fun��o utilizada para exportar dados para excel
Private Sub ExportarExcel()

  Dim appExcel As New Excel.Application
  Dim rsExpParametros As Recordset
  Dim rsExpSaidas As Recordset
  Dim rsExpSaidasProdutos As Recordset
  Dim rsExpSaidasServicos As Recordset
  Dim rsExpSaidasDigitadorOperador As Recordset
  Dim strSQL As String
  Dim strRange As String
  Dim intContador As Integer
  Dim strCampo As String
  Dim strValor As String
  
  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Encontre um registro antes."
    Set appExcel = Nothing
    Exit Sub
  End If
  
  If gsArquivoExcelSaida = "" Then
    Beep
    DisplayMsg "Arquivo modelo para exporta��o de dados n�o est� configurado, favor verificar as configura��es no arquivo config.ini no diret�rio padr�o do Quick Store"
    Set appExcel = Nothing
    Exit Sub
  End If
  
  If MsgBox("Deseja exportar dados atual para Excel?", vbYesNo + vbQuestion, "Exportar Dados para Excel") = vbYes Then
    
    Call StatusMsg("Aguarde, exportando dados...")
    MousePointer = vbHourglass
  
    'Inicia Excel
    'appExcel.Application.Visible = True
    appExcel.ScreenUpdating = False
    'Abre o arquivo modelo para exporta��o
    appExcel.Workbooks.Open gsReportPath & gsArquivoExcelSaida
    
    'Seleciona C�lula A1
    appExcel.Range("A1").Select

    'Parametros da empresa Filial
    strSQL = ""
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM [Par�metros Filial] "
    strSQL = strSQL & "WHERE Filial=" & rsSaidas("Filial")
    
    Set rsExpParametros = db.OpenRecordset(strSQL)
    
    'Exporta Cabe�alho
    With rsExpParametros
      If Not (.BOF And .EOF) Then
              
        For intContador = 0 To .Fields.Count - 1
          strCampo = Mid(.Fields(intContador).Name, InStr(1, .Fields(intContador).Name, ".") + 1)
          strCampo = Replace(sTranslateInvalidChar(.Fields(intContador).SourceTable & "_" & strCampo), " ", "_")
          If .Fields(intContador).Type = dbCurrency Or .Fields(intContador).Type = dbDecimal Or .Fields(intContador).Type = dbDouble Or .Fields(intContador).Type = dbFloat Or .Fields(intContador).Type = dbSingle Then
            strValor = Replace("" & .Fields(intContador), ",", ".")
          Else
            strValor = "" & .Fields(intContador)
          End If
          Call ExcelSubstituirCampo("[" & strCampo & "]", strValor, gsArquivoExcelSaida, appExcel)
        Next
        
      End If
      .Close
    End With
    
    'Cabe�alho da Sa�da
    strSQL = ""
    strSQL = strSQL & "SELECT Sa�das.*, Cli_For.*, [Opera��es Sa�da].*, [Caixas em Uso].*, Transportadoras.* "
    strSQL = strSQL & "FROM (((Sa�das LEFT JOIN Cli_For ON Sa�das.Cliente = Cli_For.C�digo) LEFT JOIN [Caixas em Uso] ON Sa�das.Caixa = [Caixas em Uso].Caixa) LEFT JOIN Transportadoras ON Sa�das.obs_Transportadora = Transportadoras.Nome) LEFT JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
    strSQL = strSQL & "WHERE Sa�das.Filial=" & rsSaidas("Filial") & " AND Sa�das.Sequ�ncia = " & rsSaidas("Sequ�ncia") & " "
    strSQL = strSQL & "ORDER BY Sa�das.Filial, Sa�das.Data, Sa�das.Sequ�ncia "


    Set rsExpSaidas = db.OpenRecordset(strSQL)
    
    'Exporta Cabe�alho
    With rsExpSaidas
      If Not (.BOF And .EOF) Then
              
        For intContador = 0 To .Fields.Count - 1
          strCampo = Mid(.Fields(intContador).Name, InStr(1, .Fields(intContador).Name, ".") + 1)
          strCampo = Replace(sTranslateInvalidChar(.Fields(intContador).SourceTable & "_" & strCampo), " ", "_")
          If .Fields(intContador).Type = dbCurrency Or .Fields(intContador).Type = dbDecimal Or .Fields(intContador).Type = dbDouble Or .Fields(intContador).Type = dbFloat Or .Fields(intContador).Type = dbSingle Then
            strValor = Replace("" & .Fields(intContador), ",", ".")
          Else
            strValor = "" & .Fields(intContador)
          End If
          Call ExcelSubstituirCampo("[" & strCampo & "]", strValor, gsArquivoExcelSaida, appExcel)
        Next
        
      End If
      .Close
    End With
    
    'Cabe�alho Digitador - Operador
    strSQL = ""
    strSQL = strSQL & "SELECT Sa�das.Filial, Sa�das.Sequ�ncia, Digitador.*, Operador.* "
    strSQL = strSQL & "FROM (Sa�das LEFT JOIN Funcion�rios AS Digitador ON Sa�das.Digitador = Digitador.C�digo) LEFT JOIN Funcion�rios AS Operador ON Sa�das.Operador = Operador.C�digo "
    strSQL = strSQL & "WHERE Sa�das.Filial=" & rsSaidas("Filial") & " AND Sa�das.Sequ�ncia = " & rsSaidas("Sequ�ncia") & " "
    strSQL = strSQL & "ORDER BY Sa�das.Filial, Sa�das.Data, Sa�das.Sequ�ncia "


    Set rsExpSaidasDigitadorOperador = db.OpenRecordset(strSQL)
    
    'Exporta Cabe�alho
    With rsExpSaidasDigitadorOperador
      If Not (.BOF And .EOF) Then
              
        For intContador = 0 To .Fields.Count - 1
          strCampo = Replace(.Fields(intContador).Name, ".", "_")
          strCampo = Replace(sTranslateInvalidChar(strCampo), " ", "_")
          If .Fields(intContador).Type = dbCurrency Or .Fields(intContador).Type = dbDecimal Or .Fields(intContador).Type = dbDouble Or .Fields(intContador).Type = dbFloat Or .Fields(intContador).Type = dbSingle Then
            strValor = Replace("" & .Fields(intContador), ",", ".")
          Else
            strValor = "" & .Fields(intContador)
          End If
          Call ExcelSubstituirCampo("[" & strCampo & "]", strValor, gsArquivoExcelSaida, appExcel)
        Next
        
      End If
      .Close
    End With
    
    'Detalhe da Sa�da
    strSQL = ""
    strSQL = strSQL & "SELECT [Sa�das - Produtos].*, Produtos.*, Classes.*, [Sub Classes].*, GrupoFiscal.* "
    strSQL = strSQL & "FROM ((([Sa�das - Produtos] LEFT JOIN Produtos ON [Sa�das - Produtos].C�digo = Produtos.C�digo) LEFT JOIN Classes ON Produtos.Classe = Classes.C�digo) LEFT JOIN [Sub Classes] ON Produtos.[Sub Classe] = [Sub Classes].C�digo) LEFT JOIN GrupoFiscal ON Produtos.GrupoFiscal = GrupoFiscal.C�digo "
    strSQL = strSQL & "WHERE [Sa�das - Produtos].Filial=" & rsSaidas("Filial") & " AND [Sa�das - Produtos].Sequ�ncia = " & rsSaidas("Sequ�ncia") & " "
    strSQL = strSQL & "ORDER BY [Sa�das - Produtos].Filial, [Sa�das - Produtos].Sequ�ncia, [Sa�das - Produtos].Linha "
    
    Set rsExpSaidasProdutos = db.OpenRecordset(strSQL)
    
    'Exporta Detalhe
    With rsExpSaidasProdutos
      If Not (.BOF And .EOF) Then
                
        Do Until .EOF
          'Localiza campo para inserir linha para acrescentar produto
          strRange = ""
          strRange = ExcelLocalizarCampo("[PROXIMO_PRODUTO]", gsArquivoExcelSaida, appExcel)
          
          'Se n�o tiver campo [PROXIMO_PRODUTO], o sistema n�o insere os produtos
          If strRange <> "" Then
            'Seleciona [PROXIMO_PRODUTO]
            appExcel.Range(strRange).Select
            'Seleciona linha
            appExcel.Rows(appExcel.ActiveCell.Row & ":" & appExcel.ActiveCell.Row).Select
            'Insere Linha
            appExcel.Selection.Insert Shift:=xlDown
            'Seleciona [PROXIMO_PRODUTO]
            appExcel.Range(strRange).Select
            'Seleciona linha para copiar modelo
            appExcel.Rows(appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1).Select
            'Copia modelo
            appExcel.Selection.Copy
            'Seleciona [PROXIMO_PRODUTO]
            appExcel.Range(strRange).Select
            'Seleciona Linha
            appExcel.Rows(appExcel.ActiveCell.Row & ":" & appExcel.ActiveCell.Row).Select
            'Cola Linha
            appExcel.ActiveSheet.Paste
            'Desativa mode de copia
            appExcel.CutCopyMode = False
            'Seleciona [PROXIMO_PRODUTO]
            appExcel.Range(strRange).Select
            
            'Altera conte�do das c�lulas
            For intContador = 0 To .Fields.Count - 1
              strCampo = Mid(.Fields(intContador).Name, InStr(1, .Fields(intContador).Name, ".") + 1)
              strCampo = Replace(sTranslateInvalidChar(.Fields(intContador).SourceTable & "_" & strCampo), " ", "_")
              
              If .Fields(intContador).Type = dbCurrency Or .Fields(intContador).Type = dbDecimal Or .Fields(intContador).Type = dbDouble Or .Fields(intContador).Type = dbFloat Or .Fields(intContador).Type = dbSingle Then
                strValor = Replace("" & .Fields(intContador), ",", ".")
              Else
                strValor = "" & .Fields(intContador)
              End If

              Call ExcelSubstituirCampo("[" & strCampo & "]", strValor, gsArquivoExcelSaida, appExcel, appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1)
            Next
            
          End If
          
          .MoveNext
          
        Loop
        
        'Seleciona [PROXIMO_PRODUTO]
        strRange = ""
        strRange = ExcelLocalizarCampo("[PROXIMO_PRODUTO]", gsArquivoExcelSaida, appExcel)
        
        'Seleciona [PROXIMO_PRODUTO]
        appExcel.Range(strRange).Select
        'Seleciona linha
        appExcel.Rows(appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1).Select
        'Exclui linha modelo
        appExcel.Selection.Delete Shift:=xlUp
        'Seleciona [PROXIMO_PRODUTO]
        appExcel.Range(strRange).Select
        'Limpa campo [PROXIMO_PRODUTO]
        Call ExcelSubstituirCampo("[PROXIMO_PRODUTO]", "", gsArquivoExcelSaida, appExcel)
        
      End If
      
      .Close
      
    End With
    
    'Detalhe da Sa�da Servi�os
    strSQL = ""
    strSQL = strSQL & "SELECT [Sa�das - Servi�os].*, Servi�os.*"
    strSQL = strSQL & "FROM [Sa�das - Servi�os] LEFT JOIN Servi�os ON [Sa�das - Servi�os].C�digo = Servi�os.C�digo "
    strSQL = strSQL & "WHERE [Sa�das - Servi�os].Filial=" & rsSaidas("Filial") & " AND [Sa�das - Servi�os].Sequ�ncia = " & rsSaidas("Sequ�ncia") & " "
    strSQL = strSQL & "ORDER BY [Sa�das - Servi�os].Filial, [Sa�das - Servi�os].Sequ�ncia, [Sa�das - Servi�os].Linha "
    
    Set rsExpSaidasServicos = db.OpenRecordset(strSQL)
    
    'Exporta Detalhe Servi�os
    With rsExpSaidasServicos
      If Not (.BOF And .EOF) Then
                
        Do Until .EOF
          'Localiza campo para inserir linha para acrescentar produto
          strRange = ""
          strRange = ExcelLocalizarCampo("[PROXIMO_SERVICO]", gsArquivoExcelSaida, appExcel)
          
          'Se n�o tiver campo [PROXIMO_SERVICO], o sistema n�o insere os produtos
          If strRange <> "" Then
            'Seleciona [PROXIMO_SERVICO]
            appExcel.Range(strRange).Select
            'Seleciona linha
            appExcel.Rows(appExcel.ActiveCell.Row & ":" & appExcel.ActiveCell.Row).Select
            'Insere Linha
            appExcel.Selection.Insert Shift:=xlDown
            'Seleciona [PROXIMO_SERVICO]
            appExcel.Range(strRange).Select
            'Seleciona linha para copiar modelo
            appExcel.Rows(appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1).Select
            'Copia modelo
            appExcel.Selection.Copy
            'Seleciona [PROXIMO_SERVICO]
            appExcel.Range(strRange).Select
            'Seleciona Linha
            appExcel.Rows(appExcel.ActiveCell.Row & ":" & appExcel.ActiveCell.Row).Select
            'Cola Linha
            appExcel.ActiveSheet.Paste
            'Desativa mode de copia
            appExcel.CutCopyMode = False
            'Seleciona [PROXIMO_SERVICO]
            appExcel.Range(strRange).Select
            
            'Altera conte�do das c�lulas
            For intContador = 0 To .Fields.Count - 1
              strCampo = Mid(.Fields(intContador).Name, InStr(1, .Fields(intContador).Name, ".") + 1)
              strCampo = Replace(sTranslateInvalidChar(.Fields(intContador).SourceTable & "_" & strCampo), " ", "_")
              
              If .Fields(intContador).Type = dbCurrency Or .Fields(intContador).Type = dbDecimal Or .Fields(intContador).Type = dbDouble Or .Fields(intContador).Type = dbFloat Or .Fields(intContador).Type = dbSingle Then
                strValor = Replace("" & .Fields(intContador), ",", ".")
              Else
                strValor = "" & .Fields(intContador)
              End If

              Call ExcelSubstituirCampo("[" & strCampo & "]", strValor, gsArquivoExcelSaida, appExcel, appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1)
            Next
            
          End If
          
          .MoveNext
          
        Loop
        
        'Seleciona [PROXIMO_SERVICO]
        strRange = ""
        strRange = ExcelLocalizarCampo("[PROXIMO_SERVICO]", gsArquivoExcelSaida, appExcel)
        
        'Seleciona [PROXIMO_SERVICO]
        appExcel.Range(strRange).Select
        'Seleciona linha
        appExcel.Rows(appExcel.ActiveCell.Row - 1 & ":" & appExcel.ActiveCell.Row - 1).Select
        'Exclui linha modelo
        appExcel.Selection.Delete Shift:=xlUp
        'Seleciona [PROXIMO_SERVICO]
        appExcel.Range(strRange).Select
        'Limpa campo [PROXIMO_SERVICO]
        Call ExcelSubstituirCampo("[PROXIMO_SERVICO]", "", gsArquivoExcelSaida, appExcel)
        
      End If
      
      .Close
      
    End With
    
    If gsSaveExcelSaida = "" Then
      appExcel.Visible = True
      With appExcel.FileDialog(2)
        .InitialFileName = rsSaidas("Sequ�ncia")
        .Show
      End With
      appExcel.ActiveWorkbook.SaveAs appExcel.FileDialog(2).InitialFileName & ".xls"
    Else
      appExcel.DisplayAlerts = False
      appExcel.ActiveWorkbook.SaveAs gsSaveExcelSaida & rsSaidas("Sequ�ncia") & ".xls"
      appExcel.DisplayAlerts = True
    End If
    
    appExcel.ScreenUpdating = True
    appExcel.Application.Quit
  
    MsgBox "Opera��o conclu�da com sucesso!", vbExclamation, "Exportar Dados"
  
  End If
  
  Set rsExpParametros = Nothing
  Set rsExpSaidas = Nothing
  Set rsExpSaidasProdutos = Nothing
  Set rsExpSaidasServicos = Nothing
  Set rsExpSaidasDigitadorOperador = Nothing
  Set appExcel = Nothing
  
  Call StatusMsg("")
  MousePointer = vbDefault

End Sub

'27/09/2007 - Anderson
'Implementa��o da impress�o de carn� com c�digo de barras
'Solicitado por: Naativa
Private Sub ImprimirCarne()

  Dim F As Form

  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre ou grave uma movimenta��o."
    Exit Sub
  End If
 
  If g_bolCarneCodigoBarras Then
    Set F = New frmImprimeCarneCodigoBarrasConfirmar
    F.Caption = "Impress�o de Carn�s"
    F.intFilial = rsSaidas("Filial")
    F.lngSeq = rsSaidas("Sequ�ncia")
    F.Show vbModal
    Exit Sub
  End If
End Sub

'27/05/2010 - mpdea
'Retorna o CFOP do produto
Private Function GetCfopProduto(ByVal strCodigo As String)
  Dim strRet As String
  
  If cboOper.Text <> "" Then
    rsProdutoCFOP.Index = "PrimaryKey"
    rsProdutoCFOP.Seek "=", strCodigo, cboOper.Text
    If rsProdutoCFOP.NoMatch Then
      rsOp_Sa�da.Index = "C�digo"
      rsOp_Sa�da.Seek "=", cboOper.Text
      If Not rsOp_Sa�da.NoMatch Then
        strRet = rsOp_Sa�da("C�digo Fiscal") & ""
      End If
    Else
      strRet = rsProdutoCFOP("CFOP") & ""
    End If
  End If
  
  GetCfopProduto = strRet
End Function

Private Function UpdateComanda() As Integer
'15/05/2013-Alexandre Afornali
'Case DiskEmbalages
  Dim rsComandas As Recordset
  Set rsComandas = db.OpenRecordset("SaidasComandas")
  Dim countrs As Long
  Dim verifica As Boolean
  Dim verifica2 As Boolean
  verifica = True
  verifica2 = True
  countrs = 0
  While Not rsComandas.EOF
    countrs = countrs + 1
    rsComandas.MoveNext
  Wend
  If (countrs > 0) Then
    rsComandas.MoveFirst
  End If
  While Not rsComandas.EOF
    If (rsComandas("CodComanda") = txtComanda.Text And rsComandas("Filial") = gnCodFilial) Then
      If (rsComandas("CodSaida") <> txtSeq.Text) Then
        verifica = False
        rsComandas.MoveLast
      End If
      verifica2 = False
    End If
    rsComandas.MoveNext
  Wend
  If (verifica = True) Then
    If (verifica2 = True) Then
      rsComandas.AddNew
      rsComandas("CodSaida") = txtSeq.Text
      rsComandas("CodComanda") = txtComanda.Text
      rsComandas("Filial") = gnCodFilial
      rsComandas.Update
      rsComandas.Close
    End If
  Else
    MsgBox ("Comanda ja utilizada com outra sequencia, favor utilizar outra!")
    txtComanda.Text = ""
  End If
End Function

Public Function CarregaComanda()
'15/05/2013-Alexandre Afornali
'Case DiskEmbalagens
  txtComanda.Text = ""
  Dim rsComandas As Recordset
  Set rsComandas = db.OpenRecordset("SaidasComandas")
  Dim countrs As Long
  countrs = 0
  While Not rsComandas.EOF
    countrs = countrs + 1
    rsComandas.MoveNext
  Wend
  If (countrs > 0) Then
    rsComandas.MoveFirst
  End If
  While Not rsComandas.EOF
    If (rsComandas("CodSaida") = txtSeq.Text) Then
      txtComanda.Text = rsComandas("CodComanda")
      rsComandas.MoveLast
    End If
    rsComandas.MoveNext
  Wend
End Function

Private Function UpdateTotalNCM()
  rsSaidas.Edit
  rsSaidas_Prod.OpenRecordset
  Dim totalNCM As Double 'Total em R$ de imposto pago na movimenta��o
  Dim Valor_Aprox_Impostos As Double
  Dim rsAliquotas As Recordset 'Tabela que filtra todos os produtos da sequencia
  Dim rsProdutos3 As Recordset 'Tabela que filtra produto por produto da movimenta��o
  totalNCM = 0#
  Set rsAliquotas = db.OpenRecordset("SELECT [C�digo Sem Grade],[Pre�o Final],[Valor_Aprox_Impostos],[MotivoDesoneracaoICMS] FROM [Sa�das - Produtos] WHERE Filial = " & gnCodFilial & " and [Sequ�ncia] = " & txtSeq.Text, dbOpenDynaset)
  On Error GoTo UpdateExit
  rsAliquotas.MoveFirst
  While Not rsAliquotas.EOF
    Set rsProdutos3 = db.OpenRecordset("SELECT [AliqNCM],[MotivoDesoneracaoICMS] FROM [Produtos] WHERE [C�digo] = '" & rsAliquotas("C�digo Sem Grade") & "'", dbOpenDynaset)
    rsProdutos3.MoveFirst
    If (rsProdutos3("AliqNCM") <> "" Or rsProdutos3("AliqNCM") = 0) Then
      Valor_Aprox_Impostos = (rsProdutos3("AliqNCM") * rsAliquotas("Pre�o Final") / 100)
      Valor_Aprox_Impostos = FormatNumber(Valor_Aprox_Impostos, 2)
      totalNCM = totalNCM + (rsProdutos3("AliqNCM") * rsAliquotas("Pre�o Final") / 100)
      totalNCM = FormatNumber(totalNCM, 2)
      rsAliquotas.Edit
      rsAliquotas("Valor_Aprox_Impostos") = Valor_Aprox_Impostos
      rsAliquotas("MotivoDesoneracaoICMS") = rsProdutos3("MotivoDesoneracaoICMS")
      rsAliquotas.Update
    Else
      rsAliquotas.Edit
      rsAliquotas("MotivoDesoneracaoICMS") = rsProdutos3("MotivoDesoneracaoICMS")
      rsAliquotas.Update
      'MsgBox "O produto " & rsAliquotas("C�digo Sem Grade") & " n�o possui aliquota de NCM", vbExclamation
    End If
    rsAliquotas.MoveNext
    
    rsProdutos3.Close
  Wend
  rsSaidas("TotalNCM") = totalNCM
  'rsSaidas("TotalNCM") = FormataValorTexto(rsSaidas("TotalNCM"), 2)
  rsSaidas.Update
  
  rsAliquotas.Close
  Set rsAliquotas = Nothing
  
UpdateExit:
End Function

Private Function UpdateTotalNCM_2(ByVal sCodProduto As String)
On Error GoTo UpdateExit
  
  Dim Valor_Aprox_Impostos As Double
  Dim rsProdutos3 As Recordset 'Tabela que filtra produto por produto da movimenta��o
  
  Set rsProdutos3 = db.OpenRecordset("SELECT [AliqNCM],[MotivoDesoneracaoICMS] FROM [Produtos] WHERE [C�digo] = '" & sCodProduto & "'") ', dbOpenDynaset)
  rsProdutos3.MoveFirst
  If (rsProdutos3("AliqNCM") <> "" Or rsProdutos3("AliqNCM") = 0) Then
      Valor_Aprox_Impostos = (rsProdutos3("AliqNCM") * rsSaidas_Prod("Pre�o Final") / 100)
      Valor_Aprox_Impostos = FormatNumber(Valor_Aprox_Impostos, 2)
      totalNCM_2 = totalNCM_2 + (rsProdutos3("AliqNCM") * rsSaidas_Prod("Pre�o Final") / 100)
      totalNCM_2 = FormatNumber(totalNCM_2, 2)
      
      rsSaidas_Prod("Valor_Aprox_Impostos") = Valor_Aprox_Impostos
      rsSaidas_Prod("MotivoDesoneracaoICMS") = rsProdutos3("MotivoDesoneracaoICMS")
  Else
      rsSaidas_Prod("MotivoDesoneracaoICMS") = rsProdutos3("MotivoDesoneracaoICMS")
  End If
    
  rsProdutos3.Close
  
UpdateExit:
End Function


'Formata o valor de acordo com o n�mero de casas decimais e substitui separador decimal por ponto
Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Function Retorno_PDV()
  Dim GestoBD As Database
  Dim Cfisc_Base As Recordset
  Dim Cfisc_Pgto As Recordset
  Dim TipoRecebimpgto As Recordset
  Dim DocumentoFiscal As Recordset
  Dim QuickBD As Database
  Dim Caixa As Recordset
  Dim CaixaAnterior As Recordset
  Dim Resumo_Di�rio_Financeiro As Recordset
  Dim Resumo_Di�rio As Recordset
  Dim Contas_Receber As Recordset
  Dim produtos As Recordset
  Dim cad_prod As Recordset
  Dim Estoque_Final As Recordset
  Dim Estoque As Recordset
  Dim Estoque_Anterior As Recordset
  If frmParametros.VerificaPAF = True Then
    'Atualiza Financeiro vindo do PAF
    Set rsParametros = db.OpenRecordset("Select * from [Par�metros Filial] Where Filial = " & gnCodFilial & "")
    Set GestoBD = OpenDatabase(rsParametros("BancoPDV").Value & "\Gesto.mde", False, False)
    Set DocumentoFiscal = GestoBD.OpenRecordset("Select * from DocumentoFiscal where Num_Docto = " & txtSeq.Text & "")
    If DocumentoFiscal.EOF Then
      MsgBox "Cupom n�o encontrado, favor verificar"
      Exit Function
    End If
    Set Cfisc_Pgto = GestoBD.OpenRecordset("Select * From Cfisc_Pgto where FIS_NRO = " & DocumentoFiscal("Num_Docto_Fiscal") & "")
    Set TipoRecebimpgto = GestoBD.OpenRecordset("Select * From TipoRecebimpgto Where Cint(cod_Pdv) = " & Cfisc_Pgto("Tipo_Pagto") & "")
    Set Cfisc_Base = GestoBD.OpenRecordset("Select * From Cfisc_Base Where FIS_NRO = " & Cfisc_Pgto("FIS_NRO") & "")
    Cfisc_Base.Edit
    Cfisc_Base("Importado_Retaguarda") = True
    Cfisc_Base.Update
    'Cfisc_Base = Nothing
    Set Caixa = db.OpenRecordset("Select * from Caixa where Filial = " & gnCodFilial & " and Data = #" & L_Dia.Caption & "# order by Ordem")
    If Caixa.EOF Then
      Caixa.AddNew
      Set CaixaAnterior = db.OpenRecordset("Select * from Caixa where Filial = " & gnCodFilial & " order by Data, Ordem")
      CaixaAnterior.MoveLast
      Caixa!Filial = gnCodFilial
      Caixa!Data = L_Dia.Caption
      Caixa!Caixa = 1
      Caixa!Ordem = 1
      Caixa!Funcion�rio = Combo_Operador.Text
      Caixa!Hora = Format(Time, "hh:mm:ss")
      Caixa("Saldo Anterior") = 0
      Caixa!Dinheiro = CaixaAnterior("Total Dinheiro")
      Caixa!Cheques = CaixaAnterior("Total Cheques")
      Caixa("Cheques Pr�") = CaixaAnterior("Total Cheques Pr�")
      Caixa!Cart�es = CaixaAnterior("Total Cart�es")
      Caixa!Vales = CaixaAnterior("Total Vales")
      Caixa!Parcelamento = CaixaAnterior("Total Parcelamento")
      Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro")
      Caixa("Total Cheques") = CaixaAnterior("Total Cheques")
      Caixa("Total Cheques Pr�") = CaixaAnterior("Total Cheques Pr�")
      Caixa("Total Cart�es") = CaixaAnterior("Total Cart�es")
      Caixa("Total Vales") = CaixaAnterior("Total Vales")
      Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
      Caixa!Final = CaixaAnterior("Final")
      Caixa!Descri��o = "In�cio do dia"
      Caixa.Update
    End If
    Set CaixaAnterior = db.OpenRecordset("Select * from Caixa where Filial = " & gnCodFilial & " and Data = #" & L_Dia.Caption & "# order by Ordem")
    CaixaAnterior.MoveLast
    Select Case TipoRecebimpgto("id")
      Case 1
        Caixa.AddNew
        Caixa!Filial = gnCodFilial
        Caixa!Data = L_Dia.Caption
        Caixa!Caixa = 1
        Caixa!Ordem = CaixaAnterior("Ordem") + 1
        Caixa!Funcion�rio = Combo_Operador.Text
        Caixa!Hora = Format(Time, "hh:mm:ss")
        Caixa("Saldo Anterior") = CaixaAnterior("Final")
        Caixa!Dinheiro = Cfisc_Pgto("Valor_Pagto")
        Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Cheques = 0
        Caixa("Cheques Pr�") = 0
        Caixa!Cart�es = 0
        Caixa!Vales = 0
        Caixa!Parcelamento = 0
        Caixa("Total Cheques") = CaixaAnterior("Total Cheques")
        Caixa("Total Cheques Pr�") = CaixaAnterior("Total Cheques Pr�")
        Caixa("Total Cart�es") = CaixaAnterior("Total Cart�es")
        Caixa("Total Vales") = CaixaAnterior("Total Vales")
        Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
        Caixa!Final = CaixaAnterior("Final") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Descri��o = "Sa�da nr. " & txtSeq.Text
        Caixa.Update
        rsSaidas.Edit
        rsSaidas("Recebe - Dinheiro") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas.Update
      Case 2
        Set Contas_Receber = db.OpenRecordset("Select * from [Contas a Receber] where Sequ�ncia = " & txtSeq.Text & "")
        Caixa.AddNew
        Caixa!Filial = gnCodFilial
        Caixa!Data = L_Dia.Caption
        Caixa!Caixa = 1
        Caixa!Ordem = CaixaAnterior("Ordem") + 1
        Caixa!Funcion�rio = Combo_Operador.Text
        Caixa!Hora = Format(Time, "hh:mm:ss")
        Caixa("Saldo Anterior") = CaixaAnterior("Final")
        Caixa!Dinheiro = 0
        Caixa!Cheques = Cfisc_Pgto("Valor_Pagto")
        Caixa("Total Cheques") = CaixaAnterior("Total Cheques") + Cfisc_Pgto("Valor_Pagto")
        Caixa("Cheques Pr�") = 0
        Caixa!Cart�es = 0
        Caixa!Vales = 0
        Caixa!Parcelamento = 0
        Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro")
        Caixa("Total Cheques Pr�") = CaixaAnterior("Total Cheques Pr�")
        Caixa("Total Cart�es") = CaixaAnterior("Total Cart�es")
        Caixa("Total Vales") = CaixaAnterior("Total Vales")
        Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
        Caixa!Final = CaixaAnterior("Final") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Descri��o = "Sa�da nr. " & txtSeq.Text
        Caixa.Update
        rsSaidas.Edit
        rsSaidas("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas("Tipo Parcela") = "B"
        rsSaidas.Update
        Contas_Receber.AddNew
        Contas_Receber("Filial") = gnCodFilial
        Contas_Receber("Cliente") = cboCliente.Text
        Contas_Receber!Sequ�ncia = txtSeq.Text
        Contas_Receber!Tipo = "C"
        Contas_Receber("Vencimento") = L_Dia.Caption
        Contas_Receber!Valor = Cfisc_Pgto("Valor_Pagto")
        Contas_Receber("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        Contas_Receber("Data Recebimento") = L_Dia.Caption
        Contas_Receber("Vendedor") = cboDigitador.Text
        Contas_Receber!Processado = True
        Contas_Receber.Update
      Case 3
        Set Contas_Receber = db.OpenRecordset("Select * from [Contas a Receber] where Sequ�ncia = " & txtSeq.Text & "")
        Caixa.AddNew
        Caixa!Filial = gnCodFilial
        Caixa!Data = L_Dia.Caption
        Caixa!Caixa = 1
        Caixa!Ordem = CaixaAnterior("Ordem") + 1
        Caixa!Funcion�rio = Combo_Operador.Text
        Caixa!Hora = Format(Time, "hh:mm:ss")
        Caixa("Saldo Anterior") = CaixaAnterior("Final")
        Caixa!Dinheiro = 0
        Caixa!Cheques = 0
        Caixa("Cheques Pr�") = Cfisc_Pgto("Valor_Pagto")
        Caixa("Total Cheques Pr�") = CaixaAnterior("Total Cheques Pr�") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Cart�es = 0
        Caixa!Vales = 0
        Caixa!Parcelamento = 0
        Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro")
        Caixa("Total Cheques") = CaixaAnterior("Total Cheques")
        Caixa("Total Cart�es") = CaixaAnterior("Total Cart�es")
        Caixa("Total Vales") = CaixaAnterior("Total Vales")
        Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
        Caixa!Final = CaixaAnterior("Final") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Descri��o = "Sa�da nr. " & txtSeq.Text
        Caixa.Update
        rsSaidas.Edit
        rsSaidas("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas("Tipo Parcela") = "B"
        rsSaidas.Update
        Contas_Receber.AddNew
        Contas_Receber("Filial") = gnCodFilial
        Contas_Receber("Cliente") = cboCliente.Text
        Contas_Receber!Sequ�ncia = txtSeq.Text
        Contas_Receber!Tipo = "C"
        Contas_Receber("Vencimento") = L_Dia.Caption
        Contas_Receber!Valor = Cfisc_Pgto("Valor_Pagto")
        Contas_Receber("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        Contas_Receber("Data Recebimento") = L_Dia.Caption
        Contas_Receber("Vendedor") = cboDigitador.Text
        Contas_Receber!Processado = True
        Contas_Receber.Update
      Case 5, 8, 9, 16
        Caixa.AddNew
        Caixa!Filial = gnCodFilial
        Caixa!Data = L_Dia.Caption
        Caixa!Caixa = 1
        Caixa!Ordem = CaixaAnterior("Ordem") + 1
        Caixa!Funcion�rio = Combo_Operador.Text
        Caixa!Hora = Format(Time, "hh:mm:ss")
        Caixa("Saldo Anterior") = CaixaAnterior("Final")
        Caixa!Dinheiro = 0
        Caixa!Cheques = 0
        Caixa("Cheques Pr�") = 0
        Caixa!Cart�es = Cfisc_Pgto("Valor_Pagto")
        Caixa("Total Cart�es") = CaixaAnterior("Total Cart�es") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Vales = 0
        Caixa!Parcelamento = 0
        Caixa("Total Dinheiro") = CaixaAnterior("Total Dinheiro")
        Caixa("Total Cheques") = CaixaAnterior("Total Cheques")
        Caixa("Total Cheques Pr�") = CaixaAnterior("Total Cheques Pr�")
        Caixa("Total Vales") = CaixaAnterior("Total Vales")
        Caixa("Total Parcelamento") = CaixaAnterior("Total Parcelamento")
        Caixa!Final = CaixaAnterior("Final") + Cfisc_Pgto("Valor_Pagto")
        Caixa!Descri��o = "Sa�da nr. " & txtSeq.Text
        Caixa.Update
        rsSaidas.Edit
        rsSaidas("Recebe - Cart�o") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas("Valor Recebido") = Cfisc_Pgto("Valor_Pagto")
        rsSaidas.Update
    End Select
    Set Resumo_Di�rio = db.OpenRecordset("Select * from [Resumo Di�rio] where Filial = " & gnCodFilial & " and Data = #" & L_Dia.Caption & "#")
    If Resumo_Di�rio.EOF Then
      Resumo_Di�rio.AddNew
      Resumo_Di�rio!Filial = gnCodFilial
      Resumo_Di�rio!Data = L_Dia.Caption
      Resumo_Di�rio("Valor Vendas") = L_Tot_Pagar.Text
      Resumo_Di�rio.Update
    Else
      Resumo_Di�rio.Edit
      Resumo_Di�rio!Filial = gnCodFilial
      Resumo_Di�rio!Data = L_Dia.Caption
      Resumo_Di�rio("Valor Vendas") = Resumo_Di�rio("Valor Vendas") + L_Tot_Pagar.Text
      Resumo_Di�rio.Update
    End If
    Set Resumo_Di�rio_Financeiro = db.OpenRecordset("Select * from [Resumo Di�rio] where Filial = " & gnCodFilial & " and Data = #" & L_Dia.Caption & "#")
    If Resumo_Di�rio_Financeiro.EOF Then
      Resumo_Di�rio_Financeiro.AddNew
      Resumo_Di�rio_Financeiro!Filial = gnCodFilial
      Resumo_Di�rio_Financeiro!Data = L_Dia.Caption
      Resumo_Di�rio_Financeiro("Valor Vendas") = L_Tot_Pagar.Text
      Resumo_Di�rio_Financeiro.Update
    Else
      Resumo_Di�rio_Financeiro.Edit
      Resumo_Di�rio_Financeiro!Filial = gnCodFilial
      Resumo_Di�rio_Financeiro!Data = L_Dia.Caption
      Resumo_Di�rio_Financeiro("Valor Vendas") = Resumo_Di�rio("Valor Vendas") + L_Tot_Pagar.Text
      Resumo_Di�rio_Financeiro.Update
    End If
    'Atualiza estoque PAF
    Set produtos = db.OpenRecordset("Select * from [Sa�das - Produtos] where Filial = " & gnCodFilial & " and Sequ�ncia = " & txtSeq.Text & "")
    Do Until produtos.EOF
      Set cad_prod = db.OpenRecordset("Select * from Produtos where C�digo = '" & produtos("C�digo sem Grade") & "'")
      If cad_prod("Tipo") = "N" Then
        Set Estoque_Final = db.OpenRecordset("Select * From [Estoque Final] where Filial = " & gnCodFilial & " and Produto = '" & produtos("C�digo sem Grade") & "'")
        If Estoque_Final.EOF Then
          MsgBox "O produto " & cad_prod("DESCRICAO") & " esta com estoque n�o inicializado. N�o foi possivel dar baixa no estoque"
        Else
          Estoque_Final.Edit
          Estoque_Final("Estoque Atual") = Estoque_Final("Estoque Atual") - produtos("Qtde")
          Estoque_Final("�ltima Data") = L_Dia.Caption
          Estoque_Final.Update
        End If
        Set Estoque_Anterior = db.OpenRecordset("Select * From Estoque where Filial = " & gnCodFilial & " and Produto = '" & produtos("C�digo sem Grade") & "' order by data")
        Estoque_Anterior.MoveLast
        Set Estoque = db.OpenRecordset("Select * From Estoque where Filial = " & gnCodFilial & " and Produto = '" & produtos("C�digo sem Grade") & "' And Data = #" & L_Dia.Caption & "#")
        If Estoque.EOF Then
          Estoque.AddNew
          Estoque!Filial = gnCodFilial
          Estoque!Data = L_Dia.Caption
          Estoque!Produto = produtos("C�digo sem Grade")
          Estoque!Tamanho = 0
          Estoque!Cor = 0
          Estoque!Edi��o = 0
          Estoque!Classe = cad_prod("Classe")
          Estoque("Sub Classe") = produtos("Sub Classe")
          Estoque("Estoque Anterior") = Estoque_Anterior("Estoque Final")
          Estoque!Vendas = produtos("Qtde")
          Estoque("Valor Vendas") = produtos("Pre�o Final")
          Estoque.Update
        Else
          Estoque.Edit
          Estoque("Vendas") = Estoque("Vendas") + produtos("Qtde")
          Estoque("Valor Vendas") = Estoque("Valor Vendas") + produtos("Pre�o Final")
          Estoque("Estoque Final") = Estoque("Estoque Final") - produtos("Qtde")
          Estoque.Update
        End If
        'atualiza estoque de produto com grade PAF
      ElseIf cad_prod("Tipo") = "G" Then
          Tamanho = 0
          Cor = 0
          Edicao = 0
          Tipo = 1
          Erro = 0
        modFuncoes.Acha_Produto produtos("C�digo"), produtos("C�digo"), Tamanho, Cor, Edicao, Tipo, Erro
        Set Estoque_Final = db.OpenRecordset("Select * From [Estoque Final] where Filial = " & gnCodFilial & " and Produto = '" & produtos("C�digo sem Grade") & "' AND Cor = " & Cor & " And Tamanho = " & Tamanho & "")
        If Estoque_Final.EOF Then
          MsgBox "O produto " & cad_prod("DESCRICAO") & " esta com estoque n�o inicializado. N�o foi possivel dar baixa no estoque"
        Else
          Estoque_Final.Edit
          Estoque_Final("Estoque Atual") = Estoque_Final("Estoque Atual") - produtos("Qtde")
          Estoque_Final("�ltima Data") = L_Dia.Caption
          Estoque_Final.Update
        End If
        Set Estoque_Anterior = db.OpenRecordset("Select * From Estoque where Filial = " & gnCodFilial & " and Produto = '" & produtos("C�digo sem Grade") & "' AND Cor = " & Cor & " And Tamanho = " & Tamanho & " order by data")
        Estoque_Anterior.MoveLast
        Set Estoque = db.OpenRecordset("Select * From Estoque where Filial = " & gnCodFilial & " and Produto = '" & produtos("C�digo sem Grade") & "' AND Cor = " & Cor & " And Tamanho = " & Tamanho & "")
        If Estoque.EOF Then
          Estoque.AddNew
          Estoque!Filial = gnCodFilial
          Estoque!Data = L_Dia.Caption
          Estoque!Produto = produtos("C�digo sem Grade")
          Estoque!Tamanho = Left(Right(produtos("C�digo"), 3), 3)
          Estoque!Cor = Right(produtos("C�digo"), 3)
          Estoque!Edi��o = 0
          Estoque!Classe = cad_prod("Classe")
          Estoque("Sub Classe") = cad_prod("Sub Classe")
          Estoque("Estoque Anterior") = Estoque_Anterior("Estoque Final")
          Estoque!Vendas = produtos("Qtde")
          Estoque("Valor Vendas") = produtos("Pre�o Final")
          Estoque.Update
        Else
          Estoque.Edit
          Estoque("Vendas") = Estoque("Vendas") + produtos("Qtde")
          Estoque("Valor Vendas") = Estoque("Valor Vendas") + produtos("Pre�o Final")
          Estoque("Estoque Final") = Estoque("Estoque Final") - produtos("Qtde")
          Estoque.Update
        End If
      End If
      produtos.MoveNext
    Loop
    End If
    'marca venda como efetivada
    rsSaidas.Edit
    rsSaidas("Efetivada") = True
    rsSaidas("Recebimento") = True
    rsSaidas("Cupom Fiscal Impresso") = True
    rsSaidas("Nota Impressa") = 1
    rsSaidas.Update
    L_Efetivada.Visible = True
    
End Function
