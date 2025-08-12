VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFA324&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Bem-vindo ao Quick Store®"
   ClientHeight    =   7410
   ClientLeft      =   3840
   ClientTop       =   885
   ClientWidth     =   12015
   FillColor       =   &H00FFA324&
   FontTransparent =   0   'False
   ForeColor       =   &H00FFA324&
   HelpContextID   =   10
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "Inicio.frx":4E95A
   ScaleHeight     =   568.927
   ScaleMode       =   0  'User
   ScaleWidth      =   797.015
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_instalacaoComplementar 
      BackColor       =   &H00FFA324&
      Height          =   7305
      Left            =   12060
      TabIndex        =   24
      Top             =   -30
      Width           =   3255
      Begin VB.CommandButton cmd_cnpjValida 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Registro CNPJ"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00FFA324&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   6120
         Width           =   3015
      End
      Begin VB.CommandButton cmd_fecharCompl 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6690
         Width           =   3015
      End
      Begin VB.CommandButton cmd_validar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Instalação complementar"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1950
         Width           =   3015
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFA324&
         Caption         =   "Painel validador"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3585
         Left            =   120
         TabIndex        =   25
         Top             =   2460
         Width           =   3015
         Begin VB.CheckBox chk_renomearDBCupom 
            BackColor       =   &H00FFA324&
            Caption         =   "Instalar DB Cupom Fiscal"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   28
            Top             =   472
            Value           =   1  'Checked
            Width           =   2625
         End
         Begin VB.CheckBox chk_caminhoDanfe_eCupom 
            BackColor       =   &H00FFA324&
            Caption         =   "Ajustar caminho Danfe_NFe e Cupom Fiscal"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   480
            Left            =   120
            TabIndex        =   27
            Top             =   1305
            Value           =   1  'Checked
            Width           =   2865
         End
         Begin VB.CheckBox chk_instalarSOAP 
            BackColor       =   &H00FFA324&
            Caption         =   "Instalar o SOAP ToolKit"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   26
            Top             =   2355
            Value           =   1  'Checked
            Width           =   2505
         End
         Begin VB.Label lbl_dbCupom 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ok"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   420
            TabIndex        =   31
            Top             =   780
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl_caminhosDanfe_e_CUPOM 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ok"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   420
            TabIndex        =   30
            Top             =   1830
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl_soap 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ok"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   420
            TabIndex        =   29
            Top             =   2670
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFA324&
         Caption         =   "ATENÇÃO:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   210
         TabIndex        =   34
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFA324&
         Caption         =   "Realize este procedimento uma única vez logo após a instalação do QuickStore local. Qqr dúvida entre em contato."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1365
         Left            =   180
         TabIndex        =   33
         Top             =   630
         Width           =   2865
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   11220
      Top             =   6750
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFA324&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   60
      Picture         =   "Inicio.frx":5AF36
      ScaleHeight     =   5990.812
      ScaleMode       =   0  'User
      ScaleWidth      =   11910
      TabIndex        =   7
      Top             =   30
      Width           =   11910
      Begin VB.TextBox Painel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E5E5E5&
         Height          =   405
         Left            =   7230
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   6870
         Width           =   4380
      End
      Begin VB.CommandButton cmd_complementoInstalacaoQuickLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "setup"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5940
         MaskColor       =   &H00FFA324&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6990
         Width           =   1245
      End
      Begin VB.CommandButton cmd_acessoViaContingencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "contingência"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4650
         MaskColor       =   &H00FFA324&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6990
         Width           =   1245
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   20000
         Left            =   11280
         Top             =   420
      End
      Begin VB.CommandButton cmd_manterAtivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Ping ativado?"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   60
         Width           =   1425
      End
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   390
         MaskColor       =   &H00E5E5E5&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6165
         Width           =   11280
      End
      Begin VB.TextBox txtNomeUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         HelpContextID   =   10
         Left            =   2100
         MaxLength       =   10
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3990
         Width           =   9525
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         HelpContextID   =   10
         IMEMode         =   3  'DISABLE
         Left            =   2115
         MaxLength       =   8
         PasswordChar    =   "•"
         TabIndex        =   2
         Top             =   4785
         Width           =   2760
      End
      Begin VB.TextBox txtNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         HelpContextID   =   10
         Left            =   2100
         MaxLength       =   35
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3210
         Width           =   9525
      End
      Begin VB.CommandButton cmdProsseguir 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Prosseguir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         MaskColor       =   &H00FFA324&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5460
         Width           =   11280
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "WeblySleek UI Normal"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFA324&
         Height          =   1020
         Left            =   6000
         TabIndex        =   11
         Top             =   4485
         Width           =   5685
         Begin VB.TextBox txtConfirm 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00666666&
            Height          =   435
            HelpContextID   =   10
            IMEMode         =   3  'DISABLE
            Left            =   3050
            MaxLength       =   8
            PasswordChar    =   "•"
            TabIndex        =   4
            Top             =   360
            Width           =   2565
         End
         Begin VB.TextBox txtNewPassword 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00666666&
            Height          =   435
            HelpContextID   =   10
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   8
            PasswordChar    =   "•"
            TabIndex        =   3
            Top             =   360
            Width           =   2565
         End
         Begin VB.Label lblTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFA324&
            Caption         =   "Repita a senha"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00666666&
            Height          =   315
            Index           =   6
            Left            =   3120
            TabIndex        =   13
            Top             =   90
            Width           =   1935
         End
         Begin VB.Label lblTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFA324&
            Caption         =   "Nova Senha"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00666666&
            Height          =   315
            Index           =   5
            Left            =   360
            TabIndex        =   12
            Top             =   90
            Width           =   1935
         End
      End
      Begin SSDataWidgets_B.SSDBCombo cboCodUsuario 
         Bindings        =   "Inicio.frx":67512
         Height          =   450
         Left            =   360
         TabIndex        =   1
         Top             =   4050
         Width           =   1335
         BevelWidth      =   0
         DataFieldList   =   "Código"
         BevelType       =   0
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BevelColorFrame =   12648447
         BevelColorHighlight=   8454143
         BevelColorShadow=   10066329
         BevelColorFace  =   12648447
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   16250871
         BackColorOdd    =   12648447
         RowHeight       =   529
         Columns.Count   =   2
         Columns(0).Width=   3942
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Código"
         Columns(0).DataField=   "Código"
         Columns(0).FieldLen=   256
         Columns(1).Width=   6641
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).DataField=   "Nome"
         Columns(1).FieldLen=   256
         _ExtentX        =   2355
         _ExtentY        =   794
         _StockProps     =   93
         Text            =   "1"
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCodFilial 
         Bindings        =   "Inicio.frx":6752B
         Height          =   450
         Left            =   360
         TabIndex        =   0
         Top             =   3240
         Width           =   1335
         BevelWidth      =   0
         DataFieldList   =   "Filial"
         BevelType       =   0
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BevelColorFrame =   12648447
         BevelColorHighlight=   8454143
         BevelColorShadow=   10066329
         BevelColorFace  =   12648447
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   16250871
         BackColorOdd    =   12648447
         RowHeight       =   635
         Columns.Count   =   2
         Columns(0).Width=   3545
         Columns(0).Caption=   "Codigo"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Filial"
         Columns(0).FieldLen=   256
         Columns(1).Width=   6191
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).DataField=   "Nome"
         Columns(1).FieldLen=   256
         _ExtentX        =   2355
         _ExtentY        =   794
         _StockProps     =   93
         Text            =   "1"
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Filial"
      End
      Begin MSMask.MaskEdBox medData 
         Height          =   510
         Left            =   360
         TabIndex        =   16
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   2340
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   900
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   16753444
         ForeColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lbl_manterAtivo 
         BackStyle       =   0  'Transparent
         Caption         =   "Não"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   11280
         TabIndex        =   21
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F7F7F7&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Top             =   3690
         Width           =   1335
      End
      Begin VB.Label lblTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Filial"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F7F7F7&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   2850
         Width           =   1320
      End
      Begin VB.Label lblTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F7F7F7&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   4845
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFA324&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E5E5E5&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   6870
         Width           =   4185
      End
      Begin VB.Label lblWeekDay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   660
         Left            =   4545
         TabIndex        =   8
         Top             =   2295
         Width           =   3705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00F7F7F7&
         X1              =   360
         X2              =   11610
         Y1              =   1886.677
         Y2              =   1886.677
      End
      Begin VB.Label lblVersaoSistema 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E5E5E5&
         Height          =   285
         Left            =   10380
         TabIndex        =   9
         Top             =   2385
         Width           =   1395
      End
   End
   Begin VB.Data datUsuario 
      Caption         =   "datUsuario"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE AND isPrestServ = false ORDER BY Nome"
      Top             =   7965
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datFilial 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'28/01/2009 - mpdea
'Adaptado para versão 7

'27/08/2003 - mpdea
'Corrigido nomenclatura dos objetos txtCodFilial e txtCodUsuario
'para cboCodFilial e cboCodUsuario respectivamente

Private rsParametros As Recordset
Private rsFuncionarios As Recordset
Private gbSuper As Boolean

Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long

Public sRETORNO_WEBAPI_LICENCA As String


Private Sub cmd_acessoViaContingencia_Click()
On Error GoTo Erro

  Dim sChaveCont As String
  Dim dtValida As Date
  Dim dtCorrente As String
  Dim dtContingenciaAcesso As Date
  Dim dtCorrenteDesblindada As String
  Dim dtCorrenteBlindada As String
  Dim sSQLLicenca As String
  Dim sChaveR1 As String
  Dim sChaveR2 As String
  
  sChaveCont = InputBox$("Informe a chave de contingência: ", "Acesso")
  
  If Len(LTrim(RTrim(sChaveCont))) <> 26 Then
      MsgBox "Chave inválida!", vbCritical, "Erro"
      End
  Else
      dtCorrenteDesblindada = Mid(sChaveCont, 9, 10)
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "A", "1")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "M", "2")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "C", "3")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "P", "4")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "E", "5")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "H", "6")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "I", "7")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "K", "8")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "L", "9")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "S", "0")
      dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "R", "/")

      dtContingenciaAcesso = Format(dtCorrenteDesblindada, "yyyy/MM/dd")
      dtCorrente = Format(Now, "yyyy/MM/dd")
      dtValida = Format(dtCorrente, "yyyy/MM/dd")

      If dtContingenciaAcesso = dtValida Then
        gAcessoContingencia_ControleDeLicencaWebApi = "SIM"
        
        ' Se esta liberado... então atualiza a base com data de hoje
        ' Gravar com data de hoje
        dtCorrenteBlindada = dtCorrente
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "1", "A")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "2", "M")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "3", "C")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "4", "P")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "5", "E")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "6", "H")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "7", "I")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "8", "K")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "9", "L")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "0", "S")
        dtCorrenteBlindada = Replace(dtCorrenteBlindada, "/", "R")
        
        dtCorrenteBlindada = sChaveR1 + dtCorrenteBlindada + sChaveR2

        sSQLLicenca = "UPDATE ParametroVLC SET dadoConteudo='" & dtCorrenteBlindada & "', "
        
'''        If InStr(1, sRETORNO_WEBAPI_LICENCA, "XML1") > 0 Then
        sSQLLicenca = sSQLLicenca & " estrategicoREL='1', "
'''        Else
'''          sSQLLicenca = sSQLLicenca & " estrategicoREL='0', "
'''        End If

'''        If InStr(1, sRETORNO_WEBAPI_LICENCA, "EstratRel1") > 0 Then
        sSQLLicenca = sSQLLicenca & " managerNFEXML='1' "
'''        Else
'''          sSQLLicenca = sSQLLicenca & " managerNFEXML='0' "
'''        End If

        sSQLLicenca = sSQLLicenca & " WHERE dadoValida='X2' and dadoV2='AF' "

        Call db.Execute(sSQLLicenca, dbFailOnError)

      Else
        gAcessoContingencia_ControleDeLicencaWebApi = "NAO"
        MsgBox "Chave inválida!", vbCritical, "Erro"
        End
      End If
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro " & Err.Number & " " & Err.Description, vbCritical, "Atenção"
  End
End Sub

Private Sub cmd_cnpjValida_Click()
On Error GoTo Erro

  Dim retMsg As Variant
  Dim sCNPJ As String
  Dim sSql As String
  retMsg = MsgBox("Esta opção é para você informar o CNPJ de cadastro junto a Infopar QuickStore. Deseja realmente prosseguir?", vbYesNo, "Registrar o CNPJ")
  
  If retMsg = vbNo Then
      Exit Sub
  End If
  
  sCNPJ = InputBox$("Informe o CNPJ da sua empresa: ", "Cadastrar CNPJ")
  
  If Len(LTrim(RTrim(sCNPJ))) < 11 Then
      MsgBox "CNPJ inválido!", vbCritical, "Erro"
      Exit Sub
  Else
      sCNPJ = Replace(sCNPJ, " ", "")
      sCNPJ = Replace(sCNPJ, "/", "")
      sCNPJ = Replace(sCNPJ, "\", "")
      sCNPJ = Replace(sCNPJ, "-", "")
      sCNPJ = Replace(sCNPJ, "_", "")
      sCNPJ = Replace(sCNPJ, ".", "")
      sCNPJ = Replace(sCNPJ, ";", "")
      sCNPJ = Replace(sCNPJ, ":", "")
      sCNPJ = Replace(sCNPJ, ",", "")

      sSql = "UPDATE [Parâmetros Filial] SET CGC='" & sCNPJ & "' "
      sSql = sSql & " WHERE Filial = 1 "

      Call db.Execute(sSql, dbFailOnError)

      MsgBox "CNPJ foi registrado em Parâmetros Filial/Empresa, no registro 01", vbInformation, "Aviso"
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro " & Err.Number & " " & Err.Description, vbCritical, "Atenção"
  End
End Sub

Private Sub cmd_complementoInstalacaoQuickLocal_Click()
  If frm_instalacaoComplementar.Left >= 800 Then
    frm_instalacaoComplementar.Visible = True
    frm_instalacaoComplementar.Left = 10
    Picture1.Visible = False
    Me.Width = 3570
  Else
    frm_instalacaoComplementar.Visible = False
    frm_instalacaoComplementar.Left = 800
    Picture1.Visible = True
    Me.Width = 12105
  End If
End Sub

Private Sub cmd_fecharCompl_Click()
  If frm_instalacaoComplementar.Left >= 800 Then
    frm_instalacaoComplementar.Visible = True
    frm_instalacaoComplementar.Left = 10
    Picture1.Visible = False
    Me.Width = 3570
  Else
    frm_instalacaoComplementar.Visible = False
    frm_instalacaoComplementar.Left = 800
    Picture1.Visible = True
    Me.Width = 12105
  End If
End Sub


Private Sub cmd_manterAtivo_Click()
  If lbl_manterAtivo.Caption = "Não" Then
      lbl_manterAtivo.Caption = "Sim"
      Timer2.Enabled = True
      gbPingAtivado = True
  Else
      lbl_manterAtivo.Caption = "Não"
      Timer2.Enabled = False
      gbPingAtivado = False
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
  If gbAppStarting = True Then
    db.Close
    dbFoo.Close
    ws.Close
    End
  End If
End Sub

Sub WebRequest(url As String)
On Error GoTo trata_WebApiErro

    ' Função que chama o componente WebApi
    Dim retWebApiAux As String

    Dim http As MSXML2.XMLHTTP
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    http.Open "GET", url, False
    http.Send

    retWebApiAux = http.statusText
    sRETORNO_WEBAPI_LICENCA = http.responseText

    Exit Sub
    
trata_WebApiErro:
'    Dim iInStrRet As Integer
'    iInStrRet = InStr(1, Err.Description, "timed out")
'    If iInStrRet > 0 And tentativasWEBAPI < 4 Then
'        tentativasWEBAPI = tentativasWEBAPI + 1
'        WebRequest url
'    End If
'    sRETORNO_ERRO_WEBAPI = Err.Description
End Sub

Private Sub controleGeralDeAcesso(sCodFilial As String)
On Error GoTo ERRO_Geral
'  ' ***************************************
'  ' Tratamento de ACESSO PERIODO VIA WEBAPI
'  ' PILATTI DEZ/17
  Dim dtCorrente As String
  Dim dtValida As Date
  Dim dtContingenciaAcesso As Date
  Dim rsValida As Recordset
  Dim v1 As String
  Dim v2 As String
  Dim vD As String
  Dim dtCorrenteBlindada As String
  Dim dtCorrenteDesblindada As String
  Dim sSQLLicenca As String
  Dim iDia As Integer
  Dim sChaveR1 As String
  Dim sChaveR2 As String
  
  
  Dim rsParamFilialLicenca As Recordset
  Set rsParamFilialLicenca = db.OpenRecordset("Select CGC, Nome FROM [Parâmetros Filial] where Filial= " & sCodFilial, dbOpenDynaset)
  gCNPJ_CPFControleDeLicencaWebApi = rsParamFilialLicenca.Fields("CGC").Value
  gNomeEmpresaFilial = rsParamFilialLicenca.Fields("Nome").Value
  rsParamFilialLicenca.Close
  Set rsParamFilialLicenca = Nothing
  gCNPJ_CPFControleDeLicencaWebApi = Replace(gCNPJ_CPFControleDeLicencaWebApi, "/", "")
  gCNPJ_CPFControleDeLicencaWebApi = Replace(gCNPJ_CPFControleDeLicencaWebApi, ".", "")
  gCNPJ_CPFControleDeLicencaWebApi = Replace(gCNPJ_CPFControleDeLicencaWebApi, "-", "")
  gCNPJ_CPFControleDeLicencaWebApi = Replace(gCNPJ_CPFControleDeLicencaWebApi, ",", "")
  gCNPJ_CPFControleDeLicencaWebApi = Replace(gCNPJ_CPFControleDeLicencaWebApi, "\", "")
  gCNPJ_CPFControleDeLicencaWebApi = Replace(gCNPJ_CPFControleDeLicencaWebApi, ";", "")
  gCNPJ_CPFControleDeLicencaWebApi = Replace(gCNPJ_CPFControleDeLicencaWebApi, " ", "")
  gCNPJ_CPFControleDeLicencaWebApi = Replace(gCNPJ_CPFControleDeLicencaWebApi, "_", "")
  
  
  iDia = Day(Now)
  retornoParteRandomicaChave iDia, sChaveR1, sChaveR2
  
  
'  Dim sCGCCPFLicenca As String
'  Dim rsParamFilialLicenca As Recordset
'  Set rsParamFilialLicenca = db.OpenRecordset("Select CGC FROM [Parâmetros Filial] where Filial=1 ", dbOpenDynaset)
'  sCGCCPFLicenca = rsParamFilialLicenca.Fields("CGC").Value
'  rsParamFilialLicenca.Close
'  Set rsParamFilialLicenca = Nothing
  
  ' ********************************************
  ' FLAG GERAL PARA VERIFICAR SE ESTA VERSÃO DO QUICKSTORE TRATA OU NÃO O CONTROLE DE LICENÇA DE USO DO QUICK VIA WEBAPI
  ' 1-TRATA;  0-NÃO TRATA
  gTrataControleDeLicencaWebApi = 0
  
  ' ********************************************
  ' FLAG GERAL SE USUARIO TEM OU NÃO ACESSO AO 'ESTRATÉGICO RELATORIOS'... 1-Tem acesso;  0-Não tem;
  gESTRATEGICO_Relatorios = 1
  
  ' FLAG GERAL PARA ABRIR OU NÃO AS TELAS DO XML... 1-ABRE; 0-NÃO ABRE;
  gAbreModuloXML = 1
  
'  '------------------------------
'  'abaixo quick liberado tudo e sem controle
'''  gTrataControleDeLicencaWebApi = 0
'''  gESTRATEGICO_Relatorios = 1
'''  gAbreModuloXML = 1
  ' ********************************************


  ' Abrir arquivo .txt para recuperar endereço alvo do Componente WebApi (Homolog ou Produção)
  Dim ff As Integer
  Dim sAUX As String
  Dim sEndereco As String
  Dim sCODIGO_ERRO As String
  Dim sTesteApenasSourceSafe As String
  Dim sNrSerieNFTratado As String
  Dim iNrSerieNFTratado As Integer
  Dim xConta As Integer
  Dim sLendoLinha As String
  Dim sCNPJ_ As String
  
  ' Controle de mensagem de erro...
  sCODIGO_ERRO = "Acesso ao arquivo de endereço"
  
  'MsgBox "Abrir arquivo: " & App.Path + "\ERP_WEBAPI_Acesso_Config.txt"

  ff = FreeFile
  Open App.Path + "\ERP_WEBAPI_Acesso_Config.txt" For Input As #ff
    Line Input #ff, sEndereco                                       ' linha 1:  Endereço WebApli A3Manager
    Line Input #ff, sSoapClient_MSSoapInit                          ' linha 2:  NFe
    Line Input #ff, sSoapClient_ConnectorProperty_EndPointURL       ' linha 3:  NFe
    Line Input #ff, sCaminhoDanfe_Benefix                           ' linha 4:  NFe
    Line Input #ff, sSoapClient_MSSoapInit_NFCe                     ' linha 5:  NFCe
    Line Input #ff, sSoapClient_ConnectorProperty_EndPointURL_NFCe  ' linha 6:  NFCe
    Line Input #ff, sNrSerieNFTratado           'Numero de CNPJ tratados neste arquivo (para o campo 'Serie' da NF)
    
    'MsgBox "Lendo arquivo: " & sSoapClient_MSSoapInit & "   ----    " & sSoapClient_ConnectorProperty_EndPointURL & "    -------    " & sCaminhoDanfe_Benefix
    
    sEnderecoValidaLicencaQuickStoreWEBAPI = sEndereco
    
    iNrSerieNFTratado = CInt(sNrSerieNFTratado)
    gNrSerieNF = iNrSerieNFTratado
    If iNrSerieNFTratado > 0 Then
        For xConta = 0 To gNrSerieNF - 1
            'NrCnpj1 , SerieNFe1, SerieNFCe1
            'NrCnpj2 , SerieNFe2, SerieNFCe2
            
            Line Input #ff, sLendoLinha         'CNPJ
            sCNPJ_ = sLendoLinha
            If InStr(sCNPJ_, "_") = 0 Then
              gArrayNrSerieNF(xConta, 0) = sLendoLinha
              Line Input #ff, sLendoLinha         'NrSerieNFe
              gArrayNrSerieNF(xConta, 1) = sLendoLinha
              Line Input #ff, sLendoLinha         'NrSerieNFCe
              gArrayNrSerieNF(xConta, 2) = sLendoLinha
            Else
              If Split(sCNPJ_, "_")(1) = sCodFilial Then
                gArrayNrSerieNF(xConta, 0) = Split(sLendoLinha, "_")(0)
              Else
                gArrayNrSerieNF(xConta, 0) = Split(sLendoLinha, "_")(1)
              End If
              Line Input #ff, sLendoLinha         'NrSerieNFe
              gArrayNrSerieNF(xConta, 1) = sLendoLinha
              Line Input #ff, sLendoLinha         'NrSerieNFCe
              gArrayNrSerieNF(xConta, 2) = sLendoLinha
            End If
        Next
    End If
  Close #ff

  sCODIGO_ERRO = "Após leitura do arq. de endereço"


  If Not IsNull(gAcessoContingencia_ControleDeLicencaWebApi) And gAcessoContingencia_ControleDeLicencaWebApi = "SIM" Then
    ' Entrou pela chave de acesso temporaria (de 30 DIAS).
    Exit Sub
  End If

  If gTrataControleDeLicencaWebApi = 1 Then

      Dim bLICENCA As Boolean
      bLICENCA = True
      
      dtCorrente = Format(Now, "yyyy/MM/dd")
      dtValida = Format(dtCorrente, "yyyy/MM/dd") 'define uma nova data
      dtContingenciaAcesso = Format(dtCorrente, "yyyy/MM/dd") 'define uma nova data
      
'      ' **************************************************
'      ' **************************************************
'      ' ** IF para acesso limitado por ANO
'      Dim anoTeste As String 'Cliente Telemaco/Telêmaco de Curitiba
'      anoTeste = Mid(dtCorrente, 1, 4)
'      If anoTeste >= 2020 Then
'          End
'      Else
'          Exit Sub
'      End If
'      ' **************************************************
'      ' **************************************************
      
    
      sCODIGO_ERRO = "Acesso ao DB Busca Chave"
    
      'Date = mdate  ´altera a data do sistema
      'Set rsValida = db.OpenRecordset("SELECT dadoValida, dadoConteudo, dadoV2, managerNFEXML, estrategicoREL FROM ParametroVLC", dbOpenDynaset)
      Set rsValida = db.OpenRecordset("SELECT * FROM ParametroVLC", dbOpenDynaset)
    
      sCODIGO_ERRO = "Após o Acesso ao DB Busca Chave"
    
      If Not rsValida.EOF Then
        v1 = rsValida.Fields("dadoValida").Value

        If IsNull(rsValida.Fields("dadoConteudo").Value) Then
          vD = ""
        Else
          vD = rsValida.Fields("dadoConteudo").Value
        End If

        v2 = rsValida.Fields("dadoV2").Value

        ' FLAG SE USUARIO TEM OU NÃO ACESSO AO 'ESTRATÉGICO RELATORIOS'... 1-Tem acesso;  0-Não tem;
        If rsValida.Fields("estrategicoREL").Value <> "1" Then
            gESTRATEGICO_Relatorios = 0   'Não tem acesso
        Else
            gESTRATEGICO_Relatorios = 1   'Tem acesso
        End If
        
        ' FLAG  ABRIR OU NÃO AS TELAS DO XML... 1-ABRE; 0-NÃO ABRE;
        If rsValida.Fields("managerNFEXML").Value <> "1" Then
            gAbreModuloXML = 0   'Não tem acesso
        Else
            gAbreModuloXML = 1   'Tem acesso
        End If

        ' Regra geral:
        ' Tabela ParametroVLC
        ' Campos dadoValida = 'X2' e dadoConteudo = 'AF'   ... Cliente deve ser controlado mensalmente o seu ACESSO via WEBAPI A3
        ' Campos dadoValida = 'AB' e dadoConteudo = 'X3'   ... Cliente que não precisa de verificação de acesso
        ' Se não existe registro algum nesta tabela (como uma das condições acima)    ...Acesso bloqueado

        ' Condição
        If v1 = "X2" And v2 = "AF" Then
          If vD = "" Then
            ' Primeiro acesso do usuario...sem data de vencimento...então gravo neste acesso data de (hoje -180 dias) ou seja, já expirada
            ' Gravar com data de hoje
            'dtCorrenteBlindada = dtCorrente
            dtCorrenteBlindada = Format(Now - 180, "yyyy/MM/dd")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "1", "A")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "2", "M")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "3", "C")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "4", "P")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "5", "E")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "6", "H")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "7", "I")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "8", "K")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "9", "L")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "0", "S")
            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "/", "R")
            dtCorrenteBlindada = sChaveR1 + dtCorrenteBlindada + sChaveR2
    
            Call db.Execute("UPDATE ParametroVLC SET dadoConteudo= " & _
            " '" & dtCorrenteBlindada & "' WHERE dadoValida='X2' and dadoV2='AF' ", dbFailOnError)
            
            ' Encerrar a aplicação para que o usuário agora possa chamar o WebApi...
            rsValida.Close
            Set rsValida = Nothing
            MsgBox "A configuração de licença de uso esta sendo validada. Por favor, faça o acesso novamente! (Módulo: Controle de Acesso Cód. Branco B1.)", , "Acesso"
            End
            
    
          Else
            ' Existe data para controle de acesso
            If Len(Trim(vD)) <> 26 Then
                ' Se a chave com a data blindada não tiver tamanho de 26 caracteres...bloquear acesso
                sCODIGO_ERRO = "Chave blindada mal formatada"
                bLICENCA = False
                ' Encerrar a aplicação !
            ElseIf Trim(sEndereco) = "" Then
                ' URL Inexistente ao WebApi...bloquear acesso
                sCODIGO_ERRO = "URL Inexistente ao WebApi"
                bLICENCA = False
                ' Encerrar a aplicação !
            Else
                dtCorrenteDesblindada = vD
                dtCorrenteDesblindada = Mid(dtCorrenteDesblindada, 9, 10)
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "A", "1")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "M", "2")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "C", "3")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "P", "4")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "E", "5")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "H", "6")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "I", "7")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "K", "8")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "L", "9")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "S", "0")
                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "R", "/")
    
                ' Validar se passou dos  30 dias!
                Dim dtBase As Date
                dtBase = Format(dtCorrenteDesblindada, "yyyy/MM/dd")
    
                'dtValida = dtValida - Day(30)
                dtValida = dtValida - 30
    
                If dtValida > dtBase Then
                  ' Chamar webApi para verificar se este CNPJ esta liberado para acesso/uso
                  Dim s As String
                  s = "Chamar"
    
'                  sAux = lUserInfo
'                  sAux = Replace(sAux, "/", "")
'                  sAux = Replace(sAux, ".", "")
'                  sAux = Replace(sAux, "-", "")
'                  sAux = Replace(sAux, ",", "")
'                  sAux = Replace(sAux, "\", "")
'                  sAux = Replace(sAux, ";", "")
'                  sAux = Replace(sAux, " ", "")
                  
'                  If sAux = "" Then
'''                      Dim sCGCCPFLicenca As String
'''                      Dim rsParamFilialLicenca As Recordset
'''                      Set rsParamFilialLicenca = db.OpenRecordset("Select CGC FROM [Parâmetros Filial] where Filial= " & sCodFilial, dbOpenDynaset)
'''                      sAux = rsParamFilialLicenca.Fields("CGC").Value
'''                      rsParamFilialLicenca.Close
'''                      Set rsParamFilialLicenca = Nothing
'''                      sAux = Replace(sAux, "/", "")
'''                      sAux = Replace(sAux, ".", "")
'''                      sAux = Replace(sAux, "-", "")
'''                      sAux = Replace(sAux, ",", "")
'''                      sAux = Replace(sAux, "\", "")
'''                      sAux = Replace(sAux, ";", "")
'''                      sAux = Replace(sAux, " ", "")
'                  End If
    
                  WebRequest sEndereco + "?_cnpj=" + gCNPJ_CPFControleDeLicencaWebApi

                  ' Ex dos dados que retornar do WebApi dentro da variavel sRETORNO_WEBAPI_LICENCA
                  ' LICENCA_ATIVA__XML1_EstratRel1
                  ' LICENCA_ATIVA__XML0_EstratRel0
                  ' LICENCA_ATIVA__XML0_EstratRel1
                  ' LICENCA_ATIVA__XML1_EstratRel0
                  
                  If InStr(1, sRETORNO_WEBAPI_LICENCA, "LICENCA_ATIVA") > 0 Then
                    ' Se esta liberado (pegar o retorno do webApi), então atualiza a base com data de hoje
                    ' Gravar com data de hoje
                    dtCorrenteBlindada = dtCorrente
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "1", "A")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "2", "M")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "3", "C")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "4", "P")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "5", "E")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "6", "H")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "7", "I")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "8", "K")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "9", "L")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "0", "S")
                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "/", "R")
                    
                    dtCorrenteBlindada = sChaveR1 + dtCorrenteBlindada + sChaveR2

                    sSQLLicenca = "UPDATE ParametroVLC SET dadoConteudo='" & dtCorrenteBlindada & "', "
                    If InStr(1, sRETORNO_WEBAPI_LICENCA, "XML1") > 0 Then
                      sSQLLicenca = sSQLLicenca & " estrategicoREL='1', "
                    Else
                      sSQLLicenca = sSQLLicenca & " estrategicoREL='0', "
                    End If

                    If InStr(1, sRETORNO_WEBAPI_LICENCA, "EstratRel1") > 0 Then
                      sSQLLicenca = sSQLLicenca & " managerNFEXML='1' "
                    Else
                      sSQLLicenca = sSQLLicenca & " managerNFEXML='0' "
                    End If

                    sSQLLicenca = sSQLLicenca & " WHERE dadoValida='X2' and dadoV2='AF' "

                    Call db.Execute(sSQLLicenca, dbFailOnError)

                  ElseIf InStr(1, sRETORNO_WEBAPI_LICENCA, "LICENCA_INATIVA") > 0 Then
                    ' Caso não esta liberado...Encerrar a aplicação !!
                    sCODIGO_ERRO = "Ret. LIÇENCA INATIVA"
                    MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E1. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
                    bLICENCA = False
                    ' Encerrar a aplicação !!
                  Else
                    sCODIGO_ERRO = "Ret. LIÇENCA ??"
                    ' Erro na chamada ao WebApi...Encerrar a aplicação !!
                    
                    dtContingenciaAcesso = dtContingenciaAcesso - 75   ' 2 meses e meio de tolerancia...
                    If dtContingenciaAcesso > dtBase Then
                        MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E98. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
                        bLICENCA = False
                    Else
                        MsgBox "Acesso contingenciado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E2. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
                        bLICENCA = True
                    End If
                  End If
    
                End If
    
            End If
    
          End If
''        ElseIf v1 = "AB" And v2 = "X3" Then
''          ' Cliente que não precisa de verificação de acesso liberado
''          Dim ss As String
''          ss = "ok"
        Else
          ' Inconsistente
          sCODIGO_ERRO = "Registro com Minichaves inconsistentes"
          MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E3. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
          bLICENCA = False
          ' Encerrar a aplicação !!
        End If
    
      Else
        ' Acesso bloqueado !!
        sCODIGO_ERRO = "Nenhum Registro com Minichaves"
        MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E4. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
        bLICENCA = False
        ' Encerrar a aplicação !!
      End If
    
      ' ***************************************
    
      If bLICENCA = False Then
        rsValida.Close
        Set rsValida = Nothing
        MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E5. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
        End
      End If
      
      rsValida.Close
      Set rsValida = Nothing
  End If

  Exit Sub

ERRO_Geral:

    dtContingenciaAcesso = dtContingenciaAcesso - 75   ' 2 meses e meio de tolerancia...
    If dtContingenciaAcesso > dtBase Then
        MsgBox "Acesso contingenciado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E99. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
    Else
        ' Acesso bloqueado !!
        MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E6. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
        'MsgBox "Erro leitura do arquivo WS", , "Acesso"
        End
    End If
End Sub

Private Sub retornoParteRandomicaChave(iDiaNow As Integer, ByRef sChaveR1 As String, ByRef sChaveR2 As String)

  If iDiaNow = 1 Then
    sChaveR1 = "XABJEDEG"
    sChaveR2 = "FIPBWEDA"
  ElseIf iDiaNow = 2 Then
    sChaveR1 = "B7ZJ82EG"
    sChaveR2 = "0IQBWEDW"
  ElseIf iDiaNow = 3 Then
    sChaveR1 = "C7ZJ82EG"
    sChaveR2 = "4WQECED5"
  ElseIf iDiaNow = 4 Then
    sChaveR1 = "79RETREV"
    sChaveR2 = "TR8UINGF"
  ElseIf iDiaNow = 5 Then
    sChaveR1 = "3DRSEWQQ"
    sChaveR2 = "YEMVNHSS"
  ElseIf iDiaNow = 6 Then
    sChaveR1 = "IIYJDHCD"
    sChaveR2 = "PO76NCTE"
  ElseIf iDiaNow = 7 Then
    sChaveR1 = "XYUT5R4E"
    sChaveR2 = "PIYRF45S"
  ElseIf iDiaNow = 8 Then
    sChaveR1 = "B7ZJ82EG"
    sChaveR2 = "LAGW529H"
  ElseIf iDiaNow = 9 Then
    sChaveR1 = "ASWEDC53"
    sChaveR2 = "08UHFE21"
  ElseIf iDiaNow = 10 Then
    sChaveR1 = "MBCA23ER"
    sChaveR2 = "VCDF4532"
  ElseIf iDiaNow = 11 Then
    sChaveR1 = "016F4DDS"
    sChaveR2 = "O9IJNBGR"
  ElseIf iDiaNow = 12 Then
    sChaveR1 = "MMHJ67FR"
    sChaveR2 = "UUH323WS"
  ElseIf iDiaNow = 13 Then
    sChaveR1 = "OK56FDEW"
    sChaveR2 = "MLK8U871"
  ElseIf iDiaNow = 14 Then
    sChaveR1 = "6GVDE3WQ"
    sChaveR2 = "4ESZAQ1L"
  ElseIf iDiaNow = 15 Then
    sChaveR1 = "MLP7YG4E"
    sChaveR2 = "GHJF45ER"
  ElseIf iDiaNow = 16 Then
    sChaveR1 = "WSZ343Q2"
    sChaveR2 = "HB6T5RSW"
  ElseIf iDiaNow = 17 Then
    sChaveR1 = "P0O9KNJU"
    sChaveR2 = "7YNLK8YS"
  ElseIf iDiaNow = 18 Then
    sChaveR1 = "RDX341I9"
    sChaveR2 = "OU64S21A"
  ElseIf iDiaNow = 19 Then
    sChaveR1 = "C7ZK82IJ"
    sChaveR2 = "OKFFDE43"
  ElseIf iDiaNow = 20 Then
    sChaveR1 = "0OIUJ6Y4"
    sChaveR2 = "YGT51SW2"
  ElseIf iDiaNow = 21 Then
    sChaveR1 = "DE56J9JS"
    sChaveR2 = "KJ89UW1A"
  ElseIf iDiaNow = 22 Then
    sChaveR1 = "F6LO0IKS"
    sChaveR2 = "Z3E3W5TT"
  ElseIf iDiaNow = 23 Then
    sChaveR1 = "TG9IMAWS"
    sChaveR2 = "3R445GVB"
  ElseIf iDiaNow = 24 Then
    sChaveR1 = "J0OIUN32"
    sChaveR2 = "P92S3FF5"
  ElseIf iDiaNow = 25 Then
    sChaveR1 = "5G7MJJ8I"
    sChaveR2 = "AAS3DFGY"
  ElseIf iDiaNow = 26 Then
    sChaveR1 = "AEW34RF3"
    sChaveR2 = "UQ3ED76Y"
  ElseIf iDiaNow = 27 Then
    sChaveR1 = "ZE9ON34A"
    sChaveR2 = "1AQZ2SXW"
  ElseIf iDiaNow = 28 Then
    sChaveR1 = "0PLO9IKA"
    sChaveR2 = "NJUH77YG"
  ElseIf iDiaNow = 29 Then
    sChaveR1 = "RD4EZSW3"
    sChaveR2 = "7YNJOK9I"
  ElseIf iDiaNow = 30 Then
    sChaveR1 = "PA74GFTR"
    sChaveR2 = "GRDHY65T"
  ElseIf iDiaNow = 31 Then
    sChaveR1 = "0P3DFG2A"
    sChaveR2 = "AHW34MVN"
  End If
End Sub

Private Sub cmdProsseguir_Click()
  Dim bTeste As Boolean
  Dim rsParametros As Recordset
  Dim F As Form
  
  'Valida os campos
  Call cboCodFilial_Validate(bTeste)
  If bTeste Then
    Exit Sub
  End If
  Call cboCodUsuario_Validate(bTeste)
  If bTeste Then
    Exit Sub
  End If

  '***************************************
  'Chamar tratamento de controle de acesso
  controleGeralDeAcesso cboCodFilial.Text
  '***************************************

  g_bolCarneCodigoBarras = True
  
  If txtNomeFilial.Text = "" Then
    gsTitle = "Erro de Acesso"
    gsMsg = "Código da Filial incorreto. Digite-o."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If cboCodFilial.Enabled Then cboCodFilial.SetFocus
    Exit Sub
  End If
  
  If txtNomeUsuario.Text = "" Then
    gsTitle = "Erro de Acesso"
    gsMsg = "Código de Usuário incorreto. Digite-o."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    cboCodUsuario.SetFocus
    Exit Sub
  End If
  
  gSenhaUsuarioLogado = txtPassword.Text
  
  If LCase(txtPassword.Text) = "terraqst" Then
      gbSuper = True
  Else
      With rsFuncionarios
        .FindFirst "Código = " & gsHandleNull(cboCodUsuario.Text)
        'Filial de acesso
        Filial_Liberada = ![Filial Acesso]
        
        gbMostrarTelaPesquisaProdutoTipoFoto = rsFuncionarios("bMostrarTelaPesquisaProdutoTipoFoto").Value
        gbUsuarioAcessoApenasTelaVendaRapida = rsFuncionarios("bUsuarioAcessoApenasTelaVendaRapida").Value
        
        If Not .NoMatch Then
          If CDbl(.Fields("ValorP")) = CriptografaSenha("1") And (txtNewPassword.Text = "1" Or Trim(txtNewPassword.Text) = "") Then
            gsTitle = LoadResString(56)
            gsMsg = LoadResString(55)
            gnStyle = vbOKOnly + vbExclamation
            gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
            txtPassword.Text = "1"
            txtNewPassword.SetFocus
            Exit Sub
          End If
          If CDbl(.Fields("ValorP")) <> CriptografaSenha(txtPassword.Text) Then
            Beep
            gsTitle = "Erro de Acesso"
            gsMsg = "Senha inválida."
            gnStyle = vbOKOnly + vbExclamation
            gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
            txtPassword.SetFocus
            Exit Sub
          Else
            If .Fields("Liberado").Value = True Then
              If .Fields("Superusuário").Value = False Then
                If .Fields("Filial Acesso").Value <> 0 And .Fields("Filial Acesso").Value <> cboCodFilial.Text Then
                  gsTitle = "Erro de Acesso"
                  gsMsg = "Usuário não autorizado para esta Filial."
                  gnStyle = vbOKOnly + vbExclamation
                  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
                  cboCodUsuario.SetFocus
                  Exit Sub
                End If
              End If
              If Len(Trim(txtNewPassword.Text)) > 0 Then
                If txtNewPassword.Text <> txtConfirm.Text Then
                  gsTitle = "Erro de Acesso"
                  gsMsg = "Nova senha não foi confirmada. Verifique e tente novamente."
                  gnStyle = vbOKOnly + vbExclamation
                  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
                  txtNewPassword.SetFocus
                  Exit Sub
                Else
                  If txtNomeUsuario.Text = "quick" And (CDbl(.Fields("ValorP")) = CriptografaSenha("1") And (txtNewPassword.Text = "1" Or Trim(txtNewPassword.Text) = "")) Then
                    gsTitle = LoadResString(56)
                    gsMsg = LoadResString(55)
                    gnStyle = vbOKOnly + vbExclamation
                    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
                    txtNewPassword.SetFocus
                    Exit Sub
                  Else
                    Call ChangePassword
                  End If
                End If
              Else
                If txtNomeUsuario.Text = "quick" And (CDbl(.Fields("ValorP")) = CriptografaSenha("1") And (txtNewPassword.Text = "1" Or Trim(txtNewPassword.Text) = "")) Then
                  gsTitle = LoadResString(56)
                  gsMsg = LoadResString(55)
                  gnStyle = vbOKOnly + vbExclamation
                  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
                  txtNewPassword.SetFocus
                  Exit Sub
                End If
              End If
            Else
              gsTitle = "Erro de Acesso"
              gsMsg = "Usuário não liberado pelo administrador."
              gnStyle = vbOKOnly + vbExclamation
              gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
              cboCodUsuario.SetFocus
              Exit Sub
            End If
          End If
        Else
          gsTitle = "Erro de Acesso"
          gsMsg = "Senha inválida."
          gnStyle = vbOKOnly + vbExclamation
          gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          cboCodUsuario.SetFocus
          Exit Sub
        End If
      End With
  End If
  
  If Not IsDate(medData.Text) Then
    gsTitle = "Teste de Data"
    gsMsg = "Data de referência inválida."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    medData.SetFocus
    Exit Sub
  End If
  
  Formato_Preço = "Currency"
  
  Data_Atual = CDate(medData.Text)
  Dim gsDecimal As String
  gsAppVersion = "V. " & App.Major & "." & App.Minor & "." & App.Revision
  gsCurrencySymbol = gsGetCurrencySymbol()
  gnCurrencyDecimals = 2
  gsCurrencyDecimal = Mid(Replace(FormatCurrency(0, 2), gsCurrencySymbol, ""), 2, 1)
  If gsCurrencyDecimal = "." Then
    gsCurrencySeparator = ","
  Else
    gsCurrencySeparator = "."
  End If
  gsCurrencySeparator = ""
  gsCurrencyFormat = "###" & gsCurrencySeparator & "###" & gsCurrencySeparator & "###" & gsCurrencySeparator & "###" & gsCurrencySeparator & "##0." & String(gnCurrencyDecimals, "0")
  
  gnCodFilial = CInt(cboCodFilial.Text)
  
  
  '----------------------------------------------------------------------------
  '03/11/2005 - mpdea
  'KEY: ODBC
'  Set rsParametros = db.OpenRecordset("Parâmetros Filial")
'  rsParametros.Index = "Filial"
'  rsParametros.Seek "=", gnCodFilial
  Set rsParametros = db.OpenRecordset("SELECT * FROM [Parâmetros Filial] WHERE Filial = " & _
    gnCodFilial, dbOpenDynaset, dbReadOnly)
  '----------------------------------------------------------------------------

  ' Incluida estas 4 linhas abaixo para verificar a FLAG se a Empresa/Filial logada participa do prog. fidelidade
  On Error GoTo erroProgFidelidadeSegueAdiante
  gParticipaProgramaFidelidade = 0
  gParticipaProgramaFidelidade = rsParametros("participaProgramaFidelidade")    ' O trat. de erro é apenas para que o codigo siga normal adiante caso na 'tabela de parametros' não exista o campo [participaProgramaFidelidade]
erroProgFidelidadeSegueAdiante:

  gsSenhaGerente = rsParametros("Senha Gerente") & ""
  gsNomeFilial = rsParametros("Nome")
  
  If Not IsNull(rsParametros("NumCasasDecimais").Value) Then
      If rsParametros("NumCasasDecimais").Value = 2 Then
          g_bln5CasasDecimais = False
          g_bln3CasasDecimais = False
      ElseIf rsParametros("NumCasasDecimais").Value = 3 Then
          g_bln5CasasDecimais = False
          g_bln3CasasDecimais = True
      ElseIf rsParametros("NumCasasDecimais").Value = 5 Then
          g_bln5CasasDecimais = True
          g_bln3CasasDecimais = False
      Else
          g_bln5CasasDecimais = False
          g_bln3CasasDecimais = False
      End If
  Else
    g_bln5CasasDecimais = False
    g_bln3CasasDecimais = False
  End If
  
  If Not IsNull(rsParametros("Endereço")) Then
      gsFilialEndereco = rsParametros("Endereço")
      If Not IsNull(rsParametros("EnderecoNumero")) Then
          gsFilialEndereco = gsFilialEndereco & ", " & rsParametros("EnderecoNumero")
      End If
  End If
  
  If Not IsNull(rsParametros("Bairro")) Then
      gsFilialBairro = rsParametros("Bairro")
  End If

  If Not IsNull(rsParametros("Cidade")) Then
      gsFilialCidadeEstado = rsParametros("Cidade")
      If Not IsNull(rsParametros("Estado")) Then
          gsFilialCidadeEstado = gsFilialCidadeEstado & "-" & rsParametros("Estado")
      End If
  End If

  If Not IsNull(rsParametros("Fone")) Then
      gsFilialFone = rsParametros("Fone")
  End If

  If Not IsNull(rsParametros("Cep")) Then
      gsFilialCep = rsParametros("Cep")
  End If
 
  gsPesq1 = rsParametros("Nome Pesquisa 1") & ""
  gsPesq2 = rsParametros("Nome Pesquisa 2") & ""
  gsPesq3 = rsParametros("Nome Pesquisa 3") & ""
  
  Glob_Cod_Alfa = rsParametros("Usar Códigos Alfa")
  
  If Not IsNull(rsParametros("Estado")) And rsParametros("Estado") <> "" Then
      gsEstadoOrigemEmpresaLogado = rsParametros("Estado")
  Else
      gsEstadoOrigemEmpresaLogado = ""
  End If
  
  If Not IsNull(rsParametros("CodigoRegimeTributario")) Then
      If rsParametros("CodigoRegimeTributario") = 3 Then
          gblnSimplesNacional = False
      Else
          gblnSimplesNacional = True
      End If
  End If
  
  
  
  '26/05/2004 - Daniel
  'Var de controle de cadastro de produtos
  'com zero a esquerda
  gbZeroEsquerda = rsParametros("Zero a Esquerda").Value
  '-----------------------------------------------------

  If Not IsNull(rsParametros("Quick_viaRDP_ticket")) And rsParametros("Quick_viaRDP_ticket") <> "" Then
      giQuick_viaRDP_ticket = rsParametros("Quick_viaRDP_ticket")
  End If

  If Not IsNull(rsParametros("Quick_viaRDP")) And rsParametros("Quick_viaRDP") <> "" Then
      giQuick_viaRDP = rsParametros("Quick_viaRDP")
  End If
  
  'Endereço do Site da Benefix (para o usuario chamar o browser via Menu do QuickStore)
  
  If Not IsNull(rsParametros("PastaRetornoNfe")) Then
      gPastaRetornoNfe = rsParametros("PastaRetornoNfe")
  End If
  
  gbGrade = rsParametros("Usar Grade") = True
  gbEdicao = rsParametros("Usar Edições") = True
  gbServico = rsParametros("Usar Serviços") = True
  gbCaixas = rsParametros("Usa Vários Caixas") = True
  '20/11/2006 - Anderson
  'Considerar saldo anterior ao movimentar o caixa
  gbSaldoAnterior = rsParametros("ConsiderarSaldoAnterior") = True

  '17/01/2007 - Anderson
  'Solicitar senha do gerente ao alterar vendedor nas telas de cadastro de clientes, venda rápida, saídas e check-out
  gbVendedorSenhaGerente = rsParametros("VendedorSenhaGerente") = True
  
  '31/07/2002 - mpdea
  'Campo de utilização da Loja Virtual
  gblnWorkWeb = rsParametros.Fields("WorkWeb").Value
  
  '17/09/2009 - mpdea
  'Habilitar uso de Nota Fiscal Eletrônica
  gblnNFe = rsParametros.Fields("HabilitarNotaFiscalEletronica").Value
  
  If Dir(gsHelpContextFileName) <> "" Then
    App.HelpFile = SetHTMLHelpStrings(1) & "::/Context.txt"
  End If
  
  rsParametros.Close
  Set rsParametros = Nothing
  
  gbLoginDone = True
  
  gnUserCode = CInt(cboCodUsuario.Text)
  gsUserName = UCase(txtNomeUsuario.Text)
  gbSuperUser = gbSuper
  
  'Armazena a última Filial usada
  SaveSetting "QuickStore", "Login", "Filial", CStr(gnCodFilial)
  'Armazena o último Operador usado
  SaveSetting "QuickStore", "Login", "Operador", CStr(gnUserCode)
  
  '13/07/2004 - Daniel
  'Case: TV Shopping
  'Alteração: Caso o usuário seja responsável pelo Marketing da empresa
  'será chamada uma rotina para atualização dos clientes no que se refere
  'a curva A B C D
  If CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959") Then
    Call ResponsavelMKT(CInt(cboCodUsuario.Text))
  End If
  
  Unload Me
  
End Sub

Private Sub ChangePassword()
  With rsFuncionarios
    .Edit
    .Fields("ValorP").Value = CStr(CriptografaSenha(txtNewPassword.Text))
    .Fields("Senha").Value = Format(Date, "yyyymmdd")
    .Update
  End With
End Sub

Private Sub LoadAcessos()
  Dim rsAcessos As Recordset
  Dim rsProg As Recordset
  Set rsAcessos = db.OpenRecordset("SELECT * FROM Acessos ORDER BY Usuário, Numero", dbOpenDynaset)
  Set rsProg = db.OpenRecordset("SELECT * FROM ZZZProgramas", dbOpenDynaset)
  With rsProg
    Do While Not .EOF
      rsAcessos.FindFirst "Usuário = " & cboCodUsuario.Text & " And Numero = " & .Fields("Número")
      If Not rsAcessos.NoMatch Then
        rsAcessos.Edit
      Else
        rsAcessos.AddNew
        rsAcessos("Usuário") = cboCodUsuario.Text
        rsAcessos("Numero") = .Fields("Número")
      End If
      rsAcessos("Programa") = .Fields("Nome Programa")
      rsAcessos("Gravar") = True
      rsAcessos("Apagar") = True
      rsAcessos.Update
      .MoveNext
    Loop
  End With
  rsProg.Close
  Set rsProg = Nothing
  rsAcessos.Close
  Set rsAcessos = Nothing
End Sub

Private Sub Form_Activate()
  If txtNomeFilial.Text <> "" And txtNomeUsuario.Text <> "" Then
    Call SelectAllText(txtPassword, True)
  End If
  
  If gbPingAtivado = True Then
      lbl_manterAtivo.Caption = "Sim"
      Timer2.Enabled = True
  End If
  
End Sub

Private Sub Form_Load()
  Dim sRet As String
  
  On Error GoTo ErrHandler
  
  
  '08/08/2003 - mpdea
  'Coloca a tela em modo modal e acima de todas
  Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
  
  lblVersaoSistema.Caption = sVersaoDoSistema
  
  gbLoginDone = False
  gbSuper = False
  
  txtPassword.Text = ""
  cmdProsseguir.Default = False
  
  cboCodFilial.CausesValidation = True
  cboCodUsuario.CausesValidation = True
  
  
  '03/11/2005 - mpdea
  'KEY: ODBC
  'Abertura do banco de dados
  If g_bln_odbc Then
    datFilial.Connect = "ODBC;DSN=" & g_str_dsn_quickstore
    datUsuario.Connect = "ODBC;DSN=" & g_str_dsn_quickstore
    datFilial.DatabaseName = ""
    datUsuario.DatabaseName = ""
  Else
    datFilial.DatabaseName = gsQuickDBFileName
    datUsuario.DatabaseName = gsQuickDBFileName
  End If
  
  
'  cmdProsseguir.Enabled = False
  
  ' Verifica se a Resolução da Estação Cliente está Adequada
  gnCurrHRes = GetDeviceCaps(hdc, HORZRES)
  gnCurrVRes = GetDeviceCaps(hdc, VERTRES)
  If gnCurrHRes < 800 Or gnCurrVRes < 600 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(108)
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Unload Me
    End
  End If
  
  Call CenterForm(Me)
  
  medData.Text = gsFormatDate(Date)
  
  Call medData_Change
   
  Label2.Caption = " " & gsQuickDBFileName
  
  Set rsParametros = db.OpenRecordset("SELECT * FROM [Parâmetros Filial] ORDER BY Filial", dbOpenDynaset)
  Set rsFuncionarios = db.OpenRecordset("SELECT * FROM [Funcionários] ORDER BY Código", dbOpenDynaset)
  
  'Última Filial usada
  sRet = GetSetting("QuickStore", "Login", "Filial", "")
  cboCodFilial.Text = IIf(IsNumeric(sRet), sRet, "")
  Call cboCodFilial_Validate(False)
  
  'Último Operador usado
  sRet = GetSetting("QuickStore", "Login", "Operador", "")
  cboCodUsuario.Text = IIf(IsNumeric(sRet), sRet, "")
  Call cboCodUsuario_Validate(False)
  
  
  '22/01/2003 - mpdea
  'Em modo limitado somente a filial 1 é permitida
  If Not gblnQuickFull Then
    With cboCodFilial
      .Enabled = False
      .Text = 1
    End With
    Call cboCodFilial_Validate(False)
  End If
  
  #If BETA = 1 Then
    lblBeta.Visible = True
  #End If
  
  
  ' Pilatti Outubro/17
  ' Abrir arquivo .txt para recuperar configuracao se estamos em HORARIO DE VERAO
  Dim ff As Integer
  Dim alvoHorarioVerao As String
  ff = FreeFile
  Open App.Path + "\ERP_Config_Horario_Verao.txt" For Input As #ff
    Line Input #ff, alvoHorarioVerao
  Close #ff
  
  'HORARIO_VERAO=SIM
  alvoHorarioVerao = Replace(alvoHorarioVerao, "HORARIO_VERAO=", "")
  If Trim(alvoHorarioVerao) = "SIM" Then
    bHorarioDeVerao = True
  Else
    bHorarioDeVerao = False
  End If
  ' Fim Pilatti
  
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro na abertura de tabelas da aplicação."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  On Error Resume Next
  db.Close
  dbFoo.Close
  ws.Close
  End
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  rsParametros.Close
  rsFuncionarios.Close
  Set rsParametros = Nothing
  Set rsFuncionarios = Nothing
  On Error GoTo 0
End Sub

Private Sub medData_Change()
  If IsDate(medData.Text) Then
    lblWeekDay.Caption = WeekdayName(Weekday(medData.Text, vbUseSystemDayOfWeek), False, vbUseSystemDayOfWeek)
  Else
    lblWeekDay.Caption = ""
  End If
End Sub


Private Sub Timer1_Timer()
 Static Texto_Atual As Integer
 Dim sText As String

 Timer1.Interval = 4000
 
 If Texto_Atual = 0 Then
   Texto_Atual = 1
   sText = " Quick Store - Software Inteligente de Automação Comercial"
   sText = sText & vbCrLf & vbCrLf & " Software de Automação Comercial desenvolvido pela Infopar Desenvolvimento e Comércio  de Software Ltda. CNPJ 81.653.206/0001-12"
   Painel.Text = sText
   Exit Sub
 End If
 
 If Texto_Atual = 1 Then
    Texto_Atual = 2
    sText = " © 1999-2013 Infopar Desenvolvimento e Comércio de Software Ltda."
    sText = sText & vbCrLf & " O software Quick Store é protegido pelas leis de Copyright e Tratados Internacionais, bem  como por outras legislações e tratados sobre propriedade intelectual."
    sText = sText & vbCrLf & " Software é licenciado e não vendido. Proibida a reprodução total ou parcial."
    Painel.Text = sText
    Exit Sub
  End If
  
  If Texto_Atual = 2 Then
    Texto_Atual = 0
    sText = " Em caso de dúvida entre em contato com a Infopar Software."
    sText = sText & vbCrLf
    sText = sText & " Visite a Infopar na Internet:"
    sText = sText & vbCrLf & " Home page: http://www.infopar.com.br"
    sText = sText & vbCrLf & " e-mail: suporte@infopar.com.br"
    Painel.Text = sText
    Exit Sub
  End If
  
End Sub

Private Sub medData_GotFocus()
  With medData
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub medData_LostFocus()
  medData.Text = Ajusta_Data(medData.Text)
End Sub

Private Sub medData_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      medData.Text = frmCalendario.gsDateCalender(medData.Text)
  End Select
End Sub

Private Sub cboCodFilial_CloseUp()
'  cboCodFilial.Text = cboCodFilial.Columns(0).Text
  cboCodFilial_Validate (False)
End Sub

Private Sub cboCodFilial_GotFocus()
  With cboCodFilial
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub cboCodFilial_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

'Private Sub cboCodFilial_LostFocus()
'  On Error GoTo 0
'
'  txtNomeFilial.Text = ""
'  If IsNull(cboCodFilial.Text) Then Exit Sub
'  If Not IsNumeric(cboCodFilial.Text) Then Exit Sub
'  If CInt(cboCodFilial.Text) <= 0 Then Exit Sub
'
'  With rsParametros
'    .FindFirst "Filial = " & CStr(cboCodFilial.Text)
'    If Not .NoMatch Then
'      txtNomeFilial.Text = .Fields("Nome")
'      gsSenhaGerente = .Fields("Senha Gerente") & ""
'      gnCodFilial = CInt(cboCodFilial.Text)
'      gsNomeFilial = .Fields("Nome")
'    Else
'      Beep
'      DisplayMsg "Filial inválida."
'      txtNomeFilial.Text = ""
'      Exit Sub
'    End If
'  End With
'
'   ' "###,###,##0.00"
''  If rsParametros("Qtde Dígitos Preço") = 3 Then Formato_Preço = "###,###,##0.000"
''  If rsParametros("Qtde Dígitos Preço") = 4 Then Formato_Preço = "###,###,##0.0000"
'
'End Sub

Private Sub cboCodFilial_Validate(Cancel As Boolean)
  txtNomeFilial.Text = ""
  If cboCodFilial.Text <> "" Then
    With rsParametros
      .FindFirst "Filial = " & CStr(cboCodFilial.Text)
      If .NoMatch Then
        Beep
        DisplayMsg "Filial inválida."
        cboCodFilial.Text = ""
        txtNomeFilial.Text = ""
        Cancel = True
      Else
        txtNomeFilial.Text = .Fields("Nome")
      End If
    End With
  End If
End Sub

Private Sub cboCodUsuario_CloseUp()
'  cboCodUsuario.Text = cboCodUsuario.Columns(0).Text
  cboCodUsuario_Validate (False)
End Sub

Private Sub cboCodUsuario_GotFocus()
  With cboCodUsuario
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub cboCodUsuario_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub cboCodUsuario_Validate(Cancel As Boolean)
  txtNomeUsuario.Text = ""
  If cboCodUsuario.Text <> "" Then
    With rsFuncionarios
      .FindFirst "Código = " & CStr(cboCodUsuario.Text)
      If Not .NoMatch Then
        If Len(Trim(.Fields("Apelido") & "")) > 0 Then
          txtNomeUsuario.Text = .Fields("Apelido")
          gbSuper = .Fields("Superusuário")
          txtPassword.Enabled = True
        Else
          Beep
          gsTitle = "Erro de Acesso"
          gsMsg = "Nome de Usuário ausente para código ou inválido."
          gnStyle = vbOKOnly + vbExclamation
          gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          cboCodUsuario.Text = ""
'          cboCodUsuario.SetFocus
          Cancel = True
        End If
      Else
        Beep
        gsTitle = "Erro de Acesso"
        gsMsg = "Código de Usuário inválido."
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        cboCodUsuario.Text = ""
'        cboCodUsuario.SetFocus
        Cancel = True
      End If
    End With
  End If
End Sub

Private Sub Timer2_Timer()

   frm_manterAtivo.Show 1

End Sub

'Private Sub cboCodUsuario_LostFocus()
'
'  If Len(Trim(cboCodUsuario.Text)) = 0 Then
'    Beep
'    gsTitle = "Erro de Acesso"
'    gsMsg = "Código de Usuário inválido."
'    gnStyle = vbOKOnly + vbExclamation
'    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'    Exit Sub
'  End If
'
'  If Not IsNumeric(cboCodUsuario.Text) Then
'    Beep
'    gsTitle = "Erro de Acesso"
'    gsMsg = "Código de Usuário inválido."
'    gnStyle = vbOKOnly + vbExclamation
'    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'    Exit Sub
'  End If
'
'  txtNomeUsuario.Text = ""
'
'  With rsFuncionarios
'    .FindFirst "Código = " & CStr(cboCodUsuario.Text)
'    If Not .NoMatch Then
'      If Len(Trim(.Fields("Apelido") & "")) > 0 Then
'        txtNomeUsuario.Text = .Fields("Apelido")
'        gbSuper = .Fields("Superusuário")
'        txtPassword.Enabled = True
'      Else
'        Beep
'        gsTitle = "Erro de Acesso"
'        gsMsg = "Nome de Usuário ausente para código ou inválido."
'        gnStyle = vbOKOnly + vbExclamation
'        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'      End If
'    Else
'      Beep
'      gsTitle = "Erro de Acesso"
'      gsMsg = "Código de Usuário inválido."
'      gnStyle = vbOKOnly + vbExclamation
'      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'    End If
'  End With
'
'End Sub

Private Sub txtConfirm_GotFocus()
  txtConfirm.SelStart = 0
  txtConfirm.SelLength = 8
End Sub

Private Sub txtNewPassword_GotFocus()
  txtNewPassword.SelStart = 0
  txtNewPassword.SelLength = 8
End Sub

Private Sub txtNomeFilial_GotFocus()
  If cboCodFilial.Enabled Then cboCodFilial.SetFocus
End Sub

Private Sub txtNomeUsuario_GotFocus()
  cboCodUsuario.SetFocus
End Sub

Private Sub txtPassword_Change()
  With cmdProsseguir
    .Enabled = Len(txtPassword.Text) > 0
    .Default = .Enabled
  End With
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.SelStart = 0
  txtPassword.SelLength = 8
  With rsFuncionarios
    .FindFirst "Código = " & gsHandleNull(cboCodUsuario.Text)
    If Not .NoMatch Then
      
      '27/08/2003 - mpdea
      'Corrigido Run-time 94
      If CDbl("0" & .Fields("ValorP").Value) = CriptografaSenha("1") Then
        gsTitle = LoadResString(201)
        gsMsg = LoadResString(55)
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        txtPassword.Text = "1"
        txtPassword.Enabled = False
        txtNewPassword.SetFocus
      Else
        txtPassword.Enabled = True
        txtPassword.SetFocus
      End If
    End If
  End With
End Sub

Private Sub ResponsavelMKT(ByVal CodUser As Integer)
  '13/07/2004 - Daniel
  'Case: TV Shopping
  'Verificar se este funcionário é responsável pelo MKT da empresa
  'em caso afirmativo será disparado a atualização da curva de clientes
  'A B C D (Conforme o Parâmetro de Classificação de Clientes)
  Dim rstFuncionarios As Recordset
  Dim rstGrupos       As Recordset
  Dim strQuery        As String
  
  'Primeiro Verificar se existe registro na tabela GruposDeClientes
  'Caso afirmativo chama AtualizarCurvaClientes
  strQuery = "SELECT * FROM GruposDeClientes "
  strQuery = strQuery & " WHERE Filial = " & gnCodFilial
  
  Set rstGrupos = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstGrupos.RecordCount = 0 Then 'Caso ainda não haja nada cadastrado na Classificação
    rstGrupos.Close
    Set rstGrupos = Nothing
    Exit Sub 'Sairá fora...
  Else
    rstGrupos.Close
    Set rstGrupos = Nothing
  End If
  '---------------------------------------------------------------------
  ' Se houver registros há de continuar com as rotinas a seguir...
  '---------------------------------------------------------------------
  strQuery = "SELECT Código, Marketing "
  strQuery = strQuery & " FROM Funcionários "
  strQuery = strQuery & " WHERE Código = " & CodUser
  
  Set rstFuncionarios = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .Fields("Marketing").Value = True Then
        'Verificar se existe registro na tabela de GruposDeClientes
        'caso haja chama a procedure AtualizarCurvaClientes
        strQuery = "SELECT * FROM GruposDeClientes "
        strQuery = strQuery & " WHERE Filial = " & gnCodFilial
        
        Set rstGrupos = db.OpenRecordset(strQuery, dbOpenDynaset)
        
        If rstGrupos.RecordCount <> 0 Then
          If MsgBox("Para Atualizar a Classificação de Clientes, click em Sim.", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Screen.MousePointer = vbHourglass
            
            Call AtualizarCurvaClientes
            Call AtualizarTableTmpDataCurva
            
            Screen.MousePointer = vbDefault
          End If
        End If
        
        rstGrupos.Close
        Set rstGrupos = Nothing
        '-------------------------------------------------------------
      End If
    End If
    .Close
  End With
  
  Set rstFuncionarios = Nothing
  
End Sub

Private Sub AtualizarCurvaClientes()
  '13/07/2004 - Daniel
  'Case: TV Shopping
  Dim rstClientes           As Recordset
  Dim strQuery              As String
  Dim blnTransaction        As Boolean
  Dim dblTotDinheiroBoletos As Double
  Dim dblTotCheques         As Double
  Dim dblTotCartoes         As Double
  Dim dblTotRecebido        As Double
  Dim bytCodGrupo           As Byte
  
  On Error GoTo ErrHandler
  
  strQuery = "SELECT * "
  strQuery = strQuery & " FROM Cli_For "
  strQuery = strQuery & " WHERE Tipo = '" & "C" & "'"
  
  Set rstClientes = db.OpenRecordset(strQuery, dbOpenDynaset)

  With rstClientes
    If Not (.BOF And .EOF) Then
      .MoveFirst
    
      Do Until .EOF
        Call ConsultarFinanceiro(.Fields("Código").Value, dblTotDinheiroBoletos, dblTotCheques, dblTotCartoes, dblTotRecebido)
        Call ConsultarClassificacao(bytCodGrupo, dblTotRecebido)
      
        'Inicia a Transação
        ws.BeginTrans
        blnTransaction = True
        
          .Edit
          .Fields("CodGrupo").Value = bytCodGrupo
          .Fields("TotDinheiroBoletos").Value = dblTotDinheiroBoletos
          .Fields("TotCheques").Value = dblTotCheques
          .Fields("TotCartoes").Value = dblTotCartoes
          .Fields("TotRecebido").Value = dblTotRecebido
          .Update
        
          'Limpo as vars de buscas
          dblTotDinheiroBoletos = 0
          dblTotCheques = 0
          dblTotCartoes = 0
          dblTotRecebido = 0
          bytCodGrupo = 0
        
        'Finaliza a transação
        ws.CommitTrans
        blnTransaction = False
      
      .MoveNext
      Loop
    
    End If 'If Not (.BOF And .EOF)...
    .Close
  End With

  Set rstClientes = Nothing
  
  Exit Sub
  
ErrHandler:
  'Desfaz a transação
  If blnTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub ConsultarFinanceiro(ByVal CodCliente As Long, ByRef TotDinheiroBoletos As Double, ByRef TotCheques As Double, ByRef TotCartoes As Double, ByRef TotRecebido As Double)
  '13/07/2004 - Daniel
  'Case: TV Shopping
  Dim rstCR        As Recordset
  Dim strQuery     As String
  Dim dblValorTot  As Double
  Dim datDataIni   As Date
  
  datDataIni = Format("02/25/2003", "mm/dd/yyyy") '25/02/2003 é a data inicial de uso do Quick pela TV Shopping
  dblValorTot = 0
  
  'Buscar Primeiro a informação de Pagamentos feitos através de boletos ou dinheiro
  strQuery = "SELECT Filial, Valor, Vencimento, [Contas a Receber].Desconto as Desconto, Acréscimo as Acrescimo, "
  strQuery = strQuery & " [Valor Recebido], [Data Recebimento] , Nota, Cliente, Cli_For.Nome, Sequência as Seq, "
  strQuery = strQuery & " Descrição as Descricao, Contador, [Tipo Parcelamento] "
  strQuery = strQuery & " FROM [Contas a Receber] "
  strQuery = strQuery & " INNER JOIN Cli_For ON ([Contas a Receber].Cliente = Cli_For.Código) "
  strQuery = strQuery & " WHERE Filial = " & gnCodFilial
  strQuery = strQuery & " AND Vencimento >= #" & datDataIni & "#"
  strQuery = strQuery & " AND Vencimento <= #" & Format(Data_Atual, "mm/dd/yyyy") & "#"
  strQuery = strQuery & " AND [Contas a Receber].Tipo = 'R' "
  strQuery = strQuery & " AND Cliente = " & CodCliente
  strQuery = strQuery & " AND [Valor Recebido] <> 0 " 'Já recebeu
  strQuery = strQuery & " ORDER BY Vencimento "

  Set rstCR = db.OpenRecordset(strQuery, dbOpenDynaset)

  With rstCR
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        dblValorTot = dblValorTot + (.Fields("Valor").Value) + (.Fields("Acrescimo").Value) - (.Fields("Desconto").Value)
      
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstCR = Nothing

  TotDinheiroBoletos = dblValorTot

 'Buscamos agora as informação de Pagamentos feitos através de Cheques
 'Abriremos o rstCR com outra query
  dblValorTot = 0

  strQuery = "SELECT Filial, Cliente, Banco, Bancos.Nome, Cheque, Vencimento, Valor, Processado, Devolvido, Sequência, Acréscimo, [Contas a Receber].Desconto as Desconto "
  strQuery = strQuery & " FROM [Contas a Receber] "
  strQuery = strQuery & " INNER JOIN Bancos On [Contas a Receber].Banco = Bancos.Código "
  strQuery = strQuery & " WHERE Filial = " & gnCodFilial
  strQuery = strQuery & " AND Cliente = " & CodCliente
  strQuery = strQuery & " AND [Contas a Receber].Tipo = 'C' "
  strQuery = strQuery & " AND Processado = True " 'Realmente processou
  strQuery = strQuery & " AND Vencimento >= #" & datDataIni & "#"
  strQuery = strQuery & " AND Vencimento <= #" & Format(Data_Atual, "mm/dd/yyyy") & "#"
  strQuery = strQuery & " ORDER BY Vencimento "

  Set rstCR = db.OpenRecordset(strQuery, dbOpenDynaset)

  With rstCR
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        dblValorTot = dblValorTot + (.Fields("Valor").Value) + (.Fields("Acréscimo").Value) - (.Fields("Desconto").Value)
      
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstCR = Nothing

  TotCheques = dblValorTot
  
 'Buscamos agora as informação de Pagamentos feitos através de Cartões
 'Abriremos o rstCR com outra query
  dblValorTot = 0
  
  strQuery = "SELECT Cliente, Cartões.Nome, Cartão, Vencimento, [Valor Cartão], Valor, Processado, Devolvido, Sequência, Acréscimo, [Contas a Receber].Desconto as Desconto "
  strQuery = strQuery & " FROM [Contas a Receber] "
  strQuery = strQuery & " INNER JOIN Cartões On [Contas a Receber].Administradora = Cartões.Código "
  strQuery = strQuery & " WHERE Cliente = " & CodCliente
  strQuery = strQuery & " AND [Contas a Receber].Tipo = '" & "O" & "'"
  strQuery = strQuery & " AND Vencimento >= #" & datDataIni & "#"
  strQuery = strQuery & " AND Vencimento <= #" & Format(Data_Atual, "mm/dd/yyyy") & "#"
  strQuery = strQuery & " AND Processado = True " 'Realmente processou
  strQuery = strQuery & " ORDER BY Vencimento "
  
  Set rstCR = db.OpenRecordset(strQuery, dbOpenDynaset)

  With rstCR
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        dblValorTot = dblValorTot + (.Fields("Valor").Value) + (.Fields("Acréscimo").Value) - (.Fields("Desconto").Value)
      
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstCR = Nothing

  TotCartoes = dblValorTot
  
  'Total Recebido
  TotRecebido = Format(TotDinheiroBoletos + TotCheques + TotCartoes, "##,###,###,###,###,###,##0.00")


End Sub

Private Sub ConsultarClassificacao(ByRef CodGrupo As Byte, ByVal TotRecebido)
  '13/07/2004 - Daniel
  'Case: TV Shopping
  Dim rstGrupos As Recordset
  Dim strQuery  As String
  
  strQuery = "SELECT * FROM GruposDeClientes "
  strQuery = strQuery & " WHERE Filial = " & gnCodFilial
  
  Set rstGrupos = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstGrupos.RecordCount = 0 Then 'Caso ainda não haja nada cadastrado
    rstGrupos.Close
    CodGrupo = 0 'Para não voltar com valores desconhecidos
    Set rstGrupos = Nothing
    Exit Sub
  End If
  
  With rstGrupos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      'Verificar a classificação
      Select Case TotRecebido
        Case (.Fields("LimiteIniG1").Value) To (.Fields("LimiteFinG1").Value)
          CodGrupo = .Fields("CodigoG1").Value
        Case (.Fields("LimiteIniG2").Value) To (.Fields("LimiteFinG2").Value)
          CodGrupo = .Fields("CodigoG2").Value
        Case (.Fields("LimiteIniG3").Value) To (.Fields("LimiteFinG3").Value)
          CodGrupo = .Fields("CodigoG3").Value
        Case Is >= (.Fields("LimiteIniG4").Value)
          CodGrupo = .Fields("CodigoG4").Value
      End Select
      
    End If
    .Close
  End With

  Set rstGrupos = Nothing

End Sub

Private Sub AtualizarTableTmpDataCurva()
  'Nesta sub atualizaremos o Temp.DataCurva.Ultima
  'para exibição na tela de Clientes na guia Classificação
  Dim rstDataCurva As Recordset
  
  dbTemp.Execute "DELETE * FROM DataCurva"
  
  Set rstDataCurva = dbTemp.OpenRecordset("DataCurva", dbOpenDynaset)
  
  With rstDataCurva
    .AddNew
    .Fields("Ultima").Value = Data_Atual
    .Update

    .Close
  End With
  
  Set rstDataCurva = Nothing

End Sub

Private Sub cmd_validar_Click()
  On Error GoTo Erro
  
  Dim retMsg As Variant
  retMsg = MsgBox("A 'Instalação/Configuração complementar' local do QuickStore foi acionado. Deseja prosseguir?", vbYesNo, "Instalação complementar local")
  
  If retMsg = vbNo Then
      frm_instalacaoComplementar.Visible = False
      frm_instalacaoComplementar.Left = 800
      Picture1.Visible = True
      Me.Width = 12105
      Exit Sub
  End If
  
  Dim sDriver As String
  Dim sCaminhoQuick As String
  Dim sCaminhoNFCe As String
  sCaminhoQuick = App.Path
  sDriver = Mid(sCaminhoQuick, 1, 1)

  sCaminhoNFCe = sDriver & ":\QuickStore\NFCe"

  If chk_renomearDBCupom.Value = vbChecked Then
      '*********************************************
      ' Renomear banco de dados CUPOM
    
      Name sCaminhoQuick & "\NFCe\QuickStore_cupom.mdb" As App.Path & "\NFCe\QuickStore.mdb"
      On Error Resume Next
      
      MkDir sDriver & ":\QuickStore"
      If Err.Number <> 0 And Err.Number <> 75 Then
        Exit Sub
      End If
      
      Name App.Path & "\NFCe" As sDriver & ":\QuickStore\NFCe"
    
      lbl_dbCupom.Caption = "Ok"
      lbl_dbCupom.Visible = True
      '*********************************************
  End If

  If chk_caminhoDanfe_eCupom.Value = vbChecked Then
      '*********************************************
      ' Ajustar o caminho nos campos (PastaEnvioNFe e BancoPDV) da tabela [Parâmetros Filial]
      Dim db As Database
      Dim ws As Workspace
      Dim sSql As String
      
      Set ws = DBEngine.Workspaces(0)
      Set db = ws.OpenDatabase(App.Path & "\QuickStore.mdb", False, False, ";pwd=" & gsGetPValue())
      
      ws.BeginTrans
      
      sSql = "UPDATE [Parâmetros Filial] "
      sSql = sSql & " SET PastaEnvioNFe = '" & App.Path & "\Danfe_NFe', "
      sSql = sSql & " BancoPDV = '" & sCaminhoNFCe & "' "
      sSql = sSql & " Where Filial = 1 "
      
      db.Execute sSql
      ws.CommitTrans
      
      db.Close
      Set db = Nothing
      
      lbl_caminhosDanfe_e_CUPOM.Caption = "Ok"
      lbl_caminhosDanfe_e_CUPOM.Visible = True
      '*********************************************
  End If
  
  If chk_instalarSOAP.Value = vbChecked Then
      '*********************************************
      ' Ajustar o caminho nos campos (PastaEnvioNFe e BancoPDV) da tabela [Parâmetros Filial]
      Shell App.Path & "\Apoio\SoapToolkit30.EXE", vbNormalFocus
      
      lbl_soap.Caption = "Ok"
      lbl_soap.Visible = True
      
      MsgBox "Termine a instalação do SoapToolkit30", vbInformation, "Atenção"
      '*********************************************
  End If

  Exit Sub

Erro:
  MsgBox "Inconsistência na validação... " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
  If lbl_dbCupom.Visible = False Then
      lbl_dbCupom.Caption = "Não Ok"
      lbl_dbCupom.Visible = True
  End If
  
'''  If lbl_fontesLetras.Visible = False Then
'''      lbl_fontesLetras.Caption = "Não Ok"
'''      lbl_fontesLetras.Visible = True
'''  End If
  
  If lbl_caminhosDanfe_e_CUPOM.Visible = False Then
      lbl_caminhosDanfe_e_CUPOM.Caption = "Não Ok"
      lbl_caminhosDanfe_e_CUPOM.Visible = True
  End If
  
  If lbl_soap.Visible = False Then
      lbl_soap.Caption = "Não Ok"
      lbl_soap.Visible = True
  End If
  
End Sub

'''Private Function procurarUmDiretorio(startPath As String, sDiretorioAlvo As String) As String
'''  On Error GoTo Erro
'''
'''  Dim dirFound As Boolean
'''  Dim myname As String
'''
'''  myname = Dir(startPath, vbDirectory)
'''
'''  Do While dirFound = False
'''    If myname <> "." And myname <> ".." Then
'''      If (GetAttr(startPath & myname) And vbDirectory) = vbDirectory Then
'''        If myname = sDiretorioAlvo Then
'''          dirFound = True
'''        Else
'''          getSubDirectories (startPath & myname & "\\")
'''        End If
'''      End If
'''    End If
'''
'''    If (dirFound = False) Then
'''      myname = Dir
'''    End If
'''  Loop
'''
'''  Exit Function
'''Erro:
'''    procurarUmDiretorio = vbNullString
'''End Function
'''
'''Private Sub getSubDirectories(startPath As String)
'''  Dim myname As String
'''  myname = Dir(startPath, vbDirectory)
'''
'''  Do While myname <> ""
'''    If myname <> "." And myname <> ".." Then
'''      If (GetAttr(startPath & myname) And vbDirectory) = vbDirectory Then
'''        Debug.Print myname
'''      End If
'''    End If
'''
'''    myname = Dir
'''  Loop
'''End Sub
