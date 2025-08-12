VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProgramacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programações"
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
   Icon            =   "frmProgramacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11760
   Begin VB.Frame fraFaturame 
      Caption         =   "Parcelamento"
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
      Height          =   2865
      Left            =   5880
      TabIndex        =   98
      Top             =   3240
      Width           =   5655
      Begin VB.CommandButton cmdFaturamento 
         BackColor       =   &H0080C0FF&
         Caption         =   "Confirmar &Recebimento"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   2400
         Width           =   2235
      End
      Begin VB.CheckBox chkCancel1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   106
         Top             =   768
         Width           =   255
      End
      Begin VB.CheckBox chkCancel2 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   108
         Top             =   1176
         Width           =   255
      End
      Begin VB.CheckBox chkCancel3 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   110
         Top             =   1584
         Width           =   255
      End
      Begin VB.CheckBox chkCancel4 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   111
         Top             =   1992
         Width           =   255
      End
      Begin VB.CheckBox chkStatus4 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3700
         TabIndex        =   103
         Top             =   1992
         Width           =   255
      End
      Begin VB.CheckBox chkStatus3 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3700
         TabIndex        =   101
         Top             =   1584
         Width           =   255
      End
      Begin VB.CheckBox chkStatus2 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3700
         TabIndex        =   99
         Top             =   1176
         Width           =   255
      End
      Begin VB.CheckBox chkStatus1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3700
         TabIndex        =   97
         Top             =   768
         Width           =   255
      End
      Begin VB.TextBox txtValor4 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   49
         Top             =   1992
         Width           =   1095
      End
      Begin VB.TextBox txtValor3 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   47
         Top             =   1584
         Width           =   1095
      End
      Begin VB.TextBox txtValor2 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   45
         Top             =   1176
         Width           =   1095
      End
      Begin VB.TextBox txtValor1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   43
         Top             =   768
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskVencimento1 
         Height          =   315
         Left            =   1500
         TabIndex        =   44
         ToolTipText     =   "Pressione F2 para carregar o calendário"
         Top             =   765
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskVencimento2 
         Height          =   315
         Left            =   1500
         TabIndex        =   46
         ToolTipText     =   "Pressione F2 para carregar o calendário"
         Top             =   1176
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox mskVencimento3 
         Height          =   315
         Left            =   1500
         TabIndex        =   48
         ToolTipText     =   "Pressione F2 para carregar o calendário"
         Top             =   1590
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskVencimento4 
         Height          =   315
         Left            =   1500
         TabIndex        =   50
         ToolTipText     =   "Pressione F2 para carregar o calendário"
         Top             =   1995
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
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
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Cancelado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   4560
         TabIndex        =   109
         Top             =   420
         Width           =   870
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Faturado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3480
         TabIndex        =   107
         Top             =   420
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   1440
         TabIndex        =   105
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor R$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   120
         TabIndex        =   102
         Top             =   420
         Width           =   705
      End
   End
   Begin VB.Frame fraX 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   825
      Left            =   120
      TabIndex        =   94
      Top             =   120
      Width           =   11415
      Begin VB.ComboBox cboMesX 
         BackColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "frmProgramacao.frx":058A
         Left            =   840
         List            =   "frmProgramacao.frx":0597
         TabIndex        =   0
         Top             =   280
         Width           =   975
      End
      Begin VB.TextBox txtProgramacao 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   3720
         MaxLength       =   25
         TabIndex        =   1
         Top             =   280
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mês"
         Height          =   195
         Left            =   300
         TabIndex        =   96
         Top             =   345
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Programação"
         Height          =   195
         Left            =   2520
         TabIndex        =   95
         Top             =   345
         Width           =   945
      End
   End
   Begin VB.Frame fraDetalhes 
      Caption         =   "Detalhes da Programação"
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
      Height          =   2865
      Left            =   120
      TabIndex        =   86
      Top             =   3240
      Width           =   5655
      Begin VB.TextBox txtCondicoesPagto 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   42
         Top             =   1992
         Width           =   2775
      End
      Begin VB.ComboBox cboFrequencia 
         Height          =   315
         ItemData        =   "frmProgramacao.frx":05A4
         Left            =   1560
         List            =   "frmProgramacao.frx":05B1
         TabIndex        =   40
         Top             =   1176
         Width           =   855
      End
      Begin VB.TextBox txtDuracao 
         Height          =   315
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   41
         Top             =   1584
         Width           =   615
      End
      Begin VB.TextBox txtFaixaIni 
         Height          =   315
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   38
         Top             =   768
         Width           =   1215
      End
      Begin VB.TextBox txtFaixaFin 
         Height          =   315
         Left            =   3120
         MaxLength       =   7
         TabIndex        =   39
         Top             =   768
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskPeriodoIni 
         Height          =   315
         Left            =   1545
         TabIndex        =   36
         ToolTipText     =   "Pressione F2 para carregar o calendário"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   65535
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
      Begin MSMask.MaskEdBox mskPeriodoFin 
         Height          =   315
         Left            =   3120
         TabIndex        =   37
         ToolTipText     =   "Pressione F2 para carregar o calendário"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   65535
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
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Cond. Pagto."
         Height          =   195
         Left            =   360
         TabIndex        =   100
         Top             =   2052
         Width           =   960
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Período"
         Height          =   195
         Left            =   780
         TabIndex        =   93
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   2895
         TabIndex        =   92
         Top             =   420
         Width           =   90
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Faixa"
         Height          =   195
         Left            =   930
         TabIndex        =   91
         Top             =   828
         Width           =   390
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   2880
         TabIndex        =   90
         Top             =   828
         Width           =   90
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Frequência"
         Height          =   195
         Left            =   525
         TabIndex        =   89
         Top             =   1236
         Width           =   795
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Duração"
         Height          =   195
         Left            =   720
         TabIndex        =   88
         Top             =   1644
         Width           =   600
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "(seg)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2175
         TabIndex        =   87
         Top             =   1644
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Grade de Dias"
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
      Height          =   2150
      Left            =   120
      TabIndex        =   51
      Top             =   960
      Width           =   11415
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   15
         Left            =   10920
         MaxLength       =   2
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   2
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   16
         Left            =   120
         MaxLength       =   2
         TabIndex        =   18
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   840
         MaxLength       =   2
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   8760
         MaxLength       =   2
         TabIndex        =   14
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   13
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   14
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   17
         Left            =   840
         MaxLength       =   2
         TabIndex        =   19
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   18
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   19
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   21
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   20
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   22
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   21
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   23
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   22
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   24
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   23
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   24
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   26
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   25
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   27
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   26
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   28
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   27
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   29
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   28
         Left            =   8760
         MaxLength       =   2
         TabIndex        =   30
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   29
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   31
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   30
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   32
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtTotalInsercoes 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   33
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtValorUnitario 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3720
         MaxLength       =   8
         TabIndex        =   34
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtValorTotal 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   35
         ToolTipText     =   "Valor Total do Mês (Programação)"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "01"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "02"
         Height          =   255
         Left            =   840
         TabIndex        =   84
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "03"
         Height          =   255
         Left            =   1560
         TabIndex        =   83
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "04"
         Height          =   255
         Left            =   2280
         TabIndex        =   82
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "05"
         Height          =   255
         Left            =   3000
         TabIndex        =   81
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "06"
         Height          =   255
         Left            =   3720
         TabIndex        =   80
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "07"
         Height          =   255
         Left            =   4440
         TabIndex        =   79
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "08"
         Height          =   255
         Left            =   5160
         TabIndex        =   78
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "09"
         Height          =   255
         Left            =   5880
         TabIndex        =   77
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "10"
         Height          =   255
         Left            =   6600
         TabIndex        =   76
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "11"
         Height          =   255
         Left            =   7320
         TabIndex        =   75
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "12"
         Height          =   255
         Left            =   8040
         TabIndex        =   74
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "13"
         Height          =   255
         Left            =   8760
         TabIndex        =   73
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "14"
         Height          =   255
         Left            =   9480
         TabIndex        =   72
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "15"
         Height          =   255
         Left            =   10200
         TabIndex        =   71
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "16"
         Height          =   255
         Left            =   10920
         TabIndex        =   70
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "17"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "18"
         Height          =   255
         Left            =   840
         TabIndex        =   68
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "19"
         Height          =   255
         Left            =   1560
         TabIndex        =   67
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "20"
         Height          =   255
         Left            =   2280
         TabIndex        =   66
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "21"
         Height          =   255
         Left            =   3000
         TabIndex        =   65
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "22"
         Height          =   255
         Left            =   3720
         TabIndex        =   64
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "23"
         Height          =   255
         Left            =   4440
         TabIndex        =   63
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "24"
         Height          =   255
         Left            =   5160
         TabIndex        =   62
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "25"
         Height          =   255
         Left            =   5880
         TabIndex        =   61
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "26"
         Height          =   255
         Left            =   6600
         TabIndex        =   60
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "27"
         Height          =   255
         Left            =   7320
         TabIndex        =   59
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "28"
         Height          =   255
         Left            =   8040
         TabIndex        =   58
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "29"
         Height          =   255
         Left            =   8760
         TabIndex        =   57
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   255
         Left            =   9480
         TabIndex        =   56
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "31"
         Height          =   255
         Left            =   10200
         TabIndex        =   55
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label36 
         Caption         =   "Total de Inserções"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label37 
         Caption         =   "Valor Unit."
         Height          =   255
         Left            =   2880
         TabIndex        =   53
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total do Mês"
         Height          =   195
         Left            =   5040
         TabIndex        =   52
         Top             =   1740
         Width           =   1320
      End
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   120
      Top             =   7080
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
      Bands           =   "frmProgramacao.frx":05C0
   End
End
Attribute VB_Name = "frmProgramacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private varNumRegistro  As Variant
Private rstProgramacao  As Recordset
Private m_intDia(30)    As Integer
Private m_strAuxi       As String
Private m_QtdeParcelas  As Integer
Private m_lngCliente    As Long

Private Sub cboMesX_LostFocus()
  If Not IsNumeric(cboMesX.Text) Then Exit Sub
  
  rstProgramacao.FindFirst "[Num Autorizacao] = " & frmAutorizacaoPublicidade.g_lngNumAutorizacao & " AND [MesX] = " & CInt(cboMesX.Text)
  If Not rstProgramacao.NoMatch Then
    Call ShowRecord
  Else
    varNumRegistro = Null
  End If

End Sub

Private Sub cmdFaturamento_Click()
  Dim rstProgramacao As Recordset
  Dim strQuery       As String
  Dim dblValor       As Double
  
  If Not IsNumeric(cboMesX.Text) Then Exit Sub
  
  If IsNull(varNumRegistro) Then
    MsgBox "Grave esta Programação antes de Confirmar o Recebimento.", vbExclamation, "Quick Store"
    Exit Sub
  End If
  
  If VerificarFaturamento Then
    MsgBox "Esta Programação já foi parcelada.", vbExclamation, "Quick Store"
    Exit Sub
  End If
    
  m_QtdeParcelas = 0
    
  If IsDate(mskVencimento1.Text) Then m_QtdeParcelas = m_QtdeParcelas + 1
  If IsDate(mskVencimento2.Text) Then m_QtdeParcelas = m_QtdeParcelas + 1
  If IsDate(mskVencimento3.Text) Then m_QtdeParcelas = m_QtdeParcelas + 1
  If IsDate(mskVencimento4.Text) Then m_QtdeParcelas = m_QtdeParcelas + 1

  Call StatusMsg("Criando saídas, saídas serviços e contas correntes...")
  
  Screen.MousePointer = vbHourglass
  
  'Passamos o número da autorização e o mêsX
  'Conforme instrução da STC cada Programação corresponderá uma
  'Nota Fiscal (um faturamento); Dentro de cada Programação as Parcelas
  'corresponderão o Conta Corrente
  dblValor = CDbl(txtValorTotal.Text)
  Call CriarSaidas(frmAutorizacaoPublicidade.g_lngNumAutorizacao, CInt(cboMesX.Text), dblValor, frmAutorizacaoPublicidade.g_intVendedor)
  '-------------------------------------------------------------------------------------------------------------------------------------
  
  m_QtdeParcelas = 0
  
  Screen.MousePointer = vbDefault
    
  '------------------------------
  'Atualizando os campos StatusX
  '------------------------------
  strQuery = "SELECT [Num Autorizacao], MesX, Status1, Status2, Status3, Status4 "
  strQuery = strQuery & " FROM Programacao "
  strQuery = strQuery & " WHERE [Num Autorizacao] = " & frmAutorizacaoPublicidade.g_lngNumAutorizacao
  strQuery = strQuery & " AND MesX = " & CInt(cboMesX.Text)
  
  Set rstProgramacao = db.OpenRecordset(strQuery, dbOpenDynaset)

  With rstProgramacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Edit
      'StatusX
      If IsDate(mskVencimento1.Text) Then .Fields("Status1").Value = True
      If IsDate(mskVencimento2.Text) Then .Fields("Status2").Value = True
      If IsDate(mskVencimento3.Text) Then .Fields("Status3").Value = True
      If IsDate(mskVencimento4.Text) Then .Fields("Status4").Value = True
      
      .Update
    End If
    .Close
  End With

  Set rstProgramacao = Nothing

  MsgBox "Recebimento Confirmado com sucesso.", vbInformation, "Quick Store"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Set rstProgramacao = db.OpenRecordset("SELECT * FROM Programacao WHERE [Num Autorizacao] = " & frmAutorizacaoPublicidade.g_lngNumAutorizacao & " ORDER BY MesX ", dbOpenDynaset)
  
  Call CarregarLabelFraX
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
  
End Sub

Private Sub CarregarLabelFraX()
  Dim rstClientes As Recordset
  
  If Len(frmAutorizacaoPublicidade.cboCliente.Text) <= 0 Then Exit Sub
  
  Set rstClientes = db.OpenRecordset("SELECT Código, Nome FROM Cli_For WHERE Código = " & CLng(frmAutorizacaoPublicidade.cboCliente.Text), dbOpenDynaset)
  
  With rstClientes
    If Not (.BOF And .EOF) Then
      fraX.Caption = .Fields("Código").Value & " - " & .Fields("Nome").Value & "   -   " & "Nº da Autorização: " & frmAutorizacaoPublicidade.g_lngNumAutorizacao
      m_lngCliente = .Fields("Código").Value
    End If
    .Close
  End With
  
  Set rstClientes = Nothing

End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstProgramacao
    .MoveFirst
    
    If .BOF Then Beep
    If Not .BOF Then Call ShowRecord
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  
  With rstProgramacao
    .MoveLast
    
    If .EOF Then Beep
    If Not .EOF Then Call ShowRecord
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  
  With rstProgramacao
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
  
  With rstProgramacao
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
  
  strAux = "Deseja realmente apagar esta Programação?"
  intResposta = MsgBox(strAux, 20, "ATENÇÃO")
  
  If intResposta = 6 Then
    'Verificar se já não foi faturado
    If rstProgramacao.Fields("Faturado").Value = True Then
      MsgBox "Impossível excluir esta Programação pois ela já foi faturada.", vbExclamation, "Quick Store"
      Exit Sub
    End If
    
    'Verificar se foi cancelada
    If rstProgramacao.Fields("Cancela Contrato").Value = True Then
      MsgBox "Impossível excluir esta Programação pois ela foi cancelada.", vbExclamation, "Quick Store"
      Exit Sub
    End If
  
    rstProgramacao.Delete
    varNumRegistro = Null
    Call ClearScreen
    
  End If
  
End Sub

Private Sub UpdateRecord()
  Dim blnErro   As Boolean
  Dim intDia    As Integer
    
  AtualizaSaldo
    
  On Error GoTo Processa_Erro
  
  'Verificar se os percentuais sobre serviços estão ok
  If Not Percentuais Then Exit Sub
  'Verificar se existe algo cadastrado em ParamFaturameAuto
  If Not ExisteParamFaturameAuto Then Exit Sub
  
  If Not IsNumeric(cboMesX.Text) Then
    MsgBox "Mês de faturamento incorreto, verifique.", vbExclamation, "Quick Store"
    cboMesX.SetFocus
    Exit Sub
  End If
  
  If IsNull(varNumRegistro) Then 'Novo record
  
    With rstProgramacao
      If .RecordCount <> 0 Then 'Já existe alguém com este Num Autorizacao
        .MoveFirst
        
          Do Until .EOF
            If .Fields("MesX").Value = CInt(cboMesX.Text) Then
              MsgBox "Já existe uma Programação com o Número de Autorização " & frmAutorizacaoPublicidade.g_lngNumAutorizacao & " e com o Mês " & CInt(cboMesX.Text) & ".", vbExclamation, "Quick Store"
              Exit Sub
            End If
            
            .MoveNext
          Loop
      End If
    End With
    
  End If
  
  'Tratamento para o campo Valor Total
  If Not IsNumeric(txtValorTotal.Text) Then
    MsgBox "Favor preencher o campo Valor Total.", vbExclamation, "Quick Store"
    Exit Sub
  End If
  
  If blnVerificaCampos Then
    MsgBox "O campo " & m_strAuxi & " não foi preenchido corretamente, verifique.", vbExclamation, "Quick Store"
    
    Select Case m_strAuxi
      Case "Programação"
        txtProgramacao.SetFocus
      Case "Período Inicial"
        mskPeriodoIni.SetFocus
      Case "Período Final"
        mskPeriodoFin.SetFocus
    End Select
    
    Exit Sub
  End If
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
  With rstProgramacao
    If IsNull(varNumRegistro) Then
      .AddNew
      .Fields("Num Autorizacao").Value = frmAutorizacaoPublicidade.g_lngNumAutorizacao
      .Fields("MesX").Value = CInt(cboMesX.Text)
      .Fields("Faturado").Value = False 'Será editado somente depois de confirmar o faturamento
      .Fields("Status1").Value = False  'Em primeiro momento os campos StatusX estarão como False
      .Fields("Status2").Value = False  'tornarão True somente depois de confirmar o recebimento
      .Fields("Status3").Value = False
      .Fields("Status4").Value = False
      'As somas de cancelamentos ocorrem futuramente
      .Fields("SomaCancelamento").Value = 0
    Else
      .Edit
    End If
    
    .Fields("Programacao") = txtProgramacao.Text & ""
    
    For intDia = 0 To 30
      .Fields("Dia " & Right(String(2, "0") & (intDia + 1), 2)) = m_intDia(intDia)   'txtDia(intDia).Text
    Next
    
    If Not IsNumeric(txtTotalInsercoes.Text) Then
      .Fields("Total de Insercoes").Value = 0
    Else
      .Fields("Total de Insercoes").Value = txtTotalInsercoes.Text
    End If
    
    If Not IsNumeric(txtValorUnitario.Text) Then
      .Fields("Valor Unitario").Value = 0
    Else
      .Fields("Valor Unitario").Value = txtValorUnitario.Text
    End If
    
    If Not IsNumeric(txtValorTotal.Text) Then
      .Fields("Valor Total").Value = 0
    Else
      .Fields("Valor Total").Value = txtValorTotal.Text
    End If
    
    .Fields("Periodo Ini").Value = Format(mskPeriodoIni.Text, "dd/mm/yyyy")
    .Fields("Periodo Fin").Value = Format(mskPeriodoFin.Text, "dd/mm/yyyy")
    .Fields("Faixa Ini").Value = Trim(txtFaixaIni.Text) & ""
    .Fields("Faixa Fin").Value = Trim(txtFaixaFin.Text) & ""
    .Fields("Frequencia").Value = cboFrequencia.Text & ""
    .Fields("Duracao").Value = txtDuracao.Text & ""
    Select Case Month(Data_Atual)
      Case 1
        .Fields("Mes").Value = "JAN"
      Case 2
        .Fields("Mes").Value = "FEV"
      Case 3
        .Fields("Mes").Value = "MAR"
      Case 4
        .Fields("Mes").Value = "ABR"
      Case 5
        .Fields("Mes").Value = "MAI"
      Case 6
        .Fields("Mes").Value = "JUN"
      Case 7
        .Fields("Mes").Value = "JUL"
      Case 8
        .Fields("Mes").Value = "AGO"
      Case 9
        .Fields("Mes").Value = "SET"
      Case 10
        .Fields("Mes").Value = "OUT"
      Case 11
        .Fields("Mes").Value = "NOV"
      Case Else
        .Fields("Mes").Value = "DEZ"
    End Select
    .Fields("Condicoes Pagamento").Value = txtCondicoesPagto.Text & ""
    .Fields("Gerar Etiqueta").Value = True
    'Este campo [Cancela Contrato] não terá mais utilidade para a Fase II
    'será criado um campo cancelar para cada faturamento
    .Fields("Cancela Contrato").Value = False
    '--------------------------------------------------------------------
        
    'Valor1
    If IsNumeric(txtValor1.Text) Then
      .Fields("Valor1").Value = CDbl(txtValor1.Text)
      
      m_QtdeParcelas = m_QtdeParcelas + 1
    Else
      .Fields("Valor1").Value = 0
    End If
    'Valor2
    If IsNumeric(txtValor2.Text) Then
      .Fields("Valor2").Value = CDbl(txtValor2.Text)
    Else
      .Fields("Valor2").Value = 0
    End If
    'Valor3
    If IsNumeric(txtValor3.Text) Then
      .Fields("Valor3").Value = CDbl(txtValor3.Text)
    Else
      .Fields("Valor3").Value = 0
    End If
    'Valor4
    If IsNumeric(txtValor4.Text) Then
      .Fields("Valor4").Value = CDbl(txtValor4.Text)
    Else
      .Fields("Valor4").Value = 0
    End If
    
    'Vencimento1
    If IsDate(mskVencimento1.Text) Then .Fields("Vencimento1").Value = CDate(mskVencimento1.Text)
    'Vencimento2
    If IsDate(mskVencimento2.Text) Then .Fields("Vencimento2").Value = CDate(mskVencimento2.Text)
    'Vencimento3
    If IsDate(mskVencimento3.Text) Then .Fields("Vencimento3").Value = CDate(mskVencimento3.Text)
    'Vencimento4
    If IsDate(mskVencimento4.Text) Then .Fields("Vencimento4").Value = CDate(mskVencimento4.Text)
    
    '.Fields("Cancel1").Value
    '.Fields("Cancel2").Value
    '.Fields("Cancel3").Value
    '.Fields("Cancel4").Value
    
    
    .Update
    varNumRegistro = .LastModified
    .Bookmark = varNumRegistro
  End With 'With rstContrato
  
  Call StatusMsg("")
    
  Exit Sub
  
Processa_Erro:
  MsgBox "Erro (" & Err.Number & ") " & Err.Description, vbCritical, "Quick Store"
End Sub

Public Sub ClearScreen()
  Dim ctlControl As control
  
  Call StatusMsg("")
  
  cboMesX.Text = ""
  mskPeriodoIni.Mask = ""
  mskPeriodoIni.Text = ""
  mskPeriodoIni.Mask = "##/##/####"
  
  mskPeriodoFin.Mask = ""
  mskPeriodoFin.Text = ""
  mskPeriodoFin.Mask = "##/##/####"
  
  mskVencimento1.Mask = ""
  mskVencimento1.Text = ""
  mskVencimento1.Mask = "##/##/####"
  
  mskVencimento2.Mask = ""
  mskVencimento2.Text = ""
  mskVencimento2.Mask = "##/##/####"
  
  mskVencimento3.Mask = ""
  mskVencimento3.Text = ""
  mskVencimento3.Mask = "##/##/####"
  
  mskVencimento4.Mask = ""
  mskVencimento4.Text = ""
  mskVencimento4.Mask = "##/##/####"
  
  cboFrequencia.Text = ""
  chkStatus1.Value = vbUnchecked
  chkStatus2.Value = vbUnchecked
  chkStatus3.Value = vbUnchecked
  chkStatus4.Value = vbUnchecked
  
  chkCancel1.Value = vbUnchecked
  chkCancel2.Value = vbUnchecked
  chkCancel3.Value = vbUnchecked
  chkCancel4.Value = vbUnchecked
  
  For Each ctlControl In Controls
    If TypeOf ctlControl Is TextBox Then
      ctlControl.Text = ""
    End If
  Next
  
  varNumRegistro = Null
  
  If Not rstProgramacao.EOF Then
    On Error Resume Next
    rstProgramacao.MoveFirst
    rstProgramacao.MovePrevious
    On Error GoTo 0
  End If
  
  SelectAllText cboMesX, True
  
  
End Sub

Sub ShowRecord()
  Dim intDia     As Integer
  Dim dblRet     As Double
  
  With rstProgramacao
    cboMesX.Text = .Fields("MesX").Value
    'O .Fields("Num Autorizacao").Value fica implícito
    txtProgramacao.Text = .Fields("Programacao") & ""
    
    For intDia = 0 To 30
      txtDia(intDia).Text = .Fields("Dia " & Right(String(2, "0") & intDia + 1, 2))
    Next intDia

    txtTotalInsercoes.Text = .Fields("Total de Insercoes").Value & ""
    txtValorUnitario.Text = .Fields("Valor Unitario").Value & ""
    txtValorTotal.Text = .Fields("Valor Total").Value & ""
    mskPeriodoIni.Text = .Fields("Periodo Ini").Value & ""
    mskPeriodoFin.Text = .Fields("Periodo Fin").Value & ""
    txtFaixaIni.Text = .Fields("Faixa Ini").Value & ""
    txtFaixaFin.Text = .Fields("Faixa Fin").Value & ""
    cboFrequencia.Text = .Fields("Frequencia").Value & ""
    txtDuracao.Text = .Fields("Duracao").Value & ""
    txtCondicoesPagto.Text = .Fields("Condicoes Pagamento").Value & ""
    txtValor1.Text = .Fields("Valor1").Value & ""
    txtValor2.Text = .Fields("Valor2").Value & ""
    txtValor3.Text = .Fields("Valor3").Value & ""
    txtValor4.Text = .Fields("Valor4").Value & ""
    
    If Not IsNull(.Fields("Vencimento1").Value) Then
      mskVencimento1.Text = .Fields("Vencimento1").Value
    Else
      mskVencimento1.Mask = ""
      mskVencimento1.Text = ""
      mskVencimento1.Mask = "##/##/####"
    End If
    
    If Not IsNull(.Fields("Vencimento2").Value) Then
      mskVencimento2.Text = .Fields("Vencimento2").Value
    Else
      mskVencimento2.Mask = ""
      mskVencimento2.Text = ""
      mskVencimento2.Mask = "##/##/####"
    End If
    
    If Not IsNull(.Fields("Vencimento3").Value) Then
      mskVencimento3.Text = .Fields("Vencimento3").Value
    Else
      mskVencimento3.Mask = ""
      mskVencimento3.Text = ""
      mskVencimento3.Mask = "##/##/####"
    End If
    
    If Not IsNull(.Fields("Vencimento4").Value) Then
      mskVencimento4.Text = .Fields("Vencimento4").Value
    Else
      mskVencimento4.Mask = ""
      mskVencimento4.Text = ""
      mskVencimento4.Mask = "##/##/####"
    End If
    
    If .Fields("Status1").Value = True Then
      chkStatus1.Value = vbChecked
    Else
      chkStatus1.Value = vbUnchecked
    End If
    
    If .Fields("Status2").Value = True Then
      chkStatus2.Value = vbChecked
    Else
      chkStatus2.Value = vbUnchecked
    End If
    
    If .Fields("Status3").Value = True Then
      chkStatus3.Value = vbChecked
    Else
      chkStatus3.Value = vbUnchecked
    End If
    
    If .Fields("Status4").Value = True Then
      chkStatus4.Value = vbChecked
    Else
      chkStatus4.Value = vbUnchecked
    End If

    'Cancelamento
    If .Fields("Cancel1").Value Then
      chkCancel1.Value = vbChecked
    Else
      chkCancel1.Value = vbUnchecked
    End If

    If .Fields("Cancel2").Value Then
      chkCancel2.Value = vbChecked
    Else
      chkCancel2.Value = vbUnchecked
    End If

    If .Fields("Cancel3").Value Then
      chkCancel3.Value = vbChecked
    Else
      chkCancel3.Value = vbUnchecked
    End If

    If .Fields("Cancel4").Value Then
      chkCancel4.Value = vbChecked
    Else
      chkCancel4.Value = vbUnchecked
    End If


    varNumRegistro = .Bookmark
  End With
  
  
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
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstProgramacao.Close
  Set rstProgramacao = Nothing
End Sub

Private Function blnVerificaCampos() As Boolean

  m_strAuxi = ""
  
  If Len(txtProgramacao.Text) <= 0 Then
    blnVerificaCampos = True
    m_strAuxi = "Programação"
  End If
  
  If Not IsDate(mskPeriodoIni.Text) Then
    blnVerificaCampos = True
    m_strAuxi = "Período Inicial"
  End If
  
  If Not IsDate(mskPeriodoFin.Text) Then
    blnVerificaCampos = True
    m_strAuxi = "Período Final"
  End If
  
End Function

Private Sub mskPeriodoFin_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskPeriodoFin.Text = frmCalendario.gsDateCalender(mskPeriodoFin.Text)
  End If
End Sub

Private Sub mskPeriodoFin_LostFocus()
  mskPeriodoFin.Text = Ajusta_Data(mskPeriodoFin.Text)
End Sub

Private Sub mskPeriodoIni_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskPeriodoIni.Text = frmCalendario.gsDateCalender(mskPeriodoIni.Text)
  End If
End Sub

Private Sub mskPeriodoIni_LostFocus()
  mskPeriodoIni.Text = Ajusta_Data(mskPeriodoIni.Text)
End Sub

Private Sub AtualizaSaldo()
  Dim intTotalizador  As Integer
  Dim intBound        As Integer

  For intBound = LBound(m_intDia) To UBound(m_intDia)
    If (IsNumeric(txtDia(intBound).Text)) Then
      m_intDia(intBound) = CInt(txtDia(intBound).Text)
    Else
      m_intDia(intBound) = 0
    End If
  Next intBound

  intTotalizador = 0

  For intBound = LBound(m_intDia) To UBound(m_intDia)
    intTotalizador = intTotalizador + m_intDia(intBound)
  Next intBound

End Sub

Private Sub mskVencimento1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskVencimento1.Text = frmCalendario.gsDateCalender(mskVencimento1.Text)
  End If
End Sub

Private Sub mskVencimento1_LostFocus()
  mskVencimento1.Text = Ajusta_Data(mskVencimento1.Text)
End Sub

Private Sub mskVencimento2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskVencimento2.Text = frmCalendario.gsDateCalender(mskVencimento2.Text)
  End If
End Sub

Private Sub mskVencimento2_LostFocus()
  mskVencimento2.Text = Ajusta_Data(mskVencimento2.Text)
End Sub

Private Sub mskVencimento3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskVencimento3.Text = frmCalendario.gsDateCalender(mskVencimento3.Text)
  End If
End Sub

Private Sub mskVencimento3_LostFocus()
  mskVencimento3.Text = Ajusta_Data(mskVencimento3.Text)
End Sub

Private Sub mskVencimento4_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskVencimento4.Text = frmCalendario.gsDateCalender(mskVencimento4.Text)
  End If
End Sub

Private Sub mskVencimento4_LostFocus()
  mskVencimento4.Text = Ajusta_Data(mskVencimento4.Text)
End Sub

Private Sub CriarSaidas(ByVal NumAutorizacao As Long, ByVal MesX As Integer, ByVal Valor As Double, ByVal Vendedor As Integer)
  'Será criado nesta procedure:
  'Saídas, Saídas - Serviços e o CR
  Dim rstParametros        As Recordset
  Dim rstSaidas            As Recordset
  Dim rstSaidasServicos    As Recordset
  Dim rstParamFaturameAuto As Recordset
  Dim rstServicos          As Recordset
  Dim rstCR                As Recordset
  Dim rstMovimentoParcelas As Recordset
  
  Dim nSequencia           As Long
  Dim blnTransaction       As Boolean
  
  Dim bytFilial            As Byte
  Dim intOperacao          As Integer
  Dim intServico           As Integer
  Dim bytCaixa             As Byte
  Dim strTabela            As String
  Dim dblISS               As Double
  
  Dim strDescrServico      As String
  
  Dim dblCSLL              As Double
  Dim dblCOFINS            As Double
  Dim dblPIS               As Double
  Dim dblIRRF              As Double
  
  Dim dblValorCSLL         As Double
  Dim dblValorCOFINS       As Double
  Dim dblValorPIS          As Double
  Dim dblValorIRRF         As Double
  
  Dim nRet                 As Integer
  Dim bytCont              As Byte
  
  On Error GoTo Err_Handlel
  
  '-------------------------------------
  'Abrir a transação
  '-------------------------------------
  ws.BeginTrans
  blnTransaction = True
        
      '*** Operações com o DB
      
      'Parâmetros Filial
      Set rstParametros = db.OpenRecordset("SELECT Filial, CSLL, COFINS, PIS, IRRF FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset)
      
      With rstParametros
        If Not (.BOF And .EOF) Then
          .MoveFirst
          
          dblCSLL = .Fields("CSLL").Value
          dblCOFINS = .Fields("COFINS").Value
          dblPIS = .Fields("PIS").Value
          dblIRRF = .Fields("IRRF").Value
          
        End If
        .Close
      End With
      
      Set rstParametros = Nothing
      
      'ParamFaturameAuto
      Set rstParamFaturameAuto = db.OpenRecordset("SELECT * FROM ParamFaturameAuto WHERE Filial = " & gnCodFilial, dbOpenDynaset)
      
      With rstParamFaturameAuto
        If Not (.BOF And .EOF) Then
          .MoveFirst
          
          bytFilial = .Fields("Filial").Value
          intOperacao = .Fields("Operacao").Value
          intServico = .Fields("Servico").Value
          bytCaixa = .Fields("Caixa").Value
          strTabela = .Fields("Tabela").Value & ""
          dblISS = .Fields("ISS").Value
                    
        End If
        .Close
      End With
      
      Set rstParamFaturameAuto = Nothing
      
      'Serviços
      Set rstServicos = db.OpenRecordset("SELECT Código, Descrição FROM Serviços WHERE Código = " & intServico, dbOpenDynaset)
            
      With rstServicos
        If Not (.BOF And .EOF) Then
          .MoveFirst
          
          strDescrServico = .Fields("Descrição").Value & ""
        End If
        .Close
      End With
      
      Set rstServicos = Nothing
      
      
      'Buscar uma próxima Sequência
      nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("Última Movimentação") + 1

      'Saídas
      Set rstSaidas = db.OpenRecordset("Saídas", dbOpenDynaset)
      
      With rstSaidas
        .AddNew
        .Fields("Filial").Value = gnCodFilial
        .Fields("Data").Value = Data_Atual
        .Fields("Sequência").Value = nSequencia
        .Fields("Operação").Value = intOperacao
        .Fields("Caixa").Value = bytCaixa
        .Fields("Tabela").Value = strTabela
        .Fields("Digitador").Value = Vendedor
        .Fields("Operador").Value = Vendedor
        .Fields("Cliente").Value = m_lngCliente
        .Fields("Num Autorizacao").Value = NumAutorizacao
        .Fields("MesX").Value = MesX
        .Fields("Observações").Value = "Contrato Nº " & NumAutorizacao & " - Mês: " & MesX
        .Fields("Serviços").Value = Valor
        .Fields("Base ISS").Value = 0
        .Fields("Valor ISS").Value = Format((CDbl(Valor * dblISS / 100)), "##,###,###,##0.00")
        .Fields("Percentual CSLL").Value = dblCSLL
        .Fields("Percentual COFINS").Value = dblCOFINS
        .Fields("Percentual PIS").Value = dblPIS
        .Fields("Percentual IRRF").Value = dblIRRF
        
        '27/12/2007 - Anderson
        'O valor mínimo do cálculo para o IRRF é de R$ 10,00.
        'esta alteração é para considerar este valor.
        'IRRF
        'If Valor > 666 Then
        If CDbl(Valor * dblIRRF / 100) >= 10 Then
        'IRRF
          dblValorIRRF = Format((CDbl(Valor * dblIRRF / 100)), "##,###,###,##0.00")
        Else
          dblValorIRRF = 0
        End If
        .Fields("Total IRRF").Value = dblValorIRRF
        
        'CSLL
         dblValorCSLL = Format((CDbl(Valor * dblCSLL / 100)), "##,###,###,##0.00")
         .Fields("Total CSLL").Value = dblValorCSLL
         
        'COFINS
         dblValorCOFINS = Format((CDbl(Valor * dblCOFINS / 100)), "##,###,###,##0.00")
        .Fields("Total COFINS").Value = dblValorCOFINS
        
        'PIS
        dblValorPIS = Format((CDbl(Valor * dblPIS / 100)), "##,###,###,##0.00")
        .Fields("Total PIS").Value = dblValorPIS
                
        'Este field será o valor menos os impostos sobre serviços
        .Fields("Total").Value = Format((Valor - (dblValorCSLL + dblValorCOFINS + dblValorPIS + dblValorIRRF)), "##,###,###,##0.00")
        
        .Fields("Efetivada").Value = True
        .Fields("Recebimento").Value = True
        .Fields("Nota Impressa").Value = 0
        
        .Update
        .Close
      End With
      
      Set rstSaidas = Nothing
      
      'Saídas - Serviços
      Set rstSaidasServicos = db.OpenRecordset("Saídas - Serviços", dbOpenDynaset)
      
      With rstSaidasServicos
        .AddNew
        .Fields("Filial").Value = gnCodFilial
        .Fields("Sequência").Value = nSequencia
        .Fields("Linha").Value = 1
        .Fields("Código").Value = intServico
        .Fields("Descrição").Value = strDescrServico
        .Fields("Preço").Value = Valor
        .Fields("Tempo").Value = 1
        
        .Update
        .Close
      End With
      
      Set rstSaidasServicos = Nothing

      '-------------------------------------------------------
      'EFETIVA A SAÍDA
      '-------------------------------------------------------
      Call StatusMsg("Aguarde, efetivando venda...")
  
      nRet = Efetiva_Saída(gnCodFilial, nSequencia)
  
      If nRet <> 0 Then
        Select Case nRet
          Case -1
            'Ação cancelada
            Call StatusMsg("Ação cancelada.")
          Case 5
            Call DisplayMsg("Tabela de preços inexistente.")
          Case Else
            Call DisplayMsg("Operação NÃO efetivada. Erro" & str(nRet))
        End Select
        'Cancelamento da transação
        ws.Rollback
        Exit Sub
      End If
      '-------------------------------------------------------
      'FIM DA EFETIVA A SAÍDA
      '-------------------------------------------------------

      '-------------------------------------------------------
      'Neste trecho deve entrar o código para tratamento em CR
      '-------------------------------------------------------
      Set rstCR = db.OpenRecordset("Contas a Receber", dbOpenDynaset)
      
      For bytCont = 1 To m_QtdeParcelas
      
        With rstCR
          .AddNew
          .Fields("Filial").Value = gnCodFilial
          .Fields("Cliente").Value = m_lngCliente
          .Fields("Sequência").Value = nSequencia
          .Fields("Tipo").Value = "R"
          .Fields("Tipo Parcelamento").Value = "C"
          .Fields("Data Emissão").Value = Data_Atual
          .Fields("Valor Recebido").Value = 0
          .Fields("Vendedor").Value = Vendedor
          .Fields("Nota").Value = 0
          .Fields("Processado").Value = False
         
          If bytCont = 1 Then
            .Fields("Vencimento").Value = CDate(mskVencimento1.Text)
            .Fields("Valor").Value = Format((CDbl(txtValor1.Text)), "##,###,###,##0.00")
            .Fields("Descrição").Value = "Parcela " & bytCont & "/" & m_QtdeParcelas
            .Fields("Parcela").Value = bytCont
            .Fields("Fatura").Value = ""
          End If
          
          If bytCont = 2 Then
            .Fields("Vencimento").Value = CDate(mskVencimento2.Text)
            .Fields("Valor").Value = Format((CDbl(txtValor2.Text)), "##,###,###,##0.00")
            .Fields("Descrição").Value = "Parcela " & bytCont & "/" & m_QtdeParcelas
            .Fields("Parcela").Value = bytCont
            .Fields("Fatura").Value = ""
          End If
          
          If bytCont = 3 Then
            .Fields("Vencimento").Value = CDate(mskVencimento3.Text)
            .Fields("Valor").Value = Format((CDbl(txtValor3.Text)), "##,###,###,##0.00")
            .Fields("Descrição").Value = "Parcela " & bytCont & "/" & m_QtdeParcelas
            .Fields("Parcela").Value = bytCont
            .Fields("Fatura").Value = ""
          End If
          
          If bytCont = 4 Then
            .Fields("Vencimento").Value = CDate(mskVencimento4.Text)
            .Fields("Valor").Value = Format((CDbl(txtValor4.Text)), "##,###,###,##0.00")
            .Fields("Descrição").Value = "Parcela " & bytCont & "/" & m_QtdeParcelas
            .Fields("Parcela").Value = bytCont
            .Fields("Fatura").Value = ""
          End If
          
          .Update
        End With
      
      Next bytCont

      rstCR.Close
      Set rstCR = Nothing
      
      'Criando o Movimento Parcelas
      Set rstMovimentoParcelas = db.OpenRecordset("Movimento - Parcelas", dbOpenDynaset)
      
      For bytCont = 1 To m_QtdeParcelas

        With rstMovimentoParcelas
          .AddNew
          .Fields("Filial").Value = gnCodFilial
          .Fields("Sequência").Value = nSequencia
          .Fields("Parcelas").Value = CLng(m_QtdeParcelas)
          
          If bytCont = 1 Then
            .Fields("Bom").Value = CDate(mskVencimento1.Text)
            .Fields("Valor").Value = Format((CDbl(txtValor1.Text)), "##,###,###,##0.00")
            .Fields("Ordem").Value = bytCont
          End If
          
          If bytCont = 2 Then
            .Fields("Bom").Value = CDate(mskVencimento2.Text)
            .Fields("Valor").Value = Format((CDbl(txtValor2.Text)), "##,###,###,##0.00")
            .Fields("Ordem").Value = bytCont
          End If
          
          If bytCont = 3 Then
            .Fields("Bom").Value = CDate(mskVencimento3.Text)
            .Fields("Valor").Value = Format((CDbl(txtValor3.Text)), "##,###,###,##0.00")
            .Fields("Ordem").Value = bytCont
          End If
          
          If bytCont = 4 Then
            .Fields("Bom").Value = CDate(mskVencimento4.Text)
            .Fields("Valor").Value = Format((CDbl(txtValor4.Text)), "##,###,###,##0.00")
            .Fields("Ordem").Value = bytCont
          End If
          
          .Update
        End With
      
      Next bytCont

      rstMovimentoParcelas.Close
      Set rstMovimentoParcelas = Nothing

      '-------------------------------------------------------
      'Editar o campo Programacao.Faturado
      '-------------------------------------------------------
      Call EditarFieldFaturado(frmAutorizacaoPublicidade.g_lngNumAutorizacao, CInt(cboMesX.Text))

      'Tratamento para Atualização de Parâmetros
      Set rstParametros = db.OpenRecordset(" SELECT [Última Movimentação] FROM [Parâmetros Filial] WHERE Filial =" & gnCodFilial, dbOpenDynaset)
  
        rstParametros.Edit
        rstParametros.Fields("Última Movimentação").Value = nSequencia
        rstParametros.Update
        rstParametros.Close
  
      Set rstParametros = Nothing
      'Fim do Tratamento para Atualização de Parâmetros

      '*** Final de Operações com o DB

  '-------------------------------------
  'Fechar a transação
  '-------------------------------------
  ws.CommitTrans
  blnTransaction = False
  
  Call StatusMsg("")
  
  Exit Sub

Err_Handlel:
  If blnTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Erro na transação"
  
End Sub

Private Function ExisteParamFaturameAuto() As Boolean
  Dim rstParamFaturameAuto As Recordset
  
  Set rstParamFaturameAuto = db.OpenRecordset("SELECT * FROM ParamFaturameAuto WHERE Filial = " & gnCodFilial, dbOpenDynaset)
  
  If rstParamFaturameAuto.RecordCount = 0 Then
    ExisteParamFaturameAuto = False
    
    MsgBox "Cadastre um Parâmetro para o Faturamento Automático.", vbExclamation, "Quick Store"
  Else
    ExisteParamFaturameAuto = True
  End If
  
End Function

Private Function Percentuais() As Boolean
  Dim rstParametros As Recordset
  Dim strQuery      As String
  
  Percentuais = True
  
  strQuery = "SELECT Filial, CSLL, COFINS, PIS, IRRF "
  strQuery = strQuery & " FROM [Parâmetros Filial] "
  strQuery = strQuery & " WHERE Filial = " & gnCodFilial
    
  Set rstParametros = db.OpenRecordset(strQuery, dbOpenDynaset)

  If rstParametros.RecordCount = 0 Then
    Percentuais = False
    rstParametros.Close
    Set rstParametros = Nothing
  Else
  
    With rstParametros
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        If Not IsNumeric(.Fields("CSLL").Value) Then Percentuais = False
        If Not IsNumeric(.Fields("COFINS").Value) Then Percentuais = False
        If Not IsNumeric(.Fields("PIS").Value) Then Percentuais = False
        If Not IsNumeric(.Fields("IRRF").Value) Then Percentuais = False
        
      End If
      .Close
    End With
  
    Set rstParametros = Nothing
    
  End If
  
  
  If Not Percentuais Then
    MsgBox "Cadastre em Parâmetros Filial na guia 'Outros' os Percentuais s/ Serviços.", vbExclamation, "Quick Store"
  End If
  
End Function

Private Function VerificarFaturamento() As Boolean
  Dim rstProgramacao As Recordset
  Dim strQuery       As String
  
  strQuery = "SELECT [Num Autorizacao], MesX, Faturado"
  strQuery = strQuery & " FROM Programacao "
  strQuery = strQuery & " WHERE [Num Autorizacao] = " & frmAutorizacaoPublicidade.g_lngNumAutorizacao
  strQuery = strQuery & " AND MesX = " & CInt(cboMesX.Text)
  
  Set rstProgramacao = db.OpenRecordset(strQuery, dbOpenDynaset)

  With rstProgramacao
    If Not (.BOF And .EOF) Then
      If .Fields("Faturado").Value Then
        VerificarFaturamento = True
      End If
    End If
    .Close
  End With

  Set rstProgramacao = Nothing

End Function

Private Sub EditarFieldFaturado(ByVal Numero As Long, ByVal MesX As Integer)
  Dim rstProgramacao As Recordset
  Dim strQuery       As String
  
  strQuery = "SELECT [Num Autorizacao], MesX, Faturado"
  strQuery = strQuery & " FROM Programacao "
  strQuery = strQuery & " WHERE [Num Autorizacao] = " & Numero
  strQuery = strQuery & " AND MesX = " & MesX
  
  Set rstProgramacao = db.OpenRecordset(strQuery, dbOpenDynaset)

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

End Sub
