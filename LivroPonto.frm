VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPonto 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Apontamento de Tarefas"
   ClientHeight    =   7050
   ClientLeft      =   3750
   ClientTop       =   2280
   ClientWidth     =   7995
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
   HelpContextID   =   1270
   Icon            =   "LivroPonto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7050
   ScaleWidth      =   7995
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   120
      TabIndex        =   33
      Top             =   495
      Width           =   5295
      Begin VB.TextBox Senha 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
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
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   10
         PasswordChar    =   "•"
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Func 
         Bindings        =   "LivroPonto.frx":4E95A
         DataSource      =   "Data1"
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
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
         Columns(0).Width=   3200
         _ExtentX        =   2778
         _ExtentY        =   609
         _StockProps     =   93
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
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Funcionário"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Nome_Func 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1800
         TabIndex        =   36
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         Caption         =   "Senha"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1830
         TabIndex        =   35
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Hoje 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Semana 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1800
         TabIndex        =   34
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   25
      Top             =   3060
      Width           =   7725
      Begin VB.CommandButton B_Confirma_A 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Confirmar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Confirmar Alterações"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.CommandButton B_Cancela_A 
         BackColor       =   &H00C0FFFF&
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3900
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancela Alterações"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3585
      End
      Begin MSMask.MaskEdBox A_Sai_Extra 
         Height          =   345
         Left            =   4830
         TabIndex        =   10
         Top             =   1500
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox A_Entra_Extra 
         Height          =   345
         Left            =   3600
         TabIndex        =   9
         Top             =   1500
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox A_Sai_Noite 
         Height          =   345
         Left            =   4830
         TabIndex        =   8
         Top             =   1140
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox A_Entra_Noite 
         Height          =   345
         Left            =   3600
         TabIndex        =   7
         Top             =   1140
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox A_Sai_Tarde 
         Height          =   345
         Left            =   4830
         TabIndex        =   6
         Top             =   780
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox A_Entra_Tarde 
         Height          =   345
         Left            =   3600
         TabIndex        =   5
         Top             =   780
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox A_Sai_Manhã 
         Height          =   345
         Left            =   4830
         TabIndex        =   4
         Top             =   420
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox A_Entra_Manhã 
         Height          =   345
         Left            =   3600
         TabIndex        =   3
         Top             =   420
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Período 1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   450
         Width           =   705
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Período 4"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1530
         Width           =   750
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Período 3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1170
         Width           =   750
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "Período 2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   810
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Início"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1140
         TabIndex        =   28
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Fim"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Entra_Manhã 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1140
         TabIndex        =   17
         ToolTipText     =   "Dê duplo-clique para ""bater"" o ponto."
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Sai_Manhã 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2400
         TabIndex        =   18
         ToolTipText     =   "Dê duplo-clique para ""bater"" o ponto."
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Entra_Tarde 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1140
         TabIndex        =   19
         ToolTipText     =   "Dê duplo-clique para ""bater"" o ponto."
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Sai_Tarde 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2400
         TabIndex        =   20
         ToolTipText     =   "Dê duplo-clique para ""bater"" o ponto."
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Entra_Noite 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1140
         TabIndex        =   21
         ToolTipText     =   "Dê duplo-clique para ""bater"" o ponto."
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Sai_Noite 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2400
         TabIndex        =   22
         ToolTipText     =   "Dê duplo-clique para ""bater"" o ponto."
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Entra_Extra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1140
         TabIndex        =   23
         ToolTipText     =   "Dê duplo-clique para ""bater"" o ponto."
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label Sai_Extra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2400
         TabIndex        =   24
         ToolTipText     =   "Dê duplo-clique para ""bater"" o ponto."
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Formato  hh:mm"
         Height          =   255
         Left            =   3780
         TabIndex        =   26
         Top             =   180
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.CommandButton B_Altera 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Alterar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6030
      Width           =   7695
   End
   Begin VB.CommandButton B_Limpa 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6525
      Width           =   7695
   End
   Begin VB.CommandButton B_Grava 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Gravar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5535
      Width           =   7695
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Funcionários"
      Top             =   8880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2460
      Left            =   5520
      TabIndex        =   0
      Top             =   570
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   4339
      _Version        =   393216
      ForeColor       =   -2147483647
      BackColor       =   -2147483626
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483624
      ScrollRate      =   1
      StartOfWeek     =   72613889
      TitleBackColor  =   -2147483624
      TitleForeColor  =   -2147483639
      CurrentDate     =   36172
   End
   Begin VB.Label Label11 
      Caption         =   "* Função apenas de lançamento de esforço em determinadas tarefas. Não é um registro de ponto."
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   120
      TabIndex        =   39
      Top             =   90
      Width           =   7725
   End
End
Attribute VB_Name = "frmPonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFuncionários As Recordset
Dim rsPonto As Recordset
Dim Num_Registro As Variant
Dim Hora As Variant

Private Sub A_Entra_Extra_GotFocus()
  A_Entra_Extra.SelStart = 0
  A_Entra_Extra.SelLength = A_Entra_Extra.MaxLength
End Sub

Private Sub A_Entra_Manhã_GotFocus()
  A_Entra_Manhã.SelStart = 0
  A_Entra_Manhã.SelLength = A_Entra_Manhã.MaxLength
End Sub

Private Sub A_Entra_Noite_GotFocus()
  A_Entra_Noite.SelStart = 0
  A_Entra_Noite.SelLength = A_Entra_Noite.MaxLength
End Sub

Private Sub A_Entra_Tarde_GotFocus()
  A_Entra_Tarde.SelStart = 0
  A_Entra_Tarde.SelLength = A_Entra_Tarde.MaxLength
End Sub

Private Sub A_Sai_Extra_GotFocus()
  A_Sai_Extra.SelStart = 0
  A_Sai_Extra.SelLength = A_Sai_Extra.MaxLength
End Sub

Private Sub A_Sai_Manhã_GotFocus()
  A_Sai_Manhã.SelStart = 0
  A_Sai_Manhã.SelLength = A_Sai_Manhã.MaxLength
End Sub

Private Sub A_Sai_Tarde_GotFocus()
  A_Sai_Tarde.SelStart = 0
  A_Sai_Tarde.SelLength = A_Sai_Tarde.MaxLength
End Sub

Private Sub A_Sai_Noite_GotFocus()
  A_Sai_Noite.SelStart = 0
  A_Sai_Noite.SelLength = A_Sai_Noite.MaxLength
End Sub

Private Sub B_Altera_Click()
 
 If Nome_func.Caption = "" Then
   DisplayMsg "Encontre um funcionário antes."
   Exit Sub
 End If
 
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
 
 A_Entra_Manhã.Visible = True
 A_Sai_Manhã.Visible = True
 A_Entra_Tarde.Visible = True
 A_Sai_Tarde.Visible = True
 A_Entra_Noite.Visible = True
 A_Sai_Noite.Visible = True
 A_Entra_Extra.Visible = True
 A_Sai_Extra.Visible = True
 B_Confirma_A.Visible = True
 B_Cancela_A.Visible = True
 Label3.Visible = True
 
End Sub


Private Sub B_Cancela_A_Click()

 A_Entra_Manhã.Visible = False
 A_Sai_Manhã.Visible = False
 A_Entra_Tarde.Visible = False
 A_Sai_Tarde.Visible = False
 A_Entra_Noite.Visible = False
 A_Sai_Noite.Visible = False
 A_Entra_Extra.Visible = False
 A_Sai_Extra.Visible = False
 B_Confirma_A.Visible = False
 B_Cancela_A.Visible = False
 Label3.Visible = False
 

End Sub

Private Sub B_Confirma_A_Click()

  If A_Entra_Manhã.Text <> "  :  " Then
    If IsDate(A_Entra_Manhã.Text) Then Entra_Manhã.Caption = Format(A_Entra_Manhã.Text, "hh:mm")
  End If
  If A_Sai_Manhã.Text <> "  :  " Then
    If IsDate(A_Sai_Manhã.Text) Then Sai_Manhã.Caption = Format(A_Sai_Manhã.Text, "hh:mm")
  End If
  
  If A_Entra_Tarde.Text <> "  :  " Then
    If IsDate(A_Entra_Tarde.Text) Then Entra_Tarde.Caption = Format(A_Entra_Tarde.Text, "hh:mm")
  End If
  If A_Sai_Tarde.Text <> "  :  " Then
    If IsDate(A_Sai_Tarde.Text) Then Sai_Tarde.Caption = Format(A_Sai_Tarde.Text, "hh:mm")
  End If
  
  If A_Entra_Noite.Text <> "  :  " Then
    If IsDate(A_Entra_Noite.Text) Then Entra_Noite.Caption = Format(A_Entra_Noite.Text, "hh:mm")
  End If
  If A_Sai_Noite.Text <> "  :  " Then
    If IsDate(A_Sai_Noite.Text) Then Sai_Noite.Caption = Format(A_Sai_Noite.Text, "hh:mm")
  End If

  If A_Entra_Extra.Text <> "  :  " Then
    If IsDate(A_Entra_Extra.Text) Then Entra_Extra.Caption = Format(A_Entra_Extra.Text, "hh:mm")
  End If
  If A_Sai_Extra.Text <> "  :  " Then
    If IsDate(A_Sai_Extra.Text) Then Sai_Extra.Caption = Format(A_Sai_Extra.Text, "hh:mm")
  End If

  DisplayMsg "Não se esqueça de gravar."

End Sub

Private Sub B_Grava_Click()
  Dim Erro As Integer
  Dim Horas As Double
  Dim blnInTransaction As Boolean
  
  On Error GoTo Erro:
  
  Rem Verifica código
  If IsNull(Nome_func.Caption) Then Erro = True
  If Not Erro Then If Nome_func.Caption = "" Then Erro = True
  
  If Erro Then
    DisplayMsg "Funcionário não digitado, verifique."
    Combo_Func.SetFocus
    Exit Sub
  End If
  
  If Label3.Visible = False Then
    If CriptografaSenha(Senha.Text) <> rsFuncionários("ValorP") Then
      DisplayMsg "Senha incorreta, verifique."
      Senha.SetFocus
      Exit Sub
    End If
  End If
  
  Rem Calcula o numero de horas
  Horas = 0
  If Entra_Manhã.Caption <> "" And Sai_Manhã.Caption <> "" Then
     Horas = DateDiff("s", Entra_Manhã.Caption, Sai_Manhã.Caption)
  End If
  If Entra_Tarde.Caption <> "" And Sai_Tarde.Caption <> "" Then
     Horas = Horas + DateDiff("s", Entra_Tarde.Caption, Sai_Tarde.Caption)
  End If
  If Entra_Noite.Caption <> "" And Sai_Noite.Caption <> "" Then
     Horas = Horas + DateDiff("s", Entra_Noite.Caption, Sai_Noite.Caption)
  End If
  If Entra_Extra.Caption <> "" And Sai_Extra.Caption <> "" Then
     Horas = Horas + DateDiff("s", Entra_Extra.Caption, Sai_Extra.Caption)
  End If
  
  Horas = (Horas / 3600)  'Passa de segundos para horas
  
  Call StatusMsg("Gravando ...")
  
  ' 18/07/2003 - Maikel
  '              Adicionada transação
  ws.BeginTrans
  blnInTransaction = True
  
  If IsNull(Num_Registro) Then
     rsPonto.AddNew
  Else
     rsPonto.Edit
  End If
  
  rsPonto("Data") = CDate(Hoje.Caption)
  rsPonto("Funcionário") = Val(Combo_Func.Text)
  rsPonto("Entrada Manhã") = Entra_Manhã.Caption
  rsPonto("Saída Manhã") = Sai_Manhã.Caption
  rsPonto("Entrada Tarde") = Entra_Tarde.Caption
  rsPonto("Saída Tarde") = Sai_Tarde.Caption
  rsPonto("Entrada Noite") = Entra_Noite.Caption
  rsPonto("Saída Noite") = Sai_Noite.Caption
  rsPonto("Entrada Extra") = Entra_Extra.Caption
  rsPonto("Saída Extra") = Sai_Extra.Caption
  rsPonto("Horas") = Horas
  
  rsPonto.Update
  
  g_GravaLog Data_Atual, "Hora: " & Time & ", Data: " & Date & _
                         ", Usuário: " & gnUserCode & " - " & gsUserName, "LIVRO PONTO"
                         
  Call StatusMsg("")
  ws.CommitTrans
  blnInTransaction = False
  
  B_Limpa_Click
  
  Exit Sub
  
Erro:
  If blnInTransaction Then
    ws.Rollback
    blnInTransaction = False
  End If
  
  If MsgBox("Erro: " & Err.Number & vbCrLf & vbCrLf & _
             "Descrição: " & Err.Description & vbCrLf & vbCrLf & _
             "Ao atualizar a tabela de livro ponto !", vbCritical + vbRetryCancel) = vbRetry Then
    Resume
  End If
End Sub

Private Sub B_Limpa_Click()
  Combo_Func.Text = ""
  Nome_func.Caption = ""
  Senha.Text = ""
  Entra_Manhã.Caption = ""
  Sai_Manhã.Caption = ""
  Entra_Tarde.Caption = ""
  Sai_Tarde.Caption = ""
  Entra_Noite.Caption = ""
  Sai_Noite.Caption = ""
  Entra_Extra.Caption = ""
  Sai_Extra.Caption = ""

  A_Entra_Manhã.Mask = ""
  A_Entra_Manhã.Text = "  :  "
  A_Entra_Manhã.Mask = "##:##"
    
  A_Sai_Manhã.Mask = ""
  A_Sai_Manhã.Text = "  :  "
  A_Sai_Manhã.Mask = "##:##"
  
  A_Entra_Tarde.Mask = ""
  A_Entra_Tarde.Text = "  :  "
  A_Entra_Tarde.Mask = "##:##"
  
  A_Sai_Tarde.Mask = ""
  A_Sai_Tarde.Text = "  :  "
  A_Sai_Tarde.Mask = "##:##"
  
  A_Entra_Noite.Mask = ""
  A_Entra_Noite.Text = "  :  "
  A_Entra_Noite.Mask = "##:##"
  
  A_Sai_Noite.Mask = ""
  A_Sai_Noite.Text = "  :  "
  A_Sai_Noite.Mask = "##:##"
  
  A_Entra_Extra.Mask = ""
  A_Entra_Extra.Text = "  :  "
  A_Entra_Extra.Mask = "##:##"
  
  A_Sai_Extra.Mask = ""
  A_Sai_Extra.Text = "  :  "
  A_Sai_Extra.Mask = "##:##"



  Call StatusMsg("")
  Num_Registro = Null

  Combo_Func.SetFocus


End Sub

Private Sub Combo_Func_CloseUp()
 Combo_Func.Text = Combo_Func.Columns(2).Text
 Combo_Func_LostFocus
End Sub

Private Sub Combo_Func_LostFocus()
   Nome_func.Caption = ""
   If IsNull(Combo_Func.Text) Or Combo_Func.Text = "" Then Exit Sub
   If Not IsNumeric(Combo_Func.Text) Then Exit Sub
   If Val(Combo_Func.Text) < 0 Then Exit Sub
   If Val(Combo_Func.Text) > 9999 Then Exit Sub

   rsFuncionários.Index = "Código"
   rsFuncionários.Seek "=", Val(Combo_Func.Text)
   If rsFuncionários.NoMatch Then Exit Sub
   Nome_func.Caption = rsFuncionários("Nome")


   Rem Procura dia deste sujeito
   rsPonto.Index = "Funcionário"
   rsPonto.Seek "=", Val(Combo_Func.Text), CDate(Hoje.Caption)
   If rsPonto.NoMatch Then Exit Sub
   Num_Registro = rsPonto.Bookmark

   Rem Mostra Dados
   Entra_Manhã.Caption = Format$(rsPonto("Entrada Manhã"), "hh:mm")
   Sai_Manhã.Caption = Format$(rsPonto("Saída Manhã"), "hh:mm")

   Entra_Tarde.Caption = Format$(rsPonto("Entrada Tarde"), "hh:mm")
   Sai_Tarde.Caption = Format$(rsPonto("Saída Tarde"), "hh:mm")
 
   Entra_Noite.Caption = Format$(rsPonto("Entrada Noite"), "hh:mm")
   Sai_Noite.Caption = Format$(rsPonto("Saída Noite"), "hh:mm")

   Entra_Extra.Caption = Format$(rsPonto("Entrada Extra"), "hh:mm")
   Sai_Extra.Caption = Format$(rsPonto("Saída Extra"), "hh:mm")



End Sub

Private Sub Entra_Extra_DblClick()
  Call StatusMsg("")
  If Entra_Extra.Caption <> "" Then Exit Sub

  Hora = Val(Format(Now, "hh"))
  'If Hora > 18 Then
  '  DisplayMsg "Entrada da tarde deve ser até as 18 horas."
  '  Exit Sub
  'End If

  Entra_Extra.Caption = Format$(Now, "hh:mm")

End Sub

Private Sub Entra_Manhã_DblClick()
  Call StatusMsg("")
  If Entra_Manhã.Caption <> "" Then Exit Sub

'  Hora = Val(Format(Now, "hh"))
'  If Hora > 12 Then
'    DisplayMsg "Entrada da manhã deve ser até as 12 horas."
'    Exit Sub
'  End If

  Entra_Manhã.Caption = Format$(Now, "hh:mm")

End Sub

Private Sub Entra_Noite_DblClick()
  Call StatusMsg("")
  If Entra_Noite.Caption <> "" Then Exit Sub

'  Hora = Val(Format(Now, "hh"))
'  If Hora < 17 Then
'    DisplayMsg "Entrada da noite deve ser após as 17 horas."
'    Exit Sub
'  End If

  Entra_Noite.Caption = Format$(Now, "hh:mm")
End Sub

Private Sub Entra_Tarde_DblClick()
  Call StatusMsg("")
  If Entra_Tarde.Caption <> "" Then Exit Sub

'  Hora = Val(Format(Now, "hh"))
'  If Hora > 18 Then
'    DisplayMsg "Entrada da tarde deve ser até as 18 horas."
'    Exit Sub
'  End If
'
'  If Hora < 11 Then
'    DisplayMsg "Entrada da tarde deve ser após as 11 horas."
'    Exit Sub
'  End If

  Entra_Tarde.Caption = Format$(Now, "hh:mm")

End Sub

Private Sub Form_Load()
  Dim Dia As Integer

  Call CenterForm(Me)
  
  MonthView1.Value = Format(Date, "dd/mm/yyyy")
  
  Set rsFuncionários = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsPonto = db.OpenRecordset("Livro Ponto")
  Num_Registro = Null

  Data1.DatabaseName = gsQuickDBFileName
  
  Hoje.Caption = Format$(Now, "dd/mm/yyyy")
  Dia = Weekday(Now)
  Semana.Caption = WeekdayName(Weekday(Hoje.Caption, vbUseSystemDayOfWeek), False, vbUseSystemDayOfWeek)
'  If Dia = 1 Then Semana.Caption = "Domingo"
'  If Dia = 2 Then Semana.Caption = "Segunda"
'  If Dia = 3 Then Semana.Caption = "Terça"
'  If Dia = 4 Then Semana.Caption = "Quarta"
'  If Dia = 5 Then Semana.Caption = "Quinta"
'  If Dia = 6 Then Semana.Caption = "Sexta"
'  If Dia = 7 Then Semana.Caption = "Sábado"

  'DisplayMsg "Dê um duplo-clique sobre o período desejado e pressione GRAVAR."

  Entra_Manhã.Caption = ""
  Sai_Manhã.Caption = ""
  Entra_Tarde.Caption = ""
  Sai_Tarde.Caption = ""
  Entra_Noite.Caption = ""
  Sai_Noite.Caption = ""
  Entra_Extra.Caption = ""
  Sai_Noite.Caption = ""
End Sub

Private Sub Sai_Extra_DblClick()
  Call StatusMsg("")
  If Sai_Extra.Caption <> "" Then Exit Sub
  If Entra_Extra.Caption = "" Then
     DisplayMsg "Início período 4 não preenchido."
     Exit Sub
  End If

  Sai_Extra.Caption = Format$(Now, "hh:mm")
End Sub

Private Sub Sai_Manhã_DblClick()

  Call StatusMsg("")
  If Sai_Manhã.Caption <> "" Then Exit Sub
  If Entra_Manhã.Caption = "" Then
     DisplayMsg "Início período 1 não preenchido."
     Exit Sub
  End If


  Sai_Manhã.Caption = Format$(Now, "hh:mm")

End Sub

Private Sub Sai_Noite_DblClick()

  Call StatusMsg("")
  If Sai_Noite.Caption <> "" Then Exit Sub
  If Entra_Noite.Caption = "" Then
     DisplayMsg "Início período 3 não preenchido."
     Exit Sub
  End If


  Sai_Noite.Caption = Format$(Now, "hh:mm")

End Sub

Private Sub Sai_Tarde_DblClick()
  Call StatusMsg("")
  If Sai_Tarde.Caption <> "" Then Exit Sub
  If Entra_Tarde.Caption = "" Then
     DisplayMsg "Início período 2 não preenchido."
     Exit Sub
  End If


  Sai_Tarde.Caption = Format$(Now, "hh:mm")


End Sub


