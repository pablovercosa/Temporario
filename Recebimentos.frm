VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmRecebimento 
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Recebimento"
   ClientHeight    =   8115
   ClientLeft      =   210
   ClientTop       =   420
   ClientWidth     =   13845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1880
   Icon            =   "Recebimentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8115
   ScaleWidth      =   13845
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Caption         =   "Cheques (F3)"
      ForeColor       =   &H00404040&
      Height          =   4725
      Left            =   10140
      TabIndex        =   29
      Top             =   2490
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox Qtde_Cheques 
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
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1290
         Width           =   735
      End
      Begin VB.CommandButton B_Monta_Cheques 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Dividir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3765
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Dividir valor"
         Top             =   1267
         Width           =   1650
      End
      Begin SSDataWidgets_B.SSDBGrid Grade_Cheque 
         Height          =   2535
         Left            =   750
         TabIndex        =   9
         Top             =   1800
         Width           =   4695
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelColorShadow=   15066597
         ForeColorEven   =   4210752
         ForeColorOdd    =   4210752
         BackColorEven   =   12648447
         BackColorOdd    =   15066597
         RowHeight       =   423
         ExtraHeight     =   185
         Columns.Count   =   4
         Columns(0).Width=   1058
         Columns(0).Caption=   "Banco"
         Columns(0).Name =   "Banco"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1931
         Columns(1).Caption=   "Cheque"
         Columns(1).Name =   "Cheque"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1799
         Columns(2).Caption=   "Bom Para"
         Columns(2).Name =   "Bom Para"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   7
         Columns(2).FieldLen=   256
         Columns(3).Width=   2487
         Columns(3).Caption=   "Valor"
         Columns(3).Name =   "Valor"
         Columns(3).Alignment=   1
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).NumberFormat=   "###,###,##0.00"
         Columns(3).FieldLen=   256
         UseDefaults     =   0   'False
         _ExtentX        =   8281
         _ExtentY        =   4471
         _StockProps     =   79
         Caption         =   "Cheques"
         ForeColor       =   4210752
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
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "Qtde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1095
         TabIndex        =   30
         Top             =   1350
         Width           =   585
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E5E5E5&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   30
      TabIndex        =   50
      Top             =   30
      Width           =   4935
      Begin VB.OptionButton opt_outrosPagamentos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   240
         Picture         =   "Recebimentos.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4170
         Width           =   945
      End
      Begin VB.OptionButton opt_cheque 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   240
         Picture         =   "Recebimentos.frx":4F33F
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3390
         Width           =   945
      End
      Begin VB.OptionButton opt_parcelamento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   240
         Picture         =   "Recebimentos.frx":4FCF8
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2610
         Width           =   945
      End
      Begin VB.OptionButton opt_cartaoCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   240
         Picture         =   "Recebimentos.frx":50754
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1830
         Width           =   945
      End
      Begin VB.OptionButton opt_dinheiro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   240
         Picture         =   "Recebimentos.frx":510E2
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1050
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.Label lbl_outrosAtivo 
         BackColor       =   &H00E5E5E5&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFA324&
         Height          =   315
         Left            =   2100
         TabIndex        =   76
         Top             =   4320
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbl_chequeAtivo 
         BackColor       =   &H00E5E5E5&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFA324&
         Height          =   315
         Left            =   2100
         TabIndex        =   75
         Top             =   3570
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbl_parcelamentoAtivo 
         BackColor       =   &H00E5E5E5&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFA324&
         Height          =   315
         Left            =   2100
         TabIndex        =   74
         Top             =   2760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbl_cartaoAtivo 
         BackColor       =   &H00E5E5E5&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFA324&
         Height          =   315
         Left            =   2100
         TabIndex        =   73
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbl_dinheiroAtivo 
         BackColor       =   &H00E5E5E5&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFA324&
         Height          =   315
         Left            =   2100
         TabIndex        =   72
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Vale Troca"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1260
         TabIndex        =   71
         Top             =   4620
         Width           =   1065
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Cheque"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1260
         TabIndex        =   70
         Top             =   3840
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Parcelamento"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1260
         TabIndex        =   69
         Top             =   3060
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Cartão ou VR"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1260
         TabIndex        =   68
         Top             =   2280
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Dinheiro (F2)"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1260
         TabIndex        =   67
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label Recebido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   705
         Left            =   2490
         TabIndex        =   66
         Top             =   6090
         Width           =   2295
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Recebido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   480
         Left            =   210
         TabIndex        =   65
         Top             =   6285
         Width           =   2235
      End
      Begin VB.Label Diferença 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   2490
         TabIndex        =   64
         Top             =   5265
         Width           =   2295
      End
      Begin VB.Label DifrmTro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   210
         TabIndex        =   63
         Top             =   5430
         Width           =   2145
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "R$ a receber"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   300
         TabIndex        =   62
         Top             =   405
         Width           =   1800
      End
      Begin VB.Label Total_Receber 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   705
         Left            =   2490
         TabIndex        =   61
         Top             =   210
         Width           =   2295
      End
      Begin VB.Label lbl_dinheiro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   705
         Left            =   2490
         TabIndex        =   57
         Top             =   1050
         Width           =   2295
      End
      Begin VB.Label lbl_cartaoCredito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   705
         Left            =   2490
         TabIndex        =   56
         Top             =   1830
         Width           =   2295
      End
      Begin VB.Label lbl_parcelamento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   705
         Left            =   2490
         TabIndex        =   55
         Top             =   2610
         Width           =   2295
      End
      Begin VB.Label lbl_cheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   705
         Left            =   2490
         TabIndex        =   54
         Top             =   3390
         Width           =   2295
      End
      Begin VB.Label lbl_outrosPagamentos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   705
         Left            =   2490
         TabIndex        =   53
         Top             =   4170
         Width           =   2295
      End
   End
   Begin VB.TextBox txtCredito 
      Height          =   285
      Left            =   11250
      TabIndex        =   49
      Text            =   "Credito"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtDebito 
      Height          =   285
      Left            =   10590
      TabIndex        =   48
      Text            =   "Debito"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin SSDataWidgets_B.SSDBDropDown ddwCartoes 
      Bindings        =   "Recebimentos.frx":51A76
      Height          =   735
      Left            =   10500
      TabIndex        =   47
      Top             =   1500
      Width           =   1455
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   77
      BackColor       =   -2147483633
   End
   Begin VB.Data Data3 
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
      Height          =   345
      Left            =   11820
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT Nome From Cartões"
      Top             =   1950
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SSDataWidgets_B.SSDBGrid Grade_Cartoes 
      Height          =   1725
      Left            =   11160
      TabIndex        =   46
      Top             =   3330
      Visible         =   0   'False
      Width           =   8535
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
      Col.Count       =   6
      BevelColorShadow=   15066597
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      ForeColorEven   =   4210752
      ForeColorOdd    =   4210752
      BackColorEven   =   12648447
      BackColorOdd    =   14737632
      RowHeight       =   450
      ExtraHeight     =   185
      Columns.Count   =   6
      Columns(0).Width=   4075
      Columns(0).Caption=   "Administradora"
      Columns(0).Name =   "Administradora"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1879
      Columns(1).Caption=   "Valor"
      Columns(1).Name =   "Valor"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1217
      Columns(2).Caption=   "Crédito"
      Columns(2).Name =   "Credito"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   11
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      Columns(3).Width=   1402
      Columns(3).Caption=   "Parcelas"
      Columns(3).Name =   "Qtde Parcelas"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2275
      Columns(4).Caption=   "Valor Parcelas"
      Columns(4).Name =   "Valor Parcelas"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   3201
      Columns(5).Caption=   "Numero Cartão"
      Columns(5).Name =   "Numero"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   20
      _ExtentX        =   15055
      _ExtentY        =   3043
      _StockProps     =   79
      Caption         =   "Cartões"
      ForeColor       =   4210752
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
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cartões (F5)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   10470
      TabIndex        =   35
      Top             =   4290
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox Num_Cartão 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton B_Parc_Cartão 
         Caption         =   "&Parcelar Cartão"
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
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Parcela no cartão"
         Top             =   1320
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Cartão 
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
         Bindings        =   "Recebimentos.frx":51A8A
         DataSource      =   "Data1"
         Height          =   285
         Left            =   990
         TabIndex        =   2
         Top             =   240
         Width           =   735
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
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.Label Label_Cartão2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3480
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Empresa :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Nome_Empresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cartão :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Valor :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label_Cartão4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label_Cartão3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "vezes de"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label_Cartão1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "parcelado em"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2880
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Caption         =   "Vales/ Outros (F6)"
      ForeColor       =   &H00404040&
      Height          =   2865
      Left            =   10440
      TabIndex        =   34
      Top             =   2550
      Visible         =   0   'False
      Width           =   2535
      Begin MSMask.MaskEdBox Vale 
         Height          =   705
         Left            =   1740
         TabIndex        =   6
         Top             =   2760
         Width           =   2125
         _ExtentX        =   3757
         _ExtentY        =   1244
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         ForeColor       =   4210752
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Caption         =   "Dinheiro (F2)"
      ForeColor       =   &H00404040&
      Height          =   3735
      Left            =   10320
      TabIndex        =   33
      Top             =   4050
      Width           =   3405
      Begin MSMask.MaskEdBox Dinheiro 
         Height          =   705
         Left            =   1740
         TabIndex        =   1
         Top             =   2760
         Width           =   2125
         _ExtentX        =   3757
         _ExtentY        =   1244
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         ForeColor       =   4210752
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E5E5E5&
         Caption         =   "<< TECLA TAB atualiza status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFA324&
         Height          =   315
         Left            =   270
         TabIndex        =   77
         Top             =   4920
         Width           =   2835
      End
   End
   Begin VB.CommandButton B_Retorna 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8160
      Width           =   13725
   End
   Begin VB.CheckBox Só_Leitura 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Só_Leitura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12030
      TabIndex        =   32
      Top             =   8550
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame frmParcela 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Caption         =   "Parcelamento (F7)"
      ForeColor       =   &H00404040&
      Height          =   5325
      Left            =   5220
      TabIndex        =   26
      Top             =   300
      Visible         =   0   'False
      Width           =   8415
      Begin VB.ComboBox cmb_mesInicioParcela 
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
         Height          =   330
         ItemData        =   "Recebimentos.frx":51A9E
         Left            =   5340
         List            =   "Recebimentos.frx":51AC6
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   300
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox txt_parcelamento_diaFixo 
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
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   79
         Top             =   300
         Visible         =   0   'False
         Width           =   555
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
         Left            =   5910
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Conta"
         Top             =   2700
         Visible         =   0   'False
         Width           =   1140
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Banco 
         Bindings        =   "Recebimentos.frx":51B2F
         DataSource      =   "Data2"
         Height          =   330
         Left            =   3105
         TabIndex        =   16
         Top             =   4050
         Width           =   1545
         DataFieldList   =   "Descrição"
         ListAutoValidate=   0   'False
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
         Columns(0).Width=   5821
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4180
         Columns(1).Caption=   "Conta"
         Columns(1).Name =   "Conta"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Conta"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1746
         Columns(2).Caption=   "Código"
         Columns(2).Name =   "Código"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Código"
         Columns(2).DataType=   2
         Columns(2).FieldLen=   256
         _ExtentX        =   2725
         _ExtentY        =   582
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin VB.OptionButton O_Carnet 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "Carnê"
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
         Height          =   285
         Left            =   1050
         TabIndex        =   14
         Top             =   4560
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton O_Carteira 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "C&arteira"
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
         Height          =   285
         Left            =   1050
         TabIndex        =   13
         Top             =   4080
         Width           =   1095
      End
      Begin VB.OptionButton O_Banco 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "&Banco"
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
         Height          =   285
         Left            =   2220
         TabIndex        =   15
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Qtde_Parcelas 
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
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   10
         Top             =   870
         Width           =   735
      End
      Begin VB.CommandButton B_Monta_Parcelas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Gerar Parcelas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2550
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   2100
      End
      Begin SSDataWidgets_B.SSDBGrid Grade_Parcela 
         Height          =   2535
         Left            =   1050
         TabIndex        =   12
         Top             =   1380
         Width           =   3585
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelColorShadow=   15066597
         ForeColorEven   =   4210752
         ForeColorOdd    =   4210752
         BackColorEven   =   12648447
         BackColorOdd    =   15066597
         RowHeight       =   503
         ExtraHeight     =   185
         Columns.Count   =   2
         Columns(0).Width=   2196
         Columns(0).Caption=   "Data"
         Columns(0).Name =   "Data"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   7
         Columns(0).FieldLen=   256
         Columns(1).Width=   3043
         Columns(1).Caption=   "Valor"
         Columns(1).Name =   "Valor"
         Columns(1).Alignment=   1
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).NumberFormat=   "###,###,##0.00"
         Columns(1).FieldLen=   256
         UseDefaults     =   0   'False
         _ExtentX        =   6324
         _ExtentY        =   4471
         _StockProps     =   79
         Caption         =   "Parcelas"
         ForeColor       =   4210752
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
      Begin VB.Label lbl_avisoAnoParcelamento 
         BackColor       =   &H00E5E5E5&
         Caption         =   "lbl_avisoAnoParcelamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4860
         TabIndex        =   81
         Top             =   1710
         Visible         =   0   'False
         Width           =   3075
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl_diaParcelaFixo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "Todo dia            iniciando a primeira parcela em "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1020
         TabIndex        =   78
         Top             =   345
         Visible         =   0   'False
         Width           =   4365
      End
      Begin VB.Label Nome_Banco 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   2220
         TabIndex        =   28
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         Caption         =   "Qtde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1050
         TabIndex        =   27
         Top             =   930
         Width           =   585
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   12720
      Top             =   7140
   End
   Begin VB.CheckBox Conta 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Caption         =   "Lançar o débito referente a esta compra para a conta do cliente"
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5220
      TabIndex        =   0
      Top             =   240
      Width           =   5250
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8400
      Visible         =   0   'False
      Width           =   1155
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
      Height          =   375
      Left            =   10800
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cartão"
      Top             =   7770
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton B_Cancela 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "C&ancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7560
      Width           =   13725
   End
   Begin VB.CommandButton B_Confirma 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Confirmar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Confirma Recebimento "
      Top             =   6990
      Width           =   13725
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H80000008&
      Height          =   6795
      Left            =   5100
      TabIndex        =   44
      Top             =   135
      Width           =   8655
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Para limpar a tela pressione o botão Cancelar e pressione o botão Recebimento novamente."
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   6060
         TabIndex        =   45
         Top             =   6630
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.Label Intervalo_Parc 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12420
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label L_Sequência 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sequência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12510
      TabIndex        =   25
      Top             =   7500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Max_Parcelas 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11940
      TabIndex        =   24
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Max_Cheques 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11700
      TabIndex        =   23
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Retorno 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12750
      TabIndex        =   22
      Top             =   8130
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Receber 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11490
      TabIndex        =   20
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Menu mnuConsulta 
      Caption         =   "Consulta"
      Enabled         =   0   'False
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuConsAcp 
         Caption         =   "Consulta ACP"
      End
      Begin VB.Menu mnuConfAcp 
         Caption         =   "Configurações"
      End
   End
End
Attribute VB_Name = "frmRecebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngCodigoCliente As Long
Public bytTelaChamada As Byte  ' 1 - Venda rápida, 2 - Saídas

Public strNumeroCartao As String

Private rsCartoes As Recordset
Private rsMov_Cheques As Recordset
Private rsMov_Parcelas As Recordset
Private rsParametros As Recordset
Private rsSaidas As Recordset
Private rsCliFor As Recordset
Private rsTabelas As Recordset
Private rsBancos As Recordset

'11/12/2009 - Andrea
Private rsMov_Cartoes As Recordset

Private Usa_Timer As Integer
Private Valor_A_Receber As Double
Private Valor_Recebido As Double
Private Recebido_Parc As Double
Private Recebido_Cheque As Double

Private Type Tab_1
  Banco As Integer
  Cheque As String
  Bom As String
  Valor As Double
End Type

'---------------------------------------------------------------
' DATA: 14/06/2022
' AUTOR: Pablo Verçosa Silva
' MUDANÇAS:
'    1) Incluir parâmetros de recebimento de parcelas e cheques
'    2) Ampliar o limite de parcelas e cheques para 3 dígitos
'---------------------------------------------------------------
'Private Tabe_Cheque(49) As Tab_1
Private Tabe_Cheque() As Tab_1
'---------------------------------------------------------------

Private Type Tab_2
  Dia As String
  Valor As Double
End Type

'---------------------------------------------------------------
' DATA: 14/06/2022
' AUTOR: Pablo Verçosa Silva
' MUDANÇAS:
'    1) Incluir parâmetros de recebimento de parcelas e cheques
'    2) Ampliar o limite de parcelas e cheques para 3 dígitos
'---------------------------------------------------------------
'Private Tabe_Parcela(49) As Tab_2
Private Tabe_Parcela() As Tab_2
'---------------------------------------------------------------

'06/01/2004 - Daniel
'Variáveis que armazenarão o Valor Recebido
'e o Troco para edição na tabela de Saídas
Public g_dblValorRecebidoFrmRec As Double
Public g_dblTrocoFrmRec As Double

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)


'---------------------------------------------------------------
' DATA: 14/06/2022
' AUTOR: Pablo Verçosa Silva
' MUDANÇAS:
'    1) Incluir parâmetros de recebimento de parcelas e cheques
'    2) Ampliar o limite de parcelas e cheques para 3 dígitos
'---------------------------------------------------------------
Private Sub Configura_Parcelas_e_Cheques()
  'trata parcelas
  pab_VR_Qtde_Parcela = 1
  If rsParametros("VR Permite Parcela") And rsParametros("VR Qtde Parcela") > 0 Then
    pab_VR_Qtde_Parcela = rsParametros("VR Qtde Parcela")
  End If
  ReDim Tabe_Parcela(pab_VR_Qtde_Parcela)
  Grade_Parcela.Rows = pab_VR_Qtde_Parcela
  Grade_Parcela.Refresh

  'trata cheques
  pab_VR_Qtde_Cheques = 1
  If rsParametros("VR Permite Cheques") And rsParametros("VR Qtde Cheques") > 0 Then
    pab_VR_Qtde_Cheques = rsParametros("VR Qtde Cheques")
  End If
  ReDim Tabe_Cheque(pab_VR_Qtde_Cheques)
  Grade_Cheque.Rows = pab_VR_Qtde_Cheques
  Grade_Cheque.Refresh
End Sub
'---------------------------------------------------------------


Public Function Pega_Total_Cheques() As Double
  Dim Ordem1 As Integer
  Dim Valor As Double

  Valor = 0
  ' alteração parametro cheque (Pablo)
  'For Ordem1 = 0 To 49
  For Ordem1 = 0 To pab_VR_Qtde_Cheques - 1
    If Tabe_Cheque(Ordem1).Valor <> 0 Then
      Valor = Valor + Tabe_Cheque(Ordem1).Valor
    End If
  Next Ordem1

  Pega_Total_Cheques = Valor
  Exit Function
End Function

Public Function Pega_Total_Cheques_Separado(ByVal blnPreDatado As Boolean) As Double
  Dim intX As Integer
  Dim dblTotal As Double

    ' alteração parametro cheque (Pablo)
    'For intX = 0 To 49
    For intX = 0 To pab_VR_Qtde_Cheques - 1
    If Tabe_Cheque(intX).Valor > 0 Then
      If CDate(Tabe_Cheque(intX).Bom) = CDate(Data_Atual) Then
        If Not blnPreDatado Then
          dblTotal = dblTotal + Tabe_Cheque(intX).Valor
        End If
      Else
        If blnPreDatado Then
          dblTotal = dblTotal + Tabe_Cheque(intX).Valor
        End If
      End If
    End If
  Next intX

  Pega_Total_Cheques_Separado = dblTotal

End Function

Public Sub Acerta_Tela()
  
   
  rsSaidas.Index = "Sequência"
  rsSaidas.Seek "=", gnCodFilial, Val(L_Sequência.Caption)
  
  If rsSaidas.NoMatch Then Exit Sub
  
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", rsSaidas("Cliente")
  If rsCliFor.NoMatch Then Exit Sub
  
  Max_Cheques.Caption = 0
  Max_Parcelas.Caption = 0
  If rsCliFor("Tem Conta") = False Then Conta.Enabled = False
    
  rsTabelas.Index = "Tabela"
  rsTabelas.Seek "=", rsSaidas("Tabela")
  If Not rsTabelas.NoMatch Then
    If rsTabelas("Aceita Cartão") = False Then
       Combo_Empresa.Enabled = False
       Num_Cartão.Enabled = False
       Cartão.Enabled = False
    End If
    If rsTabelas("Aceita Vale") = False Then
       Vale.Enabled = False
    End If
    If rsTabelas("Aceita Pré") = False Then
       Qtde_Cheques.Enabled = False
       Grade_Cheque.Enabled = False
    End If
    If rsTabelas("Aceita Pré") = True Then Max_Cheques.Caption = rsTabelas("Prazo Pré")
    If rsTabelas("Aceita Parcelamento") = False Then
      Qtde_Parcelas.Enabled = False
      Grade_Parcela.Enabled = False
    End If
    If rsTabelas("Aceita Parcelamento") = True Then Max_Parcelas.Caption = rsTabelas("Prazo Parcelamento")
  End If
  
  If rsCliFor("Faturado") = False Then
    Max_Cheques.Caption = 1
    Max_Parcelas.Caption = 1
  End If
  
 ' frmRecebimento.Dinheiro.SetFocus

End Sub

Function Exporta_Banco(Posição As Integer) As String
  Exporta_Banco = Tabe_Cheque(Posição).Banco
End Function

Function Exporta_Cheque(Posição As Integer) As String
  Exporta_Cheque = Tabe_Cheque(Posição).Cheque
End Function

Function Exporta_Data(Posição As Integer) As String
  Exporta_Data = Tabe_Cheque(Posição).Bom
End Function


Function Exporta_Valor(Posição As Integer) As Double
  Exporta_Valor = Tabe_Cheque(Posição).Valor
End Function


Sub Limpa_Tela(Tipo As Integer)
Dim i As Integer

 Combo_Empresa.Enabled = True
 Num_Cartão.Enabled = True
 Cartão.Enabled = True
 Vale.Enabled = True
 Qtde_Cheques.Enabled = True
 Grade_Cheque.Enabled = True
 Qtde_Parcelas.Enabled = True
 Grade_Parcela.Enabled = True
 
 '10/12/2009 - Andrea
 Grade_Cartoes.Enabled = True

If Tipo <> 1 Then
  Conta.Value = 0
  Conta.Enabled = True
End If

Grade_Cartoes.RemoveAll


'---------------------------------------------------------------
' DATA: 14/06/2022
' AUTOR: Pablo Verçosa Silva
' MUDANÇAS:
'    1) Incluir parâmetros de recebimento de parcelas e cheques
'    2) Ampliar o limite de parcelas e cheques para 3 dígitos
'---------------------------------------------------------------
'For i = 0 To 49
'   Tabe_Cheque(i).Banco = 0
'   Tabe_Cheque(i).Cheque = ""
'   Tabe_Cheque(i).Bom = ""
'   Tabe_Cheque(i).Valor = 0
   
'   Tabe_Parcela(i).Dia = ""
'   Tabe_Parcela(i).Valor = 0
' Next i
For i = 0 To pab_VR_Qtde_Cheques - 1
   Tabe_Cheque(i).Banco = 0
   Tabe_Cheque(i).Cheque = ""
   Tabe_Cheque(i).Bom = ""
   Tabe_Cheque(i).Valor = 0
 Next i
For i = 0 To pab_VR_Qtde_Parcela - 1
   Tabe_Parcela(i).Dia = ""
   Tabe_Parcela(i).Valor = 0
 Next i
 '---------------------------------------------------------------
 
 
 Dinheiro.Text = ""
 Vale.Text = ""
 Combo_Empresa.Text = ""
 Combo_Empresa_LostFocus
 Num_Cartão.Text = ""
 Cartão.Text = ""
 Label_Cartão2.Caption = ""
 Label_Cartão4.Caption = ""
 Label_Cartão2.Visible = False
 Label_Cartão4.Visible = False
 Recebido.Caption = ""
 DifrmTro.Caption = ""
 Diferença.Caption = ""
 
 Qtde_Cheques.Text = ""
 Qtde_Parcelas.Text = ""
 lbl_avisoAnoParcelamento.Caption = ""
 lbl_avisoAnoParcelamento.Visible = False
 
 Grade_Cheque.MoveLast
 Grade_Cheque.MoveFirst
 Grade_Parcela.MoveLast
 Grade_Parcela.MoveFirst
 
End Sub

Sub Mostra(Num As Long)
 Dim Erro As Integer
 Dim Ordem As Integer
 Dim sRecord As String
 Dim strSQL As String
 
 '11/12/2009 - Andrea
 '--------------------------------------------------------------------------------------------
 Grade_Cartoes.RemoveAll
 Ordem = 0
 Erro = False
 
 strSQL = "SELECT * "
 strSQL = strSQL & "FROM [Movimento - Cartoes] WHERE [Movimento - Cartoes].Filial = " & gnCodFilial & "  AND "
 strSQL = strSQL & "[Movimento - Cartoes].Sequência = " & Num & " ORDER BY [Movimento - Cartoes].Ordem "
 Set rsMov_Cartoes = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
 With rsMov_Cartoes
   If Not (.BOF And .EOF) Then
     Do Until .EOF
       sRecord = rsMov_Cartoes.Fields("Administradora").Value & vbTab & _
            rsMov_Cartoes.Fields("Valor").Value & vbTab & _
            rsMov_Cartoes.Fields("Credito").Value & vbTab & _
            rsMov_Cartoes.Fields("Parcelas").Value & vbTab & _
            rsMov_Cartoes.Fields("ValorParcelas").Value & vbTab & _
            rsMov_Cartoes.Fields("NumeroCartao").Value
      
       Grade_Cartoes.AddItem sRecord
     
     .MoveNext
      
    Loop
   End If
   .Close
 End With
 Set rsMov_Cartoes = Nothing
 '--------------------------------------------------------------------------------------------

  Combo_Empresa.Text = rsSaidas("Recebe - Emp Cartão") & ""
  Combo_Empresa_LostFocus

 '16/12/2009 - Andrea
 'Verifica se teve recebimentos em cartões anteriores a alteração de receber em vários cartões.
 'Nestes casos os valores recebidos em cartões estão na tabela de saidas.
 If rsSaidas("Recebe - Cartão").Value > 0 Then  'Teve recebimento em cartao da forma antiga
   If rsSaidas("Recebe - Emp Cartão").Value > 0 Then
  
     sRecord = Nome_Empresa.Caption & vbTab & _
        rsSaidas("Recebe - Cartão").Value & vbTab & _
        gsHandleNull(rsSaidas("Qtde Parcelas").Value & "") & vbTab & _
        gsHandleNull(rsSaidas("Valor Parcela").Value & "") & vbTab & _
        rsSaidas("Recebe - Num Cartão") & ""
  
     Grade_Cartoes.AddItem sRecord
   End If
 End If

 
 rsMov_Cheques.Index = "Ordem"
 Ordem = 0
 Erro = False
 Do
   rsMov_Cheques.Seek ">", gnCodFilial, Num, Ordem
   If rsMov_Cheques.NoMatch Then Erro = True
   If Erro = False Then If rsMov_Cheques("Filial") <> gnCodFilial Then Erro = True
   If Erro = False Then If rsMov_Cheques("Sequência") <> Num Then Erro = True
   If Erro = False Then
      Ordem = rsMov_Cheques("Ordem")
      ' alteração parametro cheque (Pablo)
      'If Ordem > 50 Then Exit Do
      If Ordem > pab_VR_Qtde_Cheques Then Exit Do
      Tabe_Cheque(Ordem - 1).Banco = rsMov_Cheques("Banco")
      Tabe_Cheque(Ordem - 1).Cheque = rsMov_Cheques("Cheque")
      Tabe_Cheque(Ordem - 1).Bom = rsMov_Cheques("Bom")
      Tabe_Cheque(Ordem - 1).Valor = rsMov_Cheques("Valor")
   End If
 Loop Until Erro = True
 
 
 
 rsMov_Parcelas.Index = "Ordem"
 Ordem = 0
 Erro = False
 Do
   rsMov_Parcelas.Seek ">", gnCodFilial, Num, Ordem
   If rsMov_Parcelas.NoMatch Then Erro = True
   If Erro = False Then If rsMov_Parcelas("Filial") <> gnCodFilial Then Erro = True
   If Erro = False Then If rsMov_Parcelas("Sequência") <> Num Then Erro = True
   If Erro = False Then
      Ordem = rsMov_Parcelas("Ordem")
      ' alteração parametro parcela (Pablo)
      'If Ordem > 50 Then Exit Do
      If Ordem > pab_VR_Qtde_Parcela Then Exit Do
      Tabe_Parcela(Ordem - 1).Dia = rsMov_Parcelas("Bom")
      Tabe_Parcela(Ordem - 1).Valor = rsMov_Parcelas("Valor")
   End If
 Loop Until Erro = True
 
 Grade_Cheque.Visible = True
 Grade_Cheque.MoveLast
 Grade_Cheque.MoveFirst
 
 
 Grade_Parcela.Visible = True
 Grade_Parcela.MoveLast
 Grade_Parcela.MoveFirst
 
 Usa_Timer = True
 
  
End Sub

Sub Mostra_Dados_Leitura()
  Dim i As Integer
 

  rsSaidas.Index = "Sequência"
  
  rsSaidas.Seek "=", gnCodFilial, L_Sequência.Caption
  
  If rsSaidas.NoMatch Then Exit Sub
  
  Conta.Value = -rsSaidas("Recebe - Conta")
  Dinheiro.Text = rsSaidas("Recebe - Dinheiro")
  
  Combo_Empresa.Text = rsSaidas("Recebe - Emp Cartão") & ""
  Combo_Empresa_LostFocus
  Num_Cartão.Text = rsSaidas("Recebe - Num Cartão") & ""
  Cartão.Text = rsSaidas("Recebe - Cartão")
  If rsSaidas("Recebe - Cartão").Value > 0 Then
    Label_Cartão1.Visible = True
    Label_Cartão3.Visible = True
    Label_Cartão2.Caption = gsHandleNull(rsSaidas("Qtde Parcelas").Value & "")
    Label_Cartão4.Caption = gsHandleNull(rsSaidas("Valor Parcela").Value & "")
    Label_Cartão2.Visible = True
    Label_Cartão4.Visible = True
  Else
    Label_Cartão1.Visible = False
    Label_Cartão3.Visible = False
    Label_Cartão2.Caption = ""
    Label_Cartão4.Caption = ""
    Label_Cartão2.Visible = False
    Label_Cartão4.Visible = False
  End If
  
  Vale.Text = rsSaidas("Recebe - Vale")
  
 Mostra L_Sequência.Caption

End Sub

Public Function Pega_Parcela(Ordem, R_Bom, R_Valor, Parcelas As Integer) As Integer
  Dim Ordem1 As Integer
  Dim i As Integer, Conta As Integer
  
  Conta = 0
  ' alteração parametro parcela (Pablo)
  'For i = 0 To 49
  For i = 0 To pab_VR_Qtde_Parcela - 1
    If Tabe_Parcela(i).Valor <> 0 Then
     If IsDate(Tabe_Parcela(i).Dia) Then Conta = Conta + 1
    End If
  Next i
    
  
  Parcelas = Conta
  
  
  Ordem1 = Ordem - 1
  If Tabe_Parcela(Ordem1).Valor = 0 Then
    Pega_Parcela = 0
    Exit Function
  End If
  
  R_Bom = Tabe_Parcela(Ordem1).Dia
  R_Valor = Tabe_Parcela(Ordem1).Valor
  
  Pega_Parcela = 1

End Function

Public Function Pega_Banco(Ordem, R_Banco, R_Cheque, R_Bom, R_Valor) As Integer
On Error GoTo Erro
  Dim Ordem1 As Integer
  
  Ordem1 = Ordem - 1
  If Tabe_Cheque(Ordem1).Valor = 0 Then
    Pega_Banco = 0
    Exit Function
  End If
  
  R_Banco = Tabe_Cheque(Ordem1).Banco
  R_Cheque = Tabe_Cheque(Ordem1).Cheque
  R_Bom = Tabe_Cheque(Ordem1).Bom
  R_Valor = Tabe_Cheque(Ordem1).Valor
  
  Pega_Banco = 1

  Exit Function

Erro:
  MsgBox "Erro na função Pega_Banco tela Recebimento " + Err.Number + " " + Err.Description, vbInformation, "Atenção"

End Function

Public Function Pega_Total_Parcelas() As Double
On Error GoTo Erro
  Dim Ordem1 As Integer
  Dim Valor As Double
  
  Valor = 0
  ' alteração parametro parcela (Pablo)
  'For Ordem1 = 0 To 49
  For Ordem1 = 0 To pab_VR_Qtde_Parcela - 1
    If Tabe_Parcela(Ordem1).Valor <> 0 Then
      Valor = Valor + Tabe_Parcela(Ordem1).Valor
    End If
  Next Ordem1
  
  Pega_Total_Parcelas = Valor

  Exit Function
Erro:
  MsgBox "Erro na função Pega_Total_Parcelas " + Err.Number + " " + Err.Description, vbInformation, "Atenção"

End Function

Sub Recalcula()
On Error GoTo Erro

  Valor_Recebido = 0
  Recebido_Parc = 0
  Recebido_Cheque = 0
  Dim i As Integer
  
  If Conta.Value = 1 Then
     Valor_Recebido = CDbl(Receber.Caption)
  End If
  
  If Not IsNull(Dinheiro.Text) Then
   If IsNumeric(Dinheiro.Text) Then
    Valor_Recebido = Valor_Recebido + CDbl(Dinheiro.Text)
   End If
  End If
  
  '-----------------------------------------------------------------------------------------------------------
  '10/12/2009 - Andrea
  'Recalcula o valor recebido em cartão
  With Grade_Cartoes
    'Verifica ocorrência
    If .Rows > 0 Then
      
      Dim lng_row As Long
      Dim var_book As Variant
      Dim dbl_valor_recebido_cartao As Double
      Dim dbl_valor As Double
      Dim str_administradora As String
      Dim int_qtde_parcelas As Double
      
      dbl_valor_recebido_cartao = 0
      
      For lng_row = 0 To .Rows - 1
          
        var_book = .AddItemBookmark(lng_row)
              
        'Verifica registro informado
        Call IsDataType(dtString, .Columns("Administradora").CellText(var_book), str_administradora)
        If str_administradora <> "" Then
          'Valores
          Call IsDataType(dtDouble, .Columns("Valor").CellText(var_book), dbl_valor)
          Call IsDataType(dtInteger, .Columns("Qtde Parcelas").CellText(var_book), int_qtde_parcelas)
          
          dbl_valor_recebido_cartao = dbl_valor_recebido_cartao + dbl_valor
        End If
      Next lng_row
      Cartão.Text = dbl_valor_recebido_cartao
      Valor_Recebido = Valor_Recebido + CDbl(Cartão.Text)
      lbl_cartaoCredito.Caption = Format(Cartão.Text, "###,###,##0.00")
    End If
  End With
'  If Not IsNull(Cartão.Text) Then
'   If IsNumeric(Cartão.Text) Then
'     Valor_Recebido = Valor_Recebido + CDbl(Cartão.Text)
'   End If
'  End If
  '-----------------------------------------------------------------------------------------------------------
  
  If Not IsNull(Vale.Text) Then
   If IsNumeric(Vale.Text) Then
     Valor_Recebido = Valor_Recebido + CDbl(Vale.Text)
   End If
  End If
  
'---------------------------------------------------------------
' DATA: 14/06/2022
' AUTOR: Pablo Verçosa Silva
' MUDANÇAS:
'    1) Incluir parâmetros de recebimento de parcelas e cheques
'    2) Ampliar o limite de parcelas e cheques para 3 dígitos
'---------------------------------------------------------------
  'For i = 0 To 49
  '  Recebido_Cheque = Recebido_Cheque + Tabe_Cheque(i).Valor
  '  Recebido_Parc = Recebido_Parc + Tabe_Parcela(i).Valor
  'Next i
  For i = 0 To pab_VR_Qtde_Parcela - 1
    Recebido_Parc = Recebido_Parc + Tabe_Parcela(i).Valor
  Next i
  For i = 0 To pab_VR_Qtde_Cheques - 1
    Recebido_Cheque = Recebido_Cheque + Tabe_Cheque(i).Valor
  Next i
'---------------------------------------------------------------
  
  Valor_Recebido = Valor_Recebido + Recebido_Parc
  Valor_Recebido = Valor_Recebido + Recebido_Cheque
  
  lbl_parcelamento.Caption = Format(Recebido_Parc, "###,###,##0.00")
  lbl_cheque.Caption = Format(Recebido_Cheque, "###,###,##0.00")
  
  Recebido.Caption = Format(Valor_Recebido, "###,###,##0.00")
  
  Valor_A_Receber = CDbl(Receber.Caption) - Recebido
  
  If Valor_A_Receber >= 0 Then
    Diferença.Caption = Format(Valor_A_Receber, "###,###,##0.00")
  End If
  
  If Valor_A_Receber < 0 Then
    Diferença.Caption = Format(-Valor_A_Receber, "###,###,##0.00")
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro em função Recalcula " + Err.Number + " " + Err.Description, vbInformation, "Atenção"
End Sub

Private Sub B_Cancela_Click()
  Retorno = "CANCELADO"
  Grade_Parcela.CancelUpdate
  Grade_Cheque.CancelUpdate
  Call StatusMsg("")
  frmRecebimento.Hide
End Sub

Private Sub B_Confirma_Click()
  Dim Linha As Integer
  Dim Erro As Integer
  Dim blnVerificaLimite As Boolean
  Dim nRow As Long
  Dim bm As Variant

  On Error GoTo TratarErro
  
  B_Confirma.Enabled = False
  
  DoEvents
  Sleep 500
  DoEvents
  
 
  '30/07/2003 - Maikel
  '             Adicionada a função abaixo que verifica o limite de crédito do cliente
  '-----------------------------------------------------------------------------------'
    Call StatusMsg("Analisando o crédito do cliente, aguarde . . . ")
    blnVerificaLimite = False

    If (bytTelaChamada = 1) Then      ' - Venda rápida
      If (rsParametros.Fields("VR Verifica Limite").Value) Then
        blnVerificaLimite = True
      End If
    ElseIf (bytTelaChamada = 2) Then  ' - Saídas
      If (rsParametros.Fields("Saída Verifica Limite").Value) Then
        blnVerificaLimite = True
      End If
    End If
    
    If blnVerificaLimite Then
      If Not AnalisaClienteComCredito Then
          B_Confirma.Enabled = True
          Exit Sub
      End If
    End If
    Call StatusMsg("")
  '-----------------------------------------------------------------------------------'
    
  DoEvents
  Recalcula
  DoEvents

  Erro = False
' alteração parametro cheque (Pablo)
'  For Linha = 0 To 49
  For Linha = 0 To pab_VR_Qtde_Cheques - 1
    If Tabe_Cheque(Linha).Valor <> 0 Then
      If Not IsDate(Tabe_Cheque(Linha).Bom) Then Erro = True
    End If
  Next Linha

  If Erro = True Then
    DisplayMsg "Existem cheques sem data digitada, verifique."
    B_Confirma.Enabled = True
    Exit Sub
  End If

  Erro = False
' alteração parametro parcela (Pablo)
'  For Linha = 0 To 49
  For Linha = 0 To pab_VR_Qtde_Parcela - 1
    If Tabe_Parcela(Linha).Valor <> 0 Then
      If Not IsDate(Tabe_Parcela(Linha).Dia) Then Erro = True
    End If
  Next Linha
  
  If Erro = True Then
    DisplayMsg "Existem parcelas sem data digitada, verifique."
    B_Confirma.Enabled = True
    Exit Sub
  End If

  If IsNull(Dinheiro.Text) Then Dinheiro.Text = 0
  If Dinheiro.Text = "" Then Dinheiro.Text = 0
  
  If IsNull(Cartão.Text) Then Cartão.Text = 0
  If Cartão.Text = "" Then Cartão.Text = 0

  DoEvents
  Sleep 500
  DoEvents
  
  '--------------------------------------------------------------------------------------------------------
  '08/12/2009 - Andrea
  'Alterada a verificação dos cartões, agora tem que verificar no grid se tem linhas
  'sem o número do cartão ou administradora
  '  If Nome_Empresa.Caption = "" Then
  '   If CDbl(Cartão.Text) <> 0 Then
  '    DisplayMsg "Encontre o cartão."
  '    Combo_Empresa.SetFocus
  '    Exit Sub
  '   End If
  '  End If
  Dim str_administradora As String
  Dim str_numero_cartao As String
  Dim bln_credito As Boolean
  
  TxtDebito.Text = 0
  txtCredito.Text = 0
  For nRow = 0 To Grade_Cartoes.Rows - 1
    
    bm = Grade_Cartoes.AddItemBookmark(nRow)
    str_administradora = Grade_Cartoes.Columns("Administradora").CellText(bm)
    str_numero_cartao = Grade_Cartoes.Columns("Numero").CellText(bm)
    strNumeroCartao = Grade_Cartoes.Columns("Numero").Text
    
    If Len(str_administradora) = 0 Then Erro = True
    'If Len(str_numero_cartao) = 0 Then Erro = True
    
    bln_credito = Grade_Cartoes.Columns("Credito").CellValue(bm)
    
    If bln_credito = True Then
      txtCredito.Text = CDbl(txtCredito.Text) + CDbl(Grade_Cartoes.Columns("Valor").CellText(bm))
    Else
      TxtDebito.Text = CDbl(TxtDebito.Text) + CDbl(Grade_Cartoes.Columns("Valor").CellText(bm))
    End If
    
    
     
  Next nRow

  If Erro = True Then
    DisplayMsg "Existem valores recebidos em cartões sem Administradora, verifique."
    B_Confirma.Enabled = True
    Exit Sub
  End If
  
  Erro = False
  '--------------------------------------------------------------------------------------------------------
  
  If IsNull(Vale.Text) Then Vale.Text = 0
  If Vale.Text = "" Then Vale.Text = 0
  
  If Valor_A_Receber > 0 Then
    Beep
    DisplayMsg "Valor recebido insuficiente, confira."
    B_Confirma.Enabled = True
    Exit Sub
  End If
  
  If Recebido_Parc <> 0 Then
    If O_Banco.Value = True Then
     If Nome_Banco.Caption = "" Then
        DisplayMsg "Banco incorreto."
        Combo_Banco.SetFocus
        B_Confirma.Enabled = True
        Exit Sub
     End If
    End If
  End If
  
  '06/01/2004 - Daniel
  'Variável que editará o campo Valor Recebido
  'na tabela de Saídas
  Me.g_dblValorRecebidoFrmRec = CDbl(Recebido.Caption)
  'Inicia a var dblTroco com zero
  Me.g_dblTrocoFrmRec = 0
  
  
  If Valor_A_Receber < 0 Then
  
    Recebido.Caption = Format((CDbl(Recebido) + Valor_A_Receber), "#############0.00")
  
    '06/01/2004 - Daniel
    'Alimentar a var g_dblTrocoFrmRec pois
    'ocorreu troco
    Me.g_dblTrocoFrmRec = CDbl(Trim(Diferença.Caption))
    '-----------------------------------------------
    gsTitle = LoadResString(201)
    gsMsg = "CLIENTE TEM TROCO DE R$ " & Trim(Diferença.Caption)
    MsgBox gsMsg, vbInformation, "CLIENTE TEM TROCO"
'''    gsMsg = "CONFIRMA DEVOLUÇÃO DO TROCO DE " & Trim(Diferença.Caption) & "?"
'''    gnStyle = vbYesNo + vbQuestion + vbDefaultButton1
'''    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    gnResponse = vbYes
    If gnResponse = vbNo Then
      Dinheiro.SetFocus
      B_Confirma.Enabled = True
      Exit Sub
    End If
    
  End If

  Dinheiro.Text = Format((CDbl(Dinheiro.Text) + Valor_A_Receber), "#############0.00")
   
  If Len(Trim(Label_Cartão2.Caption)) = 0 Then
    If CDbl(Cartão.Text) > 0 Then
      Label_Cartão2.Caption = "1"
      Label_Cartão4.Caption = CDbl(Cartão.Text)
    Else
      Label_Cartão2.Caption = ""
      Label_Cartão4.Caption = ""
    End If
  End If

  ''15/05/2013-Alexandre Afornali
  ''Case DiskEmbalagens
  'Dim rsComandas As Recordset
  'Set rsComandas = db.OpenRecordset("SaidasComandas")
  'Dim countrs As Long
  'countrs = 0
  'While Not rsComandas.EOF
  '  countrs = countrs + 1
  '  rsComandas.MoveNext
  'Wend
  'If (countrs > 0) Then
  '  rsComandas.MoveFirst
  'End If
  'While Not rsComandas.EOF
  '  If (rsComandas("CodSaida") = L_Sequência) And (rsComandas("Filial").Value = gnCodFilial) Then
  '    rsComandas.Delete
  '    rsComandas.MoveLast
  '    countrs = countrs - 1
  '  End If

  '  If countrs > 0 Then
  '      rsComandas.MoveNext
  '  End If
  'Wend

  Retorno = "OK"
  Call StatusMsg("")
  
  Erase pfParcelasFatura
  
  ReDim pfParcelasFatura(Grade_Parcela.Rows) As ParcelasFatura
  Dim intContador As Integer

  Grade_Parcela.Redraw = False
  Grade_Parcela.MoveFirst
  For intContador = 0 To UBound(pfParcelasFatura)
    pfParcelasFatura(intContador).pfDataVencimento = Grade_Parcela.Columns(0).Text
    pfParcelasFatura(intContador).pfValor = Grade_Parcela.Columns(1).Text
    
    Grade_Parcela.MoveNext
  Next intContador
  Grade_Parcela.Redraw = True
  
  frmRecebimento.Hide

  B_Confirma.Enabled = True
  Exit Sub
  
TratarErro:
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub


Private Sub B_Imprime_Click()
  Dim Str_Impre As String
  Dim Resposta As Integer

  frmImprimeCheque.Show vbModal
  Exit Sub

'  Call SetPrinterName(Nome_Impressora_Cheque.Caption, Porta_Impressora_Cheque.Caption)

 ' Str_Impre = ""
  'Str_Impre = Str_Impre + Chr$(27) + Chr$(177) + Chr$(13)
'  Printer.Print ""
'  Resposta = Escape(Printer.Hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
 
  Call SetPrinterName("CHEQUE")
  Str_Impre = Str_Impre + Chr$(27)
  Str_Impre = Str_Impre + Chr$(160) + "Secretaria do Estado da Fazenda" + Chr$(13)
  Str_Impre = Str_Impre + Chr$(27) + Chr$(161) + "Curitiba" + Chr$(13)
  Str_Impre = Str_Impre + Chr$(27) + Chr$(162) + "399" + Chr$(13)
  Str_Impre = Str_Impre + Chr$(27) + Chr$(163) + "57.91" + Chr$(13)
  Str_Impre = Str_Impre + Chr$(27) + Chr$(164) + "14/05/97" + Chr$(13)
  Str_Impre = Str_Impre + Chr$(27) + Chr$(177)
  Str_Impre = Str_Impre + Chr$(27) + Chr$(176)
  Str_Impre = Chr$(Len(Str_Impre) Mod 256) + Chr$(Len(Str_Impre) \ 256) + Str_Impre
  Printer.Print ""
  If Not IsWindowsNT() Then
    Resposta = Escape(Printer.hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
  Else
    Resposta = Escape32(Printer.hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
  End If
  Printer.EndDoc
  Call SetPrinterName("REL")

End Sub

Private Sub B_Monta_Cheques_Click()
  Dim Parc As Double
  Dim i As Integer
  
  If IsNull(Qtde_Cheques.Text) Then Exit Sub
  If Qtde_Cheques.Text = "" Then Exit Sub
  If Val(Qtde_Cheques) <= 0 Then Exit Sub
' alteração parametro cheque (Pablo)
  'If Val(Qtde_Cheques) > 50 Then Qtde_Cheques.Text = 50
  If Val(Qtde_Cheques) > pab_VR_Qtde_Cheques Then Qtde_Cheques.Text = pab_VR_Qtde_Cheques
  
  Grade_Cheque.Rows = Val(Qtde_Cheques.Text)
  Grade_Cheque.Refresh
 
  Parc = CDbl(Format(Valor_A_Receber / Val(Qtde_Cheques.Text), "########.00"))
  
  For i = 0 To (Val(Qtde_Cheques.Text) - 1)
    Tabe_Cheque(i).Valor = Parc
  Next i
 
  '-- desabilitado em 20/06/2022 (Pablo)
  'Grade_Cheque.MoveLast
  'Grade_Cheque.MoveFirst
  'Grade_Cheque.Col = 0
  Grade_Cheque.Refresh
  
  Recalcula
  
End Sub

Private Sub B_Monta_Parcelas_Click()
  Dim Parc As Double
  Dim i As Integer
  Dim Dia As Date
  Dim Diferenca As Double
  Dim ParcAux As Double
  
  
  If IsNull(Qtde_Parcelas.Text) Then Exit Sub
  If Qtde_Parcelas.Text = "" Then Exit Sub
  If Val(Qtde_Parcelas) <= 0 Then Exit Sub
' alteração parametro parcela (Pablo)
'  If Val(Qtde_Parcelas) > 50 Then Qtde_Parcelas.Text = 50
  If Val(Qtde_Parcelas) > pab_VR_Qtde_Parcela Then Qtde_Parcelas.Text = pab_VR_Qtde_Parcela
    
  Grade_Parcela.Rows = Val(Qtde_Parcelas.Text)
  Grade_Parcela.Refresh

  '--------------------------------------
  ParcAux = 0
' alteração parametro parcela (Pablo)
  'For i = 0 To 49
  For i = 0 To pab_VR_Qtde_Parcela - 1
      ParcAux = ParcAux + CDbl(Tabe_Parcela(i).Valor)
      Tabe_Parcela(i).Valor = 0
  Next
  Valor_A_Receber = Valor_A_Receber + ParcAux
  '--------------------------------------

 
  Parc = CDbl(Format(Valor_A_Receber / Val(Qtde_Parcelas.Text), "########.00"))
  
  If txt_parcelamento_diaFixo.Visible = True Then
      'Cálculo novo (para tratar dia fixo do mês)
      Dim iMes As Integer
      Dim iAno As Integer
      
      iMes = cmb_mesInicioParcela.ListIndex + 1
      
      If Month(Date) > iMes Then
          iAno = Year(Date) + 1
          lbl_avisoAnoParcelamento.Caption = "A data da 1ª parcela esta iniciando no ano de " & iAno & ". Se desejar, altere na grade ao lado."
          lbl_avisoAnoParcelamento.Visible = True
      Else
          iAno = Year(Date)
          lbl_avisoAnoParcelamento.Visible = False
      End If
      
      
      
      If Not IsDate(txt_parcelamento_diaFixo.Text & "/" & iMes & "/" & iAno) Then
          MsgBox "Data " & txt_parcelamento_diaFixo.Text & "/" & iMes & "/" & iAno & " inválida!", vbInformation, "Atenção"
          Exit Sub
      End If
      
      Dia = txt_parcelamento_diaFixo.Text & "/" & iMes & "/" & iAno
      
      Diferenca = CDbl(Format(Valor_A_Receber - Parc * Val(Qtde_Parcelas.Text), "########.00"))
      
      For i = 0 To (Val(Qtde_Parcelas.Text) - 1)
      
        If i = 0 And Diferenca > 0 Then
          Tabe_Parcela(i).Valor = Parc + Diferenca
        ElseIf i = Val(Qtde_Parcelas.Text) - 1 And Diferenca < 0 Then
          Tabe_Parcela(i).Valor = Parc + Diferenca
        Else
          Tabe_Parcela(i).Valor = Parc
        End If
        
        Tabe_Parcela(i).Dia = Dia
        
        If iMes = 12 Then
            iMes = 1
            iAno = iAno + 1
            Dia = txt_parcelamento_diaFixo.Text & "/" & iMes & "/" & iAno
        Else
            If iMes = 1 Then
                'Aqui será Fevereiro, assumirá sempre dia 28 (se o dia vencimento for entre 28 a 31)
                iMes = iMes + 1
                If CInt(txt_parcelamento_diaFixo.Text) > 27 Then
                    Dia = "28/" & iMes & "/" & iAno
                Else
                    Dia = txt_parcelamento_diaFixo.Text & "/" & iMes & "/" & iAno
                End If
            Else
                iMes = iMes + 1
                
                If (iMes = 4 Or iMes = 6 Or iMes = 9 Or iMes = 11) And txt_parcelamento_diaFixo.Text = "31" Then
                    Dia = "30/" & iMes & "/" & iAno
                ElseIf iMes = 2 And (txt_parcelamento_diaFixo.Text = "31" Or txt_parcelamento_diaFixo.Text = "30" Or txt_parcelamento_diaFixo.Text = "29") Then
                    Dia = "28/" & iMes & "/" & iAno
                Else
                    Dia = txt_parcelamento_diaFixo.Text & "/" & iMes & "/" & iAno
                End If
            End If
        End If
      Next i
  
  Else
      ' Cálculo de parcelas normal (que já tinhamos)
      Dia = Data_Atual + Val(Intervalo_Parc.Caption)
      
      Diferenca = CDbl(Format(Valor_A_Receber - Parc * Val(Qtde_Parcelas.Text), "########.00"))
      
      For i = 0 To (Val(Qtde_Parcelas.Text) - 1)
      
        If i = 0 And Diferenca > 0 Then
          Tabe_Parcela(i).Valor = Parc + Diferenca
        ElseIf i = Val(Qtde_Parcelas.Text) - 1 And Diferenca < 0 Then
          Tabe_Parcela(i).Valor = Parc + Diferenca
        Else
          Tabe_Parcela(i).Valor = Parc
        End If
        
        Tabe_Parcela(i).Dia = Dia
        'Tabe_Parcela(i).Valor = Parc
        Dia = Dia + Val(Intervalo_Parc.Caption)
      Next i
  End If
 
  '-- desabilitado em 20/06/2022 (Pablo)
  'Grade_Parcela.MoveLast
  'Grade_Parcela.MoveFirst
  'Grade_Parcela.Col = 0
  Grade_Parcela.Refresh
  
  lbl_parcelamento.Caption = Valor_A_Receber
    
  Recalcula
   
End Sub

Private Sub B_Retorna_Click()
  Me.Hide
End Sub

Private Sub Cartão_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Cartão_LostFocus()
  Cartão.Text = gsFormatCurrency(Cartão.Text, gnCurrencyDecimals)
  Recalcula
End Sub

Private Sub Combo_Banco_CloseUp()
  Combo_Banco.Text = Combo_Banco.Columns(2).Text
  Combo_Banco_LostFocus
End Sub

'05/02/2004 - mpdea
'Corrigido RT-3421 na busca pelo código do banco para conta corrente
Private Sub Combo_Banco_LostFocus()
  Dim intBanco As Integer
  
  Nome_Banco.Caption = ""
  
  Call IsDataType(dtInteger, Combo_Banco.Text, intBanco)
  If intBanco <= 0 Then Exit Sub
  
'  If IsNull(Combo_Banco.Text) Then Exit Sub
'  If Combo_Banco.Text = "" Then Exit Sub
'  If Not IsNumeric(Combo_Banco.Text) Then Exit Sub
'  If Val(Combo_Banco.Text) < 1 Then Exit Sub
  
  rsBancos.Index = "Código"
  rsBancos.Seek "=", intBanco
  If rsBancos.NoMatch Then Exit Sub
  
  Nome_Banco.Caption = rsBancos.Fields("Descrição").Value & ""

End Sub

Private Sub Combo_Empresa_CloseUp()
  Combo_Empresa.Text = Combo_Empresa.Columns(1).Text
  Combo_Empresa_LostFocus
End Sub

Private Sub Combo_Empresa_InitColumnProps()
  '04/09/2002 - mpdea
  'Incluído o redimensionamento das colunas
  With Combo_Empresa
    .Columns(0).Width = 5000
    .Columns(1).Width = 1000
  End With
End Sub

'13/12/2005 - mpdea
'Incluído tratamento de entrada de dados
Private Sub Combo_Empresa_LostFocus()
  
  Dim int_ret As Integer

  Nome_Empresa.Caption = ""

  Call IsDataType(dtInteger, Combo_Empresa.Text, int_ret)
  Combo_Empresa.Text = int_ret

'  If IsNull(Combo_Empresa.Text) Then Exit Sub
'  If Combo_Empresa.Text = "" Then Exit Sub
'  If Not IsNumeric(Combo_Empresa.Text) Then Exit Sub
'  If Val(Combo_Empresa.Text) < 0 Then Exit Sub

  rsCartoes.Index = "Código"
  rsCartoes.Seek "=", int_ret
  If rsCartoes.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsCartoes("Nome")
  
End Sub

Private Sub B_Parc_Cartão_Click()
  Dim Resp As Variant


 Call StatusMsg("")
  
  If IsNull(Cartão.Text) Then Cartão.Text = 0
  If Cartão.Text = "" Then Cartão.Text = 0
  If Not IsNumeric(Cartão.Text) Then Cartão.Text = 0
  If CDbl(Cartão.Text) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Digite o valor antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    
    
    '13/02/2003 - mpdea
    'Corrige o erro quando o controle está desabilitado
    SelectAllText Cartão, True
    'Cartão.SetFocus
    
    
    Exit Sub
  End If
  
  Resp = InputBox("Digite o número de parcelas", "Parcelamento no cartão")
  
  If IsNull(Resp) Then Exit Sub
  If Not IsNumeric(Resp) Then Exit Sub
  
  If Val(Resp) <= 0 Then
    Label_Cartão2.Caption = ""
    Label_Cartão4.Caption = ""
    Label_Cartão1.Visible = False
    Label_Cartão2.Visible = False
    Label_Cartão3.Visible = False
    Label_Cartão4.Visible = False
    Exit Sub
  End If
  
  If Val(Resp) < 1 Or Val(Resp) > 20 Then
    gsTitle = LoadResString(201)
    gsMsg = "Quantidade de parcelas inválida."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Label_Cartão2.Caption = Resp
  Label_Cartão4.Caption = Round(CDbl(Cartão.Text) / Val(Resp), 2)
  
  'Cartão.Text = Round(CInt(Resp) * CDbl(Label_Cartão4.Caption), 2)
  
  Label_Cartão1.Visible = True
  Label_Cartão2.Visible = True
  Label_Cartão3.Visible = True
  Label_Cartão4.Visible = True
  
End Sub

Private Sub Conta_Click()
'26/07/2007 - Anderson
'Impedir que seja realizada outras operações quando é selecionada a conta cliente.
 Dim bolHabilitarPagamentos As Boolean
 
 bolHabilitarPagamentos = (Conta.Value = 0)
 
 Frame1.Enabled = bolHabilitarPagamentos
 Frame2.Enabled = bolHabilitarPagamentos
 Frame4.Enabled = bolHabilitarPagamentos
 Frame3.Enabled = bolHabilitarPagamentos
 frmParcela.Enabled = bolHabilitarPagamentos
 
 If Conta.Value = 1 Then
   Limpa_Tela (1)
 End If
  
 Recalcula
End Sub

Private Sub ddwAdministradora_DropDown()
  Dim rsTemp As Recordset
  Set rsTemp = db.OpenRecordset("SELECT Nome FROM Cartões ", dbOpenSnapshot)
  If rsTemp.EOF Then
    ddwCartoes.DataFieldToDisplay = "Código"
  Else
    ddwCartoes.DataFieldToDisplay = "Nome"
  End If
  rsTemp.Close
  Set rsTemp = Nothing

End Sub

Private Sub Diferença_Change()

  If Valor_A_Receber = 0 Then
    DifrmTro.Caption = "OK"
    'Diferença.ForeColor = vbBlack '&HFF00&
    Diferença.BackColor = &HC0FFC0
    Exit Sub
  End If
  
  If Valor_A_Receber < 0 Then
    DifrmTro.Caption = "Troco"
    'Diferença.ForeColor = vbBlue '&HFFFF&
    Diferença.BackColor = &HFFFFC0
    If Recebido <> "" And Recebido <> "0" And Recebido <> "0,00" Then
        Recebido.Caption = Format(Recebido + Valor_A_Receber, "###,###,##0.00")
    End If
    Exit Sub
  End If
  
  If Valor_A_Receber > 0 Then
    DifrmTro.Caption = "A Receber"
    'Diferença.ForeColor = vbRed '&HFF&
    Diferença.BackColor = &HC0C0FF
    Exit Sub
  End If
  
End Sub

Private Sub Dinheiro_GotFocus()
    With Dinheiro
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Dinheiro_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub


Private Sub Dinheiro_LostFocus()
  Dinheiro.Text = gsFormatCurrency(Dinheiro.Text, gnCurrencyDecimals)
  lbl_dinheiro.Caption = Dinheiro.Text
  Recalcula
End Sub


Private Sub Form_Activate()
  Call CenterForm(Me)
  
  If IsNull(Receber.Caption) Then Receber.Caption = "0"
  If Receber.Caption = "" Then Receber.Caption = "0"
  Total_Receber.Caption = Format(Receber.Caption, "###,###,##0.00")
  
  Valor_A_Receber = CDbl(Receber.Caption)
  
  
  lbl_dinheiro.Caption = "0,00"
  lbl_cartaoCredito.Caption = "0,00"
  lbl_parcelamento.Caption = "0,00"
  lbl_cheque.Caption = "0,00"
  lbl_outrosPagamentos.Caption = "0,00"
  
  Grade_Cheque.MoveFirst
  Grade_Parcela.MoveFirst

  B_Imprime.Visible = True
  B_Cancela.Visible = True
  B_Confirma.Visible = True

  B_Retorna.Visible = False

  Combo_Banco_LostFocus

  If Só_Leitura.Value = 1 Then
    B_Retorna.Visible = True
    B_Imprime.Visible = False
    B_Cancela.Visible = False
    B_Confirma.Visible = False

    Mostra_Dados_Leitura
  End If
  
  If rsParametros.Fields("VR Permite Dinheiro").Value <> -1 Then
      opt_dinheiro.Visible = False
      Label4.Visible = False
      lbl_dinheiro.Visible = False
  End If
  
  If rsParametros.Fields("VR Permite Cartão").Value <> -1 Then
      opt_cartaoCredito.Visible = False
      Label5.Visible = False
      lbl_cartaoCredito.Visible = False
  End If

  If rsParametros.Fields("VR Permite Parcela").Value <> -1 Then
      opt_parcelamento.Visible = False
      Label6.Visible = False
      lbl_parcelamento.Visible = False
  End If

  If rsParametros.Fields("VR Permite Cheques").Value <> -1 Then
      opt_cheque.Visible = False
      Label7.Visible = False
      lbl_cheque.Visible = False
  End If

  If rsParametros.Fields("VR Permite Vales").Value <> -1 Then
      opt_outrosPagamentos.Visible = False
      Label8.Visible = False
      lbl_outrosPagamentos.Visible = False
  End If

  If rsParametros.Fields("VR Permite Dinheiro").Value = -1 Then
      opt_dinheiro.Value = True
      opt_dinheiro_Click
      opt_dinheiro.SetFocus
      '''Dinheiro.Text = Total_Receber.Caption
      Dinheiro.SetFocus
      'Vale.SetFocus
      Dinheiro.SetFocus
    
      If Só_Leitura.Value = 1 Then
        opt_dinheiro.SetFocus
      End If
  End If
  
  If bytTelaChamada = 1 Then
      ' 1 - Venda rápida
      If rsParametros("VR Parcela Padrão") = "B" Then O_Banco.Value = True
      If rsParametros("VR Parcela Padrão") = "C" Then O_Carteira.Value = True
      If rsParametros("VR Parcela Padrão") = "E" Then O_Carnet.Value = True
  Else
      ' 2 - Saídas
      If rsParametros("Saída Parcela Padrão") = "B" Then O_Banco.Value = True
      If rsParametros("Saída Parcela Padrão") = "C" Then O_Carteira.Value = True
      If rsParametros("Saída Parcela Padrão") = "E" Then O_Carnet.Value = True
      'If rsParametros("Saída Altera Parcela") = False Then Tipo_Parc.Enabled = False
  End If
  
  
  If Intervalo_Parc.Caption = "0" Or Intervalo_Parc.Caption = "" Or Intervalo_Parc.Caption = "30" Then
      lbl_diaParcelaFixo.Visible = True
      txt_parcelamento_diaFixo.Visible = True
      cmb_mesInicioParcela.Visible = True
      
      If Month(Date) <> 12 Then
          If IsDate(Day(Date) & "/" & Month(Date + 1) & "/" & Year(Date)) Then  'Caso seja dia 31
              txt_parcelamento_diaFixo.Text = Day(Date)
              cmb_mesInicioParcela.ListIndex = Month(Date)
          Else
              If IsDate(Day(Date - 1) & "/" & Month(Date + 1) & "/" & Year(Date)) Then
                  txt_parcelamento_diaFixo.Text = Day(Date - 1)
                  cmb_mesInicioParcela.ListIndex = Month(Date)
              Else
                  If IsDate(Day(Date - 2) & "/" & Month(Date + 1) & "/" & Year(Date)) Then
                      txt_parcelamento_diaFixo.Text = Day(Date - 2)
                      cmb_mesInicioParcela.ListIndex = Month(Date)
                  Else
                      If IsDate(Day(Date - 3) & "/" & Month(Date + 1) & "/" & Year(Date)) Then
                          txt_parcelamento_diaFixo.Text = Day(Date - 3)
                          cmb_mesInicioParcela.ListIndex = Month(Date)
                      Else
                          MsgBox "Verifique se a data sugerida é válida", vbInformation, "Atenção para a data parcelamento inicial"
                      End If
                  End If
              End If
          End If
      Else
          txt_parcelamento_diaFixo.Text = Day(Date)
          cmb_mesInicioParcela.ListIndex = 0
      End If
  End If

  '30/07/2003 - mpdea
  'Ativa timer para atualizar exibição dos grids
  Usa_Timer = True
  Timer1.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 And Frame2.Visible = True Then
    If Dinheiro.Enabled = True Then
       Dinheiro.Text = Diferença.Caption
       Dinheiro.SetFocus
       Dinheiro.SelStart = 0
       Dinheiro.SelLength = Len(Dinheiro.Text)
    End If
  End If
  
  If KeyCode = vbKeyF3 And Frame1.Visible = True Then
    If Qtde_Cheques.Enabled = True Then Qtde_Cheques.SetFocus
  End If
  If KeyCode = vbKeyF7 And frmParcela.Visible = True Then
    If Qtde_Parcelas.Enabled = True Then Qtde_Parcelas.SetFocus
  End If
'''  If KeyCode = vbKeyF5 Then
'''    If Combo_Empresa.Enabled = True Then Combo_Empresa.SetFocus
'''  End If
  
  If KeyCode = vbKeyF6 And Frame3.Visible = True Then
    If Vale.Enabled = True Then Vale.SetFocus
  End If
  
'''''''''  If KeyCode = vbKeyReturn Then
'''''''''     Call B_Confirma_Click
'''''''''  End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Set rsCartoes = db.OpenRecordset("Cartões", , dbReadOnly)
  Set rsMov_Cheques = db.OpenRecordset("Movimento - Cheques", , dbReadOnly)
  Set rsMov_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsSaidas = db.OpenRecordset("Saídas", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsTabelas = db.OpenRecordset("Tabela de Preços", , dbReadOnly)
  Set rsBancos = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
 
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  
  
  '---------------------------------------------------------------
  ' DATA: 14/06/2022
  ' AUTOR: Pablo Verçosa Silva
  ' MUDANÇAS:
  '    1) Incluir parâmetros de recebimento de parcelas e cheques
  '    2) Ampliar o limite de parcelas e cheques para 3 dígitos
  '---------------------------------------------------------------
  Call Configura_Parcelas_e_Cheques
  '---------------------------------------------------------------
  
  
  Limpa_Tela (0)
  
  Usa_Timer = True
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then Exit Sub

  If bytTelaChamada = 1 Then
      ' 1 - Venda rápida
      If rsParametros("VR Parcela Padrão") = "B" Then O_Banco.Value = True
      If rsParametros("VR Parcela Padrão") = "C" Then O_Carteira.Value = True
      If rsParametros("VR Parcela Padrão") = "E" Then O_Carnet.Value = True
  Else
      ' 2 - Saídas
      If rsParametros("Saída Parcela Padrão") = "B" Then O_Banco.Value = True
      If rsParametros("Saída Parcela Padrão") = "C" Then O_Carteira.Value = True
      If rsParametros("Saída Parcela Padrão") = "E" Then O_Carnet.Value = True
      'If rsParametros("Saída Altera Parcela") = False Then Tipo_Parc.Enabled = False
  End If

  
  '23/01/2003 - mpdea
  'Quick em modo limitado
  If Not gblnQuickFull Then
    'Parcelamento
    O_Carteira.Visible = False
    O_Carnet.Visible = True
    O_Banco.Visible = False
    Combo_Banco.Visible = False
    Nome_Banco.Visible = False
    'O_Carteira.Value = True
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  rsCartoes.Close
  rsMov_Cheques.Close
  rsMov_Parcelas.Close
  rsParametros.Close
  rsSaidas.Close
  rsCliFor.Close
  rsTabelas.Close
  rsBancos.Close
    
  Set rsCartoes = Nothing
  Set rsMov_Cheques = Nothing
  Set rsMov_Parcelas = Nothing
  Set rsParametros = Nothing
  Set rsSaidas = Nothing
  Set rsCliFor = Nothing
  Set rsTabelas = Nothing
  Set rsBancos = Nothing

End Sub

Private Sub Grade_Cartoes_AfterColUpdate(ByVal ColIndex As Integer)
  Recalcula
End Sub

Private Sub Grade_Cartoes_AfterUpdate(RtnDispErrMsg As Integer)
  Recalcula
End Sub

'08/12/2009 - Andrea
Private Sub Grade_Cartoes_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
 
 Dim Aux As Variant
 Dim Erro As Integer
 Dim Tempo As Long
 Dim dbl_valor_parcela As Double
 
 Aux = Grade_Cartoes.Columns(ColIndex).Text 'Conteúdo da célula
 
 Call StatusMsg("")
 
 If ColIndex = 0 Then 'Administradora
   Erro = False
   If IsNull(Aux) Then Erro = True
   If Erro = True Then
     DisplayMsg "Escolha uma Administradora de cartões."
     Cancel = True
     Exit Sub
    End If
 End If

 If ColIndex = 1 Then
    If Not IsNumeric(Aux) Then
      DisplayMsg "Valor inválido."
      Cancel = True
      Exit Sub
    End If
    
    If CDbl(Aux) < 0 Then
      DisplayMsg "Valor inválido."
      Cancel = True
      Exit Sub
    End If
    
    If IsNumeric(Grade_Cartoes.Columns(2).Text) Then
      If Grade_Cartoes.Columns(2).Text > 1 Then 'Número de parcelas
        dbl_valor_parcela = (CDbl(Aux) / CDbl(Grade_Cartoes.Columns(2).Text))
        Grade_Cartoes.Columns(3).Text = Format(dbl_valor_parcela, "###,###,##0.00")
      End If
    End If
  
    If Not IsNumeric(Cartão.Text) Then
      Cartão.Text = 0
    End If
    
    Cartão.Text = Cartão.Text + CDbl(Aux)
    
 End If
   
 If ColIndex = 3 Then 'Qtde Parcelas
    If IsNumeric(Aux) Then
      If CDbl(Aux) > 1 Then
        dbl_valor_parcela = (CDbl(Grade_Cartoes.Columns(1).Text) / CDbl(Aux))
        Grade_Cartoes.Columns(4).Text = Format(dbl_valor_parcela, "###,###,##0.00")
      Else
        Grade_Cartoes.Columns(4).Text = 0
      End If
    Else
        Grade_Cartoes.Columns(4).Text = 0
    End If
 End If

End Sub

Private Sub Grade_Cartoes_InitColumnProps()
  '07/12/2009 - Andrea
  Grade_Cartoes.Columns("Administradora").DropDownHwnd = ddwCartoes.hwnd
  ddwCartoes.DataFieldList = "Nome"

End Sub

Private Sub Grade_Cartoes_KeyDown(KeyCode As Integer, Shift As Integer)
    B_Confirma.Enabled = True
End Sub

Private Sub Grade_Cartoes_LostFocus()
  Grade_Cartoes.MoveNext
  Grade_Cartoes.MovePrevious
End Sub

'07/12/2009 - Andrea
'Private Sub Grade_Cartoes_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
''Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, PosicaoInicial As Variant, ByVal LeLinha As Boolean)
'
'  Dim rsCartoes As Recordset
'  Dim strSQL As String
'  Dim dr As Integer
'  Dim linha_num As Integer
'  Dim r As Integer
'  Dim linhas_retornadas As Integer
'
'  ' verifica qual a direção da leitura
'  If ReadPriorRows Then
'     dr = -1
'  Else
'     dr = 1
'  End If
'
'  ' verifica se a PosicaoInicial é nulo
'  If IsNull(StartLocation) Then
'    ' Le do fim ou do inicio dos dados
'     If ReadPriorRows Then
'      ' le de a partir do final
'        linha_num = RowBuf.RowCount - 1
'     Else
'       ' le a partir do inicio
'       linha_num = 0
'     End If
'  Else
'    ' verifica onde comecamos a leitura
'    linha_num = CLng(StartLocation) + dr
'  End If
'
'  ' copia os dados da tabela de administradoras de cartão para dentro do buffer = RowBuf.
'  linhas_retornadas = 0
'
'
'  'Ler um RecordSet das Administradoras de cartões
'  'Cartões
'  strSQL = "SELECT Nome "
'  strSQL = strSQL & "FROM [Cartões] "
'  Set rsCartoes = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
'  With rsCartoes
'    r = 0
'    If Not (.BOF And .EOF) Then
'      Do Until .EOF
'
'
'        ' copia os dados para o buffer da linha
'        RowBuf.Value(r, 0) = .Fields("Nome").Value
'
'        ' usa linha_num como um bookmark.
'        RowBuf.Bookmark(r) = linha_num
'
'        linha_num = linha_num + dr
'        linhas_retornadas = linhas_retornadas + 1
'        r = r + 1
'        dr = dr + 1
'
'      .MoveNext
'
'     Loop
'    End If
'    .Close
'  End With
'
'
'  ' define o numero de linhas retornardo
'  RowBuf.RowCount = linhas_retornadas
'
'End Sub

Private Sub Grade_Cheque_AfterUpdate(RtnDispErrMsg As Integer)
  Recalcula
End Sub

Private Sub Grade_Cheque_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
 Dim Aux As Variant
 Dim Erro As Integer
 Dim Tempo As Long
 
 Aux = Grade_Cheque.Columns(ColIndex).Text
 
 Call StatusMsg("")
 
 If ColIndex = 0 Then
   Erro = False
   If IsNull(Aux) Then Erro = True
   If Not Erro Then If Not IsNumeric(Aux) Then Erro = True
   If Not Erro Then If Val(Aux) < 0 Or Val(Aux) > 999 Then Erro = True
   If Erro = True Then
     DisplayMsg "Digite um número entre 0 e 999."
     Cancel = True
     Exit Sub
    End If
 End If

 If ColIndex = 1 Then 'Num cheque
   Aux = Left(Aux, 10)
   Grade_Cheque.Columns(1).Text = Aux
 End If
   
 If ColIndex = 2 Then 'Data
   Erro = False
   If Val(Grade_Cheque.Columns(0).Text) <> 0 Then
     If IsNull(Aux) Then Erro = True
     If Erro = False Then
        If Not IsDate(Aux) Then
          If IsNumeric(Aux) Then
             If Aux = 0 Then
                Aux = Date
                Grade_Cheque.Columns(2).Text = Aux
                Exit Sub
             End If
             If Val(Aux) < 0 Or Val(Aux) > 400 Then
                DisplayMsg "Data inválida, verifique."
                Cancel = True
                Exit Sub
             Else
                Aux = Replace(Aux, ".", "/")
                Aux = CDate(CDate(Data_Atual) + Aux)
                Grade_Cheque.Columns(2).Text = Aux
                Exit Sub
             End If
           DisplayMsg "Data inválida, verifique."
           Cancel = True
           Exit Sub
         End If
        End If
        If Not IsDate(Aux) Then
           DisplayMsg "Data inválida, verifique."
           Cancel = True
           Exit Sub
        End If
        If CDate(Aux) < Data_Atual Then
          DisplayMsg "Data não pode ser anterior a data atual."
          Cancel = True
        End If
        Aux = CDate(Aux)
        If Max_Cheques.Caption <> "0" Then
          Tempo = Aux - Date
          If Tempo > Val(Max_Cheques.Caption) Then
            DisplayMsg "Prazo muito longo."
            Cancel = True
          End If
        End If
        Grade_Cheque.Columns(2).Text = Aux
     End If
   End If
 End If
     
     
 If ColIndex = 3 Then
    If Not IsNumeric(Aux) Then
      DisplayMsg "Valor inválido."
      Cancel = True
      Exit Sub
    End If
    
    If CDbl(Aux) < 0 Then
      DisplayMsg "Valor inválido."
      Cancel = True
      Exit Sub
    End If
 End If
   
End Sub

Private Sub Grade_Cheque_DblClick()
  Grade_Cheque.MoveFirst
End Sub

Private Sub Grade_Cheque_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyF2
'      Dim frmCal As New frmCalendario
'      If Grade_Cheque.Columns(Grade_Cheque.Col).Name = "Data" Or _
'         Grade_Cheque.Columns(Grade_Cheque.Col).Name = "DataAviso" Then
'        If IsDate(Grade_Cheque.Columns(Grade_Cheque.Col).Value) Then
'          gsDate = Grade_Cheque.Columns(Grade_Cheque.Col).Value
'        End If
'        Set goField = Grade_Cheque.Columns(Grade_Cheque.Col)
'        goField.Text = Grade_Cheque.Columns(Grade_Cheque.Col).Text
'      End If
'      frmCal.Show
'  End Select
End Sub

Private Sub Grade_Cheque_LostFocus()
  Grade_Cheque.MoveNext
  Grade_Cheque.MovePrevious
End Sub

Private Sub Grade_Cheque_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
 Dim Lin, Col As Integer
 
  Lin = Grade_Cheque.Bookmark
  Col = Grade_Cheque.Col
  
  If Col = 0 Then
  If Lin > 0 Then
    If Tabe_Cheque(Lin).Valor <> 0 Then
     If IsNumeric(Tabe_Cheque(Lin - 1).Cheque) Then
       Grade_Cheque.Columns(0).Text = Tabe_Cheque(Lin - 1).Banco
       Grade_Cheque.Columns(1).Text = (Tabe_Cheque(Lin - 1).Cheque + 1)
       Grade_Cheque.Columns(2).Text = CDate(Tabe_Cheque(Lin - 1).Bom) + 30
     End If
    End If
   End If
  End If
End Sub

Private Sub Grade_Cheque_UnboundAddData(ByVal RowBuf As ssRowBuffer, NewRowBookmark As Variant)

 Dim Linha As Integer

 Linha = Grade_Cheque.Row

 Tabe_Cheque(Linha).Banco = Grade_Cheque.Columns(0).Text
 Tabe_Cheque(Linha).Cheque = Grade_Cheque.Columns(1).Text
 Tabe_Cheque(Linha).Bom = Grade_Cheque.Columns(2).Text
 Tabe_Cheque(Linha).Valor = Grade_Cheque.Columns(3).Text
 
End Sub

Private Sub Grade_Cheque_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade_Cheque.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p

End Sub


Private Sub Grade_Cheque_UnboundReadData(ByVal RowBuf As ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
On Error GoTo Erro
Dim r, i, J, p As Integer

If IsNull(StartLocation) Then
  If ReadPriorRows Then
    p = Grade_Cheque.Rows
  Else
    p = 0
  End If
 Else
  p = StartLocation
  If ReadPriorRows Then
    p = p - 1
  Else
    p = p + 1
  End If
End If

For i = 0 To RowBuf.RowCount - 1
  If p < 0 Or p >= Grade_Cheque.Rows Then Exit For
     RowBuf.Value(i, 0) = Tabe_Cheque(p).Banco
     RowBuf.Value(i, 1) = Tabe_Cheque(p).Cheque
     RowBuf.Value(i, 2) = Tabe_Cheque(p).Bom
     RowBuf.Value(i, 3) = Tabe_Cheque(p).Valor
     
   RowBuf.Bookmark(i) = p
   If ReadPriorRows Then
     p = p - 1
   Else
     p = p + 1
   End If
   
   r = r + 1
 Next i
 
 RowBuf.RowCount = r
   
  Exit Sub
Erro:
  MsgBox "Erro na sub UnboundReadData tela Recebimento " + Err.Number + " " + Err.Description, vbInformation, "Atenção"

End Sub


Private Sub Grade_Cheque_UnboundWriteData(ByVal RowBuf As ssRowBuffer, WriteLocation As Variant)
 Dim Linha As Integer
 Dim V_Cheque As Variant
 Dim Aux As Variant
 
 Linha = WriteLocation

 Call StatusMsg("")

 V_Cheque = Grade_Cheque.Columns(3).Text
 If IsNull(V_Cheque) Then
    Grade_Cheque.Columns(3).Text = 0
    V_Cheque = 0
 End If
 If Not IsNumeric(V_Cheque) Then
    Grade_Cheque.Columns(3).Text = 0
    V_Cheque = 0
 End If
 
 If CDbl(V_Cheque) <> 0 Then
   If IsNull(Grade_Cheque.Columns(0).Text) Then
     DisplayMsg "Banco incorreto."
     Exit Sub
   End If
   If Not IsNumeric(Grade_Cheque.Columns(0).Text) Then
     DisplayMsg "Banco incorreto."
     Exit Sub
   End If
   If Val(Grade_Cheque.Columns(0).Text) < 0 Or Val(Grade_Cheque.Columns(0).Text) > 999 Then
     DisplayMsg "Banco incorreto."
     Exit Sub
   End If
   
   If Not IsDate(Grade_Cheque.Columns(2).Text) Then
     DisplayMsg "Data incorreta."
     Exit Sub
   End If
 End If
 

Tabe_Cheque(Linha).Banco = Grade_Cheque.Columns(0).Text
Tabe_Cheque(Linha).Cheque = Grade_Cheque.Columns(1).Text
Tabe_Cheque(Linha).Bom = Grade_Cheque.Columns(2).Text
Tabe_Cheque(Linha).Valor = Grade_Cheque.Columns(3).Text




End Sub


Private Sub Grade_Cheque_UpdateError(ByVal ColIndex As Integer, Text As String, ErrCode As Integer, ErrString As String, Cancel As Integer)
 Beep
 DisplayMsg "Dados incorretos."
 Cancel = True

End Sub

Private Sub Grade_Parcela_AfterUpdate(RtnDispErrMsg As Integer)
 Recalcula
End Sub

Private Sub Grade_Cartoes_AfterDelete(RtnDispErrMsg As Integer)
  RtnDispErrMsg = False
  Call Recalcula
End Sub

Private Sub Grade_Cartoes_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  If Len(Trim(Grade_Cartoes.ActiveCell.Text)) = 0 Then
    If bGridBeforeDelete() = True Then
      Cancel = False
    Else
      Cancel = True
    End If
  Else
    Cancel = True
  End If
End Sub


Private Sub Grade_Parcela_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
 Dim Aux As Variant
 Dim Erro As Integer
 Dim Tempo As Long
 
 Aux = Grade_Parcela.Columns(ColIndex).Text
 
 Call StatusMsg("")
 
 If ColIndex = 0 Then 'Data
   Erro = False
   If Val(Grade_Parcela.Columns(0).Text) = 0 Then
     Grade_Parcela.Columns(0).Text = Date
'     Grade_Parcela.Columns(1).Text = 0
   End If
   If Val(Grade_Parcela.Columns(0).Text) <> 0 Then
     If IsNull(Aux) Then Erro = True
     If Erro = False Then
        If Not IsDate(Aux) Then
          If Not IsNumeric(Aux) Then
            DisplayMsg "Data inválida, verifique."
            Cancel = True
            Exit Sub
          End If
          If IsNumeric(Aux) Then
           If Val(Aux) < 0 Or Val(Aux) > 400 Then
             DisplayMsg "Data inválida, verifique."
             Cancel = True
             Exit Sub
           Else
             Aux = Replace(Aux, ".", "/")
             Aux = CDate(CDate(Data_Atual) + Aux)
             Grade_Parcela.Columns(0).Text = Aux
           End If
         End If
        End If
        If CDate(Aux) < Data_Atual Then
          DisplayMsg "Data não pode ser anterior a data atual."
          Cancel = True
        End If
        Aux = CDate(Aux)
        If Max_Parcelas.Caption <> "0" Then
          Tempo = Aux - Date
          If Tempo > Val(Max_Parcelas.Caption) Then
            DisplayMsg "Prazo muito longo para a Tabela de Preços ou Cliente não compra a prazo."
            Cancel = True
          End If
        End If
        Grade_Parcela.Columns(0).Text = Aux
     End If
   End If
 End If
     
     
     
 If ColIndex = 1 Then 'Valor
  Erro = False
  If IsNull(Aux) Then Erro = True
  If Erro = False Then If Not IsNumeric(Aux) Then Erro = True
  If Erro = False Then If CDbl(Aux) < 0 Then Erro = True
  If Erro = True Then
     DisplayMsg "Valor inválido, verifique."
     Cancel = True
     Exit Sub
  End If
 End If
 
 End Sub


Private Sub Grade_Parcela_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyF2
'      Dim frmCal As New frmCalendario
'      If Grade_Parcela.Columns(Grade_Parcela.Col).Name = "Data" Or _
'         Grade_Parcela.Columns(Grade_Parcela.Col).Name = "DataAviso" Then
'        If IsDate(Grade_Parcela.Columns(Grade_Parcela.Col).Value) Then
'          gsDate = Grade_Parcela.Columns(Grade_Parcela.Col).Value
'        End If
'        Set goField = Grade_Parcela.Columns(Grade_Parcela.Col)
'        goField.Text = Grade_Parcela.Columns(Grade_Parcela.Col).Text
'      End If
'      frmCal.Show
'  End Select
End Sub

Private Sub Grade_Parcela_LostFocus()

 Grade_Parcela.MoveNext
 Grade_Parcela.MovePrevious
 
 
  Dim intContador As Integer
  Dim vParcelasSomar As Double

  For intContador = 0 To Grade_Parcela.Rows - 1
    vParcelasSomar = vParcelasSomar + Grade_Parcela.Columns(1).Text
  Next intContador
  
  lbl_parcelamento.Caption = vParcelasSomar
 
End Sub

Private Sub Grade_Parcela_UnboundAddData(ByVal RowBuf As ssRowBuffer, NewRowBookmark As Variant)

 Dim Linha As Integer

 Linha = Grade_Parcela.Row

 Tabe_Parcela(Linha).Dia = Grade_Parcela.Columns(0).Text
 Tabe_Parcela(Linha).Valor = Grade_Parcela.Columns(1).Text
 


End Sub

Private Sub Grade_Parcela_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade_Parcela.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p


End Sub

Private Sub Grade_Parcela_UnboundReadData(ByVal RowBuf As ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim r, i, J, p As Integer

If IsNull(StartLocation) Then
  If ReadPriorRows Then
    p = Grade_Parcela.Rows
  Else
    p = 0
  End If
 Else
  p = StartLocation
  If ReadPriorRows Then
    p = p - 1
  Else
    p = p + 1
  End If
End If

For i = 0 To RowBuf.RowCount - 1
  If p < 0 Or p >= Grade_Parcela.Rows Then Exit For
     RowBuf.Value(i, 0) = Tabe_Parcela(p).Dia
     RowBuf.Value(i, 1) = Tabe_Parcela(p).Valor
     
   RowBuf.Bookmark(i) = p
   If ReadPriorRows Then
     p = p - 1
   Else
     p = p + 1
   End If
   
   r = r + 1
 Next i
 
 RowBuf.RowCount = r
   

End Sub

Private Sub Grade_Parcela_UnboundWriteData(ByVal RowBuf As ssRowBuffer, WriteLocation As Variant)
 Dim Linha As Integer
 
 Linha = WriteLocation

Tabe_Parcela(Linha).Dia = Grade_Parcela.Columns(0).Text
Tabe_Parcela(Linha).Valor = Grade_Parcela.Columns(1).Text

End Sub

Private Sub Grade_Parcela_UpdateError(ByVal ColIndex As Integer, Text As String, ErrCode As Integer, ErrString As String, Cancel As Integer)
 Beep
 DisplayMsg "Dados incorretos."
 Cancel = True
End Sub

Private Sub opt_cartaoCredito_Click()
  Grade_Cartoes.Top = 540
  Grade_Cartoes.Left = 5160
  Grade_Cartoes.Height = 6370
  Grade_Cartoes.Width = 8535  '6225
  Grade_Cartoes.Visible = True
  
  Frame1.Visible = False
  
  Frame2.Visible = False
  frmParcela.Visible = False
  Frame3.Visible = False
  
  lbl_cartaoAtivo.Visible = True
  lbl_chequeAtivo.Visible = False
  lbl_dinheiroAtivo.Visible = False
  lbl_outrosAtivo.Visible = False
  lbl_parcelamentoAtivo.Visible = False
  
  Grade_Cartoes.SetFocus
  
  B_Confirma.Enabled = True
End Sub

Private Sub opt_cheque_Click()
  Frame1.Top = 540
  Frame1.Left = 5160
  Frame1.Height = 6370
  Frame1.Width = 8535 '6225
  Frame1.Visible = True
  
  Grade_Cartoes.Visible = False
  Frame4.Visible = False
  
  Frame2.Visible = False
  frmParcela.Visible = False
  Frame3.Visible = False
  
  lbl_cartaoAtivo.Visible = False
  lbl_chequeAtivo.Visible = True
  lbl_dinheiroAtivo.Visible = False
  lbl_outrosAtivo.Visible = False
  lbl_parcelamentoAtivo.Visible = False
  
  Qtde_Cheques.SetFocus
  
  B_Confirma.Enabled = True
End Sub

Private Sub opt_dinheiro_Click()
  Frame2.Top = 540
  Frame2.Left = 5160
  Frame2.Height = 6370
  Frame2.Width = 8535 '6225
  Frame2.Visible = True
  
  Grade_Cartoes.Visible = False
  Frame4.Visible = False
  
  Frame1.Visible = False
  frmParcela.Visible = False
  Frame3.Visible = False
  
  lbl_cartaoAtivo.Visible = False
  lbl_chequeAtivo.Visible = False
  lbl_dinheiroAtivo.Visible = True
  lbl_outrosAtivo.Visible = False
  lbl_parcelamentoAtivo.Visible = False
  
  Dinheiro.SetFocus
  
  B_Confirma.Enabled = True
  
End Sub

Private Sub opt_outrosPagamentos_Click()
  Frame3.Top = 540
  Frame3.Left = 5160
  Frame3.Height = 6370
  Frame3.Width = 8535 '6225
  Frame3.Visible = True
  
  Grade_Cartoes.Visible = False
  Frame4.Visible = False
  
  Frame1.Visible = False
  Frame2.Visible = False
  frmParcela.Visible = False
  
  lbl_cartaoAtivo.Visible = False
  lbl_chequeAtivo.Visible = False
  lbl_dinheiroAtivo.Visible = False
  lbl_outrosAtivo.Visible = True
  lbl_parcelamentoAtivo.Visible = False
  
  B_Confirma.Enabled = True
End Sub

Private Sub opt_parcelamento_Click()
  frmParcela.Top = 540
  frmParcela.Left = 5160
  frmParcela.Height = 6370
  frmParcela.Width = 8535 '6225
  frmParcela.Visible = True
  
  Grade_Cartoes.Visible = False
  Frame4.Visible = False
  
  Frame1.Visible = False
  Frame2.Visible = False
  Frame3.Visible = False
  
  lbl_cartaoAtivo.Visible = False
  lbl_chequeAtivo.Visible = False
  lbl_dinheiroAtivo.Visible = False
  lbl_outrosAtivo.Visible = False
  lbl_parcelamentoAtivo.Visible = True
  
  Grade_Parcela.Enabled = True
  Qtde_Parcelas.Enabled = True
  Qtde_Parcelas.SetFocus
  
  B_Confirma.Enabled = True
End Sub

Private Sub Qtde_Cheques_KeyPress(KeyAscii As Integer)
  If KeyAscii = 8 Then Exit Sub   'backspace
  If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
    Exit Sub
  End If
End Sub

Private Sub Qtde_Parcelas_KeyPress(KeyAscii As Integer)
  If KeyAscii = 8 Then Exit Sub   'backspace
  If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
    Exit Sub
  End If
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next

  If Usa_Timer Then
    If Screen.ActiveForm.Caption = "Recebimento" Then
      Grade_Cheque.MoveLast
      Grade_Cheque.MoveFirst
      Grade_Parcela.MoveLast
      Grade_Parcela.MoveFirst
      Grade_Cheque.Refresh
      Grade_Parcela.Refresh
      Recalcula
      Usa_Timer = False
    End If
  End If
End Sub

Private Sub txt_parcelamento_diaFixo_LostFocus()
    Dim sDiaFixo As String
    txt_parcelamento_diaFixo.Text = Trim(txt_parcelamento_diaFixo.Text)

    If Len(txt_parcelamento_diaFixo.Text) > 0 And Not IsNumeric(txt_parcelamento_diaFixo.Text) Then
        MsgBox "Informe um dia fixo válido do mês.", vbInformation, "Atenção"
        txt_parcelamento_diaFixo.SetFocus
        Exit Sub
    Else
        If Len(txt_parcelamento_diaFixo.Text) > 0 And IsNumeric(txt_parcelamento_diaFixo.Text) Then
            If CInt(txt_parcelamento_diaFixo.Text) > 31 Or CInt(txt_parcelamento_diaFixo.Text) < 1 Then
                MsgBox "Informe um dia fixo válido do mês. De 01 a 31.", vbInformation, "Atenção"
                txt_parcelamento_diaFixo.SetFocus
                Exit Sub
            End If
        End If
    End If
    
'    If Not IsDate(txt_parcelamento_diaFixo.Text & "/" & cmb_mesInicioParcela.ListIndex + 1 & "/" & Year(Date)) Then
'        MsgBox "Data inválida!", vbInformation, "Atenção"
'        txt_parcelamento_diaFixo.SetFocus
'        Exit Sub
'    End If
    
    'cmb_mesInicioParcela.ListIndex = Month(Date)
End Sub

Private Sub Vale_KeyPress(KeyAscii As Integer)
   KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Vale_LostFocus()
  lbl_outrosPagamentos.Caption = Vale.Text
  Recalcula
End Sub

Private Function AnalisaClienteComCredito() As Boolean
On Error GoTo Erro

  Dim rstClientes As Recordset
  Dim strSQL As String
  Dim blnRecebimentoFaturado As Boolean
  Dim intX As Integer
  Dim dblValorRecebidoPrazo As Double
  Dim dblLimiteCredito As Double
  
  blnRecebimentoFaturado = False
  dblValorRecebidoPrazo = 0
  
  With Grade_Cheque
    .Redraw = False
    .MoveFirst
    
    For intX = 0 To .Rows - 1
      If IsNumeric(.Columns("Valor").Text) Then
        If CDbl(.Columns("Valor").Text) > 0 Then
          dblValorRecebidoPrazo = dblValorRecebidoPrazo + CDbl(.Columns("Valor").Text)
          
          If Not blnRecebimentoFaturado Then
            If IsDate(.Columns("Bom Para").Text) Then
              If CDate(.Columns("Bom Para").Text) > CDate(Data_Atual) Then
                blnRecebimentoFaturado = True
              End If
            End If
          End If
        End If
      End If
      
      .MoveNext
    Next intX
    
    .MoveFirst
    .Redraw = True
  End With
  
  With Grade_Parcela
    .Redraw = False
    .MoveFirst
    
    For intX = 0 To .Rows - 1
      If IsNumeric(.Columns("Valor").Text) Then
        If CDbl(.Columns("Valor").Text) > 0 Then
          dblValorRecebidoPrazo = dblValorRecebidoPrazo + CDbl(.Columns("Valor").Text)
          
          If Not blnRecebimentoFaturado Then
            If IsDate(.Columns("Data").Text) Then
              If CDate(.Columns("Data").Text) > CDate(Data_Atual) Then
                blnRecebimentoFaturado = True
              End If
            End If
          End If
        End If
      End If
      
      .MoveNext
    Next intX
    
    .MoveFirst
    .Redraw = True
  End With
  
  If IsNumeric(Cartão.Text) Then
    If CDbl(Cartão.Text) > 0 Then
      dblValorRecebidoPrazo = dblValorRecebidoPrazo + CDbl(Cartão.Text)
      'blnRecebimentoFaturado = True
    End If
  End If
  
  If Not blnRecebimentoFaturado Then
    blnRecebimentoFaturado = Conta.Value
  End If
  
  If blnRecebimentoFaturado Then
    Set rstClientes = db.OpenRecordset(" SELECT Faturado, [Limite Crédito] FROM Cli_For " & _
                                       " WHERE Código = " & lngCodigoCliente, dbOpenSnapshot)
    
    With rstClientes
      If Not (.BOF And .EOF) Then
        If (.Fields("Faturado") And .Fields("Limite Crédito") = 0) Then
          AnalisaClienteComCredito = True
        Else
          If (Not .Fields("Faturado")) And (dblValorRecebidoPrazo > 0) Then
            MsgBox "O cliente ao qual você está fazendo recebimento não pode fazer compra faturada. Para mudar essa opção entre no cadastro de clientes e marque a opção [Compra a Prazo]", vbCritical, "Quick Store"
            AnalisaClienteComCredito = False
          Else
            dblLimiteCredito = (.Fields("Limite Crédito").Value - Pega_Limite_Usado(lngCodigoCliente))
            '---------------------------------------------------------------------------
            '20/07/2006 - Andrea
            'Trocado campo arquivo .Fields("Limite Crédito").Value do If abaixo
            'pela variável dblLimiteCredito, pois estava errado a verificacao do limite
            'do cliente, no final do if, precisava ser comparado o valor que o cliente
            'estava tentando incluir na conta de cliente com o variável limite de crédito
            'e não com o que estava no arquivo, porque a variável leva em consideração
            'o que já foi usado de crédito pelo cliente e o campo do arquivo não.
            '---------------------------------------------------------------------------
            'If (CDbl(dblValorRecebidoPrazo) > CDbl(dblLimiteCredito)) Or (CBool(Conta.Value) And (CDbl(Total_Receber.Caption) > CDbl(.Fields("Limite Crédito").Value))) Then
            If (CDbl(dblValorRecebidoPrazo) > CDbl(dblLimiteCredito)) Or (CBool(Conta.Value) And (CDbl(Total_Receber.Caption) > CDbl(dblLimiteCredito))) Then
            '---------------------------------------------------------------------------
              If dblValorRecebidoPrazo > 0 Then
                MsgBox "O cliente ao qual você está fazendo o recebimento tem R$ " & _
                       Format(dblLimiteCredito, FORMAT_VALUE) & " de saldo para novas compras. O recebimento parcelado é de R$ " & _
                       Format(dblValorRecebidoPrazo, FORMAT_VALUE) & ". Não é possivel continuar !! ", vbCritical, "Quick Store"
              Else
                MsgBox "O cliente ao qual você está fazendo o recebimento tem R$ " & _
                       Format(dblLimiteCredito, FORMAT_VALUE) & " e você está tentando receber R$ " & Total_Receber.Caption & " em conta de cliente. Não é possivel continuar !! ", vbCritical, "Quick Store"
              End If
              
              AnalisaClienteComCredito = False
            Else
              AnalisaClienteComCredito = True
            End If
          End If
        End If
      End If
      
      If Not rstClientes Is Nothing Then .Close
      Set rstClientes = Nothing
    End With
  Else
    AnalisaClienteComCredito = True
  End If
  
  Call StatusMsg("")
  
  Exit Function
Erro:
  MsgBox "Erro na sub AnaliseClienteComCredito tela Recebimento " + Err.Number + " " + Err.Description, vbInformation, "Atenção"

End Function


