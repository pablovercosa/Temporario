VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmAcertaEmpSaida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Acerto de Empr�stimos - Sa�das"
   ClientHeight    =   8190
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   15420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AcertaEmpSaida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   15420
   Begin VB.Frame fraR 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   4260
      TabIndex        =   53
      Top             =   7050
      Width           =   2775
      Begin VB.CommandButton cmdExtrato 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E&xtrato"
         Height          =   435
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprime extrato de produtos consignados consolidado"
         Top             =   480
         Width           =   1995
      End
      Begin VB.OptionButton optSintetico 
         Appearance      =   0  'Flat
         Caption         =   "Sint�tico"
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
         Height          =   255
         Left            =   1410
         TabIndex        =   20
         Top             =   180
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton optAnal�tico 
         Appearance      =   0  'Flat
         Caption         =   "Anal�tico"
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
         Height          =   255
         Left            =   330
         TabIndex        =   19
         Top             =   180
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport crptExtrato 
      Left            =   4200
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraQtdeVendAcumu 
      Caption         =   "Qtde Vendida Acumulada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   420
      TabIndex        =   48
      Top             =   8280
      Visible         =   0   'False
      Width           =   2080
      Begin VB.TextBox txtQtdeVendidaAcumulada 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "CANCELADO 17/01/2005"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   480
         TabIndex        =   52
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"AcertaEmpSaida.frx":4E95A
         ForeColor       =   &H80000008&
         Height          =   1420
         Left            =   120
         TabIndex        =   51
         Top             =   550
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Qtde."
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   255
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   8280
      TabIndex        =   10
      Top             =   60
      Width           =   5145
      Begin VB.TextBox txtSequencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2190
         TabIndex        =   58
         Top             =   660
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton optOrdemItensUnicaSequencia 
         Appearance      =   0  'Flat
         Caption         =   "Itens da Sequ�ncia"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   57
         Top             =   750
         Width           =   1725
      End
      Begin VB.OptionButton optOrdemSequencia 
         Appearance      =   0  'Flat
         Caption         =   "Sequ�ncia"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optOrdemProduto 
         Appearance      =   0  'Flat
         Caption         =   "Produto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   510
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   6585
      Begin VB.OptionButton O_Empr�stimo 
         Appearance      =   0  'Flat
         Caption         =   "&Empr�stimo"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1350
         TabIndex        =   2
         Top             =   210
         Width           =   1140
      End
      Begin VB.OptionButton O_Acerto 
         Appearance      =   0  'Flat
         Caption         =   "&Acerto"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3990
         TabIndex        =   4
         Top             =   210
         Width           =   810
      End
      Begin VB.OptionButton O_Todas_Datas 
         Appearance      =   0  'Flat
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   300
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   870
      End
      Begin MSMask.MaskEdBox Data_Ace 
         Height          =   315
         Left            =   4800
         TabIndex        =   5
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   165
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox Data_Emp 
         Height          =   315
         Left            =   2490
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   165
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   6690
      TabIndex        =   7
      Top             =   60
      Width           =   1545
      Begin VB.OptionButton O_Aberto 
         Appearance      =   0  'Flat
         Caption         =   "Em &Aberto"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton O_Conclu�da 
         Appearance      =   0  'Flat
         Caption         =   "&Conclu�das"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   9
         Top             =   750
         Width           =   1080
      End
   End
   Begin VB.CommandButton B_Monta 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Pesquisar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13530
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   690
      Width           =   1845
   End
   Begin VB.CommandButton B_Atualiza_Empr�stimos 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Atualizar &Empr�stimos"
      Height          =   435
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7710
      Width           =   1785
   End
   Begin VB.CommandButton B_Atualiza_Tudo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gerar Venda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Atualizar os Empr�stimos e Gerar Sa�da com as Vendas"
      Top             =   7110
      Width           =   1785
   End
   Begin VB.CommandButton B_Atualiza 
      Caption         =   "&Atualizar Total"
      Height          =   400
      Left            =   2760
      TabIndex        =   22
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Imprimir Tabela"
      Height          =   435
      Left            =   2145
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7530
      Width           =   1995
   End
   Begin VB.Frame Frame_Mov 
      BackColor       =   &H00FFA324&
      Caption         =   "Contagem dos produtos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   60
      TabIndex        =   23
      Top             =   4980
      Width           =   7995
      Begin VB.CommandButton B_Confirma_Mov 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Confirmar "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1410
         Width           =   2145
      End
      Begin VB.CommandButton B_Cancela_Mov 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Cancelar"
         Height          =   465
         Left            =   5730
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1410
         Width           =   2175
      End
      Begin VB.TextBox Vendas_Prod 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   24
         Top             =   765
         Width           =   855
      End
      Begin VB.TextBox Dev_Prod 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   25
         Top             =   1155
         Width           =   855
      End
      Begin VB.TextBox Emp_Prod 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1770
         MaxLength       =   6
         TabIndex        =   26
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSMask.MaskEdBox Valor_Prod 
         Height          =   345
         Left            =   1620
         TabIndex        =   27
         Top             =   285
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   12648447
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "unidades"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2520
         TabIndex        =   55
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "unidades"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2520
         TabIndex        =   54
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "Pre�o Unit�rio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   450
         TabIndex        =   45
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "Cliente comprou"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   255
         TabIndex        =   37
         Top             =   810
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "devolveu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   855
         TabIndex        =   36
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Novo Empr�stimo"
         Height          =   255
         Left            =   300
         TabIndex        =   35
         Top             =   3900
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFA324&
         Caption         =   "Saldo empr�stimo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   34
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Label Saldo_Prod 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1620
         TabIndex        =   33
         Top             =   1560
         Width           =   855
      End
   End
   Begin VB.CheckBox O_Mostra_Detalhe 
      Caption         =   "Mostrar detalhes para cada linha"
      Height          =   255
      Left            =   90
      TabIndex        =   16
      Top             =   7170
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Caption         =   "Detalhes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8160
      TabIndex        =   15
      Top             =   5010
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Verificar Datas"
      Height          =   435
      Left            =   90
      Style           =   1   'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7530
      Width           =   1995
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2   'UseODBC
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
      Left            =   705
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1   'Dynaset
      RecordSource    =   ""
      Top             =   9300
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2   'UseODBC
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
      Left            =   2895
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1   'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   9300
      Visible         =   0   'False
      Width           =   1965
   End
   Begin MSMask.MaskEdBox Data_Acerto 
      Height          =   315
      Left            =   9660
      TabIndex        =   30
      ToolTipText     =   "Pressione F2 para Calend�rio"
      Top             =   7770
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Height          =   3705
      Left            =   60
      TabIndex        =   14
      Top             =   1170
      Width           =   15315
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
      UseGroups       =   -1  'True
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   12648447
      RowHeight       =   423
      ExtraHeight     =   185
      Groups(0).Width =   26088
      Groups(0).Caption=   "Produtos emprestados"
      Groups(0).Columns.Count=   14
      Groups(0).Columns(0).Width=   1667
      Groups(0).Columns(0).Caption=   "Sequ�ncia"
      Groups(0).Columns(0).Name=   "Sequ�ncia"
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   8
      Groups(0).Columns(0).FieldLen=   256
      Groups(0).Columns(0).Locked=   -1  'True
      Groups(0).Columns(1).Width=   2805
      Groups(0).Columns(1).Caption=   "Produto"
      Groups(0).Columns(1).Name=   "Produto"
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   256
      Groups(0).Columns(1).Locked=   -1  'True
      Groups(0).Columns(2).Width=   9234
      Groups(0).Columns(2).Caption=   "Nome"
      Groups(0).Columns(2).Name=   "Nome"
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   8
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(2).Locked=   -1  'True
      Groups(0).Columns(3).Width=   1032
      Groups(0).Columns(3).Caption=   "Tam"
      Groups(0).Columns(3).Name=   "Tamanho"
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   2
      Groups(0).Columns(3).FieldLen=   256
      Groups(0).Columns(3).Locked=   -1  'True
      Groups(0).Columns(4).Width=   926
      Groups(0).Columns(4).Caption=   "Cor"
      Groups(0).Columns(4).Name=   "Cor"
      Groups(0).Columns(4).DataField=   "Column 4"
      Groups(0).Columns(4).DataType=   2
      Groups(0).Columns(4).FieldLen=   256
      Groups(0).Columns(4).Locked=   -1  'True
      Groups(0).Columns(5).Width=   1058
      Groups(0).Columns(5).Caption=   "Edi��o"
      Groups(0).Columns(5).Name=   "Edi��o"
      Groups(0).Columns(5).DataField=   "Column 5"
      Groups(0).Columns(5).DataType=   3
      Groups(0).Columns(5).FieldLen=   256
      Groups(0).Columns(5).Locked=   -1  'True
      Groups(0).Columns(6).Width=   1217
      Groups(0).Columns(6).Caption=   "Ordem"
      Groups(0).Columns(6).Name=   "Ordem"
      Groups(0).Columns(6).DataField=   "Column 6"
      Groups(0).Columns(6).DataType=   3
      Groups(0).Columns(6).FieldLen=   256
      Groups(0).Columns(6).Locked=   -1  'True
      Groups(0).Columns(7).Width=   1826
      Groups(0).Columns(7).Caption=   "Data"
      Groups(0).Columns(7).Name=   "Data Opera��o"
      Groups(0).Columns(7).DataField=   "Column 7"
      Groups(0).Columns(7).DataType=   7
      Groups(0).Columns(7).FieldLen=   256
      Groups(0).Columns(7).Locked=   -1  'True
      Groups(0).Columns(8).Width=   2408
      Groups(0).Columns(8).Caption=   "Pre�o Unit�rio"
      Groups(0).Columns(8).Name=   "Pre�o Unit�rio"
      Groups(0).Columns(8).Alignment=   1
      Groups(0).Columns(8).DataField=   "Column 8"
      Groups(0).Columns(8).DataType=   8
      Groups(0).Columns(8).NumberFormat=   "0.00"
      Groups(0).Columns(8).FieldLen=   256
      Groups(0).Columns(8).Locked=   -1  'True
      Groups(0).Columns(9).Width=   1720
      Groups(0).Columns(9).Caption=   "Emprestou"
      Groups(0).Columns(9).Name=   "Saldo_Final"
      Groups(0).Columns(9).Alignment=   1
      Groups(0).Columns(9).DataField=   "Column 9"
      Groups(0).Columns(9).DataType=   8
      Groups(0).Columns(9).NumberFormat=   "###,##0"
      Groups(0).Columns(9).FieldLen=   256
      Groups(0).Columns(9).Locked=   -1  'True
      Groups(0).Columns(10).Width=   2196
      Groups(0).Columns(10).Caption=   "E agora ficou?"
      Groups(0).Columns(10).Name=   "Saldo_Prod"
      Groups(0).Columns(10).Alignment=   1
      Groups(0).Columns(10).DataField=   "Column 10"
      Groups(0).Columns(10).DataType=   8
      Groups(0).Columns(10).FieldLen=   256
      Groups(0).Columns(10).Locked=   -1  'True
      Groups(0).Columns(11).Width=   1191
      Groups(0).Columns(11).Visible=   0   'False
      Groups(0).Columns(11).Caption=   "Vendas_Prod"
      Groups(0).Columns(11).Name=   "Vendas_Prod"
      Groups(0).Columns(11).DataField=   "Column 11"
      Groups(0).Columns(11).DataType=   8
      Groups(0).Columns(11).FieldLen=   256
      Groups(0).Columns(12).Width=   2328
      Groups(0).Columns(12).Visible=   0   'False
      Groups(0).Columns(12).Caption=   "Dev_Prod"
      Groups(0).Columns(12).Name=   "Dev_Prod"
      Groups(0).Columns(12).DataField=   "Column 12"
      Groups(0).Columns(12).DataType=   8
      Groups(0).Columns(12).FieldLen=   256
      Groups(0).Columns(13).Width=   3493
      Groups(0).Columns(13).Visible=   0   'False
      Groups(0).Columns(13).Caption=   "Emp_Prod"
      Groups(0).Columns(13).Name=   "Emp_Prod"
      Groups(0).Columns(13).DataField=   "Column 13"
      Groups(0).Columns(13).DataType=   8
      Groups(0).Columns(13).FieldLen=   256
      _ExtentX        =   27014
      _ExtentY        =   6535
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "AcertaEmpSaida.frx":4E9EA
      DataSource      =   "Data1"
      Height          =   375
      Left            =   690
      TabIndex        =   0
      Top             =   120
      Width           =   1305
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
      Columns(0).Width=   9472
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2037
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2302
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBGrid Grade2 
      Bindings        =   "AcertaEmpSaida.frx":4E9FE
      Height          =   1980
      Left            =   8160
      TabIndex        =   46
      Top             =   5010
      Width           =   7215
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOdd    =   16777152
      RowHeight       =   423
      ExtraHeight     =   185
      Columns.Count   =   7
      Columns(0).Width=   1402
      Columns(0).Caption=   "Data Opera��o"
      Columns(0).Name =   "Data Opera��o"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Data Opera��o"
      Columns(0).DataType=   7
      Columns(0).FieldLen=   256
      Columns(1).Width=   1640
      Columns(1).Caption=   "Ordem"
      Columns(1).Name =   "Ordem"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Ordem"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   1349
      Columns(2).Caption=   "Saldo Anterior"
      Columns(2).Name =   "Saldo Anterior"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Saldo Anterior"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   256
      Columns(3).Width=   1508
      Columns(3).Caption=   "Comprou"
      Columns(3).Name =   "Vendas Cliente"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Vendas Cliente"
      Columns(3).DataType=   3
      Columns(3).FieldLen=   256
      Columns(4).Width=   1746
      Columns(4).Caption=   "Devolveu"
      Columns(4).Name =   "Devolu��o"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "Devolu��o"
      Columns(4).DataType=   3
      Columns(4).FieldLen=   256
      Columns(5).Width=   2302
      Columns(5).Caption=   "Novo Empr�stimo"
      Columns(5).Name =   "Novo Empr�stimo"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   1
      Columns(5).DataField=   "Novo Empr�stimo"
      Columns(5).DataType=   3
      Columns(5).FieldLen=   256
      Columns(6).Width=   1693
      Columns(6).Caption=   "Saldo Atual"
      Columns(6).Name =   "Saldo Atual"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   1
      Columns(6).DataField=   "Saldo Atual"
      Columns(6).DataType=   3
      Columns(6).FieldLen=   256
      _ExtentX        =   12726
      _ExtentY        =   3492
      _StockProps     =   79
      Caption         =   "Detalhes"
   End
   Begin VB.Label lbl_NomeCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2070
      TabIndex        =   56
      Top             =   150
      Width           =   4575
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "ou"
      Height          =   195
      Left            =   11280
      TabIndex        =   47
      Top             =   7530
      Width           =   180
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   11130
      X2              =   11580
      Y1              =   7950
      Y2              =   7950
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   11100
      X2              =   11580
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   11580
      Picture         =   "AcertaEmpSaida.frx":4EA12
      Top             =   7800
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   11610
      Picture         =   "AcertaEmpSaida.frx":4EAEA
      Top             =   7200
      Width           =   165
   End
   Begin VB.Shape Shape1 
      Height          =   705
      Left            =   7110
      Top             =   7395
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   225
      Left            =   90
      TabIndex        =   44
      Top             =   210
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor de Compras "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7110
      TabIndex        =   43
      Top             =   7095
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Linha atual"
      Height          =   195
      Left            =   7230
      TabIndex        =   42
      Top             =   7470
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   7290
      TabIndex        =   41
      Top             =   7095
      Width           =   360
   End
   Begin VB.Label Valor_Linha 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8115
      TabIndex        =   40
      Top             =   7425
      Width           =   1050
   End
   Begin VB.Label Valor_Total 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8115
      TabIndex        =   39
      Top             =   7740
      Width           =   1050
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pr�ximo acerto"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9660
      TabIndex        =   38
      Top             =   7155
      Width           =   1335
   End
End
Attribute VB_Name = "frmAcertaEmpSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSequenciaEmprestimo As String

Dim m_numRegistrosEmprestimo As Long
Dim m_numRegistrosEmprestimoDaUnicaSequencia As Long

Private rsClientes As Recordset
Private rsProdutos As Recordset
Private rsEmprestimos As Recordset
Private rsEstoque As Recordset
Private rsEstoque_Final As Recordset
Private rsResumo_Di�rio As Recordset
Private rsParametros As Recordset
Private rsSaidas As Recordset
Private rsSaidas_Prod As Recordset

Private Type Tab_Emp
  Sequ�ncia As Long
  Produto As String
  Nome As String
  Tamanho As Integer
  Cor As Integer
  Edi��o As Long
  Ordem As Long
  Data As Date
  Saldo_Ant As Double
  Valor_Unit As Double
  
  Novo_Emp As Double
  Nova_Venda As Double
  Nova_Devol As Double
  Novo_Saldo As Double
  
  '27/08/2003 - mpdea
  'Exibi��o por ordem de c�digo do produto
  Ordenacao As String
  
End Type

Private Type Tab_Venda
  sProduto As String
  nVenda As Long
End Type

'01/10/2003 - mpdea
'Redimensionado tamanho m�ximo do array (1000 -> 5000)
Private Const EMP_ARRAY_SIZE As Integer = 5000
Private Empr�stimos(EMP_ARRAY_SIZE) As Tab_Emp

Private Linha As Integer

Private Estado As String
Private rsEstados As Recordset

Private Function Acha_Reg_Estoque(Filial As Integer, Dia As Date, Produto As String, Tamanho As Integer, Cor As Integer, Edi��o As Long) As Variant
  Dim Est_Final As Single
  Dim Erro As Boolean
   
  Est_Final = 0
  rsEstoque.Index = "Produto"
  rsEstoque.Seek "=", Filial, Dia, Produto, Tamanho, Cor, Edi��o
                        
  If Not rsEstoque.NoMatch Then
    Acha_Reg_Estoque = rsEstoque.Bookmark
    Exit Function
  End If
       
       
  rsProdutos.Index = "C�digo"
  rsProdutos.Seek "=", Produto
       
           
  'N�o tem registro no dia atual
  rsEstoque.Index = "Data2"
  Erro = False
  rsEstoque.Seek ">", Filial, Produto, Tamanho, Cor, Edi��o, Dia
   
  If rsEstoque.NoMatch Then Erro = True
  If Erro = False Then If rsEstoque("Filial") <> Filial Then Erro = True
  If Erro = False Then If rsEstoque("Produto") <> Produto Then Erro = True
  If Erro = False Then If rsEstoque("Tamanho") <> Tamanho Then Erro = True
  If Erro = False Then If rsEstoque("Cor") <> Cor Then Erro = True
  If Erro = False Then If rsEstoque("Edi��o") <> Edi��o Then Erro = True
  
  If Erro = False Then  'j� tinha em outro dia
    Est_Final = rsEstoque("Estoque Final")
    rsEstoque.AddNew
    rsEstoque("Filial") = gnCodFilial
    rsEstoque("Data") = Dia
    rsEstoque("Produto") = Produto
    rsEstoque("Tamanho") = Tamanho
    rsEstoque("Cor") = Cor
    rsEstoque("Edi��o") = Edi��o
    rsEstoque("Classe") = rsProdutos("Classe")
    rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
    rsEstoque("Estoque Anterior") = Est_Final
    '30/11/2004 - Daniel
    'Puros Tabacos - RS ??
    '----------------------------------------
    rsEstoque.Update
    rsEstoque.Index = "Produto"
    rsEstoque.Seek "=", Filial, Dia, Produto, Tamanho, Cor, Edi��o
    Acha_Reg_Estoque = rsEstoque.Bookmark
    Exit Function
  End If
   
  If Erro = True Then  'Cria novo para o produto
    rsEstoque.AddNew
    rsEstoque("Filial") = gnCodFilial
    rsEstoque("Data") = Dia
    rsEstoque("Produto") = Produto
    rsEstoque("Tamanho") = Tamanho
    rsEstoque("Cor") = Cor
    rsEstoque("Edi��o") = Edi��o
    rsEstoque("Classe") = rsProdutos("Classe")
    rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
    rsEstoque("Estoque Anterior") = 0
    rsEstoque.Update
    rsEstoque.Index = "Produto"
    rsEstoque.Seek "=", Filial, Dia, Produto, Tamanho, Cor, Edi��o
    Acha_Reg_Estoque = rsEstoque.Bookmark
    Exit Function
  End If

End Function

'02/10/2003 - mpdea
'Implementado tratamento de erro e transa��o
Private Function Atu_Empr�stimo() As Integer
  Dim i As Integer
  Dim Qtde_Dev As Double
  Dim Qtde_Emp As Double
  Dim Qtde_Vendas As Double
  Dim Erro As Boolean
  Dim Est_Final As Single
  Dim Num_Reg As Variant
  Dim Tot_Vendas As Double
  Dim Tot_Devolu��o As Double
  Dim Tot_Empr�stimos As Double
  
  Dim blnInTransaction As Boolean
  
    
  On Error GoTo ErrHandler
  
  
  Qtde_Dev = 0
  Qtde_Emp = 0
  Qtde_Vendas = 0
  
  Call StatusMsg("")
  
  For i = 0 To Grade1.Rows
    If Empr�stimos(i).Nova_Devol <> 0 Then Qtde_Dev = Qtde_Dev + Empr�stimos(i).Nova_Devol
    If Empr�stimos(i).Novo_Emp <> 0 Then Qtde_Emp = Qtde_Emp + Empr�stimos(i).Novo_Emp
    If Empr�stimos(i).Nova_Venda <> 0 Then Qtde_Vendas = Qtde_Vendas + Empr�stimos(i).Nova_Venda
  Next i
  
  If Qtde_Dev = 0 And Qtde_Emp = 0 And Qtde_Vendas = 0 Then
    DisplayMsg "N�o existe nenhum movimento de vendas, devolu��es ou empr�stimos."
    Atu_Empr�stimo = 1
    Exit Function
  End If
  
  If IsNumeric(Qtde_Emp) Then
    If Qtde_Emp > 0 Then
      If Not IsDate(Data_Acerto.Text) Then
        DisplayMsg "Digite a data para o pr�ximo acerto."
        Data_Acerto.SetFocus
        Atu_Empr�stimo = 1
        Exit Function
      End If
    Else
      Data_Acerto.Text = Data_Atual
    End If
  End If
  
  If CDate(Data_Acerto.Text) < CDate(Data_Atual) Then
    DisplayMsg "Data de acerto inv�lida, verifique."
    Data_Acerto.SetFocus
    Atu_Empr�stimo = 1
    Exit Function
  End If
  
  '02/10/2003 - mpdea
  'Inicia transa��o
  ws.BeginTrans
  blnInTransaction = True
  
  rsProdutos.Index = "C�digo"
  
  Dim rstOperacoesSaida As Recordset
  Dim rstSaidas As Recordset
  Dim blnBaixaEstoque As Boolean
  
  For i = 0 To Grade1.Rows
    blnBaixaEstoque = False
    
    Set rstSaidas = db.OpenRecordset(" SELECT * FROM Sa�das " & _
                                     " WHERE Filial = " & gnCodFilial & _
                                     " AND Sequ�ncia = " & Empr�stimos(i).Sequ�ncia, dbOpenSnapshot)
    With rstSaidas
      If Not (.BOF And .EOF) Then
        Set rstOperacoesSaida = db.OpenRecordset("SELECT * FROM [Opera��es Sa�da] WHERE C�digo = " & .Fields("Opera��o").Value & "", dbOpenSnapshot)
        
        With rstOperacoesSaida
          If Not (.BOF And .EOF) Then
            blnBaixaEstoque = .Fields("Estoque").Value
          End If
          
          .Close
          Set rstOperacoesSaida = Nothing
        End With
      End If
      
      .Close
      Set rstSaidas = Nothing
    End With
  
    If Empr�stimos(i).Nova_Devol <> 0 Or _
       Empr�stimos(i).Novo_Emp <> 0 Or _
       Empr�stimos(i).Nova_Venda <> 0 Then
    
      rsProdutos.Seek "=", Empr�stimos(i).Produto
      
      Rem Posiciona no registro do estoque
      Num_Reg = Acha_Reg_Estoque(gnCodFilial, CDate(Data_Atual), _
        Empr�stimos(i).Produto, Empr�stimos(i).Tamanho, _
        Empr�stimos(i).Cor, Empr�stimos(i).Edi��o)
      
      
      rsEstoque.Bookmark = Num_Reg
      
      Rem Neste ponto tem o registro de estoque no buffer
      
      If blnBaixaEstoque Then
        Dim strSQL As String
        Dim Estoque_Final As Double
        Dim rstEstoque  As Recordset
        '-------------------------------------------------------------------------------------
		
		'10/10/2003 - Maikel
		'Modificada a forma de analisar a tabela de estoque. Da forma antiga gerava erro 3022 ao efetuar movimenta��o com data retroativa.
		strSQL = "SELECT * FROM Estoque WHERE " & _
				 " Filial = " & gnCodFilial & _
				 " AND Produto = '" & Empr�stimos(i).Produto & "'" & _
				 " AND Tamanho = " & Empr�stimos(i).Tamanho & _
				 " AND Cor = " & Empr�stimos(i).Cor & _
				 " AND Edi��o = " & Empr�stimos(i).Edi��o & _
				 " ORDER BY Data "
		
		Set rstEstoque = db.OpenRecordset(strSQL, dbOpenSnapshot)
		
		With rstEstoque
		  If Not (.BOF And .EOF) Then
			.MoveFirst
			.MoveLast
			Estoque_Final = .Fields("Estoque Final")
		  Else
			Estoque_Final = 0
		  End If
		  
		  .Close
		  Set rstEstoque = Nothing
		End With
		
		strSQL = "SELECT * FROM Estoque WHERE " & _
				 " Filial = " & gnCodFilial & _
				 " AND Produto = '" & Empr�stimos(i).Produto & "'" & _
				 " AND Tamanho = " & Empr�stimos(i).Tamanho & _
				 " AND Cor = " & Empr�stimos(i).Cor & _
				 " AND Edi��o = " & Empr�stimos(i).Edi��o & _
				 " AND Data = #" & Format(Data_Atual, "mm/dd/yyyy") & "#"
				
		Set rstEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
		
		With rstEstoque
		  If (.BOF And .EOF) Then
			.AddNew
			.Fields("Filial").Value = gnCodFilial
			.Fields("Data").Value = Data_Atual
			.Fields("Produto").Value = Empr�stimos(i).Produto
			.Fields("Tamanho").Value = Empr�stimos(i).Tamanho
			.Fields("Cor").Value = Empr�stimos(i).Cor
			.Fields("Edi��o").Value = Empr�stimos(i).Edi��o
			.Fields("Classe").Value = rsProdutos("Classe").Value
			.Fields("Sub Classe").Value = rsProdutos("Sub Classe").Value
			.Fields("Estoque Anterior").Value = Estoque_Final
			.Update
			.Requery
		  End If
		End With
        '-------------------------------------------------------------------------------------

        rstEstoque.Edit
        rstEstoque("Empre Entra") = rstEstoque("Empre Entra") + Empr�stimos(i).Nova_Devol
        rstEstoque("Valor Empre Entra") = rstEstoque("Valor Empre Entra") + (Empr�stimos(i).Nova_Devol * Empr�stimos(i).Valor_Unit)
        
        Estoque_Final = rstEstoque("Estoque Anterior") - rsEstoque("Vendas") + rstEstoque("Compras")
        Estoque_Final = Estoque_Final - rstEstoque("Transf Sa�da") + rstEstoque("Transf Entra")
        Estoque_Final = Estoque_Final - rstEstoque("Ajuste Sa�da") + rstEstoque("Ajuste Entra")
        Estoque_Final = Estoque_Final - rstEstoque("Gr�tis Sa�da") + rstEstoque("Gr�tis Entra")
        Estoque_Final = Estoque_Final - rstEstoque("Empre Sa�da") + rstEstoque("Empre Entra")
        
        '08/08/2003 - maikel
        'Descomentada a soma da coluna Devolu��o para resolver o problema de estoque
        Estoque_Final = Estoque_Final - rstEstoque("Quebras") + rstEstoque("Devolu��o")
  
        If rsProdutos("Estoque") = False Then
          Estoque_Final = 0
        End If
  
        rstEstoque("Estoque Final") = Estoque_Final
        rstEstoque.Update
        
        Rem Acerta Estoque Final
        Grava_Estoque_Final gnCodFilial, Empr�stimos(i).Produto, _
              Empr�stimos(i).Tamanho, Empr�stimos(i).Cor, Empr�stimos(i).Edi��o, _
              Estoque_Final, CDate(Data_Atual)
              
        rstEstoque.Close
        Set rstEstoque = Nothing
      End If
      
      Rem Acerta o resumo di�rio
      rsResumo_Di�rio.Index = "Data"
      rsResumo_Di�rio.Seek "=", gnCodFilial, Data_Atual
      If rsResumo_Di�rio.NoMatch Then
        rsResumo_Di�rio.AddNew
        rsResumo_Di�rio("Filial") = gnCodFilial
        rsResumo_Di�rio("Data") = Data_Atual
        rsResumo_Di�rio("Valor E Sa�da") = 0
        rsResumo_Di�rio("Valor E Entrada") = 0
      Else
        rsResumo_Di�rio.Edit
      End If
      
      rsResumo_Di�rio("Valor E Sa�da") = rsResumo_Di�rio("Valor E Sa�da") + (Empr�stimos(i).Novo_Emp * Empr�stimos(i).Valor_Unit)
      rsResumo_Di�rio("Valor E Entrada") = rsResumo_Di�rio("Valor E Entrada") + (Empr�stimos(i).Nova_Devol * Empr�stimos(i).Valor_Unit)
      
      rsResumo_Di�rio.Update
      
      Rem Grava Status OK para linha atual
      rsEmprestimos.Index = "Cliente"
      rsEmprestimos.Seek "=", gnCodFilial, Empr�stimos(i).Sequ�ncia, (Combo_Cliente.Text), Empr�stimos(i).Produto, Empr�stimos(i).Tamanho, Empr�stimos(i).Cor, Empr�stimos(i).Edi��o, Empr�stimos(i).Ordem
      If rsEmprestimos.NoMatch Then
        
        '02/10/2003 - mpdea
        'Desfaz transa��o
        If blnInTransaction Then ws.Rollback
        
        MsgBox "Erro ao encontrar empr�stimo."
        Atu_Empr�stimo = 1
        Exit Function
      End If
      
      Tot_Vendas = 0
      Tot_Devolu��o = 0
      Tot_Empr�stimos = 0
      
      rsEmprestimos.Edit
        rsEmprestimos("Conclu�do") = True
        Est_Final = rsEmprestimos("Saldo Atual")
      rsEmprestimos.Update
      
      rsEmprestimos.AddNew
        rsEmprestimos("Filial") = gnCodFilial
        rsEmprestimos("Sequ�ncia") = Empr�stimos(i).Sequ�ncia
        rsEmprestimos("Cliente") = Val(Combo_Cliente.Text)
        rsEmprestimos("Produto") = Empr�stimos(i).Produto
        rsEmprestimos("Tamanho") = Empr�stimos(i).Tamanho
        rsEmprestimos("Cor") = Empr�stimos(i).Cor
        rsEmprestimos("Edi��o") = Empr�stimos(i).Edi��o
        rsEmprestimos("Ordem") = Empr�stimos(i).Ordem + 1
        rsEmprestimos("Data Opera��o") = Data_Atual
        rsEmprestimos("Saldo Anterior") = Est_Final
        rsEmprestimos("Pre�o Unit�rio") = Empr�stimos(i).Valor_Unit
        
        rsEmprestimos("Vendas Cliente") = Empr�stimos(i).Nova_Venda
        rsEmprestimos("Novo Empr�stimo") = Empr�stimos(i).Novo_Emp
        rsEmprestimos("Devolu��o") = Empr�stimos(i).Nova_Devol
        
        rsEmprestimos("Saldo Atual") = Est_Final - Empr�stimos(i).Nova_Venda - Empr�stimos(i).Nova_Devol + Empr�stimos(i).Novo_Emp
        rsEmprestimos("Data Cobran�a") = Data_Acerto.Text
        rsEmprestimos("Data Altera��o") = Format(Date, "dd/mm/yyyy")
        If rsEmprestimos("Saldo Atual") = 0 Then
          rsEmprestimos("Conclu�do") = True
        End If
      rsEmprestimos.Update
    
    End If
    
  Next i
  
  '02/10/2003 - mpdea
  'Finaliza transa��o
  ws.CommitTrans
  blnInTransaction = False
  
  Atu_Empr�stimo = 0
  
  Exit Function
  
ErrHandler:
  '02/10/2003 - mpdea
  'Desfaz transa��o
  If blnInTransaction Then ws.Rollback
  
  Atu_Empr�stimo = -1
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

Private Sub Atualiza_Detalhes()
  Dim Rec_SQL As Recordset
  Dim sSql As String
  
  sSql = "Select [Data Opera��o], Ordem, [Saldo Anterior], [Vendas Cliente], [Devolu��o], [Novo Empr�stimo] , [Saldo Atual]"
  sSql = sSql + " From [Consigna��o Sa�da] "
  sSql = sSql + " Where [Consigna��o Sa�da].Filial = " + str(gnCodFilial)
  sSql = sSql + " And [Consigna��o Sa�da].Sequ�ncia = " + str(Grade1.Columns(0).Text)
  sSql = sSql + " And [Consigna��o Sa�da].Produto = '" + Grade1.Columns(1).Text + "'"
  sSql = sSql + " And [Consigna��o Sa�da].Tamanho = " + Grade1.Columns(3).Text
  sSql = sSql + " And [Consigna��o Sa�da].Cor = " + Grade1.Columns(4).Text
  sSql = sSql + " And [Consigna��o Sa�da].Edi��o = " + Grade1.Columns(5).Text
  sSql = sSql + " Order By -Ordem"
  
  Set Rec_SQL = db.OpenRecordset(sSql, dbOpenDynaset)
  
  Grade2.DataMode = 1
  
  Set Data2.Recordset = Rec_SQL
  
  Grade2.DataMode = 0
  
  Grade2.ReBind
   
  Grade2.Columns(0).Width = 1100
  Grade2.Columns(0).Locked = True
  
  Grade2.Columns(1).Width = 600
  Grade2.Columns(1).Locked = True
   
  Grade2.Columns(2).Width = 900
  Grade2.Columns(2).Caption = "Saldo Ant."
  Grade2.Columns(2).Locked = True
  
  Grade2.Columns(3).Width = 700
  Grade2.Columns(3).Locked = True
  
  Grade2.Columns(4).Width = 600
  Grade2.Columns(4).Caption = "Dev."
  Grade2.Columns(4).Locked = True
  
  Grade2.Columns(5).Width = 600
  Grade2.Columns(5).Caption = "Empr."
  Grade2.Columns(5).Locked = True
  
  Grade2.Columns(6).Width = 1100
  Grade2.Columns(6).Caption = "Saldo Final"
  Grade2.Columns(6).Locked = True
  
  Grade2.MoveLast
  Grade2.MoveFirst
  
End Sub

Private Sub Recalcula_Saldo()
  Dim Saldo As Double
  
  Saldo = 0
  If IsNumeric(Grade1.Columns(10).Text) Then
    Saldo = Grade1.Columns(10).Text
  
    Saldo = Saldo - Val(Replace(Vendas_Prod.Text, ",", "."))
    Saldo = Saldo - Val(Replace(Dev_Prod.Text, ",", "."))
    Saldo = Saldo + Val(Replace(Emp_Prod.Text, ",", "."))
  
    Saldo_Prod.Caption = Saldo
  End If
  
End Sub

'01/10/2003 - mpdea
'Modificado for next para at� o limite do array Empr�stimos
Private Sub B_Atualiza_Click()
  Dim i As Integer
  Dim Aux_Dbl As Double

  Aux_Dbl = 0
  For i = 0 To UBound(Empr�stimos)
   If Empr�stimos(i).Nova_Venda <> 0 Then
     Aux_Dbl = Aux_Dbl + (Empr�stimos(i).Nova_Venda * Empr�stimos(i).Valor_Unit)
   End If
  Next i

  Valor_Total.Caption = Format(Aux_Dbl, "###,###,###,##0.00")

End Sub

Private Sub B_Atualiza_Empr�stimos_Click()
  Dim Resp As Integer
  Resp = Atu_Empr�stimo
  
  If Resp = 0 Then B_Monta_Click
  
  Vendas_Prod.Text = 0
  Dev_Prod.Text = 0
  Emp_Prod.Text = 0
  Saldo_Prod.Caption = 0
  Valor_Prod.Text = 0
  
  '-------------------------------------------------------------------------
  ' C�DIGO CANCELADO !!!
  '-------------------------------------------------------------------------
  '14/01/2005 - Daniel
  '
  'Projeto.......: Tratamento da Quantidade Vendida Acumulada
  'Finalidade....: Corre��o do totalizador do valor da compra
  'Solicitante...: Aura Prata
  'txtQtdeVendidaAcumulada.Text = 0
  'Valor_Linha.Caption = "0,00"
  'Valor_Total.Caption = "0,00"
  '-------------------------------------------------------------------------
  
  Grade1.Refresh
  
End Sub

'02/10/2003 - mpdea
'Implementado tratamento de erro e transa��o
Private Sub B_Atualiza_Tudo_Click()
  Dim i As Integer
  Dim Qtde_Vendas As Long
  Dim Erro As Boolean
  Dim Est_Final As Single
  Dim Num_Reg As Variant
  Dim Tot_Vendas As Long
  Dim Tot_Devolu��o As Long
  Dim Tot_Empr�stimos As Long
  Dim Mov As Long
  Dim Linha As Integer
  Dim Produto As String
  Dim Tamanho As String
  Dim Cor As String
  Dim Prod_Sem_Grade As String
  Dim Total As Double
  Dim Texto As String
  Dim sOperacaoSaida As String
  Dim sTabelaPrecos As String
  Dim sObservacoesOrigem As String
 
  Dim nAuxICM As Double
 
  Dim blnInTransaction As Boolean

  
  On Error GoTo ErrHandler
 
 
  Qtde_Vendas = 0
  
  Call StatusMsg("")
  
  For i = 0 To Grade1.Rows - 1
    If Empr�stimos(i).Nova_Venda <> 0 Then Qtde_Vendas = Qtde_Vendas + Empr�stimos(i).Nova_Venda
  Next i
  
  If Qtde_Vendas = 0 Then
    DisplayMsg "N�o existem vendas. Use o bot�o Atualiza Empr�stimo."
    Exit Sub
  End If
  
  i = Atu_Empr�stimo
  If i <> 0 Then Exit Sub
 
 
  '02/10/2003 - mpdea
  'Inicia transa��o
  ws.BeginTrans
  blnInTransaction = True
  
  
  ' Pelo codigo_sequencia eu vou na tabela de saidas e recupero o codigo_operacao, ent�o vou na tabela de
  ' opera��es sa�das e vejo se o campo 'estoque' esta ativo (ou seja, movimento estoque)
  ' caso sim escrevo no campo observa�oes: 'origem emprestimo mov estoque'
  ' caso n�o escrevo no campo observa��es: 'origem emprestimo sem mov estoque'
  Dim rsSaidasAuxiliar As Recordset
  Dim sSqlAux As String
  sSqlAux = "Select Estoque From Sa�das S, [Opera��es Sa�da] Op "
  sSqlAux = sSqlAux + " Where S.Sequ�ncia = " + sSequenciaEmprestimo
  sSqlAux = sSqlAux + " And S.Opera��o = Op.C�digo "
  
  Set rsSaidasAuxiliar = db.OpenRecordset(sSqlAux, dbOpenDynaset)
  
  sObservacoesOrigem = ""
  With rsSaidasAuxiliar
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If IsNumeric(.Fields("Estoque").Value) Then
        sObservacoesOrigem = "Venda gerada por Empr�stimo (MovEst)" ' MovEst = movimentou estoque
      Else
        sObservacoesOrigem = "Venda gerada por Empr�stimo (SemMovEst)" ' SemMovEst = n�o movimentou estoque
      End If
      
    End If
    .Close
  End With
  Set rsSaidasAuxiliar = Nothing
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then
    '02/10/2003 - mpdea
    'Desfaz transa��o
    If blnInTransaction Then ws.Rollback
    
    MsgBox ("Erro ao encontrar par�metros")
    Exit Sub
  End If
  
  sOperacaoSaida = "0"
  sOperacaoSaida = rsParametros("Consignacao_OpFechamento")
  
  If sOperacaoSaida = "" Then
      sOperacaoSaida = "0"
      MsgBox "Vincule na tela de 'Par�metros da Empresa/Filial', aba 'Sa�das', no campo 'Opera��o para Venda Estadual de consignado'. Caso ainda n�o exista esta opera��o, crie uma pelo menu Cadastro, op��o Sa�das."
  End If
  
  sTabelaPrecos = ""
  sTabelaPrecos = rsParametros("Consignacao_TabelaPrecos")
  
  rsParametros.Edit
    rsParametros("�ltima Movimenta��o") = gnGetNextSequencia(gnCodFilial) 'rsParametros("�ltima Movimenta��o") + 1
    Mov = rsParametros("�ltima Movimenta��o")
  rsParametros.Update
 
  Total = 0
  
  Linha = 1
  For i = 0 To Grade1.Rows - 1
    If Empr�stimos(i).Nova_Venda <> 0 Then
      Produto = Empr�stimos(i).Produto
      Prod_Sem_Grade = Produto
      If Empr�stimos(i).Tamanho <> 0 Then
         Tamanho = "000" + Trim(str(Empr�stimos(i).Tamanho))
         Tamanho = Right(Tamanho, 3)
         
         Cor = "000" + Trim(str(Empr�stimos(i).Cor))
         Cor = Right(Cor, 3)
         Produto = Produto + Tamanho + Cor
      End If
      
      rsSaidas_Prod.AddNew
        rsSaidas_Prod("Filial") = gnCodFilial
        rsSaidas_Prod("Sequ�ncia") = Mov
        rsSaidas_Prod("Linha") = Linha
        rsSaidas_Prod("C�digo") = Produto
        rsSaidas_Prod("Qtde") = Empr�stimos(i).Nova_Venda
        rsSaidas_Prod("Pre�o") = Empr�stimos(i).Valor_Unit
        rsSaidas_Prod("Desconto") = 0
        
        nAuxICM = gvGetValueInTable("Produtos", "[Percentual ICM]", ftNumero, "C�digo", ftTexto, Produto)
        'Mostra ICM do Estado
        If Estado = "" Then
          rsSaidas_Prod("ICM") = nAuxICM
        ElseIf Estado <> "" Then
          rsEstados.Index = "Estado"
          rsEstados.Seek "=", Estado
          If rsEstados.NoMatch Then
            rsSaidas_Prod("ICM") = nAuxICM
          ElseIf Not rsEstados.NoMatch Then
            If rsEstados("ICM") = -1 Then
              rsSaidas_Prod("ICM") = nAuxICM
            Else
              rsSaidas_Prod("ICM") = rsEstados("ICM")
            End If
          End If
        End If
        
        rsSaidas_Prod("IPI") = gvGetValueInTable("Produtos", "[Percentual IPI]", ftNumero, "C�digo", ftTexto, Produto)
        rsSaidas_Prod("Pre�o Final") = rsSaidas_Prod("Qtde") * rsSaidas_Prod("Pre�o")
        rsSaidas_Prod("Etiqueta") = False
        rsSaidas_Prod("C�digo Sem Grade") = Prod_Sem_Grade
        
        Total = Total + rsSaidas_Prod("Pre�o Final")
        
      rsSaidas_Prod.Update
      Linha = Linha + 1
    End If
  Next i
 
  Rem Grava Sa�da
  rsSaidas.AddNew
  rsSaidas("Filial") = gnCodFilial
  rsSaidas("Data") = Data_Atual
  rsSaidas("Sequ�ncia") = Mov
  rsSaidas("Opera��o") = sOperacaoSaida
  rsSaidas("Caixa") = "1"
  rsSaidas("Tabela") = sTabelaPrecos
  rsSaidas("Digitador") = "0"
  rsSaidas("Operador") = "0"
  rsSaidas("Cliente") = Val(Combo_Cliente.Text)
  rsSaidas("Observa��es") = sObservacoesOrigem
  rsSaidas("Produtos") = Total
  rsSaidas("Total") = Total
  rsSaidas("Efetivada") = False
  
  rsSaidas.Update
 
  '02/10/2003 - mpdea
  'Finaliza transa��o
  ws.CommitTrans
  blnInTransaction = False
 
  Texto = "A movimenta��o " + str(Mov) + " foi criada."
  Texto = Texto & vbCrLf & Chr(13)
  Texto = Texto + "Voc� DEVE entrar na tela de SA�DAS e verificar a movimenta��o, os valores, impostos e quantidades de produtos. Se os produtos vendidos tem ICM ou IPI verifique tamb�m estes impostos."
  MsgBox (Texto), vbInformation, "Infoma��o"
 
  '-------------------------------------------------------------------------
  '02/10/2003 - mpdea
  'Atualiza informa��es na tela
  Call B_Monta_Click
  
  Vendas_Prod.Text = 0
  Dev_Prod.Text = 0
  Emp_Prod.Text = 0
  Saldo_Prod.Caption = 0
  Valor_Prod.Text = 0
  
  '-------------------------------------------------------------------------
  ' C�DIGO CANCELADO !!!
  '-------------------------------------------------------------------------
  '14/01/2005 - Daniel
  '
  'Projeto.......: Tratamento da Quantidade Vendida Acumulada
  'Finalidade....: Corre��o do totalizador do valor da compra
  'Solicitante...: Aura Prata
  'txtQtdeVendidaAcumulada.Text = 0
  'Valor_Linha.Caption = "0,00"
  'Valor_Total.Caption = "0,00"
  '-------------------------------------------------------------------------
  
  Grade1.Refresh
  '-------------------------------------------------------------------------
 
  Exit Sub
  
ErrHandler:
  '02/10/2003 - mpdea
  'Desfaz transa��o
  If blnInTransaction Then ws.Rollback
  
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub B_Cancela_Mov_Click()
  '-------------------------------------------------------------------------
  ' C�DIGO CANCELADO !!!
  '-------------------------------------------------------------------------
  '14/01/2005 - Daniel
  '
  'Projeto.......: Tratamento da Quantidade Vendida Acumulada
  'Finalidade....: Corre��o do totalizador do valor da compra
  'Solicitante...: Aura Prata
  'Call ZerarQtdeVendidaAcumulada
  '-------------------------------------------------------------------------
  
  Vendas_Prod.Text = 0
  Dev_Prod.Text = 0
  Emp_Prod.Text = 0
  
End Sub

Private Sub B_Confirma_Mov_Click()
  
 Call StatusMsg("")

 If Grade1.SelBookmarks.Count = 0 Then
   DisplayMsg "Selecione uma linha antes."
   Exit Sub
 End If
 
 If CDbl(Saldo_Prod.Caption) < 0 Then
   DisplayMsg "Estoque n�o pode ficar negativo."
   Exit Sub
 End If

 '-------------------------------------------------------------------------
 ' C�DIGO CANCELADO !!!
 '-------------------------------------------------------------------------
 '14/01/2005 - Daniel
 '
 'Projeto.......: Tratamento da Quantidade Vendida Acumulada
 'Finalidade....: Corre��o do totalizador do valor da compra
 'Solicitante...: Aura Prata
 'If IsNumeric(Vendas_Prod.Text) Then
 '  Call AtualizarQtdeVendidaAcumulada(Vendas_Prod.Text)
 'End If
 
 Grade1.Columns(11).Text = CDbl(Vendas_Prod.Text)
 Grade1.Columns(12).Text = CDbl(Dev_Prod.Text)
 Grade1.Columns(13).Text = CDbl(Emp_Prod.Text)
 Grade1.Columns(10).Text = CDbl(Saldo_Prod.Caption)
 Grade1.Columns(8).Text = CDbl(Valor_Prod.Text)
 
 Grade1.Update

 B_Atualiza_Click
 
 '-------------------------------------------------------------------------
 ' C�DIGO CANCELADO !!!
 '-------------------------------------------------------------------------
 '14/01/2005 - Daniel
 '
 'Projeto.......: Tratamento da Quantidade Vendida Acumulada
 'Finalidade....: Corre��o do totalizador do valor da compra
 'Solicitante...: Aura Prata
 'Valor_Linha.Caption = Format(CDbl(txtQtdeVendidaAcumulada.Text) * CDbl(Grade1.Columns(8).Text), FORMAT_VALUE)
 '-------------------------------------------------------------------------

End Sub

Private Sub B_Imprime_Click()
  Grade1.PrintData ssPrintAllRows, True, True
End Sub

Private Sub B_Monta_Click()
  Dim Aux_Cliente As Long
  Dim Aux_Produto As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edi��o As Integer
  Dim Aux_ordem As Long
  Dim Aux_Seq As Long
  
  Call StatusMsg("")
  
  If optOrdemItensUnicaSequencia.Value = True Then
      If Trim(txtSequencia.Text) = "" Then
          MsgBox "Informe uma Sequ�ncia v�lida.", vbInformation, "Aten��o"
          txtSequencia.SetFocus
          Exit Sub
      End If
  End If
 
  If Len(Combo_Cliente.Text) = 0 Then
    DisplayMsg "Cliente incorreto."
    Combo_Cliente.SetFocus
  End If

  If O_Acerto.Value = True And Not IsDate(Data_Ace.Text) Then
    DisplayMsg "Data inv�lida, verifique."
    Data_Ace.SetFocus
    Exit Sub
  End If
  
  If O_Empr�stimo.Value = True And Not IsDate(Data_Emp.Text) Then
    DisplayMsg "Data inv�lida, verifique."
    Data_Emp.SetFocus
    Exit Sub
  End If
  
  '02/10/2003 - mpdea
  'Status
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde...")

  Aux_Cliente = 0
  Aux_Produto = 0
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edi��o = 0
  Aux_ordem = 0
  Aux_Seq = 0
  
  Linha = 0
  
  '-------------------------------------------------------------------------
  ' C�DIGO CANCELADO !!!
  '-------------------------------------------------------------------------
  '14/01/2005 - Daniel
  '
  'Projeto.......: Tratamento da Quantidade Vendida Acumulada
  'Finalidade....: Corre��o do totalizador do valor da compra
  'Solicitante...: Aura Prata
  'txtQtdeVendidaAcumulada.Text = 0
  '-------------------------------------------------------------------------
  
  Erase Empr�stimos
  
  '02/10/2003 - mpdea
  'Zerado as linhas e modificado a atualiza��o do grid
  With Grade1
    .Rows = 0
    .MoveLast
    .MoveFirst
    .Refresh
    .Redraw = False
  End With
    
  rsEmprestimos.Index = "Cliente"
  rsProdutos.Index = "C�digo"
Lp1:
  rsEmprestimos.Seek ">", gnCodFilial, Aux_Seq, Aux_Cliente, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edi��o, Aux_ordem
  If rsEmprestimos.NoMatch Then GoTo Fim_Lp1
  
  Aux_Seq = rsEmprestimos("Sequ�ncia")
  Aux_Produto = rsEmprestimos("Produto")
  Aux_Tamanho = rsEmprestimos("Tamanho")
  Aux_Cor = rsEmprestimos("Cor")
  Aux_Edi��o = rsEmprestimos("Edi��o")
  Aux_ordem = rsEmprestimos("Ordem")
  Aux_Cliente = rsEmprestimos("Cliente")
  
  If rsEmprestimos("Filial") <> gnCodFilial Then GoTo Fim_Lp1
  If rsEmprestimos("Cliente") <> Val(Combo_Cliente.Text) Then GoTo Lp1
    
  If O_Aberto.Value = True And rsEmprestimos("Conclu�do") = True Then GoTo Lp1
  If O_Aberto.Value = False And rsEmprestimos("Conclu�do") = False Then GoTo Lp1
  
  
  If O_Empr�stimo.Value = True Then
     If CDate(Data_Emp.Text) <> CDate(rsEmprestimos("Data Opera��o")) Then GoTo Lp1
  End If
  
  If O_Acerto.Value = True Then
    If CDate(Data_Ace.Text) <> CDate(rsEmprestimos("Data Cobran�a")) Then GoTo Lp1
  End If
  
  rsProdutos.Seek "=", Aux_Produto
  
  Dim sZerosTam As String
  Dim sZerosCor As String
  
  '-----------------------------------------------------------------------------
  '28/08/2003 - mpdea
  'Otimizado c�digo e adicionado ordena��o
  With Empr�stimos(Linha)
    .Sequ�ncia = Aux_Seq
    .Produto = Aux_Produto
  
    If rsProdutos.NoMatch Then
      .Nome = "Produto n�o encontrado"
      .Ordenacao = ""
    Else
      If optOrdemSequencia.Value = True Then
          If Aux_Tamanho <> 0 Then
              If Len(Trim(Aux_Tamanho)) = 1 Then
                  sZerosTam = "00" & Aux_Tamanho
              ElseIf Len(Trim(Aux_Tamanho)) = 2 Then
                  sZerosTam = "0" & Aux_Tamanho
              End If
              
              If Len(Trim(Aux_Cor)) = 1 Then
                  sZerosCor = "00" & Aux_Cor
              ElseIf Len(Trim(Aux_Cor)) = 2 Then
                  sZerosCor = "0" & Aux_Cor
              End If
              
              .Nome = Aux_Produto & sZerosTam & sZerosCor & " - " & rsProdutos.Fields("Nome").Value
          Else
              .Nome = rsProdutos.Fields("Nome").Value
          End If
      Else
          .Nome = rsProdutos.Fields("Nome").Value & ""
      End If
      .Ordenacao = rsProdutos.Fields("C�digo Ordena��o").Value & ""
    End If
    
    .Tamanho = Aux_Tamanho
    .Cor = Aux_Cor
    .Edi��o = Aux_Edi��o
    .Ordem = Aux_ordem
    .Data = rsEmprestimos.Fields("Data Opera��o").Value
    .Saldo_Ant = rsEmprestimos.Fields("Saldo Atual").Value
    '15/02/2017 Jean, Altera��o para trazer valor formatado no Grid
    .Valor_Unit = Format(rsEmprestimos.Fields("Pre�o Unit�rio").Value, "###,###,###,##0.00")
  
    .Nova_Devol = 0
    .Nova_Venda = 0
    .Novo_Emp = 0
    .Novo_Saldo = rsEmprestimos.Fields("Saldo Atual").Value
  End With
  '-----------------------------------------------------------------------------
    
  Linha = Linha + 1
  
  GoTo Lp1
  
Fim_Lp1:
  
  'Ordena a tabela por c�digo
  If optOrdemProduto.Value = True Then
      Call OrderByCode
  ElseIf optOrdemItensUnicaSequencia.Value = True Then
      Call Ordenar_PorSequenciaEOrdemItemProduto(CLng(txtSequencia.Text))
      Linha = m_numRegistrosEmprestimoDaUnicaSequencia
  End If

  'Modificado a atualiza��o do grid
  With Grade1
    '28/10/2003 - Maikel
    'Tocado todo o if acima pela linha abaixo por que o sistema sempre cortava o ultimo produto.
    .Rows = Linha
    .MoveLast
    .MoveFirst
    .Refresh
    .Redraw = True
  End With
  
    
  '-----------------------------------------------------------------------------
  '28/08/2003 - mpdea
  'Otimizado c�digo
  Grade1.Columns(3).Visible = gbGrade
  Grade1.Columns(4).Visible = gbGrade
  Grade1.Columns(5).Visible = gbEdicao
  
  Frame_Mov.Enabled = O_Aberto.Value
  '-----------------------------------------------------------------------------
    
  
  '02/10/2003 - mpdea
  'Status
  Screen.MousePointer = vbDefault
  Call StatusMsg("Pronto")
  
End Sub

Private Sub B_Mostra_Click()

  Call StatusMsg("")
   
  If Grade1.SelBookmarks.Count < 1 Then
    DisplayMsg "Selecione a linha antes."
    Exit Sub
  End If
End Sub

Private Sub cmdExtrato_Click()
  '17/01/2005 - Daniel
  '
  'Solicitante: Aura Prata
  '
  'Relat�rio de extrato de produtos consolidados consignados
  Dim rstExtrato        As Recordset
  Dim rstExtratoGroup   As Recordset
  Dim rstConsigSaida    As Recordset
  Dim intAuxi           As Integer
  Dim strSQL            As String
  Dim dblSaldo          As Double
  Dim strReport         As String
  Dim strNomeCliente    As String
  Dim intCodigo         As Integer
  
  On Error GoTo ErrHandler
  
  '---[Valida��es]---
  If Len(Combo_Cliente.Text) <= 0 Then
    MsgBox "Escolha um cliente.", vbExclamation, "Aten��o"
    Combo_Cliente.SetFocus
    Exit Sub
  End If
  
  If Grade1.Rows <= 0 Then Exit Sub
  '---[Fim Valida��es]---
  
  '---[Limpando a tabela tempor�ria]---
  dbTemp.Execute "DELETE * FROM Extrato"
  '---[Fim Limpando a tabela tempor�ria]---
  
  '---[Criando os registros na tabela tempor�ria. Esta tabela ser� chamada na exibi��o do extrato]---
  Call StatusMsg("Criando tabela tempor�ria...")
    
    Set rstExtrato = dbTemp.OpenRecordset("Extrato", dbOpenDynaset)
  
    Grade1.MoveFirst
    
    For intAuxi = 0 To (Grade1.Rows - 1)
    
      With rstExtrato
        .AddNew
        .Fields("Sequencia").Value = CLng(Grade1.Columns("Sequ�ncia").Text)
        .Fields("Produto").Value = Grade1.Columns("Produto").Text & ""
        .Fields("NomeProduto").Value = Left(Grade1.Columns("Nome").Text & "", 50)
        .Fields("Tam").Value = Grade1.Columns("Tamanho").Text & ""
        .Fields("Cor").Value = Grade1.Columns("Cor").Text & ""
        .Fields("Data").Value = CDate(Grade1.Columns("Data Opera��o").Text)
        .Fields("ValorUnitario").Value = CDbl(Grade1.Columns("Pre�o Unit�rio").Text)
        .Update
      End With
    
      Grade1.MoveNext
    Next intAuxi
  
  '---[Fim Criando os registros na tabela tempor�ria]---
 
  '---[Atualiza��o do campo QtdeEmprestada]---
  With rstExtrato
    .MoveFirst
    
    Do Until .EOF
    
      strSQL = ""
      strSQL = "SELECT [Data Opera��o], Ordem, [Saldo Anterior], [Vendas Cliente], [Devolu��o], [Novo Empr�stimo] , [Saldo Atual] AS SaldoAtual "
      strSQL = strSQL & " FROM [Consigna��o Sa�da] "
      strSQL = strSQL & " WHERE [Consigna��o Sa�da].Filial = " & gnCodFilial
      strSQL = strSQL & " AND [Consigna��o Sa�da].Sequ�ncia = " & .Fields("Sequencia").Value
      strSQL = strSQL & " AND [Consigna��o Sa�da].Produto = '" & .Fields("Produto").Value & "'"
      If Len(.Fields("Tam").Value) > 0 Then strSQL = strSQL & " AND [Consigna��o Sa�da].Tamanho = " & CInt(.Fields("Tam").Value)
      If Len(.Fields("Cor").Value) > 0 Then strSQL = strSQL & " AND [Consigna��o Sa�da].Cor = " & CInt(.Fields("Cor").Value)
      strSQL = strSQL & " ORDER BY +Ordem "
    
      Set rstConsigSaida = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
      If Not (rstConsigSaida.BOF And rstConsigSaida.EOF) Then
        rstConsigSaida.MoveFirst '18/03/2005 - Precisamos pegar o saldo !!!
        rstConsigSaida.MoveLast  'Por esta raz�o damos o MoveLast
        dblSaldo = rstConsigSaida.Fields("SaldoAtual").Value
      End If
      
      rstConsigSaida.Close
      Set rstConsigSaida = Nothing
    
      'Atualizamos Extrato.Saldo
      .Edit
      .Fields("Saldo").Value = Format(dblSaldo, FORMAT_VALUE)
      .Update
      
     .MoveNext
    Loop
   
  End With
  '---[Fim Atualiza��o do campo Saldo]---
 
  rstExtrato.Close
  Set rstExtrato = Nothing
 
  Grade1.MoveFirst
  
  '---[Rel. Sint�tico: Tratamento para agruparmos as informa��es]---
  If optSintetico.Value Then
    dbTemp.Execute "DELETE * FROM ExtratoGroup"
    dbTemp.Execute "DELETE * FROM ExtratoSeq"
    
    intCodigo = 0
    
    Set rstExtratoGroup = dbTemp.OpenRecordset("ExtratoGroup", dbOpenDynaset)
    
    strSQL = ""
    strSQL = "SELECT Produto, NomeProduto, Tam, Cor "
    strSQL = strSQL & " FROM Extrato GROUP BY Produto, NomeProduto, Tam, Cor "
    
    Set rstExtrato = dbTemp.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rstExtrato
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do Until .EOF
            intCodigo = intCodigo + 1
        
            rstExtratoGroup.AddNew
            rstExtratoGroup.Fields("CodigoMovi").Value = intCodigo
            rstExtratoGroup.Fields("Produto").Value = .Fields("Produto").Value & ""
            rstExtratoGroup.Fields("NomeProduto").Value = .Fields("NomeProduto").Value & ""
            rstExtratoGroup.Fields("Tam").Value = .Fields("Tam").Value & ""
            rstExtratoGroup.Fields("Cor").Value = .Fields("Cor").Value & ""
            rstExtratoGroup.Update
          
         .MoveNext
        Loop
        
        rstExtratoGroup.MoveFirst
        Do Until rstExtratoGroup.EOF
          'Antes de criar rstExtratoSeq precisamos checar
          'em quantas sequ�ncias o produto est� envolvido
          Call VerificarQtasSequencias(rstExtratoGroup.Fields("CodigoMovi").Value, rstExtratoGroup.Fields("Produto").Value, rstExtratoGroup.Fields("Tam").Value, rstExtratoGroup.Fields("Cor").Value)
          'Criamos rstExtratoSeq dentro de VerificarQtasSequencias
        
         rstExtratoGroup.MoveNext
        Loop
        
      End If
      .Close
    End With
  
    rstExtratoGroup.Close
    Set rstExtratoGroup = Nothing
  
    Set rstExtrato = Nothing
  End If
  
  '---[Fim Rel. Sint�tico: Tratamento para agruparmos as informa��es]---
  
  Call StatusMsg("")
  
  '---[Montamos o relat�rio]---
  Screen.MousePointer = vbHourglass
  
  'Nome do arquivo .rpt
  If optSintetico.Value Then
    strReport = gsReportPath & "rptExtratoSintetico.rpt"
  Else
    strReport = gsReportPath & "rptExtrato.rpt"
  End If
  
  Call BuscarNomeCliente(strNomeCliente)
  
  With crptExtrato
    .Reset
    .ReportFileName = strReport
    
    If optAnal�tico.Value Then
      .DataFiles(0) = gsTempDBFileName
    Else
      .DataFiles(0) = gsTempDBFileName
      .DataFiles(1) = gsTempDBFileName
      .DataFiles(2) = gsTempDBFileName
    End If
    
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a f�rmula no crystal tamb�m
    .Formulas(1) = "cliente = '" & ("CLIENTE " & (Combo_Cliente.Text) & " - " & strNomeCliente) & "'"
    If optAnal�tico.Value Then
      .SortFields(0) = "+{Extrato.Sequencia}" 'Ordena��o
    Else
      .SortFields(0) = "+{ExtratoGroup.CodigoMovi}"
      .SortFields(1) = "+{ExtratoSeq.Sequencia}"
    End If
    
    .WindowState = crptMaximized
    .Destination = crptToWindow
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relat�rio
    Call SetPrinterName("REL", crptExtrato)
  
    .Action = 1
  End With
 
  '---[Fim Montamos o relat�rio]---
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
  Exit Sub
 
ErrHandler:
  MsgBox "Erro no processo: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub VerificarQtasSequencias(ByVal CodigoMovi As Integer, ByVal Produto As String, ByVal Tam As String, ByVal Cor As String)
  '18/03/2005 - Daniel
  '
  'Projeto de emiss�o de extrato
  '
  'Solicitante: Aura Prata
  Dim rstExtrato              As Recordset
  Dim rstExtratoSeq           As Recordset
  Dim lngArraySeqs(1 To 3000) As Long
  Dim dblSaldo(1 To 3000)     As Double
  Dim dblVlUnit(1 To 3000)    As Double
  Dim intContador             As Integer
  Dim intI                    As Integer
  Dim strSQL                  As String
  
  On Error GoTo ErrHandler
  
  intContador = 0
  
  strSQL = "SELECT * FROM Extrato "
  strSQL = strSQL & " WHERE Produto = '" & Produto & "'"
  strSQL = strSQL & " AND Tam = '" & Tam & "'"
  strSQL = strSQL & " AND Cor = '" & Cor & "'"

  Set rstExtrato = dbTemp.OpenRecordset(strSQL, dbOpenSnapshot)

  With rstExtrato
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
         intContador = intContador + 1
         
         lngArraySeqs(intContador) = .Fields("Sequencia").Value
         dblSaldo(intContador) = .Fields("Saldo").Value
         dblVlUnit(intContador) = .Fields("ValorUnitario").Value
         
       .MoveNext
      Loop
      
    End If
    .Close
  End With

  Set rstExtrato = Nothing
  
  'Agora abrimos o rstExtratoSeq
  Set rstExtratoSeq = dbTemp.OpenRecordset("ExtratoSeq", dbOpenDynaset)
  
  For intI = 1 To intContador
    
    With rstExtratoSeq
      .AddNew
      .Fields("CodigoMovi").Value = CodigoMovi
      .Fields("Sequencia").Value = lngArraySeqs(intI)
      .Fields("Saldo").Value = dblSaldo(intI)
      .Fields("ValorUnitario").Value = dblVlUnit(intI)
      .Update
    End With
  
  Next intI
  
  rstExtratoSeq.Close
  Set rstExtratoSeq = Nothing

  Exit Sub

ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Sub Combo_Cliente_CloseUp()
  Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
  Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()

  If IsNull(Combo_Cliente.Text) Then
      lbl_NomeCliente.Caption = ""
      Exit Sub
  End If
  
  If Combo_Cliente.Text = "" Then
      lbl_NomeCliente.Caption = ""
      Exit Sub
  End If
  
  If Not IsNumeric(Combo_Cliente.Text) Then
      lbl_NomeCliente.Caption = ""
      Exit Sub
  End If
  
  If Val(Combo_Cliente.Text) < 1 Then
      lbl_NomeCliente.Caption = ""
      Exit Sub
  End If
  
  rsClientes.Index = "C�digo"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then
      lbl_NomeCliente.Caption = ""
      Exit Sub
  End If
  
  Call StatusMsg(rsClientes("Nome") & "")
  
  If Not IsNull(rsClientes("Nome")) Then
      lbl_NomeCliente.Caption = rsClientes("Nome")
  Else
      lbl_NomeCliente.Caption = ""
  End If
  
  Estado = ""
  rsEstados.Index = "Estado"
  If IsNull(rsClientes("Estado")) Then Exit Sub
  If rsClientes("Estado") <> "" Then
    rsEstados.Seek "=", rsClientes("Estado")
    If Not rsEstados.NoMatch Then
      Estado = rsEstados("Estado")
    End If
  End If
  
End Sub

Private Sub Command1_Click()
  With frmVerificaDatas
    .Tipo = "SA�DA"
    .Show
    .WindowState = vbNormal
  End With
End Sub

Private Sub Data_Ace_LostFocus()
  Data_Ace.Text = Ajusta_Data(Data_Ace.Text)
End Sub

Private Sub Data_Ace_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ace.Text = frmCalendario.gsDateCalender(Data_Ace.Text)
  End Select
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

Private Sub Data_Emp_LostFocus()
  Data_Emp.Text = Ajusta_Data(Data_Emp.Text)
End Sub

Private Sub Data_Emp_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Emp.Text = frmCalendario.gsDateCalender(Data_Emp.Text)
  End Select
End Sub

Private Sub Dev_Prod_Change()
  Recalcula_Saldo
End Sub

Private Sub Dev_Prod_GotFocus()
  Dev_Prod.SelStart = 0
  Dev_Prod.SelLength = Len(Dev_Prod.Text)
End Sub

Private Sub Dev_Prod_LostFocus()
If Not IsNumeric(Dev_Prod.Text) Then
      DisplayMsg "Quantidade incorreta."
      Dev_Prod.SetFocus
End If
End Sub

Private Sub Emp_Prod_Change()
  Recalcula_Saldo
End Sub

Private Sub Emp_Prod_GotFocus()
  Emp_Prod.SelStart = 0
  Emp_Prod.SelLength = Len(Emp_Prod.Text)
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)

  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsEmprestimos = db.OpenRecordset("Consigna��o Sa�da")
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque")
  Set rsEstoque_Final = db.OpenRecordset("Estoque Final")
  Set rsResumo_Di�rio = db.OpenRecordset("Resumo Di�rio")
  Set rsSaidas = db.OpenRecordset("Sa�das")
  Set rsSaidas_Prod = db.OpenRecordset("Sa�das - Produtos")
  Set rsParametros = db.OpenRecordset("Par�metros Filial")
  Set rsEstados = db.OpenRecordset("Estados", , dbReadOnly)
  
  Grade1.Columns(13).NumberFormat = Formato_Pre�o
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsEstados.Close
  Set rsEstados = Nothing
End Sub

Private Sub Grade1_Click()
  Grade1.SelBookmarks.RemoveAll
  Grade1.SelBookmarks.Add Grade1.Bookmark
End Sub

Private Sub Grade1_LostFocus()
 If Grade1.RowChanged = True Then
   Grade1.Update
 End If
End Sub

Private Sub Grade1_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
 Dim Aux_Dbl As Double

  On Error GoTo Erro
  
  Aux_Dbl = Val(Grade1.Columns(11).Text) * CDbl(Grade1.Columns(8).Text)
  
  Valor_Linha.Caption = Format(Aux_Dbl, "###,###,###,##0.00")
  
  Vendas_Prod.Text = Grade1.Columns(11).Text
  Dev_Prod.Text = Grade1.Columns(12).Text
  Emp_Prod.Text = Grade1.Columns(13).Text
  Saldo_Prod.Caption = Grade1.Columns(10).Text
  Valor_Prod.Text = Grade1.Columns(8).Text
  sSequenciaEmprestimo = Grade1.Columns(0).Text
  
  If O_Mostra_Detalhe.Value = 1 Then
    Atualiza_Detalhes
  End If
  
  '-------------------------------------------------------------------------
  ' C�DIGO CANCELADO !!!
  '-------------------------------------------------------------------------
  '14/01/2005 - Daniel
  '
  'Projeto.......: Tratamento da Quantidade Vendida Acumulada
  'Finalidade....: Corre��o do totalizador do valor da compra
  'Solicitante...: Aura Prata
  'Call BuscarQtdeVendidaAcumulada
  '
  'Valor_Linha.Caption = Format(CDbl(txtQtdeVendidaAcumulada.Text) * CDbl(Grade1.Columns(8).Text), FORMAT_VALUE)
  
  Exit Sub
  
Erro:
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub Grade1_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p  As Long
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade1.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p

End Sub


Private Sub Grade1_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Dim r As Long, i As Long, p As Long
  
  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      p = Grade1.Rows
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
    If p < 0 Or p >= Grade1.Rows Then Exit For
    RowBuf.Value(i, 0) = Empr�stimos(p).Sequ�ncia
    RowBuf.Value(i, 1) = Empr�stimos(p).Produto
    RowBuf.Value(i, 2) = Empr�stimos(p).Nome
    RowBuf.Value(i, 3) = Empr�stimos(p).Tamanho
    RowBuf.Value(i, 4) = Empr�stimos(p).Cor
    RowBuf.Value(i, 5) = Empr�stimos(p).Edi��o
    RowBuf.Value(i, 6) = Empr�stimos(p).Ordem
    RowBuf.Value(i, 7) = Empr�stimos(p).Data
    RowBuf.Value(i, 8) = Empr�stimos(p).Valor_Unit
    RowBuf.Value(i, 9) = Empr�stimos(p).Saldo_Ant

    'campos invis�veis
    RowBuf.Value(i, 10) = Empr�stimos(p).Novo_Saldo
    RowBuf.Value(i, 11) = Empr�stimos(p).Nova_Venda
    RowBuf.Value(i, 12) = Empr�stimos(p).Nova_Devol
    RowBuf.Value(i, 13) = Empr�stimos(p).Novo_Emp
             
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

Private Sub Grade1_UnboundWriteData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, WriteLocation As Variant)
On Error GoTo Erro

  Dim Linha1 As Integer
  
  If IsNull(WriteLocation) Then Exit Sub
  
  Linha1 = WriteLocation
  If Linha1 = -1 Then Linha1 = 0
  
  With Empr�stimos(Linha1)
    .Sequ�ncia = Grade1.Columns(0).Text
    .Produto = Grade1.Columns(1).Text
    .Nome = Grade1.Columns(2).Text
    .Tamanho = Grade1.Columns(3).Text
    .Cor = Grade1.Columns(4).Text
    .Edi��o = Grade1.Columns(5).Text
    .Ordem = Grade1.Columns(6).Text
    .Data = Grade1.Columns(7).Text
    .Valor_Unit = Grade1.Columns(8).Text
    .Saldo_Ant = Grade1.Columns(9).Text
    'campos invis�veis
    .Novo_Saldo = Grade1.Columns(10).Text
    .Nova_Venda = Grade1.Columns(11).Text
    .Nova_Devol = Grade1.Columns(12).Text
    .Novo_Emp = Grade1.Columns(13).Text
  End With
  
  Exit Sub
Erro:
  MsgBox "Algo deu errado. Repita o procedimento.", vbInformation, "Aten��o"
  
End Sub

Private Sub O_Mostra_Detalhe_Click()
  If O_Mostra_Detalhe.Value = 1 Then
    Grade2.Visible = True
  Else
    Grade2.Visible = False
  End If
End Sub

Private Sub optOrdemItensUnicaSequencia_Click()
  If optOrdemItensUnicaSequencia.Value = True Then
    txtSequencia.Visible = True
  Else
    txtSequencia.Visible = False
  End If
End Sub

Private Sub optOrdemProduto_Click()
  If optOrdemProduto.Value = True Then
      txtSequencia.Visible = False
  End If
End Sub

Private Sub optOrdemSequencia_Click()
  If optOrdemSequencia.Value = True Then
      txtSequencia.Visible = False
  End If
End Sub

Private Sub Valor_Prod_GotFocus()
  Valor_Prod.SelStart = 0
  Valor_Prod.SelLength = Len(Valor_Prod.Text)
End Sub

Private Sub Valor_Prod_LostFocus()
If Not IsNumeric(Valor_Prod.Text) Then
      DisplayMsg "Valor incorreto."
      Valor_Prod.SetFocus
End If
End Sub

Private Sub Vendas_Prod_Change()
 Recalcula_Saldo
End Sub

Private Sub Vendas_Prod_GotFocus()
  Vendas_Prod.SelStart = 0
  Vendas_Prod.SelLength = Len(Vendas_Prod.Text)
End Sub

'08/10/2003 - mpdea
'Corrigido esquema da ordena��o (compara��o com vazios)

'28/08/2003 - mpdea
'Ordena��o da lista por c�digo
Private Sub OrderByCode()
  Dim TEMP_Emprestimos() As Tab_Emp
  Dim intX1 As Integer
  Dim intX2 As Integer
  Dim strCodigoOrdenacao As String
  Dim intMinPos As Integer
  
  '01/10/2003 - mpdea
  'Igualado o redimensionado
  ReDim TEMP_Emprestimos(UBound(Empr�stimos)) As Tab_Emp
  
  Dim sZerosTam As String
  Dim sZerosCor As String
  
  For intX1 = LBound(Empr�stimos) To UBound(Empr�stimos)
    '�ndice do primeiro item preenchido
    intMinPos = m_intFirstRegister()
    
    'N�o h� mais itens
    If intMinPos = -1 Then Exit For
    
    'C�digo m�nimo a ser testado para ordena��o
    strCodigoOrdenacao = Empr�stimos(intMinPos).Ordenacao
    
    For intX2 = LBound(Empr�stimos) To UBound(Empr�stimos)
      'Compara ordem
      If Empr�stimos(intX2).Ordenacao <> "" And Empr�stimos(intX2).Ordenacao < strCodigoOrdenacao Then
        intMinPos = intX2
        strCodigoOrdenacao = Empr�stimos(intMinPos).Ordenacao
      End If
    Next intX2
    
    'Copia registro
    With TEMP_Emprestimos(intX1)
      .Sequ�ncia = Empr�stimos(intMinPos).Sequ�ncia
      .Produto = Empr�stimos(intMinPos).Produto
      .Nome = Empr�stimos(intMinPos).Nome
      .Tamanho = Empr�stimos(intMinPos).Tamanho
      .Cor = Empr�stimos(intMinPos).Cor
      .Edi��o = Empr�stimos(intMinPos).Edi��o
      .Ordem = Empr�stimos(intMinPos).Ordem
      .Data = Empr�stimos(intMinPos).Data
      .Saldo_Ant = Empr�stimos(intMinPos).Saldo_Ant
      .Valor_Unit = Empr�stimos(intMinPos).Valor_Unit
      .Novo_Emp = Empr�stimos(intMinPos).Novo_Emp
      .Nova_Venda = Empr�stimos(intMinPos).Nova_Venda
      .Nova_Devol = Empr�stimos(intMinPos).Nova_Devol
      .Novo_Saldo = Empr�stimos(intMinPos).Novo_Saldo
      .Ordenacao = Empr�stimos(intMinPos).Ordenacao
    End With
    
    'Zera ordena��o do registro copiado
    Empr�stimos(intMinPos).Ordenacao = ""
    
  Next intX1
  
  'Realiza a c�pia dos dados em ordem de c�digo
  Erase Empr�stimos
  intX2 = 0
  
  For intX1 = LBound(TEMP_Emprestimos) To UBound(TEMP_Emprestimos)
    If TEMP_Emprestimos(intX1).Ordenacao <> "" Then
      With Empr�stimos(intX2)
        .Sequ�ncia = TEMP_Emprestimos(intX1).Sequ�ncia
        .Produto = TEMP_Emprestimos(intX1).Produto
        
        If TEMP_Emprestimos(intX1).Tamanho <> 0 Then
        
            If Len(Trim(TEMP_Emprestimos(intX1).Tamanho)) = 1 Then
                sZerosTam = "00" & TEMP_Emprestimos(intX1).Tamanho
            ElseIf Len(Trim(TEMP_Emprestimos(intX1).Tamanho)) = 2 Then
                sZerosTam = "0" & TEMP_Emprestimos(intX1).Tamanho
            End If
            
            If Len(Trim(TEMP_Emprestimos(intX1).Cor)) = 1 Then
                sZerosCor = "00" & TEMP_Emprestimos(intX1).Cor
            ElseIf Len(Trim(TEMP_Emprestimos(intX1).Tamanho)) = 2 Then
                sZerosCor = "0" & TEMP_Emprestimos(intX1).Cor
            End If
            
            .Nome = TEMP_Emprestimos(intX1).Produto & sZerosTam & sZerosCor & " - " & TEMP_Emprestimos(intX1).Nome
        Else
            .Nome = TEMP_Emprestimos(intX1).Nome
        End If
        
        .Tamanho = TEMP_Emprestimos(intX1).Tamanho
        .Cor = TEMP_Emprestimos(intX1).Cor
        .Edi��o = TEMP_Emprestimos(intX1).Edi��o
        .Ordem = TEMP_Emprestimos(intX1).Ordem
        .Data = TEMP_Emprestimos(intX1).Data
        .Saldo_Ant = TEMP_Emprestimos(intX1).Saldo_Ant
        .Valor_Unit = TEMP_Emprestimos(intX1).Valor_Unit
        .Novo_Emp = TEMP_Emprestimos(intX1).Novo_Emp
        .Nova_Venda = TEMP_Emprestimos(intX1).Nova_Venda
        .Nova_Devol = TEMP_Emprestimos(intX1).Nova_Devol
        .Novo_Saldo = TEMP_Emprestimos(intX1).Novo_Saldo
        .Ordenacao = TEMP_Emprestimos(intX1).Ordenacao
      End With
      intX2 = intX2 + 1
    End If
  Next intX1
  
End Sub

'Obt�m o primeiro registro preenchido da tabela
Private Function m_intPrimeiroRegistroItemOrdem() As Integer
  Dim intX As Integer
  
  For intX = LBound(Empr�stimos) To UBound(Empr�stimos)
    If Empr�stimos(intX).Ordem <> 0 Then
      'Primeiro registro ocupado
      m_intPrimeiroRegistroItemOrdem = intX
      Exit Function
    End If
  Next intX
  
  'N�o h� registro
  m_intPrimeiroRegistroItemOrdem = -1
End Function

'Obt�m o primeiro registro preenchido da tabela
Private Function ContaNumRegistroEmprestimo()
  Dim intX As Integer
  
  m_numRegistrosEmprestimo = 0
  
  For intX = LBound(Empr�stimos) To UBound(Empr�stimos)
    If Empr�stimos(intX).Sequ�ncia <> 0 Then
        m_numRegistrosEmprestimo = m_numRegistrosEmprestimo + 1
    End If
  Next intX
  
End Function

Private Function ContaNumRegistroEmprestimoDaUnicaSequencia(pSequencia As Long)
  Dim intX As Integer
  
  m_numRegistrosEmprestimoDaUnicaSequencia = 0
  
  For intX = LBound(Empr�stimos) To UBound(Empr�stimos)
    If Empr�stimos(intX).Sequ�ncia = pSequencia Then
        m_numRegistrosEmprestimoDaUnicaSequencia = m_numRegistrosEmprestimoDaUnicaSequencia + 1
    End If
  Next intX
  
End Function

Private Sub Ordenar_PorSequenciaEOrdemItemProduto(pSequencia As Long)
  Dim TEMP_Emprestimos() As Tab_Emp
  Dim intX1 As Integer
  Dim intX2 As Integer
  Dim intOrdem As Integer
  Dim intMinPos As Integer

  'Igualado o redimensionado
  ReDim TEMP_Emprestimos(UBound(Empr�stimos)) As Tab_Emp
  
  'Obter total de registros a serem visualizados
  ContaNumRegistroEmprestimo
  ContaNumRegistroEmprestimoDaUnicaSequencia (pSequencia)
  
  For intX1 = 0 To m_numRegistrosEmprestimo
    
    '�ndice do primeiro item preenchido
    intMinPos = m_intPrimeiroRegistroItemOrdem()
    
    'N�o h� mais itens
    If intMinPos = -1 Then Exit For
    
    'C�digo m�nimo a ser testado para ordena��o
    intOrdem = Empr�stimos(intMinPos).Ordem

    For intX2 = LBound(Empr�stimos) To UBound(Empr�stimos)
      
      'Compara ordem
      If Empr�stimos(intX2).Ordem <> 0 And Empr�stimos(intX2).Ordem < intOrdem Then
        intMinPos = intX2
        intOrdem = Empr�stimos(intMinPos).Ordem
      End If
      
    Next intX2
    
    If pSequencia = Empr�stimos(intMinPos).Sequ�ncia Then
        'Copia registro
        With TEMP_Emprestimos(intX1)
          .Sequ�ncia = Empr�stimos(intMinPos).Sequ�ncia
          .Produto = Empr�stimos(intMinPos).Produto
          .Nome = Empr�stimos(intMinPos).Nome
          .Tamanho = Empr�stimos(intMinPos).Tamanho
          .Cor = Empr�stimos(intMinPos).Cor
          .Edi��o = Empr�stimos(intMinPos).Edi��o
          .Ordem = Empr�stimos(intMinPos).Ordem
          .Data = Empr�stimos(intMinPos).Data
          .Saldo_Ant = Empr�stimos(intMinPos).Saldo_Ant
          .Valor_Unit = Empr�stimos(intMinPos).Valor_Unit
          .Novo_Emp = Empr�stimos(intMinPos).Novo_Emp
          .Nova_Venda = Empr�stimos(intMinPos).Nova_Venda
          .Nova_Devol = Empr�stimos(intMinPos).Nova_Devol
          .Novo_Saldo = Empr�stimos(intMinPos).Novo_Saldo
          .Ordenacao = Empr�stimos(intMinPos).Ordenacao
        End With
    End If
    
    'Zera ordena��o do registro copiado
    Empr�stimos(intMinPos).Ordem = 0
    
  Next intX1
  
  
  'Realiza a c�pia dos dados em ordem de c�digo
  Erase Empr�stimos
  intX2 = 0
  
  Dim sZerosTam As String
  Dim sZerosCor As String
  
  For intX1 = LBound(TEMP_Emprestimos) To UBound(TEMP_Emprestimos)
    If TEMP_Emprestimos(intX1).Ordem <> 0 Then
      With Empr�stimos(intX2)
        .Sequ�ncia = TEMP_Emprestimos(intX1).Sequ�ncia
        .Produto = TEMP_Emprestimos(intX1).Produto
        
        If TEMP_Emprestimos(intX1).Tamanho <> 0 Then
        
            If Len(Trim(TEMP_Emprestimos(intX1).Tamanho)) = 1 Then
                sZerosTam = "00" & TEMP_Emprestimos(intX1).Tamanho
            ElseIf Len(Trim(TEMP_Emprestimos(intX1).Tamanho)) = 2 Then
                sZerosTam = "0" & TEMP_Emprestimos(intX1).Tamanho
            End If
            
            If Len(Trim(TEMP_Emprestimos(intX1).Cor)) = 1 Then
                sZerosCor = "00" & TEMP_Emprestimos(intX1).Cor
            ElseIf Len(Trim(TEMP_Emprestimos(intX1).Tamanho)) = 2 Then
                sZerosCor = "0" & TEMP_Emprestimos(intX1).Cor
            End If
            
            .Nome = TEMP_Emprestimos(intX1).Produto & sZerosTam & sZerosCor & " - " & TEMP_Emprestimos(intX1).Nome
        Else
            .Nome = TEMP_Emprestimos(intX1).Nome
        End If
        
        .Tamanho = TEMP_Emprestimos(intX1).Tamanho
        .Cor = TEMP_Emprestimos(intX1).Cor
        .Edi��o = TEMP_Emprestimos(intX1).Edi��o
        .Ordem = TEMP_Emprestimos(intX1).Ordem
        .Data = TEMP_Emprestimos(intX1).Data
        .Saldo_Ant = TEMP_Emprestimos(intX1).Saldo_Ant
        .Valor_Unit = TEMP_Emprestimos(intX1).Valor_Unit
        .Novo_Emp = TEMP_Emprestimos(intX1).Novo_Emp
        .Nova_Venda = TEMP_Emprestimos(intX1).Nova_Venda
        .Nova_Devol = TEMP_Emprestimos(intX1).Nova_Devol
        .Novo_Saldo = TEMP_Emprestimos(intX1).Novo_Saldo
        .Ordenacao = TEMP_Emprestimos(intX1).Ordenacao
      End With
      intX2 = intX2 + 1
    End If
  Next intX1
  
End Sub

'28/08/2003 - mpdea
'Obt�m o primeiro registro preenchido da tabela
Private Function m_intFirstRegister() As Integer
  Dim intX As Integer
  
  For intX = LBound(Empr�stimos) To UBound(Empr�stimos)
    If Empr�stimos(intX).Ordenacao <> "" Then
      'Primeiro registro ocupado
      m_intFirstRegister = intX
      Exit Function
    End If
  Next intX
  
  'N�o h� registro
  m_intFirstRegister = -1
End Function

Private Sub BuscarQtdeVendidaAcumulada()
  '-------------------------------------------------------------------------
  ' C�DIGO CANCELADO !!!
  '-------------------------------------------------------------------------
  '14/01/2005 - Daniel
  '
  'Projeto.......: Tratamento da Quantidade Vendida Acumulada
  'Finalidade....: Corre��o do totalizador do valor da compra
  'Solicitante...: Aura Prata
  Dim rstConsigSaidas As Recordset
  Dim strSQL          As String
  
  strSQL = "SELECT QtdeVendidaAcumulada "
  strSQL = strSQL & " FROM [Consigna��o Sa�da] "
  strSQL = strSQL & " WHERE [Consigna��o Sa�da].Filial = " & CByte(gnCodFilial)
  strSQL = strSQL & " AND [Consigna��o Sa�da].Sequ�ncia = " & CLng(Grade1.Columns(0).Text)
  strSQL = strSQL & " AND [Consigna��o Sa�da].Produto = '" & Grade1.Columns(1).Text & "'"
  strSQL = strSQL & " AND [Consigna��o Sa�da].Tamanho = " & Grade1.Columns(3).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Cor = " & Grade1.Columns(4).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Edi��o = " & Grade1.Columns(5).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Ordem = " & CLng(Grade1.Columns(6).Text)
  
  Set rstConsigSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstConsigSaidas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If IsNumeric(.Fields("QtdeVendidaAcumulada").Value) Then
        txtQtdeVendidaAcumulada.Text = .Fields("QtdeVendidaAcumulada").Value
      Else
        txtQtdeVendidaAcumulada.Text = 0
      End If
      
    End If
    .Close
  End With
  
  Set rstConsigSaidas = Nothing
  
End Sub

Private Sub AtualizarQtdeVendidaAcumulada(ByVal Qtde As Double)
  '-------------------------------------------------------------------------
  ' C�DIGO CANCELADO !!!
  '-------------------------------------------------------------------------
  '14/01/2005 - Daniel
  '
  'Projeto.......: Tratamento da Quantidade Vendida Acumulada
  'Finalidade....: Corre��o do totalizador do valor da compra
  'Solicitante...: Aura Prata
  Dim rstConsigSaidas As Recordset
  Dim strSQL          As String
  
  strSQL = "SELECT QtdeVendidaAcumulada "
  strSQL = strSQL & " FROM [Consigna��o Sa�da] "
  strSQL = strSQL & " WHERE [Consigna��o Sa�da].Filial = " & CByte(gnCodFilial)
  strSQL = strSQL & " AND [Consigna��o Sa�da].Sequ�ncia = " & CLng(Grade1.Columns(0).Text)
  strSQL = strSQL & " AND [Consigna��o Sa�da].Produto = '" & Grade1.Columns(1).Text & "'"
  strSQL = strSQL & " AND [Consigna��o Sa�da].Tamanho = " & Grade1.Columns(3).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Cor = " & Grade1.Columns(4).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Edi��o = " & Grade1.Columns(5).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Ordem = " & CLng(Grade1.Columns(6).Text)
  
  Set rstConsigSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstConsigSaidas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Edit
      .Fields("QtdeVendidaAcumulada").Value = .Fields("QtdeVendidaAcumulada").Value + Qtde
      .Update
      
      txtQtdeVendidaAcumulada.Text = .Fields("QtdeVendidaAcumulada").Value
    End If
    .Close
  End With
  
  Set rstConsigSaidas = Nothing

End Sub

Private Sub ZerarQtdeVendidaAcumulada()
  '-------------------------------------------------------------------------
  ' C�DIGO CANCELADO !!!
  '-------------------------------------------------------------------------
  '14/01/2005 - Daniel
  '
  'Projeto.......: Tratamento da Quantidade Vendida Acumulada
  'Finalidade....: Corre��o do totalizador do valor da compra
  'Solicitante...: Aura Prata
  Dim rstConsigSaidas As Recordset
  Dim strSQL          As String
  
  strSQL = "SELECT QtdeVendidaAcumulada "
  strSQL = strSQL & " FROM [Consigna��o Sa�da] "
  strSQL = strSQL & " WHERE [Consigna��o Sa�da].Filial = " & CByte(gnCodFilial)
  strSQL = strSQL & " AND [Consigna��o Sa�da].Sequ�ncia = " & CLng(Grade1.Columns(0).Text)
  strSQL = strSQL & " AND [Consigna��o Sa�da].Produto = '" & Grade1.Columns(1).Text & "'"
  strSQL = strSQL & " AND [Consigna��o Sa�da].Tamanho = " & Grade1.Columns(3).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Cor = " & Grade1.Columns(4).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Edi��o = " & Grade1.Columns(5).Text
  strSQL = strSQL & " AND [Consigna��o Sa�da].Ordem = " & CLng(Grade1.Columns(6).Text)
  
  Set rstConsigSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstConsigSaidas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Edit
      .Fields("QtdeVendidaAcumulada").Value = 0
      .Update
      
      txtQtdeVendidaAcumulada.Text = .Fields("QtdeVendidaAcumulada").Value
      
      Valor_Linha.Caption = Format((CDbl(Valor_Linha.Caption) - ((Vendas_Prod.Text) * CDbl(Grade1.Columns(8).Text))), FORMAT_VALUE)
      Valor_Total.Caption = Format((CDbl(Valor_Total.Caption) - ((Vendas_Prod.Text) * CDbl(Grade1.Columns(8).Text))), FORMAT_VALUE)
    End If
    .Close
  End With
  
  Set rstConsigSaidas = Nothing

End Sub

Private Sub BuscarNomeCliente(ByRef strNome As String)
  '17/03/2005 - Daniel
  Dim rstClientes As Recordset
  
  If Len(Combo_Cliente.Text) <= 0 Then Exit Sub
  
  Set rstClientes = db.OpenRecordset("SELECT Nome FROM Cli_For WHERE C�digo = " & CLng(Combo_Cliente.Text), dbOpenSnapshot)

  With rstClientes
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      strNome = .Fields("Nome").Value & ""
    End If
    .Close
  End With

  Set rstClientes = Nothing

End Sub

Private Sub Vendas_Prod_LostFocus()
If Not IsNumeric(Vendas_Prod.Text) Then
      DisplayMsg "Quantidade incorreta."
      Vendas_Prod.SetFocus
End If
End Sub
