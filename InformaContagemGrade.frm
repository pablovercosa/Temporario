VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmInformaContaGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Ajuste de Estoque - Produtos com Grade (tamanho e cor)"
   ClientHeight    =   8265
   ClientLeft      =   420
   ClientTop       =   1305
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1030
   Icon            =   "InformaContagemGrade.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8265
   ScaleWidth      =   13710
   Begin VB.Frame frm_acertaEstoque 
      Caption         =   "Acertar Estoque "
      Height          =   1185
      Left            =   60
      TabIndex        =   10
      Top             =   7020
      Width           =   13575
      Begin VB.OptionButton O_Todos 
         Caption         =   "S&omente produtos com a coluna ""Diferença"" diferente de zero"
         Height          =   375
         Left            =   7020
         TabIndex        =   13
         Top             =   210
         Width           =   4815
      End
      Begin VB.OptionButton O_Consertar 
         Caption         =   "&Somente produtos com a coluna ""Consertar"" marcada"
         Height          =   405
         Left            =   1845
         TabIndex        =   12
         Top             =   195
         Value           =   -1  'True
         Width           =   4305
      End
      Begin VB.CommandButton B_Acerta 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Acertar"
         Height          =   465
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   13245
      End
   End
   Begin VB.CommandButton B_Iguala 
      Caption         =   "&Igualar"
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
      Left            =   13230
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Fazer a Contagem Igual ao Estoque"
      Top             =   8220
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ajustar Estoque"
      Height          =   5235
      Left            =   60
      TabIndex        =   3
      Top             =   1710
      Width           =   13575
      Begin VB.TextBox txt_quantidade 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F5DA33&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3345
         TabIndex        =   31
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txt_NomeProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5DA33&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   26
         Top             =   240
         Width           =   3045
      End
      Begin VB.OptionButton opt_ajustePadrão 
         Appearance      =   0  'Flat
         Caption         =   "Modo Padrão"
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
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   315
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opt_ajusteComLeitor 
         Appearance      =   0  'Flat
         Caption         =   "Modo com Leitor"
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
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   690
         Width           =   1815
      End
      Begin VB.TextBox TxtLocaliza 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5DA33&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3360
         TabIndex        =   21
         Top             =   660
         Width           =   3045
      End
      Begin VB.CommandButton B_Monta 
         BackColor       =   &H00F5DA33&
         Caption         =   "Listar os produtos para ajuste de estoque"
         Height          =   465
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1140
         Width           =   6285
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1425
         Left            =   12030
         TabIndex        =   5
         Top             =   120
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   2514
         _StockProps     =   14
         Caption         =   "Ordeção produto"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox O_Classe 
            Appearance      =   0  'Flat
            Caption         =   "Classe"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   210
            TabIndex        =   23
            Top             =   960
            Width           =   795
         End
         Begin VB.OptionButton O_Código 
            Appearance      =   0  'Flat
            Caption         =   "Por código"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   7
            Top             =   300
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton O_Nome 
            Appearance      =   0  'Flat
            Caption         =   "Por nome"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   6
            Top             =   600
            Width           =   1035
         End
      End
      Begin SSDataWidgets_B.SSDBGrid Grade1 
         Bindings        =   "InformaContagemGrade.frx":4E95A
         Height          =   3495
         Left            =   120
         TabIndex        =   8
         Top             =   1650
         Width           =   13335
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
         AllowDelete     =   -1  'True
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         ForeColorEven   =   0
         BackColorOdd    =   16112179
         RowHeight       =   450
         Columns(0).Width=   3200
         UseDefaults     =   0   'False
         _ExtentX        =   23521
         _ExtentY        =   6165
         _StockProps     =   79
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
      Begin SSDataWidgets_B.SSDBCombo Combo_Classe1 
         Bindings        =   "InformaContagemGrade.frx":4E96E
         DataSource      =   "Data1"
         Height          =   345
         Left            =   7125
         TabIndex        =   27
         ToolTipText     =   "Use 0 para todas"
         Top             =   240
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
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   6244
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1455
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1561
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   16112179
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Atualmente em estoque no QuickStore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6630
         TabIndex        =   34
         Top             =   750
         Width           =   2010
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Contagem do inventário"
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
         Left            =   8730
         TabIndex        =   33
         Top             =   1230
         Width           =   1965
      End
      Begin VB.Line Line2 
         X1              =   7980
         X2              =   7980
         Y1              =   1140
         Y2              =   1590
      End
      Begin VB.Line Line1 
         X1              =   8640
         X2              =   8640
         Y1              =   1260
         Y2              =   1590
      End
      Begin VB.Line Line3 
         X1              =   7980
         X2              =   8070
         Y1              =   1620
         Y2              =   1470
      End
      Begin VB.Line Line4 
         X1              =   7980
         X2              =   7890
         Y1              =   1620
         Y2              =   1470
      End
      Begin VB.Line Line5 
         X1              =   8640
         X2              =   8730
         Y1              =   1620
         Y2              =   1470
      End
      Begin VB.Line Line6 
         X1              =   8640
         X2              =   8550
         Y1              =   1620
         Y2              =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
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
         Left            =   2370
         TabIndex        =   32
         Top             =   330
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parte do Nome"
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
         Left            =   2055
         TabIndex        =   30
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Classe"
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
         Left            =   6570
         TabIndex        =   29
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Nome_Classe1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   8055
         TabIndex        =   28
         Top             =   240
         Width           =   3915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Produto"
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
         Left            =   2040
         TabIndex        =   22
         Top             =   735
         Width           =   1275
      End
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
      Height          =   1665
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      Begin VB.Frame Frame3 
         Caption         =   "Opções"
         Height          =   915
         Left            =   6120
         TabIndex        =   17
         Top             =   120
         Width           =   5325
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Separar por classe"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   20
            Top             =   285
            Width           =   2235
         End
         Begin VB.CheckBox O_Inativos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Considerar produtos inativos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2670
            TabIndex        =   19
            Top             =   285
            Width           =   2550
         End
         Begin VB.CheckBox O_Zero 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Não considerar produtos com estoque zero"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   210
            TabIndex        =   18
            Top             =   585
            Width           =   3645
         End
      End
      Begin VB.CommandButton bt_gerarContagemEstoque 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Inicializar processo para ajuste de estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1110
         Width           =   13275
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Relatório de estoque atual"
         Height          =   795
         Left            =   11550
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Chamar o Rel. de Contagem de Estoque"
         Top             =   240
         Width           =   1875
      End
      Begin SSDataWidgets_B.SSDBCombo Combo 
         Bindings        =   "InformaContagemGrade.frx":4E982
         DataSource      =   "Data2"
         Height          =   345
         Left            =   735
         TabIndex        =   14
         Top             =   240
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
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   8467
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1561
         Columns(1).Caption=   "Filial"
         Columns(1).Name =   "Filial"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Filial"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         _ExtentX        =   1561
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
      End
      Begin VB.Label lbl_aguardando 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "...aguarde até que o processo de Inicialização se complete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   150
         TabIndex        =   35
         Top             =   750
         Visible         =   0   'False
         Width           =   5085
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Classe"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Nome_Combo 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1665
         TabIndex        =   15
         Top             =   240
         Width           =   4275
      End
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
      Height          =   345
      Left            =   12660
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   7710
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   12660
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmInformaContaGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsParametros1 As Recordset
Dim rsProdutos1 As Recordset
Dim rsEstoque_Final1 As Recordset
Dim rsClasses1 As Recordset
Dim rsSub_Classes1 As Recordset
Dim rsCores1 As Recordset
Dim rsTamanhos1 As Recordset
Dim TB2_Contagem1 As Recordset
Dim rsGrade1 As Recordset

Dim rsContagem As Recordset
Dim rsProdutos2 As Recordset
Dim rsEstoque2  As Recordset

Dim rsPreços As Recordset
Dim Rec_Contas As Recordset
Dim rsClasses As Recordset
Dim rsSub_Classes As Recordset
Dim TB2_Contagem As Recordset

Private Type FindProdutos
  sCodigo As String
  snome As String
End Type

Private tabProdutos() As FindProdutos

Dim gnColuna As Integer

Private Sub B_Acerta_Click()
Dim Resposta       As Integer
  Dim Código         As String
  Dim Tamanho        As Integer
  Dim Cor            As Integer
  Dim Conta          As Long
  Dim Criar_Registro As Integer
  Dim Estoque_Final  As Single
  Dim Mes_Atual      As Integer
  Dim Ano_Atual      As Integer
  
  gbAcertaGrade = True
  
  Call StatusMsg("")
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Este processo não poderá ser desfeito, deseja prosseguir?"
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    DisplayMsg "Estoque não foi atualizado."
    Exit Sub
  End If
  
  On Error GoTo ErrTrans
  
  Screen.MousePointer = vbHourglass
  
  Call ws.BeginTrans
  
  Código = ""
  Tamanho = 0
  Cor = 0
  Conta = 0
  rsProdutos2.Index = "Código"
  rsContagem.Index = "Código"

Lp1:
  If gbAcertaGrade = True Then
    rsContagem.Seek ">", Código, Tamanho, Cor
  Else
    rsContagem.Seek ">", Código
  End If
  
  If rsContagem.NoMatch Then GoTo Fim_Lp
  Código = rsContagem("Código")
  
  If gbAcertaGrade = True Then
    Tamanho = rsContagem("Tamanho")
    Cor = rsContagem("Cor")
  End If
  
  'Verifica se a filial de origem é a mesma que está logado
  If rsContagem("Empresa") <> gnCodFilial Then GoTo Lp1
  
  If rsContagem("Diferença") = 0 Then GoTo Lp1
  
  If O_Consertar.Value = True Then
    If rsContagem("Consertar") = False Then GoTo Lp1
  End If
  
  rsProdutos2.Seek "=", rsContagem("Código")
  If rsProdutos2.NoMatch Then GoTo Lp1
  
  Conta = Conta + 1
  
  Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
  
  Rem Acha Último Estoque deste produto
  Criar_Registro = False
  Estoque_Final = 0
  rsEstoque2.Index = "Produto"
  rsEstoque2.Seek "=", rsContagem("Empresa"), Data_Atual, rsContagem("Código"), Tamanho, Cor, 0
  
  If Not rsEstoque2.NoMatch Then
    Estoque_Final = rsEstoque2("Estoque Final")
  End If
  
  If rsEstoque2.NoMatch Then
    
    rsEstoque2.Index = "Data"
    rsEstoque2.Seek "<", rsContagem("Empresa"), rsContagem("Código"), Tamanho, Cor, 0, Data_Atual
    If rsEstoque2.NoMatch Then Criar_Registro = True
    If Not rsEstoque2.NoMatch Then
      If rsEstoque2("Filial") = rsContagem("Empresa") And rsEstoque2("Produto") = rsContagem("Código") And rsEstoque2("Tamanho") = 0 And rsEstoque2("Cor") = 0 And rsEstoque2("Edição") = 0 Then
        Criar_Registro = True
        Estoque_Final = rsEstoque2("Estoque Final")
      End If
    End If
  
    rsEstoque2.AddNew
    rsEstoque2("Filial") = rsContagem("Empresa")
    rsEstoque2("Data") = Data_Atual
    rsEstoque2("Produto") = rsContagem("Código")
    rsEstoque2("Tamanho") = Tamanho
    rsEstoque2("Cor") = Cor
    rsEstoque2("Edição") = 0
    rsEstoque2("Classe") = rsProdutos("Classe")
    rsEstoque2("Sub Classe") = rsProdutos("Sub Classe")
    rsEstoque2("Estoque Anterior") = Estoque_Final
    rsEstoque2.Update
    
    rsEstoque2.Index = "Produto"
    rsEstoque2.Seek "=", rsContagem("Empresa"), Data_Atual, rsContagem("Código"), Tamanho, Cor, 0
  
  End If
  
  'Verifica se a real diferença está correta
  If rsContagem("Qtde Estoque") <> Estoque_Final Then
    With rsContagem
      .Edit
      .Fields("Qtde Estoque") = Estoque_Final
      .Fields("Diferença") = .Fields("Digitado") - Estoque_Final
      .Update
    End With
    If gbAcertaGrade = True Then
      rsContagem.Seek "=", Código, Tamanho, Cor
    Else
      rsContagem.Seek "=", Código
    End If
  End If
  
  Rem neste ponto esta com o registro de estoque
  Rem no buffer, agora soma com os valores da movimentação
  rsEstoque2.Edit
  If rsContagem("Diferença") < 0 Then
    rsEstoque2("Ajuste Saída") = rsEstoque2("Ajuste Saída") + Abs(rsContagem("Diferença"))
  End If
  
  If rsContagem("Diferença") > 0 Then
    rsEstoque2("Ajuste Entra") = rsEstoque2("Ajuste Entra") + Abs(rsContagem("Diferença"))
  End If
  
  Estoque_Final = rsEstoque2("Estoque Anterior") - rsEstoque2("Vendas") + rsEstoque2("Compras")
  Estoque_Final = Estoque_Final - rsEstoque2("Transf Saída") + rsEstoque2("Transf Entra")
  Estoque_Final = Estoque_Final - rsEstoque2("Ajuste Saída") + rsEstoque2("Ajuste Entra")
  Estoque_Final = Estoque_Final - rsEstoque2("Grátis Saída") + rsEstoque2("Grátis Entra")
  Estoque_Final = Estoque_Final - rsEstoque2("Empre Saída") + rsEstoque2("Empre Entra")
  Estoque_Final = Estoque_Final - rsEstoque2("Quebras") + rsEstoque2("Devolução")
  
  If rsProdutos2("Estoque") = False Then
    Estoque_Final = 0
  End If
  
  rsEstoque2("Estoque Final") = Estoque_Final
  rsEstoque2.Update
  
  If gbAcertaGrade Then
    Call Grava_Estoque_Final(rsContagem("Empresa"), rsProdutos2("Código"), Tamanho, Cor, 0, Estoque_Final, CDate(Data_Atual))
  Else
    Call Grava_Estoque_Final(rsContagem("Empresa"), rsProdutos2("Código"), 0, 0, 0, Estoque_Final, CDate(Data_Atual))
  End If
  
  rsContagem.Edit
  rsContagem("Diferença") = 0
  rsContagem("Qtde Estoque") = rsContagem("Digitado")
  rsContagem("Consertar") = False
  rsContagem.Update
  
  GoTo Lp1
  
Fim_Lp:

  '---[ Gera Log do usuário ]---'
      g_GravaLog Data_Atual, "Acerto de Estoque, DQ(" & Data_Atual & "), DW(" & Date & "),Funcionário: " & _
                            gnUserCode & " - " & gsUserName, "ACERTO ESTOQUE"
  '---[ Gera Log do usuário ]---'
  
  dbTemp.Execute "Delete * From [Contagem Grade]"
  
  Call ws.CommitTrans
  Screen.MousePointer = vbDefault
  DisplayMsg "Fim de processo. Registros atualizados : " + str(Conta)
  Exit Sub
  
ErrTrans:
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Acertar Estoque."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly & vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  On Error Resume Next
  Call ws.Rollback
  Exit Sub
End Sub

Private Sub B_Iguala_Click()
' Dim Código As String
' Dim nTam As Integer
' Dim nCor As Integer
'
'  Call StatusMsg("Aguarde...")
'  Screen.MousePointer = vbHourglass
'
'  TB2_Contagem.Index = "Código"
'  Código = 0
'  nTam = 0
'  nCor = 0
'
'Lp1:
'  TB2_Contagem.Seek ">", Código, nTam, nCor
'  If TB2_Contagem.NoMatch Then GoTo Fim_Loop
'
'  Código = TB2_Contagem("Código")
'  nTam = TB2_Contagem("Tamanho")
'  nCor = TB2_Contagem("Cor")
'
'  'Verifica se a filial de origem
'  If TB2_Contagem("Empresa") <> gnCodFilial Then GoTo Lp1
'
'  TB2_Contagem.Edit
'    TB2_Contagem("Digitado") = TB2_Contagem("Qtde Estoque")
'    TB2_Contagem("Diferença") = 0
'  TB2_Contagem.Update
'
'  GoTo Lp1
'
'Fim_Loop:
'  Call StatusMsg("")
'  Screen.MousePointer = vbDefault
'
'  B_Monta_Click
  
'Novo Código
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde...")
  dbTemp.Execute "UPDATE [Contagem Grade] SET Digitado = [Qtde Estoque], Diferença = 0 WHERE Empresa = " & gnCodFilial, dbFailOnError
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  Call B_Monta_Click
  
End Sub

Private Sub B_Monta_Click()
 Dim sSql As String
 Dim i As Integer
 Dim Classe As Integer
 Dim sCodigoProduto As String
 Dim sTamanho As String
 Dim sCor As String
 
  Call StatusMsg("")
 
 On Error GoTo Processa_Erro
 
  If Len(Trim(TxtLocaliza.Text)) <= 6 Then
      MsgBox "Informe um produto com Grade", vbInformation, "Atenção"
      Exit Sub
  End If
 
  If opt_ajusteComLeitor.Value = True Then
    If Len(Trim(TxtLocaliza.Text)) <= 3 Then
        MsgBox "Informe um produto com Grade", vbInformation, "Atenção"
        Exit Sub
    Else
        sCodigoProduto = Mid(Trim(TxtLocaliza.Text), 1, Len(Trim(TxtLocaliza.Text)) - 6)
        sTamanho = Mid(Trim(TxtLocaliza.Text), Len(sCodigoProduto) + 1, 3)
        sCor = Mid(Trim(TxtLocaliza.Text), Len(sCodigoProduto) + 4, 3)
    End If
  Else
      sCodigoProduto = Trim(TxtLocaliza.Text)
  End If
  
  Classe = 0
  If Nome_Classe1.Caption <> "" Then Classe = Val(Combo_Classe1.Text)
  
  sSql = "SELECT Código, Nome, Classe, Cor, [Nome Cor], Tamanho, [Nome Tamanho], [Qtde Estoque], Digitado, Diferença, Consertar FROM [Contagem Grade]"
  
  'Verifica se a filial de origem
  sSql = sSql & " WHERE Empresa = " & gnCodFilial
  
  Dim boPulaCondicao As Boolean
  boPulaCondicao = False
  If Trim(TxtLocaliza.Text) <> "" Then
      '''sSql = sSql & " AND Código = '" & Trim(TxtLocaliza.Text) & "'"
      sSql = sSql & " AND Código = '" & sCodigoProduto & "'"
      
      If opt_ajusteComLeitor.Value = True Then
          sSql = sSql & " AND Tamanho = " & CInt(sTamanho)
          sSql = sSql & " AND Cor = " & CInt(sCor)
      End If
      
      boPulaCondicao = True
  End If

  If boPulaCondicao = False And Classe <> 0 Then
    sSql = sSql + " AND Classe = " + str(Classe)
  End If
  
  If boPulaCondicao = False And Trim(txt_NomeProduto.Text) <> "" Then
      sSql = sSql & " AND Nome like '*" & Trim(txt_NomeProduto.Text) & "*' "
  End If
  
  If O_Classe.Value = 1 Then
    If O_Código.Value Then
      sSql = sSql + " ORDER BY Classe, Código, Tamanho, Cor"
    Else
      sSql = sSql + " ORDER BY Classe, Nome, Tamanho, Cor"
    End If
  Else
    If O_Código.Value Then
      sSql = sSql + " ORDER BY Código, Tamanho, Cor"
    Else
      sSql = sSql + " ORDER BY Nome, Tamanho, Cor"
    End If
  End If
  
  
  Set Rec_Contas = dbTemp.OpenRecordset(sSql, dbOpenDynaset)

  Grade1.DataMode = 1

  Set Data2.Recordset = Rec_Contas


  Grade1.Visible = False
  
  Grade1.DataMode = 0
  
  Grade1.ReBind
 
    Grade1.Columns(0).Width = 900
    Grade1.Columns(0).Locked = True
    
    Grade1.Columns(1).Width = 2500
    Grade1.Columns(1).Locked = True
    
    Grade1.Columns(2).NumberFormat = "#####0"
    Grade1.Columns(2).Width = 650
    Grade1.Columns(2).Locked = True
    
    Grade1.Columns(3).Width = 500
    Grade1.Columns(3).Locked = True
    Grade1.Columns(4).Width = 900
    Grade1.Columns(4).Locked = True
    
    Grade1.Columns(5).Width = 800
    Grade1.Columns(5).Locked = True
    Grade1.Columns(6).Width = 900
    Grade1.Columns(6).Locked = True
    
    
    Grade1.Columns(7).Locked = True
    Grade1.Columns(7).Width = 650
    
    Grade1.Columns(8).Width = 710
    
    Grade1.Columns(9).Locked = True
    
    Grade1.Columns(10).Style = 2
   
    
  Grade1.Visible = True
  
'''  frmLocalizaProduto.Visible = True

  Call StatusMsg("")
  
  If opt_ajusteComLeitor.Value = True Then
      AtualizaEstoqueDoProdutoNaGrade Trim(sCodigoProduto), sTamanho, sCor
  End If

 
  Exit Sub
Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Informar Contagem de Estoque (Grade) - Montar")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Sub
    Case 3 'Encerrar
      End
  End Select
  
End Sub

Private Sub bt_gerarContagemEstoque_Click()
On Error GoTo Erro:

  Dim Termina As Integer
  Dim Val2 As Integer
  Dim Erro As Integer
  Dim Str1 As String
  Dim Str2 As String
  Dim Str3 As String
  Dim Str_Data1 As String
  Dim Str_Data2 As String
  Dim Str_Rel As String
  Dim Data1 As Variant
  Dim Produto As String
  Dim Aux_Produto As String
  Dim Completo As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Tipo As Integer
  Dim sSql As String
  Dim Estoque As Double
  Dim Aux_Data As Variant
  Dim Aux_Classe As Integer
  Dim Aux_Sub As Integer
  Dim Nome_Cla As String
  Dim Nome_Sub As String
  Dim Nome_Cor As String
  Dim Nome_Tam As String
    
  Call StatusMsg("")
  
  Rem apaga pesquisa anterior desta filial do arquivo temporario
  Call StatusMsg("Aguarde, preparando arquivo temporário ...")
  
  lbl_aguardando.Visible = True
  
  sSql = "Delete * From [Contagem Grade] Where Empresa = " & gnCodFilial
  dbTemp.Execute sSql
  
  Call StatusMsg("")
 
  Rem Le estoque e joga no temporário
  rsProdutos1.Index = "Código"
  rsEstoque_Final1.Index = "Produto"
  Termina = False
  Produto = ""
  Call StatusMsg("Aguarde, contando estoque...")
  
  rsClasses1.Index = "Código"
  rsSub_Classes1.Index = "Código"
  rsGrade1.Index = "Original"
  rsCores1.Index = "Código"
  rsTamanhos1.Index = "Código"
 
LP1S:
  rsProdutos1.Seek ">", Produto
  If rsProdutos1.NoMatch Then GoTo Imprime
  Produto = rsProdutos1("Código")
  If Produto = "0" Then GoTo LP1S
  
  If Nome_Combo.Caption <> "" Then
    If rsProdutos1("Classe") <> Val(Combo.Text) Then GoTo LP1S
  End If
  
  If rsProdutos1("Desativado") = True And O_Inativos.Value = 0 Then GoTo LP1S

  If rsProdutos1("Tipo") <> "G" Then GoTo LP1S

  If rsProdutos1("Fracionado") = True Then GoTo LP1S

  Rem    Tem um produto com grade disponível
  Rem    Agora deve achar todas as cores e tamanhos possíveis
  Rem    e seus respectivos estoques
  Completo = ""
LP2:
  rsGrade1.Seek ">", Produto, Completo
  If rsGrade1.NoMatch Then GoTo LP1S
  If rsGrade1("Código Original") <> Produto Then GoTo LP1S
  
  Completo = rsGrade1("Código")
  
  Acha_Produto Completo, Aux_Produto, Tamanho, Cor, Edição, Tipo, Erro
  
  If Erro <> 0 Then GoTo LP2
  
  '''Estoque = Acha_Estoque(Val(Combo.Text), Produto, Tamanho, Cor, 0, Erro)
  Estoque = Acha_Estoque(gnCodFilial, Produto, Tamanho, Cor, 0, Erro)
  If Erro > 1 Then GoTo LP2
  
  If Estoque = 0 Then
   If O_Zero.Value = 1 Then GoTo LP2
  End If
  
  
  Call StatusMsg("Aguarde, gravando arquivo temporário, produto " & (Produto))
     
     
  rsClasses1.Seek "=", rsProdutos1("Classe")
  If rsClasses1.NoMatch Then
     Nome_Cla = "Classe não cadastrada"
  Else
     Nome_Cla = rsClasses1("Nome")
  End If
  
  rsSub_Classes1.Seek "=", rsProdutos1("Sub Classe")
  If rsSub_Classes1.NoMatch Then
    Nome_Sub = "Subclasse não cadastrada"
  Else
    Nome_Sub = rsSub_Classes1("Nome")
  End If
  
  rsCores1.Seek "=", Cor
  If rsCores1.NoMatch Then
    Nome_Cor = "Cor não cadastrada"
  Else
    Nome_Cor = rsCores1("Nome")
  End If
  
  rsTamanhos1.Seek "=", Tamanho
  If rsTamanhos1.NoMatch Then
    Nome_Tam = "Tamanho não cadastrado"
  Else
    Nome_Tam = rsTamanhos1("Nome")
  End If
  
  TB2_Contagem1.AddNew
     TB2_Contagem1("Código") = Produto
     TB2_Contagem1("Código Ordenação") = rsProdutos1("Código Ordenação")
     TB2_Contagem1("Nome") = rsProdutos1("Nome")
     TB2_Contagem1("Classe") = rsProdutos1("Classe")
     TB2_Contagem1("Nome Classe") = Nome_Cla
     TB2_Contagem1("Sub Classe") = rsProdutos1("Sub Classe")
     TB2_Contagem1("Nome Sub") = Nome_Sub
     TB2_Contagem1("Unidade") = rsProdutos1("Unidade Venda")
     TB2_Contagem1("Qtde Estoque") = Estoque
     TB2_Contagem1("Empresa") = gnCodFilial
     TB2_Contagem1("Cor") = Cor
     TB2_Contagem1("Nome Cor") = Nome_Cor
     TB2_Contagem1("Tamanho") = Tamanho
     TB2_Contagem1("Nome Tamanho") = Nome_Tam
  TB2_Contagem1.Update
  
  GoTo LP2

Imprime:

  Call StatusMsg("")
  MousePointer = vbDefault

  lbl_aguardando.Visible = False
  MsgBox "Processo iniciado com sucesso", vbInformation, "Estoque"
  Exit Sub
  
Erro:
    MsgBox "Erro na realização de inicialização do processo de ajuste de estoque " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    lbl_aguardando.Visible = False
 
End Sub

Private Sub Combo_Classe1_CloseUp()
  Combo_Classe1.Text = Combo_Classe1.Columns(1).Text
  Combo_Classe1_LostFocus
End Sub

Private Sub Combo_Classe1_LostFocus()
  Dim Aux As Variant
  
  Nome_Classe1.Caption = ""
  Aux = Combo_Classe1.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsClasses.Index = "Código"
  rsClasses.Seek "=", Val(Aux)
  If rsClasses.NoMatch Then Exit Sub
  
  Nome_Classe1.Caption = rsClasses("Nome")
End Sub

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(1).Text
  Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
  Dim Aux As Variant

  Nome_Combo.Caption = ""
  Aux = Combo.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsClasses.Index = "Código"
  rsClasses.Seek "=", Val(Aux)
  If rsClasses.NoMatch Then Exit Sub
  
  Nome_Combo.Caption = rsClasses("Nome")
End Sub

Private Sub Command1_Click()
  '14/05/2005 - Daniel
  'Otimizando a chamada da tela de Rel. Contagem de
  'Estoque caso seja necessário
  frmRelContagemGrade.Show
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)
 
  Data1.DatabaseName = gsQuickDBFileName
 
  ' ======================================================================
  ' Tratando o frm_contagemEstoque
  Set rsParametros1 = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsProdutos1 = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsEstoque_Final1 = db.OpenRecordset("Estoque Final", , dbReadOnly)
  Set rsClasses1 = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSub_Classes1 = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsCores1 = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsTamanhos1 = db.OpenRecordset("Tamanhos", , dbReadOnly)
  
  Set TB2_Contagem1 = dbTemp.OpenRecordset("Contagem Grade")
  
  Set rsGrade1 = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  ' ======================================================================
 
  ' ======================================================================
  ' Tratando o frm_acertaEstoque
    Dim sSql As String
    Dim sCaption As String
  
'''    If gbAcertaGrade Then
      sSql = "Contagem Grade"
      sCaption = "(Produtos com Grade)"
'''    Else
'''      sSql = "Contagem"
'''      sCaption = ""
'''    End If
    '''Me.Caption = "Acerta Estoque " & sCaption
    Set rsContagem = dbTemp.OpenRecordset(sSql)
    Set rsProdutos2 = db.OpenRecordset("Produtos", , dbReadOnly)
    Set rsEstoque2 = db.OpenRecordset("Estoque")
  ' ======================================================================

   
 
 
  Set TB2_Contagem = dbTemp.OpenRecordset("Contagem Grade")
  '-----------------------------------------------------------------------------
  '03/08/2006 - Andrea
  'Inclui um frame : Localizador de produtos, pois para alguns clientes (Idéia Íntima)
  'por exemplo, localizar um item na grid era muito difícil, porque eles trabalham
  'com muitos produtos.
  'Quando entra, o frame está invisível, somente aparece após o usuário
  'Pesquisar os produtos (Processo que preenche a grade com informações).
'''  frmLocalizaProduto.Visible = False
  '-------------------------------------------------------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' ========================================================
  ' Objetos do frm_contagemEstoque
  rsCores1.Close
  rsTamanhos1.Close
  rsGrade1.Close
  rsParametros1.Close
  rsProdutos1.Close
  rsEstoque_Final1.Close
  rsClasses1.Close
  rsSub_Classes1.Close
  Set rsParametros1 = Nothing
  Set rsProdutos1 = Nothing
  Set rsEstoque_Final1 = Nothing
  Set rsClasses1 = Nothing
  Set rsSub_Classes1 = Nothing
  Set rsCores1 = Nothing
  Set rsTamanhos1 = Nothing
  Set rsGrade1 = Nothing
  ' ========================================================
  
  ' ========================================================
  ' Objetos do frm_acertaEstoque
  rsContagem.Close
  rsProdutos2.Close
  rsEstoque2.Close
  Set rsContagem = Nothing
  Set rsProdutos2 = Nothing
  Set rsEstoque2 = Nothing
  ' ========================================================
End Sub

Private Sub Grade1_AfterDelete(RtnDispErrMsg As Integer)
  Grade1.Scroll 0, -32767
  Grade1.Scroll 0, 32767
End Sub

Private Sub Grade1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  Dim Erro As Integer
  Dim Est, Dig, Dif As Double
  
  
  If ColIndex = 8 Then
    Erro = False
    Aux = Grade1.Columns(8).Text
    If IsNull(Aux) Then Erro = True
    If Erro = False Then If Not IsNumeric(Aux) Then Erro = True
    If Erro = False Then If Abs(CDbl(Aux)) > 999999999 Then Erro = True
  '  If Erro = False Then If CDbl(Aux) < 0 Then Erro = True
    If Erro = True Then
      DisplayMsg "Digite um valor."
      Cancel = True
      Exit Sub
    End If
    Grade1.Columns(8).Text = Format(CDbl(Aux), "#########0")
    
    Est = Grade1.Columns(7).Text
    Dig = Aux
    Dif = Abs(CDbl(Est) - Dig)
    
    If CDbl(Est) > CDbl(Dig) Then Dif = Dif * -1
   ' If Est < 0 And Dig < 0 And CDbl(Dig) < CDbl(Est) Then Dif = Dif * -1
    
    Grade1.Columns(9).Text = Format(Dif, "########0")
    
  End If
      
End Sub

Private Sub Grade1_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  If Not bGridBeforeDelete() Then
    Cancel = True
  End If
End Sub

Private Sub Grade1_BeforeUpdate(Cancel As Integer)
' Exit Sub
' On Error GoTo Deu_Erro
' Grade1.Columns(4).Text = Format(Date, "dd/mm/yyyy")
'Deu_Erro:
' On Error GoTo 0
End Sub

Private Sub Grade1_LostFocus()
  With Grade1
    If .RowChanged Then
      .Update
    End If
  End With
End Sub

Private Sub opt_ajusteComLeitor_Click()
    If opt_ajusteComLeitor.Value = True Then
        Label4.Visible = True
        Label1.Visible = True
        txt_quantidade.Visible = True
        txt_quantidade.Text = "1"
        TxtLocaliza.Text = ""
'        Label7.Visible = True
'        Label8.Visible = True
'        Line1.Visible = True
'        Line2.Visible = True
'        Line3.Visible = True
'        Line4.Visible = True
'        Line5.Visible = True
'        Line6.Visible = True
        
        Label2.Visible = False
        Label3.Visible = False
        txt_NomeProduto.Text = ""
        txt_NomeProduto.Visible = False
        Combo_Classe1.Text = ""
        Combo_Classe1.Visible = False
        Nome_Classe1.Caption = ""
        Nome_Classe1.Visible = False
        SSFrame1.Visible = False
        B_Monta.Visible = False
        Grade1.Visible = False
        
        TxtLocaliza.SetFocus
    End If
End Sub

Private Sub opt_ajustePadrão_Click()
    If opt_ajustePadrão.Value = True Then
        Label2.Visible = True
        txt_NomeProduto.Visible = True
        Label3.Visible = True
        Combo_Classe1.Visible = True
        Nome_Classe1.Visible = True
        SSFrame1.Visible = True
        B_Monta.Visible = True
    
        TxtLocaliza.Text = ""
        Label4.Visible = False
        txt_quantidade.Text = ""
        txt_quantidade.Visible = False
        Grade1.Visible = False
'        Line1.Visible = False
'        Line2.Visible = False
'        Line3.Visible = False
'        Line4.Visible = False
'        Line5.Visible = False
'        Line6.Visible = False
'        Label7.Visible = False
'        Label8.Visible = False
    End If
End Sub

Private Sub TxtLocaliza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then  'Tecla Enter
        B_Monta_Click
    End If
End Sub

Private Sub AtualizaEstoqueDoProdutoNaGrade(pCodigoProdutoGrade As String, pTamanho As String, pCor As String)
On Error GoTo Erro

  Dim Aux As Variant
  Dim Erro As Integer
  Dim Est, Dig, Dif As Double
  
  Dim nQtdeCasaDec As Integer
  Dim bFrac As Boolean
  
  Erro = False
  Aux = Grade1.Columns(8).Text
  If IsNull(Aux) Then Erro = True
  If Erro = False Then If Not IsNumeric(Aux) Then Erro = True
  If Erro = False Then If Abs(CDbl(Aux)) > 999999999 Then Erro = True
  If Erro = True Then
      DisplayMsg "Digite um valor."
      Exit Sub
  End If
    
    
  If IsNull(txt_quantidade.Text) Then Erro = True
  If Erro = False Then If Not IsNumeric(txt_quantidade.Text) Then Erro = True
  If Erro = False Then If Abs(CDbl(txt_quantidade.Text)) > 999999999 Then Erro = True
  If Erro = True Then
      DisplayMsg "Digite uma quantidade válida."
      txt_quantidade.SetFocus
      Exit Sub
  End If
    
  If gbIsFrac(Grade1.Columns(0).Text, nQtdeCasaDec) Then
    bFrac = True
    Grade1.Columns(4).Text = Round(CDbl(Aux) + CDbl(txt_quantidade.Text), nQtdeCasaDec)   'Format(CDbl(Aux), "#0.000")
  Else
    Grade1.Columns(8).Text = Format(CDbl(Aux) + CDbl(txt_quantidade.Text), "#0")
  End If
    
  Est = Grade1.Columns(7).Text
  Dig = CDbl(Aux) + CDbl(txt_quantidade.Text)
  Dif = Abs(CDbl(Est) - Dig)
    
  If CDbl(Est) > CDbl(Dig) Then Dif = Dif * -1
  
  If bFrac Then
    Grade1.Columns(9).Text = Round(Dif, nQtdeCasaDec)  'Format(Dif, "#0.000")
  Else
    Grade1.Columns(9).Text = Format(Dif, "#0")
  End If
  
  Grade1.Columns(10).Text = vbChecked
  Grade1.Update
  
  txt_quantidade.Text = "1"
  TxtLocaliza.Text = ""
      
  Exit Sub
Erro:
  MsgBox "Erro na atualização do número estoque na grade " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
      
End Sub
