VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRegInvent 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Registro de Inventário"
   ClientHeight    =   5715
   ClientLeft      =   3240
   ClientTop       =   1680
   ClientWidth     =   8415
   ForeColor       =   &H80000008&
   Icon            =   "RegInvent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   8415
   Begin Crystal.CrystalReport Rel 
      Left            =   4680
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame5 
      Height          =   3015
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   8175
      Begin VB.CheckBox chk_dataAte 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5610
         TabIndex        =   31
         Top             =   240
         Width           =   195
      End
      Begin VB.CommandButton cmd_calendarioDtFim 
         Height          =   420
         Left            =   7485
         Picture         =   "RegInvent.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   172
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frame1 
         Caption         =   "Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4080
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
         Begin VB.OptionButton B_Vídeo 
            Appearance      =   0  'Flat
            Caption         =   "Vídeo"
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
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton B_Impressora 
            Appearance      =   0  'Flat
            Caption         =   "Impressora"
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
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6840
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
         Begin VB.OptionButton O_Normal 
            Appearance      =   0  'Flat
            Caption         =   "Normal"
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
            Height          =   225
            Left            =   105
            TabIndex        =   14
            Top             =   210
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton O_Grade 
            Appearance      =   0  'Flat
            Caption         =   "Grade"
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
            Height          =   225
            Left            =   105
            TabIndex        =   15
            Top             =   420
            Width           =   1065
         End
         Begin VB.OptionButton O_Edições 
            Appearance      =   0  'Flat
            Caption         =   "Edições"
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
            Height          =   225
            Left            =   105
            TabIndex        =   16
            Top             =   630
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5640
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
         Begin VB.OptionButton O_Código 
            Appearance      =   0  'Flat
            Caption         =   "Código"
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
            Height          =   225
            Left            =   105
            TabIndex        =   11
            Top             =   315
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton O_Nome 
            Appearance      =   0  'Flat
            Caption         =   "Nome"
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
            Height          =   225
            Left            =   105
            TabIndex        =   12
            Top             =   630
            Width           =   855
         End
      End
      Begin VB.CheckBox O_Preço 
         Appearance      =   0  'Flat
         Caption         =   "Não imprimir produtos que não tenham preço."
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox O_Estoque 
         Appearance      =   0  'Flat
         Caption         =   "Não imprimir produtos que não tenham estoque"
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
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2190
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox O_Classes 
         Appearance      =   0  'Flat
         Caption         =   "Relatório separado por classes e sub classes"
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
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   2745
         Width           =   3795
      End
      Begin VB.CheckBox optInativo 
         Appearance      =   0  'Flat
         Caption         =   "Não imprimir produtos inativos"
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2460
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6120
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Parâmetro"
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox Lista 
         Appearance      =   0  'Flat
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
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   1800
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Cli_For ORDER BY Nome"
         Top             =   1260
         Visible         =   0   'False
         Width           =   1815
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
         Bindings        =   "RegInvent.frx":4F23C
         DataSource      =   "Data2"
         Height          =   315
         Left            =   1095
         TabIndex        =   2
         Top             =   1260
         Width           =   1800
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   9313
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2037
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
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
      Begin SSDataWidgets_B.SSDBCombo Combo 
         Bindings        =   "RegInvent.frx":4F250
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1095
         TabIndex        =   1
         Top             =   825
         Width           =   1800
         DataFieldList   =   "Filial"
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
         Columns.Count   =   2
         Columns(0).Width=   8414
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
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Filial"
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   6285
         TabIndex        =   29
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   225
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Até"
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
         Index           =   8
         Left            =   5910
         TabIndex        =   30
         Top             =   277
         Width           =   300
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   8040
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Nome_Empresa 
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
         Height          =   315
         Left            =   2970
         TabIndex        =   27
         Top             =   825
         Width           =   4935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Filial"
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
         Height          =   285
         Left            =   135
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Tabela"
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
         Height          =   285
         Left            =   135
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Fornecedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   1275
         Width           =   960
      End
      Begin VB.Label Nome_Fornecedor 
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
         Height          =   315
         Left            =   2970
         TabIndex        =   23
         Top             =   1260
         Width           =   4935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   -120
      TabIndex        =   19
      Top             =   -120
      Width           =   8535
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   240
         Picture         =   "RegInvent.frx":4F264
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Registro de inventário"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   21
         Top             =   210
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"RegInvent.frx":510CC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   675
         Left            =   1980
         TabIndex        =   20
         Top             =   600
         Width           =   6015
      End
   End
   Begin ComctlLib.ProgressBar pgbStatus 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   8175
   End
End
Attribute VB_Name = "frmRegInvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private contador_arrayClasses     As Long
Private contador_arraySubClasses  As Long
Private contador_arrayTamanho     As Long
Private contador_arrayCor         As Long
Private contador_arrayPrecos      As Long
Dim arrayClasses()                As Variant
Dim arraySubClasses()             As Variant
Dim arrayTamanho()                As Variant
Dim arrayCor()                    As Variant
Dim arrayPrecos()                 As Variant

'13/06/2005 - Daniel
'Variável modular que identificará o cliente Irmãos Ambrózio (Nyron)
'para o devido tratamento do campo Classificação Fiscal
Dim m_blnAmbrozio As Boolean
Dim rstProdutos2 As Recordset

Private Sub B_Imprime_Click()
  Dim rstInventario     As Recordset
  Dim rstEstoqueFinal   As Recordset
  Dim rstProdutos       As Recordset
  Dim rstClasse         As Recordset
  Dim rstSubClasse      As Recordset
  Dim rstTamanhos       As Recordset
  Dim rstCor            As Recordset
  Dim rstPrecos         As Recordset
  
  Dim strCodigoProduto    As String
  Dim strNomeProduto      As String
  Dim strCodigoOrdenacao  As String
  Dim blnFracionado       As Boolean
  Dim strUnidadeVenda     As String
  Dim blnInativo          As Boolean
  Dim intCodigoClasse     As String
  Dim strNomeClasse       As String
  Dim intCodigoSubClasse  As Integer
  Dim strNomeSubClasse    As String
  Dim intCodigoTamanho    As Integer
  Dim strNomeTamanho      As String
  Dim intCodigoCor        As Integer
  Dim strNomeCor          As String
  Dim strCodigoEdicao     As String
  Dim strNomeEdicao       As String
  Dim dblPreco            As Double
  
  Dim blnInTransaction    As Boolean
  Dim blnGeraRegistro     As Boolean
  Dim strSQL              As String
  
  
  If Data_Fim.Text <> "  /  /    " Then
      Dim retMsg As Variant
      retMsg = MsgBox("Este procedimento pode ser BEM DEMORADO. Sugerimos fazê-lo fora do horário comercial e quando nenhum outro usuário estiver usando o QuickStore. Deseja prosseguir agora com este procedimento?", vbYesNo, "Registro de Inventário por data limite")
      
      If retMsg = vbNo Then
          Exit Sub
      Else
          retMsg = MsgBox("Tem certeza que deseja prosseguir agora com este procedimento?", vbYesNo, "Registro de Inventário por data limite")
          
          If retMsg = vbNo Then
              Exit Sub
          End If
      End If
  End If
  
  If Len(Trim(Lista.Text)) <= 0 Then
    MsgBox "Tabela de preço inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Len(Trim(Nome_Empresa.Caption)) <= 0 Then
    MsgBox "Você deve escolher uma filial !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  On Error GoTo Erro:
  
  Call StatusMsg("Gerando o arquivo temporário . . . ")
  dbTemp.Execute "DELETE * FROM Inventário"
  
  '==============================================================
  'Carga em array de TABELA DE PREÇO, conforme tabela selecionada
  Dim lContador As Long
  Dim rsPrecos  As Recordset
  Dim sPrecoAux As String
  
  lContador = 0
  Set rsPrecos = db.OpenRecordset("select * From Preços WHERE Tabela = '" & Lista.Text & "' ", dbOpenDynaset)
  
  If Not (rsPrecos.EOF And rsPrecos.BOF) Then
      rsPrecos.MoveLast
      rsPrecos.MoveFirst
      
      ReDim arrayPrecos(rsPrecos.RecordCount, 2)
      contador_arrayPrecos = rsPrecos.RecordCount
      While Not rsPrecos.EOF
          arrayPrecos(lContador, 0) = rsPrecos.Fields("Produto").Value
          If Not IsNull(rsPrecos.Fields("Preço").Value) Then
              arrayPrecos(lContador, 1) = rsPrecos.Fields("Preço").Value
          Else
              arrayPrecos(lContador, 1) = ""
          End If
          lContador = lContador + 1
          rsPrecos.MoveNext
      Wend
  End If
  rsPrecos.Close
  Set rsPrecos = Nothing
  
  '==============================================================
  
  
  
  
  ws.BeginTrans
  blnInTransaction = True
  
  If Data_Fim.Text = "  /  /    " Then
      Set rstEstoqueFinal = db.OpenRecordset("SELECT * FROM [Estoque Final] WHERE Filial = " & Combo.Text, dbOpenSnapshot)
  
      With rstEstoqueFinal
        If Not (.BOF And .EOF) Then
          .MoveLast
          .MoveFirst
          
          pgbStatus.Max = .RecordCount + 1
          pgbStatus.Value = 0
          
          Do While Not .EOF
            blnGeraRegistro = True
            pgbStatus.Value = .AbsolutePosition
            
            '---[ Informações sobre o produto ]---'
              strCodigoProduto = .Fields("Produto")
              intCodigoTamanho = .Fields("Tamanho")
              intCodigoCor = .Fields("Cor")
              strCodigoEdicao = .Fields("Edição")
              
              If Len(Trim(Nome_Fornecedor.Caption)) > 0 Then
                strSQL = " SELECT Produtos.Código, Produtos.Nome, Produtos.[Código Ordenação], Produtos.[Unidade Venda], Produtos.Classe, Produtos.[Sub Classe], Produtos.Fracionado, Produtos.[Desativado], Forn_Prod.Fornecedor " & _
                         " FROM Produtos INNER JOIN Forn_Prod ON Produtos.Código = Forn_Prod.Produto " & _
                         " WHERE Produtos.Código = '" & strCodigoProduto & "' AND Forn_Prod.Fornecedor = " & CLng(Combo_Fornecedor.Text)
              Else
                strSQL = " SELECT Código, Nome, [Código Ordenação], [Unidade Venda], Classe, [Sub Classe], Fracionado, Desativado FROM Produtos " & _
                         " WHERE Código = '" & strCodigoProduto & "'"
              End If
                                  
              If O_Grade.Value Then
                strSQL = strSQL & " AND Tipo = 'G' "
              ElseIf O_Edições.Value Then
                strSQL = strSQL & " AND Tipo = 'E' "
              End If
              
              Set rstProdutos = db.OpenRecordset(strSQL, dbOpenSnapshot)
              
              If Not (rstProdutos.BOF And rstProdutos.EOF) Then
                strNomeProduto = rstProdutos.Fields("Nome") & ""
                strCodigoOrdenacao = rstProdutos.Fields("Código Ordenação") & ""
                blnFracionado = rstProdutos.Fields("Fracionado")
                strUnidadeVenda = rstProdutos.Fields("Unidade Venda") & ""
                intCodigoClasse = rstProdutos.Fields("Classe") & ""
                
                '20/01/2005 - Daniel
                'BUG: Run-time error '94' invalid use of Null
                '
                'Na base do cliente PIETRA BELLA C.BIJOUTERIAS - SP
                'estava aparecendo alguns produtos com a sub classe Null
                'isto gerava o erro mencionado acima, para corrigir
                'criamos o tratamento abaixo:
                If Not IsNull(rstProdutos.Fields("Sub Classe")) Then
                  intCodigoSubClasse = rstProdutos.Fields("Sub Classe")
                Else
                  intCodigoSubClasse = 0
                End If
                '-------------------------------------------------------
                
                blnInativo = rstProdutos.Fields("Desativado")
                
                '---[ Informações sobre a classe ]---'
                  Set rstClasse = db.OpenRecordset("SELECT * FROM Classes WHERE Código = " & intCodigoClasse, dbOpenSnapshot)
                  
                  With rstClasse
                    If Not (.BOF And .EOF) Then
                      strNomeClasse = .Fields("Nome") & ""
                    Else
                      strNomeClasse = "<Classe_sem_nome>"
                    End If
                    If Not rstClasse Is Nothing Then .Close
                    Set rstClasse = Nothing
                  End With
                '---[ Informações sobre a classe ]---'
                
                '---[ Informações sobre a sub classe ]---'
                  Set rstSubClasse = db.OpenRecordset("SELECT * FROM [Sub Classes] WHERE Código = " & intCodigoSubClasse, dbOpenSnapshot)
                  
                  With rstSubClasse
                    If Not (.BOF And .EOF) Then
                      strNomeSubClasse = .Fields("Nome") & ""
                    Else
                      strNomeSubClasse = "<Sub_Classe_sem_nome>"
                    End If
                    If Not rstSubClasse Is Nothing Then .Close
                    Set rstSubClasse = Nothing
                  End With
                '---[ Informações sobre a sub classe ]---'
                
                '---[ Informações sobre o tamanho ]---'
                  Set rstTamanhos = db.OpenRecordset("SELECT * FROM Tamanhos WHERE Código = " & intCodigoTamanho, dbOpenSnapshot)
                  
                  With rstTamanhos
                    If Not (.BOF And .EOF) Then
                      strNomeTamanho = .Fields("Nome") & ""
                    Else
                      strNomeTamanho = "<Produto_sem_tamanho>"
                    End If
                    If Not rstTamanhos Is Nothing Then .Close
                    Set rstTamanhos = Nothing
                  End With
                '---[ Informações sobre o tamanho ]---'
                
                '---[ Informações sobre a cor ]---'
                  Set rstCor = db.OpenRecordset("SELECT * FROM Cores WHERE Código = " & intCodigoCor, dbOpenSnapshot)
                  
                  With rstCor
                    If Not (.BOF And .EOF) Then
                      strNomeCor = .Fields("Nome") & ""
                    Else
                      strNomeCor = "<produto_sem_cor>"
                    End If
                    If Not rstCor Is Nothing Then .Close
                    Set rstCor = Nothing
                  End With
                '---[ Informações sobre a cor ]---'
                
                Set rstPrecos = db.OpenRecordset("SELECT * FROM Preços WHERE Tabela = '" & Lista.Text & "' AND Produto = '" & .Fields("Produto") & "'", dbOpenSnapshot)
                
                With rstPrecos
                  If Not (.BOF And .EOF) Then
                    dblPreco = .Fields("Preço")
                  Else
                    dblPreco = 0
                  End If
                  
                  If Not rstPrecos Is Nothing Then .Close
                  Set rstPrecos = Nothing
                End With
                
                If optInativo.Value = vbChecked Then
                  If blnInativo Then blnGeraRegistro = False
                End If
                
                If blnGeraRegistro Then
                  If O_Preço.Value = vbChecked Then
                    If dblPreco <= 0 Then blnGeraRegistro = False
                  End If
                End If
                
                If blnGeraRegistro Then
                  If O_Estoque.Value = vbChecked Then
                    If .Fields("Estoque Atual") <= 0 Then blnGeraRegistro = False
                  End If
                End If
                
                If blnGeraRegistro Then
                  If O_Normal.Value Then
                    Set rstInventario = dbTemp.OpenRecordset(" SELECT * FROM Inventário WHERE " & _
                                                             " Empresa = " & .Fields("Filial") & _
                                                             " AND Produto = '" & strCodigoProduto & "'", dbOpenDynaset)
                                                             
                    If (rstInventario.BOF And rstInventario.EOF) Then
                      rstInventario.AddNew
                      
                      rstInventario.Fields("Empresa") = .Fields("Filial")
                      rstInventario.Fields("Produto") = strCodigoProduto
                      '13/06/2005 - Daniel
                      'Tratamento do campo Classificação Fiscal
                      'Solicitante: Irmãos Ambrózio
                      If m_blnAmbrozio Then rstInventario.Fields("Classificação Fiscal").Value = GetClassificacaoFiscal(strCodigoProduto) & ""
                      rstInventario.Fields("Ordenação") = strCodigoOrdenacao
                      rstInventario.Fields("Nome") = strNomeProduto
                      rstInventario.Fields("Unidade Venda") = strUnidadeVenda
                      rstInventario.Fields("Fracionado") = blnFracionado
                      rstInventario.Fields("Classe") = intCodigoClasse
                      rstInventario.Fields("Nome Classe") = strNomeClasse
                      rstInventario.Fields("Sub Classe") = intCodigoSubClasse
                      rstInventario.Fields("Nome Sub") = strNomeSubClasse
                      
                      rstInventario.Fields("Tamanho") = intCodigoTamanho
                      rstInventario.Fields("Nome Tamanho") = strNomeTamanho
                      
                      rstInventario.Fields("Cor") = intCodigoCor
                      rstInventario.Fields("Nome Cor") = strNomeCor
                      
                      rstInventario.Fields("Edição") = strCodigoEdicao
                      rstInventario.Fields("Nome Edição") = "-"
                      
                      rstInventario.Fields("Data") = Replace(.Fields("Última Data"), ".", "/")
                      rstInventario.Fields("Preço") = dblPreco
                      rstInventario.Fields("Estoque Final") = .Fields("Estoque Atual")
                    Else
                      rstInventario.Edit
                      rstInventario.Fields("Estoque Final") = rstInventario.Fields("Estoque Final") + .Fields("Estoque Atual")
                    End If
                  ElseIf O_Grade.Value Then
                    Set rstInventario = dbTemp.OpenRecordset(" SELECT * FROM Inventário ", dbOpenDynaset)
                  
                    rstInventario.AddNew
                    
                    rstInventario.Fields("Empresa") = .Fields("Filial")
                    rstInventario.Fields("Produto") = strCodigoProduto
                    '13/06/2005 - Daniel
                    'Tratamento do campo Classificação Fiscal
                    'Solicitante: Irmãos Ambrózio
                    If m_blnAmbrozio Then rstInventario.Fields("Classificação Fiscal").Value = GetClassificacaoFiscal(strCodigoProduto) & ""
                    rstInventario.Fields("Ordenação") = strCodigoOrdenacao
                    rstInventario.Fields("Nome") = strNomeProduto
                    rstInventario.Fields("Unidade Venda") = strUnidadeVenda
                    rstInventario.Fields("Fracionado") = blnFracionado
                    rstInventario.Fields("Classe") = intCodigoClasse
                    rstInventario.Fields("Nome Classe") = strNomeClasse
                    rstInventario.Fields("Sub Classe") = intCodigoSubClasse
                    rstInventario.Fields("Nome Sub") = strNomeSubClasse
                    
                    rstInventario.Fields("Tamanho") = intCodigoTamanho
                    rstInventario.Fields("Nome Tamanho") = strNomeTamanho
                    
                    rstInventario.Fields("Cor") = intCodigoCor
                    rstInventario.Fields("Nome Cor") = strNomeCor
                    
                    rstInventario.Fields("Edição") = strCodigoEdicao
                    rstInventario.Fields("Nome Edição") = "-"
                    
                    rstInventario.Fields("Data") = Replace(.Fields("Última Data"), ".", "/")
                    rstInventario.Fields("Preço") = dblPreco
                    rstInventario.Fields("Estoque Final") = .Fields("Estoque Atual")
                  End If
                  
                  rstInventario.Update
                  rstInventario.Close
                  Set rstInventario = Nothing
                End If
              End If
              
              If Not rstProdutos Is Nothing Then rstProdutos.Close
              Set rstProdutos = Nothing
            '---[ Informações sobre o produto ]---'
            
            .MoveNext
          Loop
        End If
      End With
  Else
  
      '=====================================================================================
      ' Posição de estoque na data específica
  
      ' ======================
      ' Selecionar os produtos
      
      strSQL = "SELECT * FROM Produtos "
      
      If optInativo.Value = vbChecked Then
          strSQL = strSQL & " Where Desativado = 0 "
      End If
      
      Set rstProdutos2 = db.OpenRecordset(strSQL, dbOpenSnapshot)
      ' ======================
      
      ' ======================
      ' Selecionar o estoque detalhadao
      strSQL = "SELECT * "
      strSQL = strSQL & " FROM Estoque WHERE Filial = " & Combo.Text
      strSQL = strSQL & " AND Data <= #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "# "
      
      ''If O_Estoque.Value = vbChecked Then
          ''strSQL = strSQL & " AND [Estoque Final] > 0 "
      ''End If
      
      strSQL = strSQL & " Order by Data Desc "
      Set rstEstoqueFinal = db.OpenRecordset(strSQL, dbOpenSnapshot)
      ' =====================
      
      If Not (rstProdutos2.BOF And rstProdutos2.EOF) Then
          rstProdutos2.MoveLast
          rstProdutos2.MoveFirst
  
          While Not rstProdutos2.EOF
  
              rstEstoqueFinal.MoveFirst
              rstEstoqueFinal.FindFirst "Produto = '" & rstProdutos2.Fields("Código").Value & "'"
      
              If rstEstoqueFinal.Fields("Estoque Final").Value > 0 Then
      
              With rstEstoqueFinal
                If Not rstEstoqueFinal.NoMatch Then
                    pgbStatus.Max = .RecordCount + 1
                    pgbStatus.Value = 0
                    
                    blnGeraRegistro = True
                    pgbStatus.Value = .AbsolutePosition
                    
                    '---[ Informações sobre o produto ]---'
                    strCodigoProduto = .Fields("Produto")
                    intCodigoTamanho = .Fields("Tamanho")
                    intCodigoCor = .Fields("Cor")
                    strCodigoEdicao = .Fields("Edição")
                    
                    ' IDEAL É FAZER A BUSCA DOS PRODUTOS POR FORNECEDOR NO SELECT DE PRODUTOS ACIMA...NÃO AQUI...
                    If Len(Trim(Nome_Fornecedor.Caption)) > 0 Then
                      strSQL = " SELECT Produtos.Código, Produtos.Nome, Produtos.[Código Ordenação], Produtos.[Unidade Venda], Produtos.Classe, Produtos.[Sub Classe], Produtos.Fracionado, Produtos.[Desativado], Forn_Prod.Fornecedor " & _
                               " FROM Produtos INNER JOIN Forn_Prod ON Produtos.Código = Forn_Prod.Produto " & _
                               " WHERE Produtos.Código = '" & strCodigoProduto & "' AND Forn_Prod.Fornecedor = " & CLng(Combo_Fornecedor.Text)
                    Else
                      strSQL = " SELECT Código, Nome, [Código Ordenação], [Unidade Venda], Classe, [Sub Classe], Fracionado, Desativado FROM Produtos " & _
                               " WHERE Código = '" & strCodigoProduto & "'"
                    End If
                                        
                    If Len(Trim(Nome_Fornecedor.Caption)) > 0 Then
                        If O_Grade.Value Then
                          strSQL = strSQL & " AND Tipo = 'G' "
                        ElseIf O_Edições.Value Then
                          strSQL = strSQL & " AND Tipo = 'E' "
                        End If
    
                        Set rstProdutos = db.OpenRecordset(strSQL, dbOpenSnapshot)

                        If Not (rstProdutos.BOF And rstProdutos.EOF) Then
                            strNomeProduto = rstProdutos.Fields("Nome") & ""
                            strCodigoOrdenacao = rstProdutos.Fields("Código Ordenação") & ""
                            blnFracionado = rstProdutos.Fields("Fracionado")
                            strUnidadeVenda = rstProdutos.Fields("Unidade Venda") & ""
                            intCodigoClasse = rstProdutos.Fields("Classe") & ""
                            
                            'Se alguns produtos com a sub classe Null...segue o tratamento abaixo:
                            If Not IsNull(rstProdutos.Fields("Sub Classe")) Then
                              intCodigoSubClasse = rstProdutos.Fields("Sub Classe")
                            Else
                              intCodigoSubClasse = 0
                            End If
                          
                            blnInativo = rstProdutos.Fields("Desativado")
                        End If
                        
                        If Not rstProdutos Is Nothing Then rstProdutos.Close
                        Set rstProdutos = Nothing
                        
                    Else
                        strNomeProduto = rstProdutos2.Fields("Nome") & ""
                        strCodigoOrdenacao = rstProdutos2.Fields("Código Ordenação") & ""
                        blnFracionado = rstProdutos2.Fields("Fracionado")
                        strUnidadeVenda = rstProdutos2.Fields("Unidade Venda") & ""
                        intCodigoClasse = rstProdutos2.Fields("Classe") & ""
                        
                        'Se alguns produtos com a sub classe Null...segue o tratamento abaixo:
                        If Not IsNull(rstProdutos2.Fields("Sub Classe")) Then
                          intCodigoSubClasse = rstProdutos2.Fields("Sub Classe")
                        Else
                          intCodigoSubClasse = 0
                        End If
                      
                        blnInativo = rstProdutos2.Fields("Desativado")
                    End If
                        
                    If O_Classes.Value = vbChecked Then
                        ' Nome da Classe e SubClasse do produto
                        strNomeClasse = AcharClasse(intCodigoClasse)
                        strNomeSubClasse = AcharSubClasse(intCodigoSubClasse)
                    Else
                        strNomeClasse = ""
                        strNomeSubClasse = ""
                    End If
                    
                    If O_Grade.Value = True Then
                        ' Nome do tamanho e cor
                        strNomeTamanho = AcharTamanho(intCodigoTamanho)
                        strNomeCor = AcharCor(intCodigoCor)
                    Else
                        strNomeTamanho = ""
                        strNomeCor = ""
                    End If

                    ' ======================================
                    ' Achar preço pela tabela selecionada
                    sPrecoAux = ""
                    sPrecoAux = AcharPreco(.Fields("Produto"))
                    
                    If sPrecoAux <> "" And sPrecoAux <> "0" Then
                      dblPreco = sPrecoAux
                    Else
                      dblPreco = 0
                    End If
                    ' ======================================
                  
                    If optInativo.Value = vbChecked Then
                        If blnInativo Then blnGeraRegistro = False
                    End If
                      
                    If blnGeraRegistro Then
                        If O_Preço.Value = vbChecked Then
                          If dblPreco <= 0 Then blnGeraRegistro = False
                        End If
                    End If
                  
                    If blnGeraRegistro Then
                        If O_Estoque.Value = vbChecked Then
                            If .Fields("Estoque Final") <= 0 Then blnGeraRegistro = False
                        End If
                    End If
                      
                    If blnGeraRegistro Then
                        If O_Normal.Value Then
                            Set rstInventario = dbTemp.OpenRecordset(" SELECT * FROM Inventário WHERE " & _
                                                               " Empresa = " & .Fields("Filial") & _
                                                               " AND Produto = '" & strCodigoProduto & "'", dbOpenDynaset)
                                                               
                            If (rstInventario.BOF And rstInventario.EOF) Then
                                rstInventario.AddNew
                                
                                rstInventario.Fields("Empresa") = .Fields("Filial")
                                rstInventario.Fields("Produto") = strCodigoProduto
                                '13/06/2005 - Daniel
                                'Tratamento do campo Classificação Fiscal
                                'Solicitante: Irmãos Ambrózio
                                If m_blnAmbrozio Then rstInventario.Fields("Classificação Fiscal").Value = GetClassificacaoFiscal(strCodigoProduto) & ""
                                rstInventario.Fields("Ordenação") = strCodigoOrdenacao
                                rstInventario.Fields("Nome") = strNomeProduto
                                rstInventario.Fields("Unidade Venda") = strUnidadeVenda
                                rstInventario.Fields("Fracionado") = blnFracionado
                                rstInventario.Fields("Classe") = intCodigoClasse
                                rstInventario.Fields("Nome Classe") = strNomeClasse
                                rstInventario.Fields("Sub Classe") = intCodigoSubClasse
                                rstInventario.Fields("Nome Sub") = strNomeSubClasse
                                
                                rstInventario.Fields("Tamanho") = intCodigoTamanho
                                rstInventario.Fields("Nome Tamanho") = strNomeTamanho
                                
                                rstInventario.Fields("Cor") = intCodigoCor
                                rstInventario.Fields("Nome Cor") = strNomeCor
                                
                                rstInventario.Fields("Edição") = strCodigoEdicao
                                rstInventario.Fields("Nome Edição") = "-"
                                
                                rstInventario.Fields("Data") = Replace(.Fields("Data"), ".", "/")
                                rstInventario.Fields("Preço") = dblPreco
                                rstInventario.Fields("Estoque Final") = .Fields("Estoque Final")
                            Else
                                rstInventario.Edit
                                rstInventario.Fields("Estoque Final") = rstInventario.Fields("Estoque Final") + .Fields("Estoque Final")
                            End If
                        ElseIf O_Grade.Value Then
                            Set rstInventario = dbTemp.OpenRecordset(" SELECT * FROM Inventário ", dbOpenDynaset)
                          
                            rstInventario.AddNew
                            
                            rstInventario.Fields("Empresa") = .Fields("Filial")
                            rstInventario.Fields("Produto") = strCodigoProduto
                            '13/06/2005 - Daniel
                            'Tratamento do campo Classificação Fiscal
                            'Solicitante: Irmãos Ambrózio
                            If m_blnAmbrozio Then rstInventario.Fields("Classificação Fiscal").Value = GetClassificacaoFiscal(strCodigoProduto) & ""
                            rstInventario.Fields("Ordenação") = strCodigoOrdenacao
                            rstInventario.Fields("Nome") = strNomeProduto
                            rstInventario.Fields("Unidade Venda") = strUnidadeVenda
                            rstInventario.Fields("Fracionado") = blnFracionado
                            rstInventario.Fields("Classe") = intCodigoClasse
                            rstInventario.Fields("Nome Classe") = strNomeClasse
                            rstInventario.Fields("Sub Classe") = intCodigoSubClasse
                            rstInventario.Fields("Nome Sub") = strNomeSubClasse
                            
                            rstInventario.Fields("Tamanho") = intCodigoTamanho
                            rstInventario.Fields("Nome Tamanho") = strNomeTamanho
                            
                            rstInventario.Fields("Cor") = intCodigoCor
                            rstInventario.Fields("Nome Cor") = strNomeCor
                            
                            rstInventario.Fields("Edição") = strCodigoEdicao
                            rstInventario.Fields("Nome Edição") = "-"
                            
                            rstInventario.Fields("Data") = Replace(.Fields("Data"), ".", "/")
                            rstInventario.Fields("Preço") = dblPreco
                            rstInventario.Fields("Estoque Final") = .Fields("Estoque Final")
                        End If
                    
                        rstInventario.Update
                        rstInventario.Close
                        Set rstInventario = Nothing
                    End If
                   
                  '.MoveNext
                End If
              End With
              
              End If
              
              rstProdutos2.MoveNext
          Wend
      End If
  
      rstProdutos2.Close
      Set rstProdutos2 = Nothing
      rstEstoqueFinal.Close
      Set rstEstoqueFinal = Nothing
      
  '=========================================================================
      
  End If
  
  
  ws.CommitTrans
  blnInTransaction = False
  
  PrintReport
  
  Exit Sub
  
Erro:
  If blnInTransaction Then
    ws.Rollback
    blnInTransaction = False
  End If
  
  pgbStatus.Value = 0
   
  '19/09/2005 - mpdea
  'Exibe a mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub


Private Function AcharClasse(ByVal iCodigoClasse As Integer) As String
  Dim lContaReg As Long
  
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 0) = iCodigoClasse Then
          AcharClasse = arrayClasses(lContaReg, 1)
          Exit Function
      End If
  Next
  
  AcharClasse = ""
End Function


Private Function AcharSubClasse(ByVal iCodigoSubClasse As Integer) As String
  Dim lContaReg As Long
  
  For lContaReg = 0 To contador_arraySubClasses - 1
      If arraySubClasses(lContaReg, 0) = iCodigoSubClasse Then
          AcharSubClasse = arraySubClasses(lContaReg, 1)
          Exit Function
      End If
  Next
  
  AcharSubClasse = ""
End Function

Private Function AcharTamanho(ByVal iCodigoTamanho As Integer) As String
  Dim lContaReg As Long
  
  For lContaReg = 0 To contador_arrayTamanho - 1
      If arrayTamanho(lContaReg, 0) = iCodigoTamanho Then
          AcharTamanho = arrayTamanho(lContaReg, 1)
          Exit Function
      End If
  Next
  
  AcharTamanho = ""
End Function

Private Function AcharCor(ByVal iCodigoCor As Integer) As String
  Dim lContaReg As Long
  
  For lContaReg = 0 To contador_arrayCor - 1
      If arrayCor(lContaReg, 0) = iCodigoCor Then
          AcharCor = arrayCor(lContaReg, 1)
          Exit Function
      End If
  Next
  
  AcharCor = ""
End Function


Private Function AcharPreco(sCodigoProduto As String) As String
  Dim lContaReg As Long
  
  For lContaReg = 0 To contador_arrayPrecos - 1
      If arrayPrecos(lContaReg, 0) = sCodigoProduto Then
          AcharPreco = arrayPrecos(lContaReg, 1)
          Exit Function
      End If
  Next
  
  AcharPreco = ""
End Function


Private Sub chk_dataAte_Click()
    If chk_dataAte.Value = vbChecked Then
        Data_Fim.Text = "  /  /    "
        Data_Fim.Visible = True
        cmd_calendarioDtFim.Visible = True
        
        Label4.Visible = False
        Combo_Fornecedor.Text = ""
        Combo_Fornecedor.Visible = False
        Nome_Fornecedor.Caption = ""
        Nome_Fornecedor.Visible = False
    Else
    
        Data_Fim.Visible = False
        cmd_calendarioDtFim.Visible = False
        
        Label4.Visible = True
        Combo_Fornecedor.Visible = True
        Nome_Fornecedor.Visible = True
    End If
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub Combo_CloseUp()
  Combo_LostFocus
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
  Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_LostFocus()
  Dim rstFornecedores As Recordset
  
  Nome_Fornecedor.Caption = ""
  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
  
  Set rstFornecedores = db.OpenRecordset("SELECT Nome FROM Cli_For WHERE Código = " & Combo_Fornecedor.Text, dbOpenSnapshot)
  
  With rstFornecedores
    If Not (.BOF And .EOF) Then
      Nome_Fornecedor.Caption = .Fields("Nome") & ""
    End If
    
    If Not rstFornecedores Is Nothing Then .Close
    Set rstFornecedores = Nothing
  End With
End Sub

Private Sub Combo_LostFocus()
  Dim rstParametros As Recordset
  
  Nome_Empresa.Caption = ""
  If Not IsNumeric(Combo.Text) Then Exit Sub
  
  Set rstParametros = db.OpenRecordset("SELECT Nome FROM [Parâmetros Filial] WHERE Filial = " & Combo.Text, dbOpenSnapshot)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      Nome_Empresa.Caption = .Fields("Nome") & ""
    End If
    
    If Not rstParametros Is Nothing Then .Close
    Set rstParametros = Nothing
  End With
End Sub

Private Sub Form_Load()
  Dim rstPrecos         As Recordset
  Dim lContador         As Long
  Dim rsClasse          As Recordset
  Dim rsSubClasse       As Recordset
  Dim rsTamanho         As Recordset
  Dim rsCor             As Recordset
  
  Call CenterForm(Me)
  
  
  lContador = 0
  Set rsClasse = db.OpenRecordset("select Código, Nome from Classes order by Nome", dbOpenDynaset)
  If Not (rsClasse.EOF And rsClasse.BOF) Then
      rsClasse.MoveLast
      rsClasse.MoveFirst
      
      ReDim arrayClasses(rsClasse.RecordCount, 2)
      contador_arrayClasses = rsClasse.RecordCount
      While Not rsClasse.EOF
          arrayClasses(lContador, 0) = rsClasse.Fields(0).Value
          If Not IsNull(rsClasse.Fields(1).Value) Then
              arrayClasses(lContador, 1) = rsClasse.Fields(1).Value
          Else
              arrayClasses(lContador, 1) = ""
          End If
          
          lContador = lContador + 1
          rsClasse.MoveNext
      Wend
  End If
  rsClasse.Close
  Set rsClasse = Nothing


  lContador = 0
  Set rsSubClasse = db.OpenRecordset("select Código, Nome from [Sub Classes] order by Nome", dbOpenDynaset)
  If Not (rsSubClasse.EOF And rsSubClasse.BOF) Then
      rsSubClasse.MoveLast
      rsSubClasse.MoveFirst
      
      ReDim arraySubClasses(rsSubClasse.RecordCount, 2)
      contador_arraySubClasses = rsSubClasse.RecordCount
      While Not rsSubClasse.EOF
          arraySubClasses(lContador, 0) = rsSubClasse.Fields(0).Value
          If Not IsNull(rsSubClasse.Fields(1).Value) Then
              arraySubClasses(lContador, 1) = rsSubClasse.Fields(1).Value
          Else
              arraySubClasses(lContador, 1) = ""
          End If
          lContador = lContador + 1
          rsSubClasse.MoveNext
      Wend
  End If
  rsSubClasse.Close
  Set rsSubClasse = Nothing
  
  
  lContador = 0
  Set rsTamanho = db.OpenRecordset("select Código, Nome from Tamanhos order by Nome", dbOpenDynaset)
  If Not (rsTamanho.EOF And rsTamanho.BOF) Then
      rsTamanho.MoveLast
      rsTamanho.MoveFirst
      
      ReDim arrayTamanho(rsTamanho.RecordCount, 2)
      contador_arrayTamanho = rsTamanho.RecordCount
      While Not rsTamanho.EOF
          arrayTamanho(lContador, 0) = rsTamanho.Fields(0).Value
          If Not IsNull(rsTamanho.Fields(1).Value) Then
              arrayTamanho(lContador, 1) = rsTamanho.Fields(1).Value
          Else
              arrayTamanho(lContador, 1) = ""
          End If
          lContador = lContador + 1
          rsTamanho.MoveNext
      Wend
  End If
  rsTamanho.Close
  Set rsTamanho = Nothing
  
  lContador = 0
  Set rsCor = db.OpenRecordset("select Código, Nome from Cores order by Nome", dbOpenDynaset)
  If Not (rsCor.EOF And rsCor.BOF) Then
      rsCor.MoveLast
      rsCor.MoveFirst
      
      ReDim arrayCor(rsCor.RecordCount, 2)
      contador_arrayCor = rsCor.RecordCount
      While Not rsCor.EOF
          arrayCor(lContador, 0) = rsCor.Fields(0).Value
          If Not IsNull(rsCor.Fields(1).Value) Then
              arrayCor(lContador, 1) = rsCor.Fields(1).Value
          Else
              arrayCor(lContador, 1) = ""
          End If
          lContador = lContador + 1
          rsCor.MoveNext
      Wend
  End If
  rsCor.Close
  Set rsCor = Nothing
  
  '13/06/2005 - Daniel
  'Variável modular que tratará o campo Classificação Fiscal
  'para Irmãos Ambrózio
  m_blnAmbrozio = CheckSerialCaseMod("QS35288-570", "QS36824-735")
  
  Set rstPrecos = db.OpenRecordset("SELECT DISTINCT Tabela FROM Preços", dbOpenSnapshot)
  
  '---[ Preenche a combo de preços ]---'
    With rstPrecos
      Lista.Clear
      
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do While Not .EOF
          Lista.AddItem .Fields("Tabela") & ""
          .MoveNext
        Loop
      End If
      
      If Not rstPrecos Is Nothing Then .Close
      Set rstPrecos = Nothing
    End With
  '---[ Preenche a combo de preços ]---'
  
  O_Grade.Enabled = gbGrade
  O_Edições.Enabled = gbEdicao
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  chk_dataAte.Value = 1
End Sub

Private Sub PrintReport()
  Dim Str1    As String
  Dim Str_Rel As String
  
  Call StatusMsg("")
  '31/10/2002 - mpdea
  'Corrigido associação com a localização das bases de dados
  With Rel
    .Reset
    .DataFiles(0) = gsTempDBFileName
    If Not O_Grade.Value Then .DataFiles(1) = gsTempDBFileName
    '29/11/2002 - mpdea
    'Associação necessária somente para relatórios com produto normal
    If O_Normal.Value = True Then .DataFiles(2) = gsQuickDBFileName
  End With
  
  Rem Saída
  If B_Vídeo = True Then Rel.Destination = crptToWindow
  If B_Impressora = True Then Rel.Destination = crptToPrinter
  
  Rel.WindowShowGroupTree = False
  If O_Classes.Value = 1 Then Rel.WindowShowGroupTree = True
  
  Rem Nome do arquivo .rpt
  If O_Classes.Value = 0 Then
    If O_Normal.Value = True Then
    
      '13/06/2005 - Daniel
      'O relatório da empresa Irmãos Ambrózio irá
      'mostrar de diferencial a coluna Class. Fiscal
      If m_blnAmbrozio Then
        Str1 = gsReportPath & "Invent1ComClassificacaoFiscal.RPT"
      Else
        Str1 = gsReportPath & "INVENT1.RPT"
      End If
    
    End If
    
    '06/12/2005 - mpdea
    'Código de seleção estava dentro da opção de relatório Normal
    'não sendo possível escolher por Grade ou Edição,
    'gerando o erro RT-20507
    If O_Grade.Value = True Then Str1 = gsReportPath & "INVENT1G.RPT"
    If O_Edições.Value = True Then Str1 = gsReportPath & "INVENT1E.RPT"
    
    
    If O_Código.Value = True Then
      Rel.SortFields(0) = "+{Inventário.Ordenação}"
      Rel.SortFields(1) = ""
      Rel.SortFields(2) = ""
    Else
      Rel.SortFields(0) = "+{Inventário.Nome}"
      Rel.SortFields(1) = ""
      Rel.SortFields(2) = ""
    End If
  End If
  
  If O_Classes.Value = 1 Then
    If O_Normal.Value = True Then Str1 = gsReportPath & "INVENT2.RPT"
    If O_Grade.Value = True Then Str1 = gsReportPath & "INVENT2G.RPT"
    If O_Edições.Value = True Then Str1 = gsReportPath & "INVENT2E.RPT"
    
    If O_Código.Value = True Then
      Rel.SortFields(0) = "+{Inventário.Classe}"
      Rel.SortFields(1) = "+{Inventário.Sub Classe}"
      Rel.SortFields(2) = "+{Inventário.Produto}"
    Else
      Rel.SortFields(0) = "+{Inventário.Classe}"
      Rel.SortFields(1) = "+{Inventário.Sub Classe}"
      Rel.SortFields(2) = "+{Inventário.Nome}"
    End If
  End If
  
  Rel.ReportFileName = Str1
  
  Rem Seleção
  
  Str_Rel = "{Inventário.Empresa} =" + Combo.Text
  ' Str_Rel = Str_Rel + " And {Caixa.Data} ="
  ' Str_Rel = Str_Rel + Str_Data1
  Rel.SelectionFormula = Str_Rel
  
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  
  Rel.Formulas(0) = Str_Rel
  
  Str_Rel = "nome_filial = '"
  Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
  Rel.Formulas(1) = Str_Rel
  
  Str_Rel = "tabela_preço = '"
  Str_Rel = Str_Rel + Lista.Text + "'"
  Rel.Formulas(2) = Str_Rel
  
  '06/05/2004 - Daniel
  'Caso seja Embalavi, formataremos o valor para
  '5 casas após a vírgula
  If g_bln5CasasDecimais Then
    If O_Normal.Value = True Then 'Normal..
      Rel.Formulas(3) = "QtdeCasasDecimais = " & "5"
    End If
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    If O_Normal.Value = True Then 'Normal..
      Rel.Formulas(3) = "QtdeCasasDecimais = " & "3"
    End If
  Else
    Rel.Formulas(3) = "QtdeCasasDecimais = " & "2"
  End If
  
  
  Call StatusMsg("Aguarde, imprimindo...")
  
  '26/10/2005 - mpdea
  'Exibe botão de configuração de impressão
  Rel.WindowShowPrintSetupBtn = True
  
  Rel.WindowState = crptMaximized
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  
  Rel.Action = 1
  
  Call StatusMsg("")
End Sub

Private Function GetClassificacaoFiscal(ByVal CodProd As String) As String
  '13/06/2005 - Daniel
  'Função que retornará a classificação fiscal (descrição da classificação) do produto
  'Solicitante: Irmãos Ambrózio (Nyron)
  'Finalidades: Atender às normas da nomeclatura brasileira de mercadoria / sistema harmonico
  'para o registro de inventário
  Dim rstProdutos    As Recordset
  Dim rstClassFiscal As Recordset
  
  GetClassificacaoFiscal = ""
  
  On Error GoTo TratarErro
  
  Set rstProdutos = db.OpenRecordset("SELECT [Classificação Fiscal] FROM Produtos WHERE Código = '" & CodProd & "'", dbOpenDynaset)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Set rstClassFiscal = db.OpenRecordset("SELECT Nome FROM [Classificação Fiscal] WHERE Código = " & (.Fields("Classificação Fiscal").Value), dbOpenDynaset)
      
      If Not (rstClassFiscal.BOF And rstClassFiscal.EOF) Then
        rstClassFiscal.MoveFirst
        'Solicitado 8 posições apenas...
        GetClassificacaoFiscal = Left(rstClassFiscal.Fields("Nome").Value, 8) & ""
        rstClassFiscal.Close
      End If
      
      Set rstClassFiscal = Nothing
      
    End If
    .Close
  End With
  
  Set rstProdutos = Nothing

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function
