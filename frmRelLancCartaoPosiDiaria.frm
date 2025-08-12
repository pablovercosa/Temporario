VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRelLancCartaoPosiDiaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Posição de Cartões"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelLancCartaoPosiDiaria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   14670
   Begin VB.Data dtaVendedor 
      Caption         =   "dtaVendedor"
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
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   7830
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmd_Detalhar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Detalhar =>"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4740
      Width           =   885
   End
   Begin VB.TextBox txt_totalLiquido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   6810
      TabIndex        =   31
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   6900
      Width           =   1605
   End
   Begin VB.TextBox txt_total 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   3420
      TabIndex        =   29
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   6900
      Width           =   1605
   End
   Begin VB.CommandButton cmd_pesquisarParaGrade 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Visualizar na Grade"
      Height          =   465
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2520
      Width           =   7095
   End
   Begin VB.Frame fraO 
      Caption         =   "Ordenação"
      Height          =   1800
      Left            =   9600
      TabIndex        =   23
      Top             =   660
      Width           =   2940
      Begin VB.OptionButton optNomeAdm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome da Administradora"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   300
         TabIndex        =   8
         Top             =   855
         Width           =   2325
      End
      Begin VB.OptionButton optValorBruto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Valor Bruto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   1245
         Width           =   1215
      End
      Begin VB.OptionButton optCodAdm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código da Administradora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   450
         Value           =   -1  'True
         Width           =   2385
      End
   End
   Begin VB.Data datCartoes 
      Caption         =   "datCartoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cartões ORDER BY Código"
      Top             =   7890
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data datCaixa 
      Caption         =   "datCaixa"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1980
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Caixa, Descrição FROM [Caixas em Uso] ORDER BY Caixa"
      Top             =   7860
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data datFilial 
      Caption         =   "datFilial"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   7860
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFechar 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   435
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7380
      Width           =   14655
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Gerar Relatório"
      Height          =   465
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   7365
   End
   Begin VB.Frame Frame4 
      Caption         =   "Saída"
      Height          =   1800
      Left            =   12570
      TabIndex        =   22
      Top             =   660
      Width           =   2070
      Begin VB.OptionButton optVideo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   5
         Top             =   450
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optImpressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período"
      Height          =   1800
      Left            =   6330
      TabIndex        =   21
      Top             =   660
      Width           =   3225
      Begin MSMask.MaskEdBox mskDia 
         Height          =   345
         Left            =   210
         TabIndex        =   3
         ToolTipText     =   "Ao teclar [F2] carrega calendário"
         Top             =   795
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDiaFim 
         Height          =   345
         Left            =   1830
         TabIndex        =   4
         ToolTipText     =   "Ao teclar [F2] carrega calendário"
         Top             =   795
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Até"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1800
         TabIndex        =   28
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "De"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   27
         Top             =   540
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1800
      Left            =   0
      TabIndex        =   14
      Top             =   660
      Width           =   6285
      Begin SSDataWidgets_B.SSDBCombo cboCartoes 
         Bindings        =   "frmRelLancCartaoPosiDiaria.frx":4E95A
         DataSource      =   "datCartoes"
         Height          =   345
         Left            =   945
         TabIndex        =   2
         Top             =   990
         Width           =   825
         DataFieldList   =   "Código"
         _Version        =   196617
         BackColorOdd    =   16777152
         Columns(0).Width=   3200
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCaixa 
         Bindings        =   "frmRelLancCartaoPosiDiaria.frx":4E973
         DataSource      =   "datCaixa"
         Height          =   345
         Left            =   945
         TabIndex        =   1
         Top             =   600
         Width           =   825
         DataFieldList   =   "Caixa"
         _Version        =   196617
         Cols            =   2
         BackColorOdd    =   16777152
         Columns(0).Width=   3200
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Caixa"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelLancCartaoPosiDiaria.frx":4E98A
         DataSource      =   "datFilial"
         Height          =   345
         Left            =   945
         TabIndex        =   0
         Top             =   210
         Width           =   825
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
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   7752
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1879
         Columns(1).Caption=   "Filial"
         Columns(1).Name =   "Filial"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Filial"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Filial"
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
         Bindings        =   "frmRelLancCartaoPosiDiaria.frx":4E9A2
         DataSource      =   "dtaVendedor"
         Height          =   345
         Left            =   945
         TabIndex        =   37
         Top             =   1380
         Width           =   825
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
         BevelColorFrame =   0
         BevelColorHighlight=   16777215
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   9208
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
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin VB.Label Nome_Vendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1800
         TabIndex        =   39
         Top             =   1380
         Width           =   4425
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cartão"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label lblNomeCartao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1800
         TabIndex        =   19
         Top             =   990
         Width           =   4425
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Caixa"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   660
         Width           =   405
      End
      Begin VB.Label lblNomeCaixa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1800
         TabIndex        =   17
         Top             =   600
         Width           =   4425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   270
         Width           =   300
      End
      Begin VB.Label lblNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1800
         TabIndex        =   15
         Top             =   210
         Width           =   4425
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   -120
      Width           =   14655
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelLancCartaoPosiDiaria.frx":4E9BC
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   510
         TabIndex        =   26
         Top             =   450
         Width           =   13965
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Exibe os saldos de um período das vendas em que foram recebidas com cartões de débito ou crédito totalizando por administradora."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   165
         Width           =   9495
      End
   End
   Begin Crystal.CrystalReport crpRel 
      Left            =   120
      Top             =   7350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid gridCartoes 
      Height          =   3585
      Left            =   0
      TabIndex        =   24
      Top             =   3270
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6324
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454143
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16250871
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid gridCartoesDetalhe 
      Height          =   3585
      Left            =   9390
      TabIndex        =   33
      Top             =   3270
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   6324
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454143
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16250871
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Lista as vendas de UMA Administradora/Bandeira de cartão"
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
      Left            =   9390
      TabIndex        =   36
      Top             =   3000
      Width           =   4785
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Lista todas as Administradoras/Bandeiras de cartões"
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
      Left            =   30
      TabIndex        =   35
      Top             =   3000
      Width           =   4200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Total Liquido R$ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   32
      Top             =   6960
      Width           =   1440
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Total Bruto R$ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2010
      TabIndex        =   30
      Top             =   6960
      Width           =   1290
   End
End
Attribute VB_Name = "frmRelLancCartaoPosiDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'24/03/2005 - Daniel
'
'Relatório desenvolvido para atender inicialmente
'o cliente Bem Me Quer. Este relatório estará
'disponível para todos clientes do Quick Store
Option Explicit

Private rsVendedor As Recordset
Public paramCodFilial As Integer

Private Sub cboCaixa_CloseUp()
  cboCaixa.Text = cboCaixa.Columns(0).Text
  cboCaixa_LostFocus
End Sub

Private Sub cboCaixa_LostFocus()
  Dim rstCaixa As Recordset
  
  lblNomeCaixa.Caption = ""
  If Not IsNumeric(cboCaixa.Text) Then Exit Sub
  
  Set rstCaixa = db.OpenRecordset("SELECT Caixa, Descrição FROM [Caixas em Uso] WHERE Caixa = " & cboCaixa.Text, dbOpenSnapshot)
  
  With rstCaixa
    If Not (.BOF And .EOF) Then
      lblNomeCaixa.Caption = .Fields("Descrição") & ""
    End If
    
    If Not rstCaixa Is Nothing Then .Close
    Set rstCaixa = Nothing
  End With

End Sub

Private Sub cboCartoes_CloseUp()
  cboCartoes.Text = cboCartoes.Columns(0).Text
  cboCartoes_LostFocus
End Sub

Private Sub cboCartoes_LostFocus()
  Dim rstCartoes As Recordset
  
  lblNomeCartao.Caption = ""
  If Not IsNumeric(cboCartoes.Text) Then Exit Sub
  
  Set rstCartoes = db.OpenRecordset("SELECT Código, Nome FROM Cartões WHERE Código = " & cboCartoes.Text, dbOpenSnapshot)
  
  With rstCartoes
    If Not (.BOF And .EOF) Then
      lblNomeCartao.Caption = .Fields("Nome") & ""
    End If
    
    If Not rstCartoes Is Nothing Then .Close
    Set rstCartoes = Nothing
  End With

End Sub

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(1).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstParame As Recordset
  
  lblNomeFilial.Caption = ""
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstParame = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & cboFilial.Text, dbOpenSnapshot)
  
  With rstParame
    If Not (.BOF And .EOF) Then
      lblNomeFilial.Caption = .Fields("Nome") & ""
    End If
    
    If Not rstParame Is Nothing Then .Close
    Set rstParame = Nothing
  End With
  
End Sub

Private Sub cmd_Detalhar_Click()
On Error GoTo Erro
  Dim strSQL As String
  Dim rstTotalCartoesDet As Recordset
  Dim sCodigoAdm As String
  
  gridCartoesDetalhe.Rows = 1

  If gridCartoes.RowSel > 0 Then
  
    sCodigoAdm = gridCartoes.TextMatrix(gridCartoes.RowSel, 1)
    
    strSQL = ""
    strSQL = "SELECT Sequencia, Vl_Bruto, Vl_Liquido FROM TotalCartoes where Administradora = " & sCodigoAdm & " And Filial = " & gnCodFilial
    strSQL = strSQL & " ORDER BY Sequencia "
 
    Set rstTotalCartoesDet = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
 
    With rstTotalCartoesDet
        If Not (.BOF And .EOF) Then
            .MoveFirst
    
            Do Until .EOF
                gridCartoesDetalhe.AddItem .Fields("Sequencia").Value & vbTab & _
                          FormataValorTexto(.Fields("Vl_Bruto").Value, 2) & "" & vbTab & _
                          FormataValorTexto(.Fields("Vl_Liquido").Value, 2)
    
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstTotalCartoesDet = Nothing
  End If

  Exit Sub
Erro:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
End Sub

Private Sub cmd_pesquisarParaGrade_Click()
  On Error GoTo ErrHandler
  
  gridCartoes.Rows = 1
  gridCartoes.Row = 0
  gridCartoesDetalhe.Rows = 1
  
  txt_total.Text = ""
  txt_totalLiquido.Text = ""
  
  'Chamada da validação
  If ValidarCampos Then Exit Sub
  
  Call StatusMsg("Totalizando os cartões...")
  Screen.MousePointer = vbHourglass
  Call TotalizarCartoes
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  
  On Error GoTo ErrHandler
  
  gridCartoes.Rows = 1
  gridCartoes.Row = 0
  
  txt_total.Text = ""
  txt_totalLiquido.Text = ""
  
  'Chamada da validação
  If ValidarCampos Then Exit Sub
  
  Call StatusMsg("Totalizando os cartões...")
  Screen.MousePointer = vbHourglass
  Call TotalizarCartoes
  
  Call StatusMsg("Montando o relatório...")
  Call MontarRelatorio
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(1).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_LostFocus()
  Call StatusMsg("")
  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) > 999 Then Exit Sub
  rsVendedor.Index = "Código"
  rsVendedor.Seek "=", Val(Combo_Vendedor.Text)
  If rsVendedor.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsVendedor("Nome")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsVendedor.Close
    Set rsVendedor = Nothing
End Sub

Private Sub gridCartoes_Click()
    gridCartoesDetalhe.Rows = 1
End Sub

Private Sub mskDia_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDia.Text = frmCalendario.gsDateCalender(mskDia.Text)
  End If
End Sub

Private Sub mskDia_LostFocus()
  mskDia.Text = Ajusta_Data(mskDia.Text)
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFilial.DatabaseName = gsQuickDBFileName
  datCaixa.DatabaseName = gsQuickDBFileName
  datCartoes.DatabaseName = gsQuickDBFileName
  dtaVendedor.DatabaseName = gsQuickDBFileName
  
  mskDia.Text = Format(Date, "DD/MM/YYYY")
  mskDiaFim.Text = Format(Date, "DD/MM/YYYY")
  
  gridCartoes.ColWidth(0) = 0
  gridCartoes.ColWidth(1) = 1400
  gridCartoes.ColWidth(2) = 3300
  gridCartoes.ColWidth(3) = 1700
  gridCartoes.ColWidth(4) = 1700
  
  gridCartoes.Row = 0
  gridCartoes.TextMatrix(0, 0) = ""
  gridCartoes.TextMatrix(0, 1) = "Administradora"
  gridCartoes.TextMatrix(0, 2) = "Nome Administradora"
  gridCartoes.TextMatrix(0, 3) = "Valor Bruto"
  gridCartoes.TextMatrix(0, 4) = "Valor Líquido"
  
  gridCartoesDetalhe.ColWidth(0) = 1500
  gridCartoesDetalhe.ColWidth(1) = 1700
  gridCartoesDetalhe.ColWidth(2) = 1700
  
  gridCartoesDetalhe.Row = 0
  gridCartoesDetalhe.TextMatrix(0, 0) = "Sequência Venda"
  gridCartoesDetalhe.TextMatrix(0, 1) = "Valor Bruto"
  gridCartoesDetalhe.TextMatrix(0, 2) = "Valor Líquido"
  
  Set rsVendedor = db.OpenRecordset("Funcionários", , dbReadOnly)
  
  If paramCodFilial > 0 Then
      cboFilial.Text = paramCodFilial
      cboCaixa.Text = 1
      cboFilial_LostFocus
      cboCaixa_LostFocus
      cmd_pesquisarParaGrade_Click
  End If
  
End Sub

Private Function ValidarCampos() As Boolean
  If Len(lblNomeFilial.Caption) <= 0 Then
    ValidarCampos = True
    MsgBox "Filial inválida, verifique.", vbExclamation, "Quick Store"
    cboFilial.SetFocus
    Exit Function
  End If
  
  If Len(lblNomeCaixa.Caption) <= 0 Then
    ValidarCampos = True
    MsgBox "Caixa inválido, verifique.", vbExclamation, "Quick Store"
    cboCaixa.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDia.Text) Then
    ValidarCampos = True
    MsgBox "Data inicial inválida, verifique.", vbExclamation, "Quick Store"
    mskDia.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDiaFim.Text) Then
    ValidarCampos = True
    MsgBox "Data final inválida, verifique.", vbExclamation, "Quick Store"
    mskDiaFim.SetFocus
    Exit Function
  End If
  
End Function

Private Sub TotalizarCartoes()
  '24/03/2005 - Daniel
  'Rotina que monta o total de cartões
  'por administradora
  Dim rstSaidas            As Recordset
  Dim rstCR                As Recordset
  Dim rstTotalCartoes      As Recordset
  Dim rstTotalCartoesGroup As Recordset
  Dim strSQL               As String
 
  Dim dTotalGrade As Double
  Dim sValorTotalGrade As String
  
  Dim dTotalGradeLiquido As Double
  Dim sValorTotalGradeLiquido As String
 
 
On Error GoTo ErrHandler
 
 dbTemp.Execute "DELETE FROM TotalCartoes Where Filial = " & gnCodFilial
 Set rstTotalCartoes = dbTemp.OpenRecordset("TotalCartoes", dbOpenDynaset)
 
 '---[Primeiro buscamos todas às saídas onde houve recebimento com cartão e usou o "caixa escolhido"]---
 strSQL = "SELECT Sequência AS Seq FROM Saídas "
 strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
 strSQL = strSQL & " AND Caixa = " & CByte(cboCaixa.Text)
 strSQL = strSQL & " AND [Recebe - Cartão] <> " & 0
 strSQL = strSQL & " AND Data >= #" & Format(CDate(mskDia.Text), "MM/DD/YYYY") & "#"
 strSQL = strSQL & " AND Data <= #" & Format(CDate(mskDiaFim.Text), "MM/DD/YYYY") & "#"
 
 If Nome_Vendedor.Caption <> "" Then
    strSQL = strSQL & " AND Digitador = " & Combo_Vendedor.Text
 End If
 
 Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
 
 With rstSaidas
  If Not (.BOF And .EOF) Then
    .MoveFirst
    
    Do Until .EOF
      strSQL = ""
      strSQL = "SELECT * "
      strSQL = strSQL & " FROM [Contas a Receber] "
      strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
      strSQL = strSQL & " AND Sequência = " & rstSaidas.Fields("Seq").Value
      strSQL = strSQL & " AND Tipo = '" & "O" & "'"
      strSQL = strSQL & " AND [Data Emissão] >= #" & Format(CDate(mskDia.Text), "MM/DD/YYYY") & "#"
      strSQL = strSQL & " AND [Data Emissão] <= #" & Format(CDate(mskDiaFim.Text), "MM/DD/YYYY") & "#"
    
      If Len(lblNomeCartao.Caption) > 0 Then strSQL = strSQL & " AND Administradora = " & CByte(cboCartoes.Text)
    
      Set rstCR = db.OpenRecordset(strSQL, dbOpenDynaset)
      
      If Not (rstCR.BOF And rstCR.EOF) Then
        rstCR.MoveFirst
         Do Until rstCR.EOF
            'Criamos o registro temporário
            rstTotalCartoes.AddNew
             rstTotalCartoes.Fields("Filial").Value = gnCodFilial
             rstTotalCartoes.Fields("Sequencia").Value = rstCR.Fields("Sequência").Value
             rstTotalCartoes.Fields("Administradora").Value = rstCR.Fields("Administradora").Value
             rstTotalCartoes.Fields("Vl_Bruto").Value = rstCR.Fields("Valor Cartão").Value
             rstTotalCartoes.Fields("Vl_Liquido").Value = rstCR.Fields("Valor").Value
            rstTotalCartoes.Update
          rstCR.MoveNext
         Loop
      End If
      rstCR.Close
      Set rstCR = Nothing
      
     .MoveNext
    Loop
    
  End If
  .Close
 End With
 
 rstTotalCartoes.Close
 Set rstTotalCartoes = Nothing
 
 Set rstSaidas = Nothing
 '---[Fim da busca em saídas]---
 
 'A partir daqui já temos às informações necessárias na tabela temporária TotalCartoes
 'onde poderemos agrupar os registros para a TotalCartoesGroup
 dbTemp.Execute "DELETE FROM TotalCartoesGroup Where Filial = " & gnCodFilial
 
 Set rstTotalCartoesGroup = dbTemp.OpenRecordset("TotalCartoesGroup", dbOpenDynaset)
 
 strSQL = ""
 strSQL = "SELECT Filial, Administradora, SUM(Vl_Bruto) AS Bruto, SUM(Vl_Liquido) AS Liquido FROM TotalCartoes "
 strSQL = strSQL & " Where Filial = " & gnCodFilial
 strSQL = strSQL & " GROUP BY Filial, Administradora "
 
 Set rstTotalCartoes = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
 
 With rstTotalCartoes
  If Not (.BOF And .EOF) Then
    .MoveFirst
    
    Do Until .EOF
      rstTotalCartoesGroup.AddNew
        rstTotalCartoesGroup.Fields("Filial").Value = gnCodFilial
        rstTotalCartoesGroup.Fields("Administradora").Value = .Fields("Administradora").Value
        rstTotalCartoesGroup.Fields("Nome").Value = getNomeAdministradora(.Fields("Administradora").Value) & ""
        rstTotalCartoesGroup.Fields("Vl_Bruto").Value = .Fields("Bruto").Value
        rstTotalCartoesGroup.Fields("Vl_Liquido").Value = .Fields("Liquido").Value
      rstTotalCartoesGroup.Update
      
      
      If .Fields("Filial").Value = gnCodFilial Then
          gridCartoes.AddItem vbTab & .Fields("Administradora").Value & vbTab & _
                getNomeAdministradora(.Fields("Administradora").Value) & "" & vbTab & _
                FormataValorTexto(.Fields("Bruto").Value, 2) & vbTab & _
                FormataValorTexto(.Fields("Liquido").Value, 2)
    
          sValorTotalGrade = Format(.Fields("Bruto").Value, FORMAT_VALUE)
          dTotalGrade = dTotalGrade + CDbl(sValorTotalGrade)
        
          sValorTotalGradeLiquido = Format(.Fields("Liquido").Value, FORMAT_VALUE)
          dTotalGradeLiquido = dTotalGradeLiquido + CDbl(sValorTotalGradeLiquido)
      End If
    
     .MoveNext
    Loop
    
  End If
  .Close
 End With
 
  txt_total.Text = Format(dTotalGrade, FORMAT_VALUE)
  txt_totalLiquido.Text = Format(dTotalGradeLiquido, FORMAT_VALUE)
 
 Set rstTotalCartoes = Nothing
 
 rstTotalCartoesGroup.Close
 Set rstTotalCartoesGroup = Nothing
 
 Exit Sub
 
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Atenção"
  Exit Sub
  
End Sub

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
  
  If lngCasasDecimais = 2 Then
      If Len(FormataValorTexto) = 7 Then  ' 9999.99     = 9.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 6)
      ElseIf Len(FormataValorTexto) = 8 Then ' 99999.99    = 99.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 6)
      ElseIf Len(FormataValorTexto) = 9 Then ' 999999.99   = 999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 6)
      ElseIf Len(FormataValorTexto) = 10 Then ' 9999999.99   = 9.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 3) + "." + Mid(FormataValorTexto, 5, 6)
      ElseIf Len(FormataValorTexto) = 11 Then ' 99999999.99   = 99.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 3) + "." + Mid(FormataValorTexto, 6, 6)
      ElseIf Len(FormataValorTexto) = 12 Then ' 999999999.99   = 999.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 3) + "." + Mid(FormataValorTexto, 7, 6)
      End If
  End If
End Function

Private Function getNomeAdministradora(ByVal CodAdmi As String) As String
  '23/03/2005 - Daniel
  Dim rstCartoes As Recordset
  
  Set rstCartoes = db.OpenRecordset("SELECT Nome FROM Cartões WHERE Código = " & CodAdmi, dbOpenDynaset)
  
  With rstCartoes
    If Not (.BOF And .EOF) Then
      .MoveFirst
      getNomeAdministradora = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  
  Set rstCartoes = Nothing
  
End Function

Private Sub MontarRelatorio()
  Dim strReport As String
  
  'Nome do arquivo .rpt
  strReport = gsReportPath & "rptCartoesCreditoPosiDiaria2.rpt"
  
  With crpRel
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crpRel
    
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsTempDBFileName
    
    '.SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .Formulas(1) = "Dia = '" & "Período: " & (mskDia.Text) & " à " & (mskDiaFim.Text) & "'"
    .Formulas(2) = "Caixa = '" & "Caixa: " & (cboCaixa.Text) & "'"
    
    .ParameterFields(0) = "pFilial;" & gnCodFilial & ";true"
    
    'Ordenação
    If optCodAdm.Value Then .SortFields(0) = "+{TotalCartoesGroup.Administradora}"
    If optNomeAdm.Value Then .SortFields(0) = "+{TotalCartoesGroup.Nome}"
    If optValorBruto.Value Then .SortFields(0) = "-{TotalCartoesGroup.Vl_Bruto}"
    
    .WindowState = crptMaximized
    .Destination = IIf(optVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crpRel)
  
    .Action = 1
  End With

  Screen.MousePointer = vbDefault

End Sub

Private Sub mskDiaFim_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDiaFim.Text = frmCalendario.gsDateCalender(mskDiaFim.Text)
  End If
End Sub

Private Sub mskDiaFim_LostFocus()
  mskDiaFim.Text = Ajusta_Data(mskDiaFim.Text)
End Sub
