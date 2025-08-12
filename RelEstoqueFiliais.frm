VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEstoqueFiliais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Estoque por Filiais"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelEstoqueFiliais.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   10185
   Begin SSDataWidgets_B.SSDBCombo cboClasse 
      Bindings        =   "RelEstoqueFiliais.frx":4E95A
      DataSource      =   "datClasses"
      Height          =   345
      Left            =   4140
      TabIndex        =   33
      Top             =   615
      Width           =   1215
      DataFieldList   =   "Código"
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
      Columns(0).Width=   3200
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Código"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   6218
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Nome"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
      Enabled         =   0   'False
   End
   Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
      Bindings        =   "RelEstoqueFiliais.frx":4E973
      DataSource      =   "datSubClasses"
      Height          =   345
      Left            =   4140
      TabIndex        =   32
      Top             =   1050
      Width           =   1215
      DataFieldList   =   "Código"
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
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   2143
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
      Enabled         =   0   'False
   End
   Begin Crystal.CrystalReport Rel2 
      Left            =   9630
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   9630
      Top             =   1290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto 
      Bindings        =   "RelEstoqueFiliais.frx":4E98F
      DataSource      =   "datCodProdutos"
      Height          =   345
      Left            =   4140
      TabIndex        =   25
      Top             =   180
      Width           =   1215
      DataFieldList   =   "Código"
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
      Columns(0).Width=   3228
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Código"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   7144
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Nome"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
      Height          =   465
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4140
      Width           =   9975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Seleção"
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   6315
      Begin VB.CheckBox chkTotalizarPorProdutos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Totalizar por produto"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   2085
      End
      Begin VB.CheckBox chkOpcoes_SepararClasse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Separar por classe/ sub classe"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2430
         TabIndex        =   23
         Top             =   660
         Width           =   2685
      End
      Begin VB.CheckBox chkOpcoes_IgnorarInativos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Não Imprimir inativos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2430
         TabIndex        =   22
         Top             =   300
         Width           =   2085
      End
      Begin VB.TextBox txtEstoque_Ate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   480
         TabIndex        =   21
         Text            =   "99999999"
         Top             =   622
         Width           =   1455
      End
      Begin VB.TextBox txtEstoque_De 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   480
         TabIndex        =   19
         Text            =   "0"
         Top             =   262
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "(Somente produtos com grade)"
         Height          =   255
         Left            =   2460
         TabIndex        =   35
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   675
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   315
         Width           =   195
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Saída"
      Height          =   1455
      Left            =   6510
      TabIndex        =   14
      Top             =   2520
      Width           =   3585
      Begin VB.Frame Frame5 
         Caption         =   "Layout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   2040
         TabIndex        =   27
         Top             =   30
         Width           =   1545
         Begin VB.OptionButton optOtimizado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Otimizado"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   750
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optCompleto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Completo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   330
            Width           =   1095
         End
      End
      Begin VB.OptionButton optSaida_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   16
         Top             =   780
         Width           =   1215
      End
      Begin VB.OptionButton optSaida_Video 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   390
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordenação"
      Height          =   705
      Left            =   6510
      TabIndex        =   10
      Top             =   1470
      Width           =   3585
      Begin VB.OptionButton optOrd_Qtde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Qtde"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   300
         Width           =   705
      End
      Begin VB.OptionButton optOrd_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1350
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optOrd_Codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   705
      Left            =   3240
      TabIndex        =   6
      Top             =   1470
      Width           =   3195
      Begin VB.OptionButton optTipo_Edicao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Edição"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1230
         TabIndex        =   9
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optTipo_Grade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Grade"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2190
         TabIndex        =   8
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optTipo_Normal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Normal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.ListBox lstFiliais 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   1380
      ItemData        =   "RelEstoqueFiliais.frx":4E9AC
      Left            =   120
      List            =   "RelEstoqueFiliais.frx":4E9AE
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Data datClasses 
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
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "con_Classe"
      Top             =   570
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data datSubClasses 
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
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "con_Sub_Classe"
      Top             =   990
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data datCodProdutos 
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
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSubClasse 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5400
      TabIndex        =   31
      Top             =   1050
      Width           =   4695
   End
   Begin VB.Label lbltitSubClasse 
      Caption         =   "Sub-Classe"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   30
      Top             =   1095
      Width           =   855
   End
   Begin VB.Label lblProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5400
      TabIndex        =   3
      Top             =   180
      Width           =   4695
   End
   Begin VB.Label lblFiliaisSelecionadas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "0"
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   2520
      TabIndex        =   26
      Top             =   2190
      Width           =   615
   End
   Begin VB.Label lblClasse 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5400
      TabIndex        =   5
      Top             =   615
      Width           =   4695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Classe"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   675
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Produto"
      Height          =   195
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "Filiais"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   495
   End
End
Attribute VB_Name = "frmRelEstoqueFiliais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsEstoque         As Recordset
Private rsEstoqueTEMP     As Recordset
Private rsParametros      As Recordset
Private rsClassesCombo    As Recordset
Private rsSubClassesCombo As Recordset
Private rsProdutosCombo   As Recordset
Private sSql              As String
Private sSQLTemp          As String
Private rsEstFilial       As Recordset
Private nI                As Integer
  
Private Sub cboClasse_CloseUp()
  lblClasse.Caption = cboClasse.Columns(1).Text
End Sub

Private Sub cboClasse_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboClasse_LostFocus()
  If cboClasse.Text = "0" Then
    lblClasse.Caption = "Todas as classes"
    Exit Sub
  End If
  
  lblClasse.Caption = ""
  If IsNull(cboClasse.Text) Then Exit Sub
  If cboClasse.Text = "" Then Exit Sub
  If Not IsNumeric(cboClasse.Text) Then Exit Sub
  If Val(cboClasse.Text) < 1 Then Exit Sub
  If Val(cboClasse.Text) > 9999 Then Exit Sub
  
  rsClassesCombo.Index = "Código"
  rsClassesCombo.Seek "=", Val(cboClasse.Text)
  If rsClassesCombo.NoMatch Then Exit Sub
  
  lblClasse.Caption = rsClassesCombo("Nome") & ""
End Sub

Private Sub cboProduto_CloseUp()
  lblProduto.Caption = cboProduto.Columns(1).Text
End Sub

Private Sub cboProduto_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboProduto_LostFocus()
  If cboProduto.Text = "0" Then
    lblProduto.Caption = "Todos os produtos"
    Exit Sub
  End If
  
  lblProduto.Caption = ""
  If IsNull(cboProduto.Text) Then Exit Sub
  If cboProduto.Text = "" Then Exit Sub

  rsProdutosCombo.Index = "Código"
  rsProdutosCombo.Seek "=", cboProduto.Text
  If rsProdutosCombo.NoMatch Then Exit Sub
  
  lblProduto.Caption = rsProdutosCombo("Nome") & ""
End Sub

Private Sub cboSubClasse_CloseUp()
  lblSubClasse.Caption = cboSubClasse.Columns(0).Text
End Sub

Private Sub cboSubClasse_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboSubClasse_LostFocus()
  If cboSubClasse.Text = "0" Then
    lblSubClasse.Caption = "Todas as Sub-Classes"
    Exit Sub
  End If
  
  lblSubClasse.Caption = ""
  If IsNull(cboSubClasse.Text) Then Exit Sub
  If cboSubClasse.Text = "" Then Exit Sub
  If Not IsNumeric(cboSubClasse.Text) Then Exit Sub
  
  rsSubClassesCombo.Index = "Código"
  rsSubClassesCombo.Seek "=", Val(cboSubClasse.Text)
  If rsSubClassesCombo.NoMatch Then Exit Sub
  
  lblSubClasse.Caption = rsSubClassesCombo("Nome") & ""
End Sub

Private Sub chkOpcoes_SepararClasse_Click()
  Label4.Enabled = chkOpcoes_SepararClasse.Value
  lbltitSubClasse.Enabled = chkOpcoes_SepararClasse.Value
  cboClasse.Enabled = chkOpcoes_SepararClasse.Value
  cboSubClasse.Enabled = chkOpcoes_SepararClasse.Value
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  'Validações nos campos digitados
  '-------------------------------
  If lstFiliais.SelCount < 1 Then
    MsgBox "Nenhuma filial selecionada, verifique.", vbCritical, "Erro ao gerar relatório"
    Exit Sub
  End If
  
  If Not IsNumeric(txtEstoque_De.Text) Then
    MsgBox "Estoque inicial inválido, verifique.", vbCritical, "Erro ao gerar relatório"
    Exit Sub
  End If
  
  If Not IsNumeric(txtEstoque_Ate.Text) Then
    MsgBox "Estoque final inválido, verifique.", vbCritical, "Erro ao gerar relatório"
    Exit Sub
  End If
  
  If Val(txtEstoque_Ate.Text) < Val(txtEstoque_De.Text) Then
    MsgBox "O Estoque final não pode ser menor que o estoque inicial, verifique.", vbCritical, "Erro ao gerar relatório"
    Exit Sub
  End If
  
  If (Val(lblFiliaisSelecionadas.Caption) > 10 And optOtimizado.Value) Then
    MsgBox "O layout de relatório que você selecionou não suporta mais do que 10 filiais, verifique.", vbCritical, "Erro ao gerar relatório"
    Exit Sub
  End If
  '-------------------------------
    
  If optCompleto.Value = True Then
    rptCompleto
  ElseIf optOtimizado.Value = True Then
    rptOtimizado
  End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Set rsParametros = db.OpenRecordset("SELECT Filial FROM [Parâmetros Filial] ORDER BY Filial ", dbOpenSnapshot)
  Set rsClassesCombo = db.OpenRecordset("Classes")
  Set rsSubClassesCombo = db.OpenRecordset("Sub Classes")
  Set rsProdutosCombo = db.OpenRecordset("Produtos")
  
  With rsParametros
    .MoveFirst
    lstFiliais.Clear
    Do While Not .EOF
      lstFiliais.AddItem .Fields("Filial")
      .MoveNext
    Loop
    If lstFiliais.ListCount = 1 Then lstFiliais.Selected(0) = True
  End With
  
  datCodProdutos.DatabaseName = gsQuickDBFileName
  datSubClasses.DatabaseName = gsQuickDBFileName
  datClasses.DatabaseName = gsQuickDBFileName
End Sub

Private Sub lstFiliais_Click()
  If lstFiliais.SelCount > 10 Then
    If optOtimizado.Value Then
      lstFiliais.Selected(lstFiliais.ListIndex) = False
      MsgBox "Número máximo de filiais excedido para esse layout!!! ", vbCritical, "Erro ao selecionar filial"
    End If
  End If
  lblFiliaisSelecionadas.Caption = lstFiliais.SelCount
End Sub

Private Sub rptCompleto()
  Dim sStrReport As String
  Dim nDataFiles As Integer
  
  With Rel2
    For nDataFiles = 0 To 8
      .DataFiles(nDataFiles) = gsQuickDBFileName
    Next nDataFiles
  End With
  
  'Saída em vídeo ou impressora
  If optSaida_Video.Value Then
    Rel2.Destination = 0
  ElseIf optSaida_Impressora.Value Then
    Rel2.Destination = 1
  End If
  
  'Indica ao Report qual layout usar
  If optTipo_Grade.Value Then
    If chkOpcoes_SepararClasse.Value = 1 Then
      Rel2.ReportFileName = gsReportPath & "EstComGC.RPT"
    Else
      Rel2.ReportFileName = gsReportPath & "EstComG.RPT"
    End If
  ElseIf optTipo_Edicao.Value Then
    If chkOpcoes_SepararClasse.Value = 1 Then
      Rel2.ReportFileName = gsReportPath & "EstComEC.RPT"
    Else
      Rel2.ReportFileName = gsReportPath & "EstComE.RPT"
    End If
  ElseIf optTipo_Normal.Value Then
    If chkOpcoes_SepararClasse.Value = 1 Then
      Rel2.ReportFileName = gsReportPath & "EstComNC.RPT"
    Else
      Rel2.ReportFileName = gsReportPath & "EstComN.RPT"
    End If
  End If
  '--------------------------------------
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel2
  
  'Seleções iniciais...
  sStrReport = "{Estoque Final.Produto} <> ""0"" and " & _
               "{Estoque Final.Filial} <> 0 and "
  
  If chkOpcoes_SepararClasse.Value = 1 Then
    '25/01/2006 - mpdea
    'Adicionado verificação de classe preenchida
    If lblClasse.Caption <> "" Then
      sStrReport = sStrReport & " {Estoque Final.Classe} = " & Val(cboClasse.Text) & " and"
    End If
    'Adicionado seleção de subclasse
    If lblSubClasse.Caption <> "" Then
      sStrReport = sStrReport & " {Estoque Final.Sub Classe} = " & Val(cboSubClasse.Text) & " and"
    End If
  End If
  
  If chkOpcoes_IgnorarInativos.Value = 1 Then
    sStrReport = sStrReport & " {Produtos.Desativado} = FALSE AND "
  End If
  
  If cboProduto.Text <> "" And _
     cboProduto.Text <> "0" Then
     sStrReport = sStrReport & " {Estoque Final.Produto} = """ & cboProduto.Text & """ and "
  End If
  
  For nI = 0 To lstFiliais.ListCount - 1
     If lstFiliais.Selected(nI) Then
       sStrReport = sStrReport & " {Estoque Final.Filial} = " & lstFiliais.List(nI) & " or "
     End If
  Next nI
  
  If optOrd_Codigo.Value Then
    Rel2.SortFields(0) = "+{Estoque Final.Produto} "
  ElseIf optOrd_Nome.Value Then
    Rel2.SortFields(0) = "+{Produtos.Nome} "
  ElseIf optOrd_Qtde.Value Then
    Rel2.SortFields(0) = "+{Estoque Final.Estoque Atual} "
  End If
  
  sStrReport = Left(sStrReport, Len(sStrReport) - 4)
  Rel2.SelectionFormula = sStrReport
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel2)
  
  
  Rel2.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault
End Sub

Private Sub rptOtimizado()
  Dim sStrReport       As String
  Dim aArrayFiliais(9) As String
  Dim nX               As Integer
  nX = 0
  
  Call StatusMsg("Aguarde, gerando relatório...")
  
  'Gerando dados para a tabela temporária
  '--------------------------------------
  
  For nI = 0 To lstFiliais.ListCount - 1
    If lstFiliais.Selected(nI) Then
      aArrayFiliais(nX) = lstFiliais.List(nI)
      nX = nX + 1
    End If
  Next nI
  
  sSql = " TRANSFORM Sum([Estoque Final].[Estoque Atual]) AS [O valor] " & _
         " SELECT [Estoque Final].Produto, "
  If optTipo_Grade.Value Then
    sSql = sSql & " [Estoque Final].Tamanho, [Estoque Final].Cor, "
  ElseIf optTipo_Edicao.Value Then
    sSql = sSql & "[Estoque Final].Edição, "
  End If
  
  sSql = sSql & _
       " Produtos.[Código Ordenação], Produtos.Nome, " & _
       " Sum([Estoque Final].[Estoque Atual]) AS Total " & _
       " FROM [Estoque Final] INNER JOIN Produtos ON " & _
       " [Estoque Final].Produto = Produtos.Código " & _
       " WHERE ([Estoque Final].Produto <> '0')"
  
  If chkOpcoes_IgnorarInativos.Value = vbChecked Then
    sSql = sSql & " AND ( Produtos.Desativado = False ) "
  End If
  
  If cboProduto.Text <> "" And cboProduto.Text <> "0" Then
    sSql = sSql & _
         " AND ( [Estoque Final].Produto = '" & cboProduto.Text & "' )"
  End If
  
  sSql = sSql & " AND ( [Estoque Final].Filial <> 0 ) AND ("
  
  For nI = LBound(aArrayFiliais) To nX - 1
    sSql = sSql & " ( [Estoque Final].Filial = " & aArrayFiliais(nI) & " ) OR "
  Next nI
  
  sSql = Left(sSql, Len(sSql) - 4) & " ) " ' )
  
  sSql = sSql & _
         " GROUP BY [Estoque Final].Produto"
  
  If optTipo_Grade.Value = True Then
    sSql = sSql & ", [Estoque Final].Tamanho, [Estoque Final].Cor"
  ElseIf optTipo_Edicao.Value = True Then
    sSql = sSql & ", [Estoque Final].Edição "
  End If
  
  sSql = sSql & ", Produtos.[Código Ordenação], Produtos.Nome "
  If optOrd_Codigo.Value Then
    sSql = sSql & " ORDER BY Produtos.[Código Ordenação]"
  ElseIf optOrd_Nome.Value Then
    sSql = sSql & " ORDER BY Produtos.Nome "
  End If
  sSql = sSql & " PIVOT [Estoque Final].Filial"
  
  dbTemp.Execute ("DELETE * FROM tbRelEstoqueFiliais_Campos")
  Set rsEstoqueTEMP = dbTemp.OpenRecordset("tbRelEstoqueFiliais_Campos")
  Set rsEstoque = db.OpenRecordset(sSql, dbOpenSnapshot)
  
  If Not rsEstoque.EOF Then rsEstoque.MoveFirst
  
  Do While Not rsEstoque.EOF
    rsEstoqueTEMP.AddNew
    rsEstoqueTEMP.Fields("proCodigo") = rsEstoque.Fields(0)
    If optTipo_Grade.Value = True Then
      rsEstoqueTEMP.Fields("proTamanho") = rsEstoque.Fields(1)
      rsEstoqueTEMP.Fields("proCor") = rsEstoque.Fields(2)
      rsEstoqueTEMP.Fields("est_Total") = IIf(rsEstoque.Fields(5) = "", "0", rsEstoque.Fields(5))

'      If (rsEstoque.Fields(5) = "") Then Stop  '(rsEstoqueTEMP.Fields("est_Total") = 0) And
      
    ElseIf optTipo_Edicao.Value = True Then
      rsEstoqueTEMP.Fields("proEdicao") = rsEstoque.Fields(1)
      rsEstoqueTEMP.Fields("est_Total") = IIf(rsEstoque.Fields(4) = "", "0", rsEstoque.Fields(4))
'      rsEstoqueTEMP.Fields("est_Total") = rsEstoque.Fields(4)
    End If
    
    
    Dim nContNomeFilial As Integer
    nContNomeFilial = 1
    For nI = LBound(aArrayFiliais) To nX - 1
      On Error Resume Next
      rsEstoqueTEMP.Fields("est_" & CStr(nContNomeFilial)) = rsEstoque.Fields(CStr(aArrayFiliais(nContNomeFilial - 1)))
      nContNomeFilial = nContNomeFilial + 1
    Next nI
    
    rsEstoqueTEMP.Update
    rsEstoque.MoveNext
  Loop
  
  
  With Rel1
    Dim nDataFiles As Integer
    .DataFiles(0) = gsTempDBFileName
    For nDataFiles = 1 To 8
      .DataFiles(nDataFiles) = gsQuickDBFileName
    Next nDataFiles
  End With

  'Saída em vídeo ou impressora
  If optSaida_Video.Value Then
    Rel1.Destination = 0
  ElseIf optSaida_Impressora.Value Then
    Rel1.Destination = 1
  End If
  
  'Indica ao Report qual layout usar
  If optTipo_Grade.Value Then
    If chkOpcoes_SepararClasse.Value Then
      If chkTotalizarPorProdutos.Value Then
        Rel1.ReportFileName = gsReportPath & "EstOtiGClassTot.RPT"
      Else
        Rel1.ReportFileName = gsReportPath & "EstOtiGClass.RPT"
      End If
    Else
      If chkTotalizarPorProdutos.Value Then
        Rel1.ReportFileName = gsReportPath & "EstOtiGTot.RPT"
      Else
        Rel1.ReportFileName = gsReportPath & "EstOtiG.RPT"
      End If
    End If
  ElseIf optTipo_Edicao.Value Then
    If chkOpcoes_SepararClasse.Value Then
      Rel1.ReportFileName = gsReportPath & "EstOtiEClass.RPT"
    Else
      Rel1.ReportFileName = gsReportPath & "EstOtiE.RPT"
    End If
  ElseIf optTipo_Normal.Value Then
    If chkOpcoes_SepararClasse.Value Then
      Rel1.ReportFileName = gsReportPath & "EstOtiNClass.RPT"
    Else
      Rel1.ReportFileName = gsReportPath & "EstOtiN.RPT"
    End If
  End If
  '--------------------------------------
  
  Dim sSelectionFormula As String
  sSelectionFormula = ""
  
  If chkOpcoes_IgnorarInativos.Value = 1 Then
    sSelectionFormula = " {Produtos.Desativado} <> TRUE AND"
  End If
  
  Dim nSelectionFormula As Integer
  
  For nSelectionFormula = 1 To 10
    sSelectionFormula = sSelectionFormula & _
                        " {tbRelEstoqueFiliais_Campos.est_" & CStr(nSelectionFormula) & "} >= " & Val(txtEstoque_De.Text) & " AND " & _
                        " {tbRelEstoqueFiliais_Campos.est_" & CStr(nSelectionFormula) & "} <= " & Val(txtEstoque_Ate.Text)
    sSelectionFormula = sSelectionFormula & " OR "
  Next nSelectionFormula
  
  sSelectionFormula = Left(sSelectionFormula, Len(sSelectionFormula) - 4)
  
  Rel1.SelectionFormula = sSelectionFormula
  
  Dim sGroupSelect As String
  
  If chkOpcoes_SepararClasse.Value Then
    If lblClasse.Caption <> "" Then
      sGroupSelect = " GroupName ({Produtos.Classe}) = '" & Val(cboClasse.Text) & "'"
    End If
    
    '25/01/2006 - mpdea
    'Adicionado seleção de subclasse independente da classe
    'If (lblSubClasse.Caption <> "") And (lblClasse.Caption <> "") Then
    If lblSubClasse.Caption <> "" Then
      If sGroupSelect <> "" Then
        sGroupSelect = sGroupSelect & " AND "
      End If
      sGroupSelect = sGroupSelect & " GroupName ({Produtos.Sub Classe}) = '" & Val(cboSubClasse.Text) & "'"
    End If
    
    Rel1.GroupSelectionFormula = sGroupSelect
  End If
  
  If optOrd_Qtde.Value Then
    Rel1.SortFields(0) = "+{@Total}"
  End If
  
  'Gerando o cabeçalho das colunas de filiais
  For nI = LBound(aArrayFiliais) To UBound(aArrayFiliais)
    Rel1.Formulas(nI) = "Fil_" & Trim(CStr(nI + 1)) & " = '" & Trim(CStr(aArrayFiliais(nI))) & "'"
  Next nI
  '------------------------------------------
  
  Rel1.Formulas(10) = "Est_DeAte = 'Estoque de [" & Val(txtEstoque_De.Text) & _
                               "], até [" & Val(txtEstoque_Ate.Text) & "]'"
  
  On Error Resume Next
  Rel1.Formulas("Total") = " {tbRelEstoqueFiliais_Campos.est_1}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_2}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_3}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_4}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_5}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_6}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_7}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_8}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_9}  + " & _
                      " {tbRelEstoqueFiliais_Campos.est_10} "
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
  
  'Verifica os campos nulos e substitui por 0 (zero)
  ZeraTabelas
  
  On Error GoTo TrataErro:
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  
  
  Rel1.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault
  
  'Fechando a conexão com a base de dados
  '--------------------------------------
  rsEstoque.Close
  rsEstoqueTEMP.Close
  
  Set rsEstoque = Nothing
  Set rsEstoqueTEMP = Nothing
  
  Exit Sub
  
TrataErro:
  MsgBox Err.Number & vbCrLf & _
         Err.Description
  Err = 0

  Call StatusMsg("")
  MousePointer = vbDefault
End Sub

'25/10/2002 - mpdea
'Especificado tipo da procedure para Private
'Comentado variáveis não utilizadas e abertura do banco de dados Temp
Private Sub ZeraTabelas()
'----------------------------------------------------------------------------------
'  Dim dbTemp As Database
'  Dim rs As Recordset

'  Dim nJ As Long
'  Dim sX As String
  
'  Set dbTemp = Workspaces(0).OpenDatabase(App.Path & "\Temp.mdb", False, False)
'----------------------------------------------------------------------------------

  Dim nX As Long
  
  For nX = 1 To 10
    dbTemp.Execute " UPDATE tbRelEstoqueFiliais_Campos SET " & _
                   " Est_" & nX & " = 0 " & _
                   " WHERE Est_" & nX & " IS NULL "
  
''''    sX = " SELECT tbRelEstoqueFiliais_Campos.est_" & nX & _
''''         " From tbRelEstoqueFiliais_Campos " & _
''''         " WHERE (((tbRelEstoqueFiliais_Campos.est_" & nX & ") Is Null))"
''''
''''    Set rs = dbTEMP.OpenRecordset(sX, dbOpenDynaset)
''''
''''    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
''''
''''    Do While Not rs.EOF
''''      rs.Edit
''''      rs.Fields("Est_" & nX) = 0
''''      rs.Update
''''      rs.MoveNext
''''    Loop
''''    rs.Close
''''    Set rs = Nothing
  Next nX
End Sub
