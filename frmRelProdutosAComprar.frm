VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelProdutosAComprar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Produtos a Comprar Por Fator"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelProdutosAComprar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   7335
   Begin VB.Data datProdutos 
      Caption         =   "datProdutos"
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
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   6840
      Width           =   2295
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   1920
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraRelatorio 
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   27
      Top             =   2640
      Width           =   3255
      Begin VB.OptionButton optSemEspacoFisico 
         Caption         =   "Não Considerar Espaço Físico"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton optComEspacoFisico 
         Caption         =   "Considerar Espaço Físico"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtNomeProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txtFator 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtNomeFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelProdutosAComprar.frx":058A
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1335
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.TextBox txtNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtNomeClasse 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtNomeSubClasse 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "frmRelProdutosAComprar.frx":05A4
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
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
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmRelProdutosAComprar.frx":05BF
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
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
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelProdutosAComprar.frx":05D7
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1335
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
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Filial"
      End
      Begin SSDataWidgets_B.SSDBCombo cboProduto 
         Bindings        =   "frmRelProdutosAComprar.frx":05F0
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   1335
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label Label8 
         Caption         =   "Dias"
         Height          =   240
         Left            =   2520
         TabIndex        =   31
         Top             =   2062
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Produto"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Fator"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2055
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Filial"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Sub-Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1710
         Width           =   975
      End
   End
   Begin VB.Data datSubClasse 
      Caption         =   "datSubClasse"
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
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM [Sub Classes] ORDER BY Nome"
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Data datClasse 
      Caption         =   "datClasse"
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
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Classes ORDER BY Nome"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Data datClientes 
      Caption         =   "datClientes"
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For ORDER BY Nome"
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Data datFiliais 
      Caption         =   "datFiliais"
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   1695
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Imprimir"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   3735
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskDataInicio 
         Height          =   315
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label4 
         Caption         =   "até:"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmRelProdutosAComprar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo ErrHandler

  Dim strNomeArquivo As String
  Dim Str_Rel As String
  Dim bolAtivarEspacoFisico As Boolean
  
  If (Not IsDate(mskDataInicio.Text)) And (Not IsDate(mskDataFinal.Text)) Then
    mskDataInicio.Text = Data_Atual
    mskDataFinal.Text = Data_Atual
  End If
  
  If Not IsDate(mskDataInicio.Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(mskDataFinal.Text) Then
    MsgBox "Data final inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(mskDataInicio.Text) > CDate(mskDataFinal.Text) Then
    MsgBox "A data inicial não pode ser maior que a data final !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If IsNull(txtNomeFilial.Text) Or txtNomeFilial.Text = "" Then
     DisplayMsg "Escolha a filial."
     cboFilial.SetFocus
     Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(cboFilial.Text) <> Filial_Liberada Then
       DisplayMsg "Funcionário não tem acesso a esta filial."
       Exit Sub
    End If
  End If

  If cboProduto.Text = "" Or IsNull(cboProduto.Text) Then
    cboProduto.Text = "0"
  End If
  
  If cboClasse.Text = "" Or IsNull(cboClasse.Text) Then
    cboClasse.Text = "0"
  End If
    
  If cboFornecedor.Text = "" Or IsNull(cboFornecedor.Text) Then
    cboFornecedor.Text = "0"
  End If
  
  If cboSubClasse.Text = "" Or IsNull(cboSubClasse.Text) Then
    cboSubClasse.Text = "0"
  End If
  
  If txtFator.Text = "" Or IsNull(txtFator.Text) Or "0" & txtFator.Text <= 0 Then
    DisplayMsg "O campo Fator não pode conter caracteres nulos ou em branco. Verifique se o campo Fator foi preenchido corretamente."
    txtFator.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Imprimindo...")
  MousePointer = vbHourglass
  
  g_bolRelatorioProdutosComprar cboFilial.Text, mskDataInicio.Text, mskDataFinal.Text, cboProduto.Text, cboFornecedor.Text, cboClasse.Text, cboSubClasse.Text, optComEspacoFisico.Value, txtFator.Text
  
  'Rem  Nome do BD
  Rel1.DataFiles(0) = gsQuickDBFileName
  Rel1.DataFiles(1) = gsTempDBFileName
  
  Rem Saída
  If B_Vídeo = True Then Rel1.Destination = 0
  If B_Impressora = True Then Rel1.Destination = 1
    
  If optComEspacoFisico = True Then
    strNomeArquivo = gsReportPath & "rptProdCompComEspFisico.rpt"
  Else
    strNomeArquivo = gsReportPath & "rptProdCompSemEspFisico.rpt"
  End If
  Rel1.ReportFileName = strNomeArquivo
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel1
  
  Str_Rel = "Filtro_Filial = '" + txtNomeFilial.Text + "'"
  Rel1.Formulas(1) = Str_Rel
  
  Rem data inicial
  Str_Rel = "Periodo = 'Período: " + mskDataInicio.Text + " à " + mskDataFinal.Text + "'"
  Rel1.Formulas(2) = Str_Rel
  
  Str_Rel = "Fator = '" + txtFator.Text + "'"
  Rel1.Formulas(3) = Str_Rel
  
  Str_Rel = "Filtro_Fornecedor = '(" + cboFornecedor.Text + ") " + txtNomeFornecedor.Text + "'"
  Rel1.Formulas(4) = Str_Rel

  Str_Rel = "Filtro_Produto = '(" + cboProduto.Text + ") " + txtNomeProduto.Text + "'"
  Rel1.Formulas(5) = Str_Rel
  
  Str_Rel = "Filtro_Classe = '(" + cboClasse.Text + ") " + txtNomeClasse.Text + "'"
  Rel1.Formulas(6) = Str_Rel
  
  Str_Rel = "Filtro_SubClasse = '(" + cboSubClasse.Text + ") " + txtNomeSubClasse.Text + "'"
  Rel1.Formulas(7) = Str_Rel

   '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  
  Rel1.Action = 1
   
  Call StatusMsg("")
  MousePointer = vbDefault
  
  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  MsgBox "Erro ao imprimir relatório: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_Load()

  datFiliais.DatabaseName = gsQuickDBFileName
  datClientes.DatabaseName = gsQuickDBFileName
  datClasse.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
  datProdutos.DatabaseName = gsQuickDBFileName
  
  Call CenterForm(Me)

End Sub


Private Sub cboClasse_CloseUp()
  cboClasse.Text = cboClasse.Columns(0).Text
  cboClasse_LostFocus
End Sub

Private Sub cboClasse_LostFocus()
  Dim rstClasses As Recordset
  
  txtNomeClasse.Text = ""
  If Not IsNumeric(cboClasse.Text) Then Exit Sub
  
  Set rstClasses = db.OpenRecordset("SELECT Código, Nome FROM Classes WHERE Código = " & cboClasse.Text, dbOpenSnapshot)
  
  With rstClasses
    If Not (.BOF And .EOF) Then
      txtNomeClasse.Text = .Fields("Nome") & ""
    End If
    
    If Not rstClasses Is Nothing Then .Close
    Set rstClasses = Nothing
  End With
End Sub

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstFiliais As Recordset
  
  txtNomeFilial.Text = ""
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstFiliais = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & cboFilial.Text, dbOpenSnapshot)
  
  With rstFiliais
    If Not (.BOF And .EOF) Then
      txtNomeFilial.Text = .Fields("Nome") & ""
    End If
    
    If Not rstFiliais Is Nothing Then .Close
    Set rstFiliais = Nothing
  End With
End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(0).Text
  cboFornecedor_LostFocus
End Sub

Private Sub cboFornecedor_LostFocus()
  txtNomeFornecedor.Text = ""
  If Not IsNumeric(cboFornecedor.Text) Then Exit Sub
  
  datClientes.Recordset.FindFirst "Código = " & cboFornecedor.Text
  
  If Not datClientes.Recordset.NoMatch Then
    txtNomeFornecedor.Text = datClientes.Recordset.Fields("Nome") & ""
  End If
End Sub

Private Sub cboProduto_CloseUp()
  cboProduto.Text = cboProduto.Columns(0).Text
  cboProduto_LostFocus
End Sub

Private Sub cboProduto_LostFocus()
  Dim rstProdutos As Recordset
  
  txtNomeProduto.Text = ""
  
  Set rstProdutos = db.OpenRecordset("SELECT Código, Nome FROM Produtos WHERE Código = '" & cboProduto.Text & "' AND Código <> '0' ", dbOpenSnapshot)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      txtNomeProduto.Text = .Fields("Nome") & ""
    End If
    
    If Not rstProdutos Is Nothing Then .Close
    Set rstProdutos = Nothing
  End With
End Sub

Private Sub cboSubClasse_CloseUp()
  cboSubClasse.Text = cboSubClasse.Columns(0).Text
  cboSubClasse_LostFocus
End Sub

Private Sub cboSubClasse_LostFocus()
  Dim rstSubClasses As Recordset
  
  txtNomeSubClasse.Text = ""
  If Not IsNumeric(cboSubClasse.Text) Then Exit Sub
  
  Set rstSubClasses = db.OpenRecordset("SELECT Código, Nome FROM [Sub Classes] WHERE Código = " & cboSubClasse.Text, dbOpenSnapshot)
  
  With rstSubClasses
    If Not (.BOF And .EOF) Then
      txtNomeSubClasse.Text = .Fields("Nome") & ""
    End If
    
    If Not rstSubClasses Is Nothing Then .Close
    Set rstSubClasses = Nothing
  End With
End Sub

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinal.Text = frmCalendario.gsDateCalender(mskDataFinal.Text)
  End If
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal.Text = Ajusta_Data(mskDataFinal.Text)
End Sub

Private Sub mskDataInicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicio.Text = frmCalendario.gsDateCalender(mskDataInicio.Text)
  End If
End Sub

Private Sub mskDataInicio_LostFocus()
  mskDataInicio.Text = Ajusta_Data(mskDataInicio.Text)
End Sub

Private Sub txtFator_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub
