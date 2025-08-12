VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelVendasPorFornecedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rel. de Vendas por Fornecedor"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelVendasPorFornecedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   7320
   Begin VB.Frame Frame7 
      Caption         =   "Oper. de Entrada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      TabIndex        =   36
      Top             =   4080
      Width           =   1935
      Begin VB.OptionButton optEmprestimo 
         Caption         =   "Empréstimo"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optCompras 
         Caption         =   "Compras"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Vendas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   2100
      TabIndex        =   35
      Top             =   4080
      Width           =   1935
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optSomenteVendidos 
         Caption         =   "Somente Vendidos"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   320
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Data datClasses 
      Caption         =   "datClasses"
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
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Classes"
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Caption         =   "Período das Saídas"
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
      Left            =   3720
      TabIndex        =   30
      Top             =   3360
      Width           =   3495
      Begin MSMask.MaskEdBox mskDataFinalSaidas 
         Height          =   315
         Left            =   2160
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
      Begin MSMask.MaskEdBox mskDataInicioSaidas 
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
      Begin VB.Label Label5 
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   1800
         TabIndex        =   31
         Top             =   240
         Width           =   300
      End
   End
   Begin Crystal.CrystalReport crtRelFornecedor 
      Left            =   3360
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data datFornecedor 
      Caption         =   "datFornecedor"
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
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' ORDER BY Nome"
      Top             =   5880
      Width           =   1455
   End
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
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0' ORDER BY Nome"
      Top             =   5880
      Width           =   1575
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Imprimir"
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
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   1575
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
      Height          =   1005
      Left            =   4080
      TabIndex        =   29
      Top             =   4080
      Width           =   1455
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   340
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período das Entradas"
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
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   3495
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
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
         TabIndex        =   4
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
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label3 
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
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
      Height          =   1890
      Left            =   120
      TabIndex        =   19
      Top             =   1455
      Width           =   7095
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1040
         Width           =   4455
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4455
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   640
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelVendasPorFornecedor.frx":058A
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
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
      Begin SSDataWidgets_B.SSDBCombo cboProduto 
         Bindings        =   "frmRelVendasPorFornecedor.frx":05A6
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   640
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
         Bindings        =   "frmRelVendasPorFornecedor.frx":05C0
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
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmRelVendasPorFornecedor.frx":05D9
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1040
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Classe"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1100
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   700
         Width           =   570
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
      Height          =   1575
      Left            =   0
      TabIndex        =   16
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "• Caso não queira utilizar algum filtro, basta não preencher o campo"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Vendas por Fornecedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   240
         Picture         =   "frmRelVendasPorFornecedor.frx":05F2
         Top             =   240
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmRelVendasPorFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim rstEntradas      As Recordset
  Dim rstRelFornecedor As Recordset
  Dim strSQL           As String
  Dim dblQtdeVendida   As Double
  Dim dblPrecoVendido  As Double
  Dim sngEstoqueAtual  As Single
  
  '25/11/2004 - Daniel
  'Criado validação para a Resultado
  'não utilizar mais este relatório
  If CheckSerialCaseMod("QS40590-987") Then
    MsgBox "Relatório não disponível.", vbInformation, "Quick Store"
    Exit Sub
  End If
  
  If Not ValidarDados Then Exit Sub
  
  Call StatusMsg("Pesquisando no banco de dados...")
  Screen.MousePointer = vbHourglass

  dbTemp.Execute "DELETE * FROM tblRelFornecedorTemp"
  dbTemp.Execute "DELETE * FROM tblRelFornecedor"

  Set rstRelFornecedor = dbTemp.OpenRecordset("tblRelFornecedorTemp", dbOpenDynaset)
  
  strSQL = "SELECT Entradas.*, [Entradas - Produtos].*, [Operações Entrada].Código AS CodOpera, [Operações Entrada].Tipo, Produtos.Código AS CodProduto "
  strSQL = strSQL & " FROM Entradas, [Entradas - Produtos], [Operações Entrada], Produtos "
  strSQL = strSQL & " WHERE Entradas.Operação = [Operações Entrada].Código "
  
  If optEmprestimo.Value Then strSQL = strSQL & " AND [Operações Entrada].Tipo = 'E' "
  If optCompras.Value Then strSQL = strSQL & " AND [Operações Entrada].Tipo = 'C' "
  
  strSQL = strSQL & " AND Entradas.Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Entradas.Data >= #" & Format(mskDataInicio.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Entradas.Data <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND [Entradas - Produtos].Sequência = Entradas.Sequência "
  
  strSQL = strSQL & " AND Produtos.Código = [Entradas - Produtos].Código "
  
  If Len(txtNomeFornecedor.Text) > 0 Then
    strSQL = strSQL & " AND Entradas.Fornecedor = " & CLng(cboFornecedor.Text)
  End If
  
  'strSQL = strSQL & " AND [Entradas - Produtos].Sequência = Entradas.Sequência "
  
  If Len(txtNomeProduto.Text) > 0 Then
    strSQL = strSQL & " AND [Entradas - Produtos].Código = '" & Trim(cboProduto.Text) & "'"
  End If

  If Len(txtNomeClasse.Text) > 0 Then
    strSQL = strSQL & " AND Produtos.Classe = " & CInt(cboClasse.Text)
  End If

  strSQL = strSQL & " ORDER BY Entradas.Sequência "

  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
    
      Do Until .EOF
        
        rstRelFornecedor.AddNew
        rstRelFornecedor.Fields("Fornecedor").Value = .Fields("Fornecedor").Value
        rstRelFornecedor.Fields("Produto").Value = .Fields("Código").Value
        rstRelFornecedor.Fields("QtdeEntrada").Value = .Fields("Qtde").Value
        rstRelFornecedor.Fields("PrecoCusto").Value = .Fields("Preço").Value
        rstRelFornecedor.Update
        
        .MoveNext
      Loop
    
    End If
    .Close
  End With

  Set rstEntradas = Nothing
  
  rstRelFornecedor.Close
  Set rstRelFornecedor = Nothing

  '-----------------------------------------------------------------
  ' Checar em Saídas
  '-----------------------------------------------------------------
  Set rstRelFornecedor = dbTemp.OpenRecordset("tblRelFornecedorTemp", dbOpenDynaset)
  
  With rstRelFornecedor
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        Call ChecarEmSaidas(.Fields("Produto").Value, dblQtdeVendida, dblPrecoVendido)
      
        .Edit
        .Fields("QtdeVendida").Value = dblQtdeVendida
        '.Fields("EstoqueAtual").Value = (.Fields("QtdeEntrada").Value) - dblQtdeVendida
        Call BuscarEstoqueAtual(.Fields("Produto").Value, sngEstoqueAtual)
        .Fields("EstoqueAtual").Value = sngEstoqueAtual
        .Fields("PrestacaoContas").Value = Format(((.Fields("PrecoCusto").Value) * dblQtdeVendida), FORMAT_VALUE)
        .Fields("Venda").Value = dblPrecoVendido
        .Update
        
        dblQtdeVendida = 0
        dblPrecoVendido = 0
        
      .MoveNext
      Loop
      
    End If
    .Close
  End With

  Set rstRelFornecedor = Nothing
  
  If optSomenteVendidos.Value Then dbTemp.Execute "DELETE * FROM tblRelFornecedorTemp WHERE QtdeVendida = 0"
  
  '-----------------------------------------------------------------
  ' Agrupar os valores para evitar redundâncias
  '-----------------------------------------------------------------
  Call AgruparValores
  
  '-----------------------------------------------------------------
  ' Montando o relatório
  '-----------------------------------------------------------------
  Call MontarRelatorio

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFiliais.DatabaseName = gsQuickDBFileName
  datProdutos.DatabaseName = gsQuickDBFileName
  datClasses.DatabaseName = gsQuickDBFileName
  datFornecedor.DatabaseName = gsQuickDBFileName
  
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
  
  datFornecedor.Recordset.FindFirst "Código = " & cboFornecedor.Text
  
  If Not datFornecedor.Recordset.NoMatch Then
    txtNomeFornecedor.Text = datFornecedor.Recordset.Fields("Nome") & ""
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

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinal.Text = frmCalendario.gsDateCalender(mskDataFinal.Text)
  End If
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal.Text = Ajusta_Data(mskDataFinal.Text)
End Sub

Private Sub mskDataFinalSaidas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinalSaidas.Text = frmCalendario.gsDateCalender(mskDataFinalSaidas.Text)
  End If
End Sub

Private Sub mskDataFinalSaidas_LostFocus()
  mskDataFinalSaidas.Text = Ajusta_Data(mskDataFinalSaidas.Text)
End Sub

Private Sub mskDataInicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicio.Text = frmCalendario.gsDateCalender(mskDataInicio.Text)
  End If
End Sub

Private Sub mskDataInicio_LostFocus()
  mskDataInicio.Text = Ajusta_Data(mskDataInicio.Text)
End Sub

Private Function ValidarDados() As Boolean
  ValidarDados = True
  
  If Len(txtNomeFilial.Text) <= 0 Then
    ValidarDados = False
    MsgBox "Filial inválida, verifique", vbExclamation, "Quick Store"
    cboFilial.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataInicio.Text) Then
    ValidarDados = False
    MsgBox "Data Inicial das Entradas inválida, verifique.", vbExclamation, "Quick Store"
    mskDataInicio.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataFinal.Text) Then
    ValidarDados = False
    MsgBox "Data Final das Entradas inválida, verifique.", vbExclamation, "Quick Store"
    mskDataFinal.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataInicioSaidas.Text) Then
    ValidarDados = False
    MsgBox "Data Inicial das Saídas inválida, verifique.", vbExclamation, "Quick Store"
    mskDataInicioSaidas.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataFinalSaidas.Text) Then
    ValidarDados = False
    MsgBox "Data Final das Saídas inválida, verifique.", vbExclamation, "Quick Store"
    mskDataFinalSaidas.SetFocus
    Exit Function
  End If
  
  
End Function

Private Sub ChecarEmSaidas(ByVal Produto As String, ByRef QtdeVendida As Double, ByRef PrecoVendido As Double)
  Dim rstSaidas As Recordset
  Dim strSQL    As String

  strSQL = "SELECT * FROM Saídas, [Saídas - Produtos]"
  strSQL = strSQL & " WHERE Saídas.Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Saídas.Data >= #" & Format(mskDataInicioSaidas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Saídas.Data <= #" & Format(mskDataFinalSaidas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Saídas.Efetivada "
  '09/10/2004 - Daniel
  strSQL = strSQL & " AND Saídas.Recebimento "
  strSQL = strSQL & " AND NOT Saídas.[Nota Cancelada] "
  strSQL = strSQL & " AND [Saídas - Produtos].Sequência = Saídas.Sequência "
  strSQL = strSQL & " ORDER BY Saídas.Sequência "
  
  Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstSaidas
    If Not (.BOF And .EOF) Then
      .MoveFirst
    
      Do Until .EOF
        
        If .Fields("Código").Value = Produto Then
          QtdeVendida = QtdeVendida + .Fields("Qtde").Value
          PrecoVendido = Format(PrecoVendido + .Fields("Preço Final").Value, FORMAT_VALUE)
        End If
      
        .MoveNext
      Loop
    
    End If
    .Close
  End With

  Set rstSaidas = Nothing

End Sub

Private Sub MontarRelatorio()
  Dim strReport As String
  
  'Nome do arquivo .rpt
  strReport = gsReportPath & "rptRelVendasFornecedor.rpt"
  
  With crtRelFornecedor
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crtRelFornecedor
    
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsTempDBFileName
    .DataFiles(3) = gsTempDBFileName
    
    '.SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .SortFields(0) = "+{tblRelFornecedor.Produto}" 'Ordenação
    
    .WindowState = crptMaximized
    .Destination = IIf(optSaidaVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crtRelFornecedor)
  
    .Action = 1
  End With

  Screen.MousePointer = vbDefault
  
  Call StatusMsg("")

End Sub

Private Sub mskDataInicioSaidas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicioSaidas.Text = frmCalendario.gsDateCalender(mskDataInicioSaidas.Text)
  End If
End Sub

Private Sub mskDataInicioSaidas_LostFocus()
  mskDataInicioSaidas.Text = Ajusta_Data(mskDataInicioSaidas.Text)
End Sub

Private Sub BuscarEstoqueAtual(ByVal Produto As String, ByRef EstoqueAtual As Single)
  Dim rstEstoqueFinal As Recordset
  Dim strSQL          As String
  
  EstoqueAtual = 0
  
  strSQL = "SELECT [Estoque Atual] FROM [Estoque Final]"
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  strSQL = strSQL & " AND Produto = '" & Produto & "'"
  
  Set rstEstoqueFinal = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEstoqueFinal
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      EstoqueAtual = .Fields("Estoque Atual").Value
    End If
    .Close
  End With
  
  Set rstEstoqueFinal = Nothing
  
End Sub

Private Sub AgruparValores()
  Dim rsttblRelFornecedor     As Recordset
  Dim rsttblRelFornecedorTemp As Recordset
  Dim strSQL                  As String
  
  strSQL = "SELECT Fornecedor, Produto, QtdeEntrada, PrecoCusto, QtdeVendida, EstoqueAtual, PrestacaoContas, Venda "
  strSQL = strSQL & " FROM tblRelFornecedorTemp "
  strSQL = strSQL & " GROUP BY Fornecedor, Produto, QtdeEntrada, PrecoCusto, QtdeVendida, EstoqueAtual, PrestacaoContas, Venda "
  
  Set rsttblRelFornecedorTemp = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
  Set rsttblRelFornecedor = dbTemp.OpenRecordset("tblRelFornecedor", dbOpenDynaset)
  
  With rsttblRelFornecedorTemp
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        rsttblRelFornecedor.AddNew
          rsttblRelFornecedor.Fields("Fornecedor").Value = .Fields("Fornecedor").Value
          rsttblRelFornecedor.Fields("Produto").Value = .Fields("Produto").Value
          rsttblRelFornecedor.Fields("QtdeEntrada").Value = .Fields("QtdeEntrada").Value
          rsttblRelFornecedor.Fields("PrecoCusto").Value = .Fields("PrecoCusto").Value
          rsttblRelFornecedor.Fields("QtdeVendida").Value = .Fields("QtdeVendida").Value
          rsttblRelFornecedor.Fields("EstoqueAtual").Value = .Fields("EstoqueAtual").Value
          rsttblRelFornecedor.Fields("PrestacaoContas").Value = .Fields("PrestacaoContas").Value
          rsttblRelFornecedor.Fields("Venda").Value = .Fields("Venda").Value
        rsttblRelFornecedor.Update
      
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rsttblRelFornecedorTemp = Nothing
  
  rsttblRelFornecedor.Close
  Set rsttblRelFornecedor = Nothing

End Sub
