VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelPrestacaoContas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rel. de Prestação de Contas com Fornecedores"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelPrestacaoContas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7080
   Begin VB.Frame fraNF 
      Caption         =   "Nota Fiscal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   0
      TabIndex        =   25
      Top             =   3360
      Width           =   3495
      Begin VB.TextBox txtNFFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   5
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtNFIni 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   480
         MaxLength       =   8
         TabIndex        =   4
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   1800
         TabIndex        =   27
         Top             =   390
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   390
         Width           =   255
      End
   End
   Begin VB.Frame fraProd 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3550
      TabIndex        =   24
      Top             =   2520
      Width           =   1695
      Begin VB.OptionButton optVendidos 
         Caption         =   "Vendidos"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optNaoVendidos 
         Caption         =   "Não Vendidos"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
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
      Height          =   810
      Left            =   0
      TabIndex        =   21
      Top             =   2520
      Width           =   3495
      Begin MSMask.MaskEdBox mskDataFinalSaidas 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   360
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
         TabIndex        =   2
         Top             =   360
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
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   1800
         TabIndex        =   22
         Top             =   420
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   3720
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
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
      Height          =   810
      Left            =   5400
      TabIndex        =   20
      Top             =   2520
      Width           =   1695
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   225
         Value           =   -1  'True
         Width           =   855
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
      Height          =   1050
      Left            =   0
      TabIndex        =   15
      Top             =   1440
      Width           =   7095
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelPrestacaoContas.frx":058A
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
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelPrestacaoContas.frx":05A6
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   300
      End
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
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' ORDER BY Nome"
      Top             =   4800
      Width           =   2415
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   4800
      Width           =   2415
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
      TabIndex        =   12
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelPrestacaoContas.frx":05BF
         ForeColor       =   &H00808080&
         Height          =   855
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Prestação de Contas com o Fornecedor"
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
         TabIndex        =   13
         Top             =   240
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   1170
         Left            =   360
         Picture         =   "frmRelPrestacaoContas.frx":0662
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1290
      End
   End
   Begin Crystal.CrystalReport crtRelPrestacao 
      Left            =   5280
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRelPrestacaoContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()

  If Not ValidarDados Then Exit Sub
  
  dbTemp.Execute "DELETE * FROM PrestacaoContasTemp"
  dbTemp.Execute "DELETE * FROM PrestacaoContas"
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde gerando o relatório...")
  
  Call CriarPrestacaoContasTemp
  Call AgruparValoresECriarPrestacaoContas
  Call AtualizarQtdeAcertada
  Call AtualizarQtdeDevolvida
  
  '15/12/2004 - Daniel
  'Atualização do Custo devido mudanças no processo
  Call AtualizarCusto
  
  Call CriarRelatorio
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFiliais.DatabaseName = gsQuickDBFileName
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

Private Sub mskDataFinalSaidas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinalSaidas.Text = frmCalendario.gsDateCalender(mskDataFinalSaidas.Text)
  End If
End Sub

Private Sub mskDataFinalSaidas_LostFocus()
  mskDataFinalSaidas.Text = Ajusta_Data(mskDataFinalSaidas.Text)
End Sub
Private Sub mskDataInicioSaidas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicioSaidas.Text = frmCalendario.gsDateCalender(mskDataInicioSaidas.Text)
  End If
End Sub

Private Sub mskDataInicioSaidas_LostFocus()
  mskDataInicioSaidas.Text = Ajusta_Data(mskDataInicioSaidas.Text)
End Sub

Private Function ValidarDados() As Boolean
  ValidarDados = True
  
  If Len(txtNomeFilial.Text) <= 0 Then
    ValidarDados = False
    MsgBox "Filial inválida, verifique", vbExclamation, "Quick Store"
    cboFilial.SetFocus
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

Private Sub CriarPrestacaoContasTemp()
  Dim rstPrestacao        As Recordset
  Dim rstPrestacaoTemp    As Recordset
  Dim strSQL              As String
  Dim blnMaisdeUmRegistro As Boolean
  Dim bytQtdeRegistros    As Byte
  
  strSQL = "SELECT * FROM PrestacaoContas"
  strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
  'PeriodoVenda
  strSQL = strSQL & " AND PeriodoVenda >= #" & Format(mskDataInicioSaidas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND PeriodoVenda <= #" & Format(mskDataFinalSaidas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Finalizado "
  
  If Len(txtNomeFornecedor.Text) > 0 Then strSQL = strSQL & " AND Fornecedor = " & CLng(cboFornecedor.Text)
  
  '18/10/2004 - Daniel
  'Adicionado filtro de notas fiscais
  If Len(txtNFIni.Text) > 0 And Len(txtNFFin.Text) > 0 Then
    If CLng(txtNFIni.Text) <= CLng(txtNFFin.Text) Then
      strSQL = strSQL & " AND NotaFiscal >= " & CLng(txtNFIni.Text)
      strSQL = strSQL & " AND NotaFiscal <= " & CLng(txtNFFin.Text)
    End If
  End If
  
  If optVendidos.Value Then
    strSQL = strSQL & " AND QtdeVendida <> 0 "
  Else
    strSQL = strSQL & " AND QtdeVendida = 0 "
  End If
  
  strSQL = strSQL & " ORDER BY Fornecedor, Filial, Sequencia, Linha "
  
  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        Call VerificarQtdeRegistros(.Fields("Filial").Value, .Fields("Fornecedor").Value, .Fields("Sequencia").Value, .Fields("Linha").Value, blnMaisdeUmRegistro)
      
        If Not blnMaisdeUmRegistro Then 'Criamos o registro
        
          Set rstPrestacaoTemp = dbTemp.OpenRecordset("PrestacaoContasTemp", dbOpenDynaset)
          
            rstPrestacaoTemp.AddNew
              rstPrestacaoTemp.Fields("Filial").Value = .Fields("Filial").Value
              rstPrestacaoTemp.Fields("Fornecedor").Value = .Fields("Fornecedor").Value
              rstPrestacaoTemp.Fields("Sequencia").Value = .Fields("Sequencia").Value
              rstPrestacaoTemp.Fields("Linha").Value = .Fields("Linha").Value
              rstPrestacaoTemp.Fields("Produto").Value = .Fields("Produto").Value
              rstPrestacaoTemp.Fields("Custo").Value = .Fields("Custo").Value
              rstPrestacaoTemp.Fields("QtdeOriginal").Value = .Fields("QtdeOriginal").Value
              rstPrestacaoTemp.Fields("QtdeDevolvida").Value = .Fields("QtdeDevolvida").Value
              rstPrestacaoTemp.Fields("QtdeVendida").Value = .Fields("QtdeVendida").Value
              rstPrestacaoTemp.Fields("QtdeComprada").Value = .Fields("QtdeComprada").Value
              rstPrestacaoTemp.Fields("DatadaGeracao").Value = .Fields("DatadaGeracao").Value
              rstPrestacaoTemp.Fields("Finalizado").Value = .Fields("Finalizado").Value
              rstPrestacaoTemp.Fields("DatadaFinalizacao").Value = .Fields("DatadaFinalizacao").Value
              rstPrestacaoTemp.Fields("ImpressoNF").Value = .Fields("ImpressoNF").Value
              rstPrestacaoTemp.Fields("Resultado").Value = .Fields("Resultado").Value
              rstPrestacaoTemp.Fields("PrestacaoFechada").Value = .Fields("PrestacaoFechada").Value
              rstPrestacaoTemp.Fields("CompraFechada").Value = .Fields("CompraFechada").Value
              rstPrestacaoTemp.Fields("PeriodoVenda").Value = .Fields("PeriodoVenda").Value
              rstPrestacaoTemp.Fields("NotaFiscal").Value = .Fields("NotaFiscal").Value
              rstPrestacaoTemp.Fields("QtdeAcertada").Value = .Fields("QtdeAcertada").Value
            rstPrestacaoTemp.Update
            rstPrestacaoTemp.Close
          
          Set rstPrestacaoTemp = Nothing
        
        End If 'If Not blnMaisdeUmRegistro
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstPrestacao = Nothing

End Sub

Private Sub VerificarQtdeRegistros(ByVal Filial As Byte, ByVal Fornecedor As Long, ByVal Sequencia As Long, ByVal Linha As Byte, ByRef MaisdeUm As Boolean)
  Dim rstPrestacao     As Recordset
  Dim rstPrestacaoTemp As Recordset
  Dim strSQL           As String
  Dim bytAuxi          As Byte
  Dim bytQtdeReg       As Byte
  
  strSQL = "SELECT * FROM PrestacaoContas"
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Fornecedor = " & Fornecedor
  strSQL = strSQL & " AND Sequencia = " & Sequencia
  strSQL = strSQL & " AND Linha = " & Linha
  strSQL = strSQL & " ORDER BY QtdeVendida "
  
  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstPrestacao.RecordCount <> 0 Then
    rstPrestacao.MoveFirst
    rstPrestacao.MoveLast
    rstPrestacao.MoveFirst
  End If
  
  If rstPrestacao.RecordCount = 1 Then
    MaisdeUm = False
    rstPrestacao.Close
    Set rstPrestacao = Nothing
    Exit Sub
  End If
  
  If rstPrestacao.RecordCount > 1 Then
    MaisdeUm = True
    bytQtdeReg = rstPrestacao.RecordCount
  End If
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do Until .EOF
  
        bytAuxi = bytAuxi + 1
        
        If bytAuxi = bytQtdeReg Then
            'Criamos com o último registro que é o mais atualizado
            Set rstPrestacaoTemp = dbTemp.OpenRecordset("PrestacaoContasTemp", dbOpenDynaset)
      
              rstPrestacaoTemp.AddNew
                rstPrestacaoTemp.Fields("Filial").Value = .Fields("Filial").Value
                rstPrestacaoTemp.Fields("Fornecedor").Value = .Fields("Fornecedor").Value
                rstPrestacaoTemp.Fields("Sequencia").Value = .Fields("Sequencia").Value
                rstPrestacaoTemp.Fields("Linha").Value = .Fields("Linha").Value
                rstPrestacaoTemp.Fields("Produto").Value = .Fields("Produto").Value
                rstPrestacaoTemp.Fields("Custo").Value = .Fields("Custo").Value
                rstPrestacaoTemp.Fields("QtdeOriginal").Value = .Fields("QtdeOriginal").Value
                rstPrestacaoTemp.Fields("QtdeDevolvida").Value = .Fields("QtdeDevolvida").Value
                rstPrestacaoTemp.Fields("QtdeVendida").Value = .Fields("QtdeVendida").Value
                rstPrestacaoTemp.Fields("QtdeComprada").Value = .Fields("QtdeComprada").Value
                rstPrestacaoTemp.Fields("DatadaGeracao").Value = .Fields("DatadaGeracao").Value
                rstPrestacaoTemp.Fields("Finalizado").Value = .Fields("Finalizado").Value
                rstPrestacaoTemp.Fields("DatadaFinalizacao").Value = .Fields("DatadaFinalizacao").Value
                rstPrestacaoTemp.Fields("ImpressoNF").Value = .Fields("ImpressoNF").Value
                rstPrestacaoTemp.Fields("Resultado").Value = .Fields("Resultado").Value
                rstPrestacaoTemp.Fields("PrestacaoFechada").Value = .Fields("PrestacaoFechada").Value
                rstPrestacaoTemp.Fields("CompraFechada").Value = .Fields("CompraFechada").Value
                rstPrestacaoTemp.Fields("PeriodoVenda").Value = .Fields("PeriodoVenda").Value
                rstPrestacaoTemp.Fields("NotaFiscal").Value = .Fields("NotaFiscal").Value
                rstPrestacaoTemp.Fields("QtdeAcertada").Value = .Fields("QtdeAcertada").Value
              rstPrestacaoTemp.Update
              rstPrestacaoTemp.Close
      
            Set rstPrestacaoTemp = Nothing

        End If

       .MoveNext
      Loop
      
    End If
    .Close
  End With

  Set rstPrestacao = Nothing

End Sub


Private Sub AgruparValoresECriarPrestacaoContas()
  Dim rstPrestacaoTemp As Recordset
  Dim rstPrestacao     As Recordset
  Dim strSQL           As String
  
  Set rstPrestacao = dbTemp.OpenRecordset("PrestacaoContas", dbOpenDynaset)
  
  strSQL = "SELECT Filial, Fornecedor, Sequencia, Linha, Produto, Custo, QtdeOriginal, QtdeDevolvida, QtdeVendida, QtdeComprada, DatadaGeracao, Finalizado, DatadaFinalizacao, ImpressoNF, Resultado, PrestacaoFechada, CompraFechada, PeriodoVenda, NotaFiscal, QtdeAcertada "
  strSQL = strSQL & " FROM PrestacaoContasTemp "
  strSQL = strSQL & " GROUP BY Filial, Fornecedor, Sequencia, Linha, Produto, Custo, QtdeOriginal, QtdeDevolvida, QtdeVendida, QtdeComprada,DatadaGeracao,Finalizado, DatadaFinalizacao, ImpressoNF, Resultado, PrestacaoFechada, CompraFechada, PeriodoVenda, NotaFiscal, QtdeAcertada"

  Set rstPrestacaoTemp = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)

  With rstPrestacaoTemp
    If Not (.BOF And .EOF) Then
      .MoveFirst
  
      Do Until .EOF
  
        rstPrestacao.AddNew
          rstPrestacao.Fields("Filial").Value = .Fields("Filial").Value
          rstPrestacao.Fields("Fornecedor").Value = .Fields("Fornecedor").Value
          rstPrestacao.Fields("Sequencia").Value = .Fields("Sequencia").Value
          rstPrestacao.Fields("Linha").Value = .Fields("Linha").Value
          rstPrestacao.Fields("Produto").Value = .Fields("Produto").Value
          rstPrestacao.Fields("Custo").Value = .Fields("Custo").Value
          rstPrestacao.Fields("QtdeOriginal").Value = .Fields("QtdeOriginal").Value
          rstPrestacao.Fields("QtdeDevolvida").Value = .Fields("QtdeDevolvida").Value
          rstPrestacao.Fields("QtdeVendida").Value = .Fields("QtdeVendida").Value
          rstPrestacao.Fields("QtdeComprada").Value = .Fields("QtdeComprada").Value
          rstPrestacao.Fields("DatadaGeracao").Value = .Fields("DatadaGeracao").Value
          rstPrestacao.Fields("Finalizado").Value = .Fields("Finalizado").Value
          rstPrestacao.Fields("DatadaFinalizacao").Value = .Fields("DatadaFinalizacao").Value
          rstPrestacao.Fields("ImpressoNF").Value = .Fields("ImpressoNF").Value
          rstPrestacao.Fields("Resultado").Value = .Fields("Resultado").Value
          rstPrestacao.Fields("PrestacaoFechada").Value = .Fields("PrestacaoFechada").Value
          rstPrestacao.Fields("CompraFechada").Value = .Fields("CompraFechada").Value
          rstPrestacao.Fields("PeriodoVenda").Value = .Fields("PeriodoVenda").Value
          rstPrestacao.Fields("NotaFiscal").Value = .Fields("NotaFiscal").Value
          rstPrestacao.Fields("QtdeAcertada").Value = .Fields("QtdeAcertada").Value
        rstPrestacao.Update
  
  
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstPrestacaoTemp = Nothing

  rstPrestacao.Close
  Set rstPrestacao = Nothing
  
End Sub

Private Sub CriarRelatorio()
  Dim strReport As String
  
  'Nome do arquivo .rpt
  strReport = gsReportPath & "rptPrestacaoContas.rpt"
  
  With crtRelPrestacao
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crtRelPrestacao
    
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsTempDBFileName
    .DataFiles(3) = gsTempDBFileName
    
    '.SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .Formulas(1) = "Periodo = '" & "Período de " & mskDataInicioSaidas.Text & " até " & mskDataFinalSaidas.Text & "'"
    .SortFields(0) = "+{PrestacaoContas.Fornecedor}" 'Ordenação
    .SortFields(1) = "+{PrestacaoContas.Sequencia}"
    
    .WindowState = crptMaximized
    .Destination = IIf(optSaidaVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crtRelPrestacao)
  
    .Action = 1
  End With

  Screen.MousePointer = vbDefault
  
  Call StatusMsg("")


End Sub

Private Sub AtualizarQtdeAcertada()
  Dim rstTemp      As Recordset
  Dim dblQtdeAcer  As Double

  Set rstTemp = dbTemp.OpenRecordset("PrestacaoContas", dbOpenDynaset)

  If rstTemp.RecordCount = 0 Then Exit Sub

  With rstTemp
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do Until .EOF

        Call BuscarQtdeAcertada(.Fields("Filial").Value, .Fields("Fornecedor").Value, .Fields("Sequencia").Value, .Fields("Linha").Value, .Fields("Produto").Value, dblQtdeAcer)

        .Edit
        .Fields("QtdeAcertada").Value = dblQtdeAcer
        .Update
        
       .MoveNext
      Loop

    End If
    .Close
  End With

  Set rstTemp = Nothing

End Sub

Private Sub BuscarQtdeAcertada(ByVal Filial As Byte, ByVal Fornecedor As Long, ByVal Seq As Long, ByVal Linha As Byte, ByVal Produto As String, ByRef QtdeAcertada As Double)
  Dim rstPrestacao As Recordset
  Dim strSQL       As String

  strSQL = "SELECT SUM(QtdeAcertada) AS TotalQtdeAcertada FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Fornecedor = " & Fornecedor
  strSQL = strSQL & " AND Sequencia = " & Seq
  strSQL = strSQL & " AND Linha = " & Linha
  strSQL = strSQL & " AND Produto = '" & Produto & "'"
  strSQL = strSQL & " AND PeriodoVenda >= #" & Format(mskDataInicioSaidas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND PeriodoVenda <= #" & Format(mskDataFinalSaidas.Text, "MM/DD/YYYY") & "#"
  
  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      QtdeAcertada = .Fields("TotalQtdeAcertada").Value
      
    End If
    .Close
  End With
  
  Set rstPrestacao = Nothing

End Sub

Private Sub AtualizarQtdeDevolvida()
  Dim rstTemp      As Recordset
  Dim dblQtdeDevo  As Double

  Set rstTemp = dbTemp.OpenRecordset("PrestacaoContas", dbOpenDynaset)

  If rstTemp.RecordCount = 0 Then Exit Sub

  With rstTemp
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do Until .EOF

        Call BuscarQtdeDevolvida(.Fields("Filial").Value, .Fields("Fornecedor").Value, .Fields("Sequencia").Value, .Fields("Linha").Value, .Fields("Produto").Value, dblQtdeDevo)

        .Edit
        .Fields("QtdeDevolvida").Value = dblQtdeDevo
        .Update
        
       .MoveNext
      Loop

    End If
    .Close
  End With

  Set rstTemp = Nothing

End Sub

Private Sub BuscarQtdeDevolvida(ByVal Filial As Byte, ByVal Fornecedor As Long, ByVal Seq As Long, ByVal Linha As Byte, ByVal Produto As String, ByRef QtdeDevolvida As Double)
  Dim rstPrestacao As Recordset
  Dim strSQL       As String

  strSQL = "SELECT SUM(QtdeDevolvida) AS TotalQtdeDevolvida FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Fornecedor = " & Fornecedor
  strSQL = strSQL & " AND Sequencia = " & Seq
  strSQL = strSQL & " AND Linha = " & Linha
  strSQL = strSQL & " AND Produto = '" & Produto & "'"
  
  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      QtdeDevolvida = .Fields("TotalQtdeDevolvida").Value
      
    End If
    .Close
  End With
  
  Set rstPrestacao = Nothing

End Sub

Private Sub AtualizarCusto()
  Dim rstPrestacao As Recordset
  
  Set rstPrestacao = dbTemp.OpenRecordset("PrestacaoContas", dbOpenDynaset)
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        .Edit
        .Fields("Custo").Value = Format((.Fields("Custo").Value * .Fields("QtdeAcertada").Value), FORMAT_VALUE)
        .Update
        
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstPrestacao = Nothing
  
End Sub

