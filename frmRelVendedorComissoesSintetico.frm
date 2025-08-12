VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelVendedorComissoesSintetico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendas por Vendedor e Comissões - Rel. Sintético"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelVendedorComissoesSintetico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   6240
   Begin VB.Frame Frame4 
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
      Height          =   735
      Left            =   0
      TabIndex        =   20
      Top             =   1770
      Width           =   3315
      Begin VB.OptionButton optImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   855
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1500
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Frame fraO 
      Caption         =   "Ordenação (Ranking)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3360
      TabIndex        =   19
      Top             =   1040
      Width           =   2805
      Begin VB.OptionButton optValor 
         Caption         =   "Valor "
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton optQtItensVendidos 
         Caption         =   "Quantidade de Ítens Vendidos"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1245
         Width           =   2535
      End
      Begin VB.OptionButton optCodVendedor 
         Caption         =   "Código do Vendedor"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optQtOperacoes 
         Caption         =   "Quantidade de Operações"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   930
         Width           =   2295
      End
      Begin VB.OptionButton optNomeVendedor 
         Caption         =   "Nome do Vendedor"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   675
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   1040
      Width           =   3315
      Begin MSMask.MaskEdBox mskDataIni 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Ao teclar [F2] carrega calendário"
         Top             =   280
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
      Begin MSMask.MaskEdBox mskDataFim 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         ToolTipText     =   "Ao teclar [F2] carrega calendário"
         Top             =   285
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "|-----|"
         Height          =   195
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Data datFilial 
      Caption         =   "datFilial"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Data datFuncionarios 
      Caption         =   "datFuncionarios"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários  WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Código"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   0
      TabIndex        =   13
      Top             =   -50
      Width           =   6180
      Begin SSDataWidgets_B.SSDBCombo cboVendedor 
         Bindings        =   "frmRelVendedorComissoesSintetico.frx":058A
         DataSource      =   "datFuncionarios"
         Height          =   315
         Left            =   975
         TabIndex        =   1
         Top             =   600
         Width           =   735
         DataFieldList   =   "Código"
         _Version        =   196617
         Cols            =   2
         Columns(0).Width=   3200
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelVendedorComissoesSintetico.frx":05A8
         DataSource      =   "datFilial"
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   735
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
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Filial"
      End
      Begin VB.Label lblNomeFilial 
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
         Height          =   315
         Left            =   1815
         TabIndex        =   17
         Top             =   240
         Width           =   4200
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   300
         Width           =   300
      End
      Begin VB.Label lblNomeVendedor 
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
         Height          =   315
         Left            =   1815
         TabIndex        =   15
         Top             =   600
         Width           =   4200
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   660
         Width           =   690
      End
   End
   Begin Crystal.CrystalReport crpRel 
      Left            =   4080
      Top             =   4560
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
End
Attribute VB_Name = "frmRelVendedorComissoesSintetico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboVendedor_CloseUp()
  cboVendedor.Text = cboVendedor.Columns(0).Text
  cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_LostFocus()
  Dim rstFuncionarios As Recordset
  
  lblNomeVendedor.Caption = ""
  If Not IsNumeric(cboVendedor.Text) Then Exit Sub
  
  Set rstFuncionarios = db.OpenRecordset("SELECT Código, Nome FROM Funcionários WHERE Código = " & CInt(cboVendedor.Text), dbOpenSnapshot)
  
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      lblNomeVendedor.Caption = .Fields("Nome") & ""
    End If
    
    If Not rstFuncionarios Is Nothing Then .Close
    Set rstFuncionarios = Nothing
  End With

End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  
  On Error GoTo ErrHandler

  If ValidarCampos Then Exit Sub
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde, gerando informações...")
  Call GerarInformacoes
  Call MontarRelatorio
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFilial.DatabaseName = gsQuickDBFileName
  datFuncionarios.DatabaseName = gsQuickDBFileName
  
  mskDataIni.Text = Format(Date, "DD/MM/YYYY")
  mskDataFim.Text = Format(Date, "DD/MM/YYYY")
  
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

Private Sub mskDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFim.Text = frmCalendario.gsDateCalender(mskDataFim.Text)
  End If
End Sub

Private Sub mskDataFim_LostFocus()
  mskDataFim.Text = Ajusta_Data(mskDataFim.Text)
End Sub

Private Sub mskDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataIni.Text = frmCalendario.gsDateCalender(mskDataIni.Text)
  End If
End Sub

Private Sub mskDataIni_LostFocus()
  mskDataIni.Text = Ajusta_Data(mskDataIni.Text)
End Sub

Private Function ValidarCampos() As Boolean
  'Filial
  If Len(lblNomeFilial.Caption) <= 0 Then
    ValidarCampos = True
    MsgBox "Filial inválida, verifique.", vbExclamation, "Atenção"
    cboFilial.SetFocus
    Exit Function
  End If
  
  'Data Ini
  If Not IsDate(mskDataIni.Text) Then
    ValidarCampos = True
    MsgBox "Data inicial inválida, verifique.", vbExclamation, "Atenção"
    mskDataIni.SetFocus
    Exit Function
  End If
  
  'Data Fim
  If Not IsDate(mskDataFim.Text) Then
    ValidarCampos = True
    MsgBox "Data final inválida, verifique.", vbExclamation, "Atenção"
    mskDataFim.SetFocus
    Exit Function
  End If
  
  'Verificação de datas
  If CDate(mskDataIni.Text) > CDate(mskDataFim.Text) Then
    ValidarCampos = True
    MsgBox "Data inicial maior que a final, verifique.", vbExclamation, "Atenção"
    mskDataIni.SetFocus
    Exit Function
  End If
  
End Function

Private Sub GerarInformacoes()
  Dim rstComissoes     As Recordset
  Dim rstControlVendas As Recordset
  Dim strSQL           As String
  Dim dblTotal         As Double

  
  'Tratamento para a table temporária
  dbTemp.Execute "DELETE * FROM ControlVendas"
  Set rstControlVendas = dbTemp.OpenRecordset("ControlVendas", dbOpenDynaset)
  'Fim Tratamento

  strSQL = "SELECT Vendedor, SUM(Qtde) AS Qtde_Itens, SUM(Valor) AS ValorTot, SUM(Comissão) AS TotComissao "
  strSQL = strSQL & " FROM Comissão "
  strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Data >= #" & Format(mskDataIni.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Data <= #" & Format(mskDataFim.Text, "MM/DD/YYYY") & "#"
  If Len(lblNomeVendedor.Caption) > 0 Then strSQL = strSQL & " AND Vendedor = " & (cboVendedor.Text)
  strSQL = strSQL & " GROUP BY Vendedor "
  
  Set rstComissoes = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstComissoes
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        rstControlVendas.AddNew
        rstControlVendas.Fields("Codigo").Value = .Fields("Vendedor").Value
        rstControlVendas.Fields("Nome").Value = GetNomeVendedor(.Fields("Vendedor").Value) & ""
        rstControlVendas.Fields("QtdeOpera").Value = GetQtdeSequencias(.Fields("Vendedor").Value) 'Trata-se da qtde de sequências feitas pelo vendedor
        rstControlVendas.Fields("QtdeItens").Value = .Fields("Qtde_Itens").Value
        dblTotal = Format(.Fields("ValorTot").Value - (GetDescontoSubTotal(.Fields("Vendedor").Value)), FORMAT_VALUE)
        rstControlVendas.Fields("Valor").Value = dblTotal
        rstControlVendas.Fields("Comissao").Value = IIf(IsNumeric(.Fields("TotComissao").Value), .Fields("TotComissao").Value, 0)
        rstControlVendas.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With

  Set rstComissoes = Nothing

  rstControlVendas.Close
  Set rstControlVendas = Nothing
  
End Sub

Private Function GetDescontoSubTotal(ByVal CodFunc As Integer) As Double
  Dim rstDescSubTotal As Recordset
  Dim strSQL          As String
  
  strSQL = "SELECT Sum(DescontoSubTotal) AS Total "
  strSQL = strSQL & " FROM Saídas "
  strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Digitador = " & CodFunc
  strSQL = strSQL & " AND Data BETWEEN #" & Format(mskDataIni.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND #" & Format(mskDataFim.Text, "MM/DD/YYYY") & "#"

  Set rstDescSubTotal = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstDescSubTotal
    Call IsDataType(dtDouble, .Fields("Total").Value, GetDescontoSubTotal)
    If Not rstDescSubTotal Is Nothing Then .Close
    Set rstDescSubTotal = Nothing
  End With

End Function

'24/04/2009 - mpdea
'Corrigido quantidade de sequências
'Nota: no access não é possível fazer distinct count
'e na versão 97 nem por subconsulta no SELECT
'http://blogs.msdn.com/access/archive/2007/09/19/writing-a-count-distinct-query-in-access.aspx
Private Function GetQtdeSequencias(ByVal CodFunc As Integer) As Integer
  Dim rstComissoes As Recordset
  Dim strSQL       As String
  Dim int_total As Integer
  
  strSQL = "SELECT DISTINCT Sequência "
  strSQL = strSQL & "FROM Comissão "
  strSQL = strSQL & "WHERE Filial = " & CByte(cboFilial.Text) & " "
  strSQL = strSQL & "AND Data >= #" & Format(mskDataIni.Text, "MM/DD/YYYY") & "# "
  strSQL = strSQL & "AND Data <= #" & Format(mskDataFim.Text, "MM/DD/YYYY") & "# "
  strSQL = strSQL & "AND Vendedor = " & CodFunc
  
  Set rstComissoes = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rstComissoes
    If Not (.BOF And .EOF) Then
      .MoveLast
      int_total = .RecordCount
    End If
    'Call IsDataType(dtLong, .Fields("Qtde").Value, GetQtdeSequencias)
    .Close
  End With
  Set rstComissoes = Nothing
  
  'Retorna quantidade
  GetQtdeSequencias = int_total
  
End Function

Private Function GetNomeVendedor(ByVal CodFunc As Integer) As String
  Dim rstFuncionarios As Recordset
  
  GetNomeVendedor = ""
  
  Set rstFuncionarios = db.OpenRecordset("SELECT Nome FROM Funcionários WHERE Código = " & CodFunc, dbOpenSnapshot)
  
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      .MoveFirst
      GetNomeVendedor = .Fields("Nome").Value & ""
    End If
    .Close
  End With

  Set rstFuncionarios = Nothing
  
End Function

Private Sub MontarRelatorio()
  Dim strReport As String
   
  'Nome do arquivo .rpt
  strReport = gsReportPath & "rptVendedorComissoesSintetico.rpt"
  
  With crpRel
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crpRel
    
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsTempDBFileName
    
    '.SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .Formulas(1) = "Periodo = '" & "Período: " & (mskDataIni.Text) & " à " & (mskDataFim.Text) & "'"
    If optCodVendedor.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Código do Vendedor" & "'"
    If optNomeVendedor.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Nome do Vendedor" & "'"
    If optQtOperacoes.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Quantidade de Operações" & "'"
    If optQtItensVendidos.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Quantidade de Ítens" & "'"
    If optValor.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Valor das Vendas" & "'"
    
    '12/05/2005 - Daniel
    'Correção para exibição dos botões de Configuração
    'de Impressoras e Botão de Pesquisas
    .WindowShowPrintSetupBtn = True
    .WindowShowSearchBtn = True
    
    'Ordenação
    If optCodVendedor.Value Then .SortFields(0) = "+{ControlVendas.Codigo}"
    If optNomeVendedor.Value Then .SortFields(0) = "+{ControlVendas.Nome}"
    If optQtOperacoes.Value Then .SortFields(0) = "-{ControlVendas.QtdeOpera}"
    If optQtItensVendidos.Value Then .SortFields(0) = "-{ControlVendas.QtdeItens}"
    If optValor.Value Then .SortFields(0) = "-{ControlVendas.Valor}"
    
    .WindowState = crptMaximized
    .Destination = IIf(optVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crpRel)
  
    .Action = 1
  End With
  
End Sub
