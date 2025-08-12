VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelComissaoComRetencao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rel. de Comissões contemplando taxas de retenções sobre cartões"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelComissaoComRetencao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmRelComissaoComRetencao.frx":058A
   ScaleHeight     =   6150
   ScaleWidth      =   6930
   Begin VB.Data datFuncionarios 
      Caption         =   "datFuncionarios"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Código"
      Top             =   7200
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame fraSintetico 
      Caption         =   "Ordenação para o Sintético (Ranking)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      TabIndex        =   25
      Top             =   4275
      Width           =   3435
      Begin VB.OptionButton optValor 
         Caption         =   "Valor "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optQtItensVendidos 
         Caption         =   "Quantidade de Ítens Vendidos"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optCodVendedor 
         Caption         =   "Código do Vendedor"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optQtOperacoes 
         Caption         =   "Quantidade de Operações"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame fraT 
      Caption         =   "Tipo do Relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   24
      Top             =   4275
      Width           =   3435
      Begin VB.OptionButton optSintetico 
         Caption         =   "Sintético"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optAnalitico 
         Caption         =   "Analítico"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   6915
      Begin SSDataWidgets_B.SSDBCombo cboVendedor 
         Bindings        =   "frmRelComissaoComRetencao.frx":38CC
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
         Bindings        =   "frmRelComissaoComRetencao.frx":38EA
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   660
         Width           =   690
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
         TabIndex        =   22
         Top             =   600
         Width           =   4920
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   300
         Width           =   300
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
         TabIndex        =   20
         Top             =   240
         Width           =   4920
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
      TabIndex        =   17
      Top             =   3480
      Width           =   3435
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
         TabIndex        =   18
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   5685
      Width           =   1500
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5685
      Width           =   1500
   End
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
      Left            =   3480
      TabIndex        =   16
      Top             =   3480
      Width           =   3435
      Begin VB.OptionButton optVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1215
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
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelComissaoComRetencao.frx":3902
         ForeColor       =   &H00808080&
         Height          =   1095
         Left            =   2520
         TabIndex        =   26
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   1170
         Left            =   240
         Picture         =   "frmRelComissaoComRetencao.frx":39F0
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelComissaoComRetencao.frx":513D
         ForeColor       =   &H00808080&
         Height          =   855
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   3855
      End
   End
   Begin Crystal.CrystalReport crpRel 
      Left            =   3960
      Top             =   7200
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
Attribute VB_Name = "frmRelComissaoComRetencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  If ValidarCampos Then Exit Sub
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde gerando às informações...")
  Call GerarInformacoes
  
  If optSintetico.Value Then AgruparDados
  
  Call StatusMsg("")
  Call MontarRelatorio
  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Call CenterForm(Me)
  
  datFilial.DatabaseName = gsQuickDBFileName
  datFuncionarios.DatabaseName = gsQuickDBFileName
  
  mskDataIni.Text = Format(Date, "DD/MM/YYYY")
  mskDataFim.Text = Format(Date, "DD/MM/YYYY")
  
  fraSintetico.Enabled = False
  optCodVendedor.Enabled = False
  optQtOperacoes.Enabled = False
  optQtItensVendidos.Enabled = False
  optValor.Enabled = False
  
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
  Dim rstComissoes         As Recordset
  Dim rstComissoesRetencao As Recordset
  Dim strSQL               As String
  Dim dblTotal             As Double
  Dim dblDescSubTotal      As Double
  
  On Error GoTo ErrHandler
  
  'Tratamento para a table temporária
  dbTemp.Execute "DELETE * FROM ComissoesRetencao"
  Set rstComissoesRetencao = dbTemp.OpenRecordset("ComissoesRetencao", dbOpenDynaset)
  'Fim Tratamento

  strSQL = "SELECT * FROM Comissão "
  strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Data >= #" & Format(mskDataIni.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Data <= #" & Format(mskDataFim.Text, "MM/DD/YYYY") & "#"
  If Len(lblNomeVendedor.Caption) > 0 Then strSQL = strSQL & " AND Vendedor = " & (cboVendedor.Text)
  strSQL = strSQL & " ORDER BY Sequência "
  
  Set rstComissoes = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstComissoes
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        rstComissoesRetencao.AddNew
        rstComissoesRetencao.Fields("Vendedor").Value = .Fields("Vendedor").Value
        rstComissoesRetencao.Fields("Sequencia").Value = .Fields("Sequência").Value
        rstComissoesRetencao.Fields("QtdeItens").Value = .Fields("Qtde").Value
        dblDescSubTotal = GetDescSubTotal(.Fields("Sequência").Value, .Fields("Vendedor").Value)
        rstComissoesRetencao.Fields("VlPagoSemCartao").Value = Format(.Fields("Valor").Value - .Fields("VlPagoEmCartao").Value - dblDescSubTotal, FORMAT_VALUE)
        rstComissoesRetencao.Fields("VlPagoComCartao").Value = .Fields("VlPagoEmCartao").Value
        rstComissoesRetencao.Fields("TaxaRetencao").Value = .Fields("TaxaRetencao").Value
        rstComissoesRetencao.Fields("DescontoSubTotal").Value = dblDescSubTotal
        rstComissoesRetencao.Fields("VlPagoComCartaoRetendo").Value = .Fields("VlPagoEmCartaoComRetencao").Value
        dblTotal = Format((.Fields("Valor").Value - .Fields("VlPagoEmCartao").Value - dblDescSubTotal) + .Fields("VlPagoEmCartaoComRetencao").Value, FORMAT_VALUE)
        rstComissoesRetencao.Fields("ValorParaCalculoComi").Value = dblTotal
        rstComissoesRetencao.Fields("Comissao").Value = GetCalcularComissao(.Fields("Vendedor").Value, dblTotal)
        rstComissoesRetencao.Fields("QtdeOperacao").Value = 1
        rstComissoesRetencao.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With

  Set rstComissoes = Nothing

  rstComissoesRetencao.Close
  Set rstComissoesRetencao = Nothing

  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Function GetDescSubTotal(ByVal lngSeq As Long, ByVal intVendedor As Integer) As Double
  Dim rstDescSubTotal As Recordset
  Dim strSQL          As String
  
  GetDescSubTotal = 0
  
  strSQL = "SELECT DescontoSubTotal "
  strSQL = strSQL & " FROM Saídas "
  strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Digitador = " & intVendedor
  strSQL = strSQL & " AND Sequência = " & lngSeq

  Set rstDescSubTotal = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  If rstDescSubTotal.RecordCount = 0 Then Exit Function
  
  With rstDescSubTotal
    Call IsDataType(dtDouble, .Fields("DescontoSubTotal").Value, GetDescSubTotal)
    If Not rstDescSubTotal Is Nothing Then .Close
    Set rstDescSubTotal = Nothing
  End With

End Function

Private Function GetCalcularComissao(ByVal intVendedor As Integer, ByVal dblValor As Double) As Double
  Dim rstFuncionario As Recordset
  
  On Error GoTo TratarErro
  
  Set rstFuncionario = db.OpenRecordset("SELECT Comissão FROM Funcionários WHERE Código = " & intVendedor, dbOpenDynaset)

  With rstFuncionario
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      GetCalcularComissao = Format((dblValor * .Fields("Comissão").Value) / 100, FORMAT_VALUE)
    End If
    .Close
  End With
  
  Set rstFuncionario = Nothing
  
  Exit Function

TratarErro:
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Function

End Function

Private Sub AgruparDados()
  Dim rstComiReteGroup As Recordset
  Dim rstComiRete      As Recordset
  Dim strSQL           As String
  
  On Error GoTo TratarErro
  
  '---[Tabela de Agrupamento]---
  dbTemp.Execute "DELETE * FROM ComissoesRetencaoGroup"
  '---[Abrindo a temp-table]
  Set rstComiReteGroup = dbTemp.OpenRecordset("ComissoesRetencaoGroup", dbOpenDynaset)
  
  
  '---[Group By]---
  strSQL = "SELECT Vendedor, SUM(QtdeItens) AS F1, SUM(VlPagoSemCartao) AS F2, SUM(VlPagoComCartao) AS F3, SUM(VlPagoComCartaoRetendo) AS F4, SUM(DescontoSubTotal) AS F5, SUM(ValorParaCalculoComi) AS F6, SUM(Comissao) AS F7, SUM(QtdeOperacao) AS F8 "
  strSQL = strSQL & " FROM ComissoesRetencao "
  strSQL = strSQL & " GROUP BY Vendedor "
  
  Set rstComiRete = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstComiRete
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        rstComiReteGroup.AddNew
          rstComiReteGroup.Fields("Vendedor").Value = .Fields("Vendedor").Value
          rstComiReteGroup.Fields("QtdeItens").Value = .Fields("F1").Value
          rstComiReteGroup.Fields("VlPagoSemCartao").Value = .Fields("F2").Value
          rstComiReteGroup.Fields("VlPagoComCartao").Value = .Fields("F3").Value
          rstComiReteGroup.Fields("VlPagoComCartaoRetendo").Value = .Fields("F4").Value
          rstComiReteGroup.Fields("DescontoSubTotal").Value = .Fields("F5").Value
          rstComiReteGroup.Fields("ValorParaCalculoComi").Value = .Fields("F6").Value
          rstComiReteGroup.Fields("Comissao").Value = .Fields("F7").Value
          rstComiReteGroup.Fields("QtdeOperacao").Value = .Fields("F8").Value
        rstComiReteGroup.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstComiRete = Nothing
  
  Exit Sub

TratarErro:
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub MontarRelatorio()
  Dim strReport As String
  
  On Error GoTo TratarErro
  
  'Nome do arquivo .rpt
  If optAnalitico.Value Then
    strReport = gsReportPath & "rptComiReteAnalitico.rpt"
  Else
    strReport = gsReportPath & "rptComiReteSintetico.rpt"
  End If
  
  With crpRel
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crpRel
    
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsTempDBFileName
    .DataFiles(2) = gsTempDBFileName
    
    '.SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .Formulas(1) = "Periodo = '" & "Período: " & (mskDataIni.Text) & " à " & (mskDataFim.Text) & "'"
    
    '12/05/2005 - Daniel
    'Correção para exibição dos botões de Configuração
    'de Impressoras e Botão de Pesquisas
    .WindowShowPrintSetupBtn = True
    .WindowShowSearchBtn = True
    
    If optSintetico.Value Then
      If optCodVendedor.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Código do Vendedor" & "'"
      If optQtOperacoes.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Quantidade de Operações" & "'"
      If optQtItensVendidos.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Quantidade de Ítens" & "'"
      If optValor.Value Then .Formulas(2) = "Ordenacao = '" & "Ordenação (Ranking): " & "Valor das Vendas" & "'"
    End If
    
    'Ordenação
    If optSintetico.Value Then
      If optCodVendedor.Value Then .SortFields(0) = "+{ComissoesRetencaoGroup.Vendedor}"
      If optQtOperacoes.Value Then .SortFields(0) = "-{ComissoesRetencaoGroup.QtdeOperacao}"
      If optQtItensVendidos.Value Then .SortFields(0) = "-{ComissoesRetencaoGroup.QtdeItens}"
      If optValor.Value Then .SortFields(0) = "-{ComissoesRetencaoGroup.ValorParaCalculoComi}"
    End If
    
    .WindowState = crptMaximized
    .Destination = IIf(optVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crpRel)
  
    .Action = 1
  End With

  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  Exit Sub

TratarErro:
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Sub optAnalitico_Click()
  fraSintetico.Enabled = False
  optCodVendedor.Enabled = False
  optQtOperacoes.Enabled = False
  optQtItensVendidos.Enabled = False
  optValor.Enabled = False
End Sub

Private Sub optSintetico_Click()
  fraSintetico.Enabled = True
  optCodVendedor.Enabled = True
  optQtOperacoes.Enabled = True
  optQtItensVendidos.Enabled = True
  optValor.Enabled = True
End Sub
