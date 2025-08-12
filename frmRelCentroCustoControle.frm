VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelCentroCustoControle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Controle de Centros de Custo"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelCentroCustoControle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   6615
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Frame fraC 
      Caption         =   "Contas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      TabIndex        =   25
      Top             =   1320
      Width           =   3135
      Begin VB.OptionButton optTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optVencidas 
         Caption         =   "Vencidas"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optAPagar 
         Caption         =   "A Pagar"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optPagas 
         Caption         =   "Pagas"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Data datPara 
      Caption         =   "datPara"
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
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Data datCC 
      Caption         =   "datCC"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Centros de Custo] WHERE Ativo ORDER BY Nome"
      Top             =   4800
      Width           =   2055
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
      Left            =   3360
      TabIndex        =   24
      Top             =   2640
      Width           =   3135
      Begin VB.OptionButton optVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   300
         Width           =   1215
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
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   3135
      Begin VB.OptionButton optSintetico 
         Caption         =   "Sintético"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optAnalitico 
         Caption         =   "Analítico"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   1575
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
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   3135
      Begin VB.OptionButton optPeriodoPagamento 
         Caption         =   "Pagamento"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optPeriodoVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskDataIni 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Ao teclar [F2] carrega calendário"
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
      Begin MSMask.MaskEdBox mskDataFim 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         ToolTipText     =   "Ao teclar [F2] carrega calendário"
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   1440
         TabIndex        =   22
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame fraX 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6375
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelCentroCustoControle.frx":058A
         DataSource      =   "datPara"
         Height          =   315
         Left            =   1440
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
      Begin SSDataWidgets_B.SSDBCombo cboCentro 
         Bindings        =   "frmRelCentroCustoControle.frx":05A0
         DataSource      =   "datCC"
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   735
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
         Columns(0).Width=   9922
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1746
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
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
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   3960
      End
      Begin VB.Label lblCC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   1185
      End
      Begin VB.Label lblNomeCC 
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
         Left            =   2280
         TabIndex        =   17
         Top             =   600
         Width           =   3960
      End
   End
   Begin Crystal.CrystalReport crpRel 
      Left            =   4680
      Top             =   4800
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
Attribute VB_Name = "frmRelCentroCustoControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCentro_CloseUp()
  cboCentro.Text = cboCentro.Columns(1).Text
  cboCentro_LostFocus
End Sub

Private Sub cboCentro_LostFocus()
  Dim rstCC As Recordset
  
  lblNomeCC.Caption = ""
  If Not IsNumeric(cboCentro.Text) Then Exit Sub
  
  Set rstCC = db.OpenRecordset("SELECT Nome FROM [Centros de Custo] WHERE Código = " & cboCentro.Text, dbOpenSnapshot)
  
  With rstCC
    If Not (.BOF And .EOF) Then
      lblNomeCC.Caption = .Fields("Nome") & ""
    End If
    
    If Not rstCC Is Nothing Then .Close
    Set rstCC = Nothing
  End With

End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  If ValidarCampos Then Exit Sub
  
  If optVencidas.Value Then MsgBox "O sistema estará trazendo contas vencidas do dia " & Format(Date - 1, "DD/MM/YYYY") & " para trás.", vbInformation, "Atenção"
  Call MontarRegistros
  If optSintetico.Value Then AgruparRegistros
  Call MontarRelatorio
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datPara.DatabaseName = gsQuickDBFileName
  datCC.DatabaseName = gsQuickDBFileName
  
  cboFilial.Text = gnCodFilial
  cboFilial_LostFocus
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
  
  If Not (optVencidas.Value) Then
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
  End If
  
End Function

'20/10/2008 - mpdea
'Incluído opção para selecionar por Data de Vencimento ou Data de Pagamento
Private Sub MontarRegistros()
  Dim rstCP          As Recordset
  Dim rstCompetencia As Recordset
  Dim strSQL         As String
  '20/10/2008 - mpdea
  Dim str_campo_periodo As String
  
  On Error GoTo TratarErro
  
  Call StatusMsg("Aguarde consultando o banco de dados...")
  
  Screen.MousePointer = vbHourglass
  
  '20/10/2008 - mpdea
  If optPeriodoVencimento.Value Then
    str_campo_periodo = "Vencimento"
  Else
    str_campo_periodo = "Pagamento"
  End If
  
  '-----[Início das interações]-----
  dbTemp.Execute "DELETE * FROM Competencia"
  
  'Montando a query
  strSQL = ""
  strSQL = "SELECT * FROM [Contas a Pagar]"
  strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
  
  If Not optVencidas.Value Then
    '20/10/2008 - mpdea
    strSQL = strSQL & " AND " & str_campo_periodo & " >= #" & Format(mskDataIni.Text, "MM/DD/YYYY") & "#"
    strSQL = strSQL & " AND " & str_campo_periodo & " <= #" & Format(mskDataFim.Text, "MM/DD/YYYY") & "#"
  End If
  
  If Len(lblNomeCC.Caption) > 0 Then strSQL = strSQL & " AND [Centro de Custo] = " & CInt(cboCentro.Text)
  
  'Tipo da conta
  If optPagas.Value Then strSQL = strSQL & " AND [Valor Pago] <> 0"
  
  If optAPagar.Value Then strSQL = strSQL & " AND [Valor Pago] = 0"
  
  If optVencidas.Value Then
    '20/10/2008 - mpdea
    strSQL = strSQL & " AND " & str_campo_periodo & " >= #" & Format("01/01/1990", "MM/DD/YYYY") & "#"
    strSQL = strSQL & " AND " & str_campo_periodo & " <= #" & Format(Date - 1, "MM/DD/YYYY") & "#"
    strSQL = strSQL & " AND [Valor Pago] = 0"
  End If
  'Fim Tipo da conta
  
  strSQL = strSQL & " ORDER BY Vencimento"
  'Fim de Montando a query
  
  Set rstCP = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstCP.RecordCount = 0 Then
    MsgBox "Não foram encontradas informações neste intervalo", vbExclamation, "Quick Store"
    rstCP.Close
    Set rstCP = Nothing
    Screen.MousePointer = vbDefault
    Call StatusMsg("")
    Exit Sub
  End If
  
  'Abrimos a temp-table
  Set rstCompetencia = dbTemp.OpenRecordset("Competencia", dbOpenDynaset)
  
  With rstCP
    .MoveFirst
    
    Do Until .EOF
      rstCompetencia.AddNew
        rstCompetencia.Fields("Centro").Value = .Fields("Centro de Custo").Value
        rstCompetencia.Fields("Emissao").Value = .Fields("Data Emissão").Value
        rstCompetencia.Fields("Vencimento").Value = .Fields("Vencimento").Value
        rstCompetencia.Fields("Descricao").Value = .Fields("Descrição").Value & ""
        rstCompetencia.Fields("Fornecedor").Value = .Fields("Fornecedor").Value
        rstCompetencia.Fields("Nota").Value = .Fields("Nota").Value & ""
        rstCompetencia.Fields("Sequencia").Value = .Fields("Sequência").Value
        rstCompetencia.Fields("Valor").Value = Format(.Fields("Valor").Value, FORMAT_VALUE)
        '21/10/2008 - mpdea
        'Incluído Data de Pagamento, Acréscimo, Desconto e Valor Pago
        rstCompetencia.Fields("Pagamento").Value = .Fields("Pagamento").Value
        rstCompetencia.Fields("Acrescimo").Value = Format("0" & .Fields("Acréscimo").Value, FORMAT_VALUE)
        rstCompetencia.Fields("Desconto").Value = Format("0" & .Fields("Desconto").Value, FORMAT_VALUE)
        rstCompetencia.Fields("ValorPago").Value = Format("0" & .Fields("Valor Pago").Value, FORMAT_VALUE)
      rstCompetencia.Update
    
     .MoveNext
    Loop
    .Close
  End With
  
  Set rstCP = Nothing
  
  'Fechamos a temp-table
  rstCompetencia.Close
  Set rstCompetencia = Nothing
  '---------------------------------
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
TratarErro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub AgruparRegistros()
  Dim rstCompetencia      As Recordset
  Dim rstCompetenciaGroup As Recordset
  Dim strSQL              As String
  
  On Error GoTo TratarErro
  
  Call StatusMsg("Agrupando os valores...")
  
  Screen.MousePointer = vbHourglass
  
  '-----[Início das interações]-----
  dbTemp.Execute "DELETE * FROM CompetenciaGroup"
  
  Set rstCompetenciaGroup = dbTemp.OpenRecordset("CompetenciaGroup", dbOpenDynaset)
  
  '21/10/2008 - mpdea
  'Incluído os totais para Acréscimo, Desconto e Valor Pago
  strSQL = "SELECT Centro, SUM(Valor) AS ValorTotal, SUM(Desconto) AS DescontoTotal, "
  strSQL = strSQL & "SUM(Acrescimo) AS AcrescimoTotal, SUM(ValorPago) AS ValorPagoTotal "
  strSQL = strSQL & "FROM Competencia "
  strSQL = strSQL & "GROUP BY Centro"
  
  Set rstCompetencia = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstCompetencia
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        rstCompetenciaGroup.AddNew
          '21/10/2008 - mpdea
          'Incluído os totais para Acréscimo, Desconto e Valor Pago
          rstCompetenciaGroup.Fields("Centro").Value = .Fields("Centro").Value
          rstCompetenciaGroup.Fields("Valor").Value = Format(.Fields("ValorTotal").Value, FORMAT_VALUE)
          rstCompetenciaGroup.Fields("Desconto").Value = Format(.Fields("DescontoTotal").Value, FORMAT_VALUE)
          rstCompetenciaGroup.Fields("Acrescimo").Value = Format(.Fields("AcrescimoTotal").Value, FORMAT_VALUE)
          rstCompetenciaGroup.Fields("ValorPago").Value = Format(.Fields("ValorPagoTotal").Value, FORMAT_VALUE)
        rstCompetenciaGroup.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  rstCompetenciaGroup.Close
  Set rstCompetenciaGroup = Nothing
  '---------------------------------
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
TratarErro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub MontarRelatorio()
  Dim strReport As String
  
  On Error GoTo TratarErro
  
  Call StatusMsg("Montando o relatório...")
  
  Screen.MousePointer = vbHourglass
  
  '-----[Montando o relatório]-----
  
  'Nome do arquivo .rpt
  If optAnalitico.Value Then
    strReport = gsReportPath & "rptCentroControleAnalitico.rpt"
  Else
    strReport = gsReportPath & "rptCentroControleSintetico.rpt"
  End If
  
  With crpRel
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crpRel
    
    If optAnalitico.Value Then
      .DataFiles(0) = gsTempDBFileName
      .DataFiles(1) = gsTempDBFileName
      .DataFiles(2) = gsQuickDBFileName
      .DataFiles(3) = gsQuickDBFileName
    Else
      .DataFiles(0) = gsTempDBFileName
      .DataFiles(1) = gsTempDBFileName
      .DataFiles(2) = gsQuickDBFileName
    End If
    
    'Fórmulas
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    
    If optVencidas.Value Then
      .Formulas(1) = "Periodo = '" & "Contas Vencidas de " & Format(Date - 1, "DD/MM/YYYY") & " para trás" & "'"
    Else
      .Formulas(1) = "Periodo = '" & "Período de Vencimento: " & (mskDataIni.Text) & " à " & (mskDataFim.Text) & "'"
    End If
    
    If optPagas.Value Then .Formulas(2) = "Conta = '" & "Tipo de contas: Contas pagas" & "'"
    If optAPagar.Value Then .Formulas(2) = "Conta = '" & "Tipo de contas: Contas a pagar" & "'"
    If optVencidas.Value Then .Formulas(2) = "Conta = '" & "Tipo de contas: Contas vencidas" & "'"
    If optTodas.Value Then .Formulas(2) = "Conta = '" & "Tipo de contas: Todas às Contas" & "'"
    
    '12/05/2005 - Daniel
    'Correção para exibição dos botões de Configuração
    'de Impressoras e Botão de Pesquisas
    .WindowShowPrintSetupBtn = True
    .WindowShowSearchBtn = True
    
    'Ordenação
    If optAnalitico.Value Then
      .SortFields(0) = "+{Competencia.Centro}"
      .SortFields(1) = "+{Competencia.Emissao}"
      .SortFields(2) = "+{Competencia.Sequencia}"
    Else
      .SortFields(0) = "+{CompetenciaGroup.Centro}"
    End If
    
    .WindowState = crptMaximized
    .Destination = IIf(optVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crpRel)
  
    .Action = 1
  End With
  '--------------------------------
  
  Call StatusMsg("")
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
TratarErro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Sub optAPagar_Click()
  mskDataIni.Enabled = True
  mskDataFim.Enabled = True

  mskDataIni.BackColor = &H80000005
  mskDataFim.BackColor = &H80000005
End Sub

Private Sub optPagas_Click()
  mskDataIni.Enabled = True
  mskDataFim.Enabled = True
  
  mskDataIni.BackColor = &H80000005
  mskDataFim.BackColor = &H80000005
End Sub

Private Sub optTodas_Click()
  mskDataIni.Enabled = True
  mskDataFim.Enabled = True
  
  mskDataIni.BackColor = &H80000005
  mskDataFim.BackColor = &H80000005
End Sub

Private Sub optVencidas_Click()
  mskDataIni.Enabled = False
  mskDataFim.Enabled = False
  
  mskDataIni.Mask = ""
  mskDataIni.Text = ""
  mskDataIni.Mask = "##/##/####"
  
  mskDataFim.Mask = ""
  mskDataFim.Text = ""
  mskDataFim.Mask = "##/##/####"
  
  mskDataIni.BackColor = &H8000000F
  mskDataFim.BackColor = &H8000000F
End Sub


