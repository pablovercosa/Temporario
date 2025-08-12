VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelPosicaoFinanceiraCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posição Financeira do Cliente"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelPosicaoFinanceiraCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   5850
   Begin VB.Frame fraY 
      Caption         =   "Situação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      TabIndex        =   23
      Top             =   3720
      Width           =   3735
      Begin VB.OptionButton optTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optNaoRecebidas 
         Caption         =   "Não Recebidas"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optRecebidas 
         Caption         =   "Recebidas"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraSaida 
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
      Left            =   3840
      TabIndex        =   22
      Top             =   2760
      Width           =   1935
      Begin VB.OptionButton optVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
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
      Height          =   885
      Left            =   0
      TabIndex        =   19
      Top             =   2760
      Width           =   3735
      Begin MSMask.MaskEdBox mskPeriodoFinal 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para obter calendário."
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
      Begin MSMask.MaskEdBox mskPeriodoInicio 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para obter calendário."
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   1920
         TabIndex        =   21
         Top             =   420
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   420
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   4320
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame fraX 
      Caption         =   "Seleção"
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
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   5775
      Begin VB.TextBox txtNomeCliente 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   720
         Width           =   3855
      End
      Begin VB.Data datCliente 
         Caption         =   "datCliente"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome, Tipo FROM Cli_For WHERE Tipo = 'C' ORDER BY Código"
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   3855
      End
      Begin VB.Data datParametros 
         Caption         =   "datParametros"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelPosicaoFinanceiraCliente.frx":058A
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   855
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
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Filial"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCliente 
         Bindings        =   "frmRelPosicaoFinanceiraCliente.frx":05A6
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   855
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
         Columns(0).Width=   3200
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Código"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   300
         Width           =   300
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   11
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelPosicaoFinanceiraCliente.frx":05BF
         ForeColor       =   &H00808080&
         Height          =   735
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Geração da posição financeira por Cliente"
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
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   3855
      End
   End
   Begin Crystal.CrystalReport rptPosicao 
      Left            =   5280
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRelPosicaoFinanceiraCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim rsttblRelPosFin1 As Recordset
  Dim rsttblRelPosFin2 As Recordset
  
Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(0).Text
  cboCliente_LostFocus
End Sub

Private Sub cboCliente_LostFocus()
  Dim rstClientes As Recordset
  
  txtNomeCliente.Text = ""

  If Not IsNumeric(cboCliente.Text) Then Exit Sub
  
  Set rstClientes = db.OpenRecordset("SELECT Código, Nome FROM Cli_For WHERE Código = " & cboCliente.Text, dbOpenDynaset)
  
  With rstClientes
    If Not (.BOF And .EOF) Then
      txtNomeCliente.Text = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  
  Set rstClientes = Nothing
  
End Sub

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstParametros As Recordset
  
  txtNomeFilial.Text = ""
  
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstParametros = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & CByte(cboFilial.Text), dbOpenSnapshot)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      txtNomeFilial.Text = .Fields("Nome").Value & ""
    End If
    .Close
  End With

  Set rstParametros = Nothing

End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim rstCRxCli        As Recordset
  Dim strSQL           As String
  Dim lngSeq1          As Long
  Dim lngSeq2          As Long
  Dim byteAuxi         As Byte
  
  'Mensagem temporária
  MsgBox "Relatório ainda em desenvolvimento...", vbExclamation, "Quick Store"
  Exit Sub
  '---------------------------------------------------------------------------
  
  If Not ValidarCampos Then Exit Sub

  Call StatusMsg("Gerando o relatório de posição financeira...")

  dbTemp.Execute "DELETE * FROM tblRelPosFin1"
  dbTemp.Execute "DELETE * FROM tblRelPosFin2"
  
  Set rsttblRelPosFin1 = dbTemp.OpenRecordset("tblRelPosFin1", dbOpenDynaset)
  Set rsttblRelPosFin2 = dbTemp.OpenRecordset("tblRelPosFin2", dbOpenDynaset)

  strSQL = "SELECT [Contas a Receber].Filial, [Contas a Receber].Sequência, [Contas a Receber].Fatura, [Contas a Receber].Vencimento, [Contas a Receber].Valor, [Contas a Receber].Nota, Cli_For.Nome "
  strSQL = strSQL & " FROM [Contas a Receber], Cli_For "
  strSQL = strSQL & " WHERE [Contas a Receber].Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND [Contas a Receber].Vencimento >= #" & Format(mskPeriodoInicio.Text, "mm/dd/yyyy") & "#"
  strSQL = strSQL & " AND [Contas a Receber].Vencimento <= #" & Format(mskPeriodoFinal.Text, "mm/dd/yyyy") & "#"
  strSQL = strSQL & " AND [Contas a Receber].Cliente = " & CLng(cboCliente.Text)

  If optRecebidas.Value Then
    strSQL = strSQL & " AND [Contas a Receber].[Valor Recebido] <> 0 "
  End If

  If optNaoRecebidas.Value Then
    strSQL = strSQL & " AND [Contas a Receber].[Valor Recebido] = 0 "
  End If

  strSQL = strSQL & " AND Cli_For.Código = [Contas a Receber].Cliente "
  strSQL = strSQL & " ORDER BY [Contas a Receber].Nota, [Contas a Receber].Fatura "

  Set rstCRxCli = db.OpenRecordset(strSQL, dbOpenDynaset)

  If rstCRxCli.RecordCount = 0 Then
    MsgBox "Não foi encontrado informações na seleção indicada, verifique.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  With rstCRxCli
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      byteAuxi = 1
      
      Do Until .EOF
        lngSeq1 = .Fields("Sequência").Value
        
        rsttblRelPosFin2.AddNew
        rsttblRelPosFin2.Fields("NF").Value = .Fields("Nota").Value
        rsttblRelPosFin2.Fields("Fatura").Value = .Fields("Fatura").Value
        rsttblRelPosFin2.Fields("Valor").Value = .Fields("Valor").Value
        rsttblRelPosFin2.Fields("Vencimento").Value = .Fields("Vencimento").Value
        rsttblRelPosFin2.Fields("Seq").Value = .Fields("Sequência").Value
        rsttblRelPosFin2.Update
        
        If byteAuxi = 1 Then
          lngSeq2 = .Fields("Sequência").Value
          byteAuxi = 0
          Call BuscarProdutos(cboFilial.Text, .Fields("Sequência").Value, .Fields("Nota").Value)
        End If
        
        If lngSeq1 <> lngSeq2 Then
          lngSeq2 = .Fields("Sequência").Value
          Call BuscarProdutos(cboFilial.Text, .Fields("Sequência").Value, .Fields("Nota").Value)
        End If
        
      .MoveNext
      Loop
    End If
    .Close
  End With

  Set rstCRxCli = Nothing
  
  rsttblRelPosFin2.Close
  rsttblRelPosFin1.Close
  Set rsttblRelPosFin2 = Nothing
  Set rsttblRelPosFin1 = Nothing
  
  'Call MontarRelatorio


End Sub

Private Sub BuscarProdutos(ByVal Filial As Byte, ByVal Seq As Long, ByVal NF As Long)
  Dim rstSaidasProdutos As Recordset
  Dim rstProdutos       As Recordset
  Dim strQuery          As String

  strQuery = "SELECT Filial, Sequência, Código "
  strQuery = strQuery & " FROM [Saídas - Produtos] "
  strQuery = strQuery & " WHERE Filial = " & Filial
  strQuery = strQuery & " AND Sequência = " & Seq

  Set rstSaidasProdutos = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstSaidasProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        
        strQuery = "SELECT Código, Nome "
        strQuery = strQuery & " FROM Produtos "
        strQuery = strQuery & " WHERE Código = '" & (.Fields("Código").Value) & "'"
        
        Set rstProdutos = db.OpenRecordset(strQuery, dbOpenDynaset)
      
          rsttblRelPosFin1.AddNew
          rsttblRelPosFin1.Fields("CodProduto").Value = rstProdutos.Fields("Código").Value & ""
          rsttblRelPosFin1.Fields("NomeProduto").Value = rstProdutos.Fields("Nome").Value & ""
          rsttblRelPosFin1.Fields("Seq").Value = .Fields("Sequência").Value
          rsttblRelPosFin1.Fields("NF").Value = NF
          rsttblRelPosFin1.Update
        
        rstProdutos.Close
        Set rstProdutos = Nothing
      
      .MoveNext
      Loop
  
    End If
    .Close
  End With
  
  Set rstSaidasProdutos = Nothing

End Sub

Private Sub MontarRelatorio()
'  Dim strSQL    As String
'  Dim strReport As String
'
'  strSQL = " {tblRelPosFin2.NF} <> '' "
'
'  'Nome do arquivo .rpt
'  strReport = gsReportPath & "rptPosicaoFinCli.rpt"
'  MousePointer = vbHourglass
'
'  With rptPosicao
'    .Reset
'    .ReportFileName = strReport
'
'    .DataFiles(0) = gsTempDBFileName
'    .DataFiles(1) = gsTempDBFileName
'    .DataFiles(2) = gsTempDBFileName
'
'    .SelectionFormula = strSQL
'    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
'    .SortFields(0) = "+{tblRelPosFin2.NF}" 'Ordenação
'    .SortFields(1) = "+{tblRelPosFin2.Fatura}"
'
'    .WindowState = crptMaximized
'    .Destination = IIf(optVideo.Value, crptToWindow, crptToPrinter)
'    Call StatusMsg("Aguarde, imprimindo...")
'
'    'Seta a impressora para relatório
'    Call SetPrinterName("REL", rptPosicao)
'
'    .Action = 1
'  End With
'
'  MousePointer = vbDefault

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datParametros.DatabaseName = gsQuickDBFileName
  datCliente.DatabaseName = gsQuickDBFileName
End Sub

Private Sub mskPeriodoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskPeriodoFinal.Text = frmCalendario.gsDateCalender(mskPeriodoFinal.Text)
  End If
End Sub

Private Sub mskPeriodoFinal_LostFocus()
  mskPeriodoFinal.Text = Ajusta_Data(mskPeriodoFinal.Text)
End Sub

Private Sub mskPeriodoInicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskPeriodoInicio.Text = frmCalendario.gsDateCalender(mskPeriodoInicio.Text)
  End If
End Sub

Private Sub mskPeriodoInicio_LostFocus()
  mskPeriodoInicio.Text = Ajusta_Data(mskPeriodoInicio.Text)
End Sub

Private Function ValidarCampos() As Boolean
  
  If Len(txtNomeFilial.Text) <= 0 Then
    ValidarCampos = False
    MsgBox "Filial inválida, verifique.", vbExclamation, "Quick Store"
    cboFilial.SetFocus
    Exit Function
  End If
  
  If Len(txtNomeCliente.Text) <= 0 Then
    ValidarCampos = False
    MsgBox "Cliente inválido, verifique.", vbExclamation, "Quick Store"
    cboCliente.SetFocus
    Exit Function
  End If
  
  If Len(mskPeriodoInicio.Text) > 0 Then
    If Not IsDate(mskPeriodoInicio.Text) Then
      ValidarCampos = False
      MsgBox "Data inicial inválida, verifique.", vbExclamation, "Quick Store"
      mskPeriodoInicio.SetFocus
      Exit Function
    End If
  End If
  
  If Len(mskPeriodoFinal.Text) > 0 Then
    If Not IsDate(mskPeriodoFinal.Text) Then
      ValidarCampos = False
      MsgBox "Data final inválida, verifique.", vbExclamation, "Quick Store"
      mskPeriodoFinal.SetFocus
      Exit Function
    End If
  End If
  
  If Len(mskPeriodoFinal.Text) > 0 And Len(mskPeriodoInicio.Text) > 0 Then
    If CDate(mskPeriodoFinal.Text) < CDate(mskPeriodoInicio.Text) Then
      ValidarCampos = False
      MsgBox "Data final menor que a inicial, verifique.", vbExclamation, "Quick Store"
      mskPeriodoFinal.SetFocus
      Exit Function
    End If
  End If
  
  ValidarCampos = True
  
End Function
