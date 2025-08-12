VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmFolhaPagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folha de Pagamento"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmFolhaPagamento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   5790
   Begin VB.Frame fraX 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   5550
      Begin VB.Data datFilial 
         Caption         =   "datFilial"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   3855
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmFolhaPagamento.frx":058A
         Height          =   315
         Left            =   600
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdGerar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Gerar"
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
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   2805
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
      Height          =   885
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   3735
      Begin MSMask.MaskEdBox mskPeriodoFinal 
         Height          =   315
         Left            =   2280
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
      Begin MSMask.MaskEdBox mskPeriodoInicio 
         Height          =   315
         Left            =   600
         TabIndex        =   1
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   420
         Width           =   270
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Geração de Arquivo para RH"
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
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmFolhaPagamento.frx":05A2
         ForeColor       =   &H00808080&
         Height          =   975
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   4935
      End
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   5160
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdDesfazer 
      Cancel          =   -1  'True
      Caption         =   "&Desfazer "
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
End
Attribute VB_Name = "frmFolhaPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstFiliais As Recordset
  
  txtNomeFilial.Text = ""
  
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstFiliais = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & CByte(cboFilial.Text), dbOpenSnapshot)
  
  With rstFiliais
    If Not (.BOF And .EOF) Then
      txtNomeFilial.Text = .Fields("Nome") & ""
    End If
    
    If Not rstFiliais Is Nothing Then .Close
    Set rstFiliais = Nothing
  End With

End Sub

Private Sub cmdDesfazer_Click()
  '16/05/2005 - Daniel
  'Adicionado rotina que desfaz às informações
  'geradas no arquivo
  
  'Validações
  If ValidaCampos Then Exit Sub
  
  MsgBox "Será desfeito o período de " & (mskPeriodoInicio.Text) & " até " & (mskPeriodoFinal.Text)
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  Else
    Screen.MousePointer = vbHourglass
    Call StatusMsg("Atualizando Status...")
    Call AtualizarStatus
    Screen.MousePointer = vbDefault
    Call StatusMsg("")
  End If
  
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdGerar_Click()
  Dim rstSaidas     As Recordset
  Dim rstFolhaPagto As Recordset
  Dim strSQL        As String
  Dim intQtdeParc   As Integer

  On Error GoTo TratarErro

  'Validações
  If ValidaCampos Then Exit Sub
  
  'Limpando a tabela temporária FolhaPagamento
  dbTemp.Execute "DELETE * FROM FolhaPagamento"
  'Abrindo a Tabela
  Set rstFolhaPagto = dbTemp.OpenRecordset("SELECT * FROM FolhaPagamento", dbOpenDynaset)

  'Início da query
  strSQL = "SELECT * FROM Saídas WHERE Filial = " & CByte(cboFilial.Text)                         'Filial
  strSQL = strSQL & " AND Data >= #" & Format(CDate(mskPeriodoInicio.Text), "MM/DD/YYYY") & "#"   'Intervalo de Datas
  strSQL = strSQL & " AND Data <= #" & Format(CDate(mskPeriodoFinal.Text), "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND NOT [Status Venda Func] "                          'Status de não encaminhado para a Folha
  strSQL = strSQL & " AND [Codigo Func Comprador] <> 0 "                     'Houve compra
  '16/05/2005 - Daniel
  'Adicionado novas cláusulas [ AND ]
  strSQL = strSQL & " AND Saídas.Efetivada "
  strSQL = strSQL & " AND NOT Saídas.[Nota Cancelada] "
  '----------------------------------
  strSQL = strSQL & " ORDER BY [Codigo Func Comprador], Data "

  Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)

  intQtdeParc = 0
  
  If rstSaidas.RecordCount = 0 Then
    MsgBox "Não foram encontradas informações dentro deste intervalo, verifique. ", vbExclamation, "Quick Store"
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde pesquisando banco de dados...")

  With rstSaidas
    If Not (.BOF And .EOF) Then
      .MoveFirst
  
      Do Until .EOF
        Call BuscaQtdeParcelas(.Fields("Sequência").Value, intQtdeParc)
        
        'Criar record em FolhaPagamento---------------
         rstFolhaPagto.AddNew
         rstFolhaPagto.Fields("Codigo").Value = .Fields("Codigo Func Comprador").Value
         rstFolhaPagto.Fields("Parcela").Value = intQtdeParc
         rstFolhaPagto.Fields("Total").Value = .Fields("Total").Value
         rstFolhaPagto.Update
        '---[Fim do Criar record em FolhaPagamento]---
        
        .Edit
        .Fields("Status Venda Func").Value = True 'Após acender o flag em Saídas, daremos baixa em CR
        'Call BaixarCR(.Fields("Sequência").Value)
        .Update
        
        intQtdeParc = 0
        .MoveNext
      Loop
    End If
    .Close
    rstFolhaPagto.Close
  End With

  Set rstSaidas = Nothing
  Set rstFolhaPagto = Nothing
  
  Call CriarArquivo
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
  Exit Sub
  
TratarErro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Sub

Private Sub BuscaQtdeParcelas(ByVal Seq As Long, ByRef QtdeParcelas As Integer)
  Dim rstCR           As Recordset
  Dim rstContaCliente As Recordset
  
  Set rstCR = db.OpenRecordset("SELECT Sequência FROM [Contas a Receber] WHERE Filial = " & gnCodFilial & " AND Sequência = " & Seq, dbOpenDynaset)  'Filial e Seq

  'Não achou nada em CR vamos procurar em Conta Cliente
  If rstCR.RecordCount = 0 Then
    Set rstContaCliente = db.OpenRecordset("SELECT Sequência FROM [Conta Cliente] WHERE Filial = " & gnCodFilial & " AND Sequência = " & Seq, dbOpenDynaset)
  
    With rstContaCliente
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do Until .EOF
          QtdeParcelas = QtdeParcelas + 1
          
          .MoveNext
        Loop
        
      End If
      .Close
    End With
    
    Set rstContaCliente = Nothing
    
  End If

  'Tratamento para CR
  With rstCR
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        QtdeParcelas = QtdeParcelas + 1
    
        .MoveNext
      Loop
      .Close
    End If
  End With
  
  Set rstCR = Nothing

End Sub

Private Sub CriarArquivo()
  Dim rstFolhaPagto  As Recordset
  Dim strAuxi1       As String
  Dim strAuxi2       As String
  Dim strNomeArquivo As String
  Dim intResp        As Integer
  Dim strCodUser     As String
  Dim strParcela     As String
  Dim strValor       As String
  
  '04/10/2005 - mpdea
  'Número do arquivo
  Dim intFileNum As Integer
  'Controle de transação
  Dim blnInTransaction As Boolean
  
  
  On Error GoTo ErrHandler
  
  
  strAuxi1 = gsReportPath & "FOLHA"
  strAuxi2 = Format(Date, "dd/mm/yy")
  strAuxi1 = strAuxi1 & Left(strAuxi2, 2)
  strAuxi1 = strAuxi1 & Mid(strAuxi2, 4, 2)
  strAuxi1 = strAuxi1 & Mid(strAuxi2, 7, 4)
  strAuxi1 = strAuxi1 & ".txt"
  
  Dialog1.FileName = strAuxi1
  
  
  With Dialog1
    .CancelError = True
    .DialogTitle = "Salvar arquivo para a Folha de Pagamento como"
    .DefaultExt = "txt"
    .InitDir = gsDefaultPath
    .Filter = "Arquivo de Folha de Pagamento | *.txt"
    .Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly & cdlOFNOverwritePrompt
    .ShowSave
    strNomeArquivo = .FileName
  End With
    
  '-------------------------------------------------------------------------------
  '04/10/2005 - mpdea
  'Comentado as linhas abaixo devido a possibilidade de Run-time com
  '"On Error GoTo 0" e a substituição de sobrescrita de arquivo através
  'da propriedade Flags do objeto Dialog1
  '
'  On Error GoTo 0
'
'    If Dir(strNomeArquivo) <> "" Then
'      intResp = MsgBox("Já existe este arquivo, deseja sobrescrever ?", vbQuestion + vbOKCancel, "Atenção")
'      If intResp = vbCancel Then
'        DisplayMsg "Geração de arquivo cancelada."
'        Exit Sub
'      End If
'    End If
  '-------------------------------------------------------------------------------

  
  '04/10/2005 - mpdea
  'Início de transação
  ws.BeginTrans: blnInTransaction = True


  'Inicialização das vars de impressão
  strCodUser = "000000000" '09 posições
  strParcela = "00"        '02 posições
  strValor = "00000000000" '11 posições

  '----------[rstFolhaPagto]----------
  Set rstFolhaPagto = dbTemp.OpenRecordset("SELECT Codigo, Sum(Parcela) as Parc, Sum(Total) as Tot FROM FolhaPagamento GROUP BY Codigo ", dbOpenDynaset)
  
  With rstFolhaPagto
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      '04/10/2005 - mpdea
      'Substituído identificação de arquivo fixa (#1)
      intFileNum = FreeFile
      
      Open strNomeArquivo For Output As #intFileNum
        
      Do Until .EOF
        strCodUser = strCodUser & CStr(.Fields("Codigo").Value)
        strParcela = strParcela & CStr(.Fields("Parc").Value)
        strValor = strValor & CStr((.Fields("Tot").Value) * 100)
        
        strCodUser = Right(strCodUser, 9)
        strParcela = Right(strParcela, 2)
        strValor = Right(strValor, 11)
      
        Print #intFileNum, strCodUser & strParcela & strValor
        
        'Formatamos novamente as vars de impressão
        'para não guardar sujeiras...
        strCodUser = "000000000" '09 posições
        strParcela = "00"        '02 posições
        strValor = "00000000000" '11 posições
      
        .MoveNext
      Loop
      
      Close #intFileNum
  
    End If
    .Close
  End With

  Set rstFolhaPagto = Nothing
  
  '04/10/2005 - mpdea
  'Fim de transação
  ws.CommitTrans: blnInTransaction = True

  DisplayMsg "Geração efetuada com Sucesso."
  Exit Sub

ErrHandler:
  
  '---------------------------------------------------------------------------------
  '04/10/2005 - mpdea
  'Verifica cancelamento da operação de salvar arquivo
  If Err.Number = cdlCancel Then Exit Sub
  'Desfaz transação e não precisa chamar a função comentada
  'Call AtualizarStatus
  If blnInTransaction Then ws.Rollback
  'Fecha arquivos abertos por Open
  Close
  'Informa o erro ocorrido
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  '---------------------------------------------------------------------------------

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  datFilial.DatabaseName = gsQuickDBFileName
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

Private Function ValidaCampos() As Boolean

  ValidaCampos = False
  
  If Len(txtNomeFilial.Text) = 0 Then
    MsgBox "Selecione uma Filial válida.", vbExclamation, "Quick Store"
    cboFilial.SetFocus
    ValidaCampos = True
    Exit Function
  End If

  If Not IsDate(mskPeriodoInicio.Text) Then
    MsgBox "Data Inicial inválida.", vbExclamation, "Quick Store"
    mskPeriodoInicio.SetFocus
    ValidaCampos = True
    Exit Function
  End If

  If Not IsDate(mskPeriodoFinal.Text) Then
    MsgBox "Data Final inválida.", vbExclamation, "Quick Store"
    mskPeriodoFinal.SetFocus
    ValidaCampos = True
    Exit Function
  End If
  
  If CDate(mskPeriodoInicio.Text) > CDate(mskPeriodoFinal.Text) Then
    MsgBox "Data Final menor que a Inicial, verifique.", vbExclamation, "Quick Store"
    mskPeriodoFinal.SetFocus
    ValidaCampos = True
    Exit Function
  End If
  
End Function

Private Sub BaixarCR(ByVal Seq As Long)
'Através da Seq daremos baixa no CR
'Nota: Cliente preferiu dar baixas manuais
End Sub

Private Sub AtualizarStatus()
  'Caso não seja gerado o arquivo teremos que atualizar
  'o campo [Status Venda Func] para False
  Dim rstSaidas As Recordset
  Dim strQuery  As String
  
  strQuery = " SELECT Filial, Data, [Codigo Func Comprador], [Status Venda Func] "
  strQuery = strQuery & " FROM Saídas "
  strQuery = strQuery & " WHERE Data >= #" & Format(mskPeriodoInicio.Text, "mm/dd/yyyy") & "#"
  strQuery = strQuery & " AND Data <= #" & Format(mskPeriodoFinal.Text, "mm/dd/yyyy") & "#"
  strQuery = strQuery & " AND [Codigo Func Comprador] <> 0 "
  
  Set rstSaidas = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstSaidas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        .Edit
        .Fields("Status Venda Func").Value = False
        .Update
      
      .MoveNext
      Loop
    End If
    .Close
  End With

  Set rstSaidas = Nothing
  
End Sub
