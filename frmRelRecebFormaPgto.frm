VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelRecebFormaPgto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Recebimentos por Forma de Pagamento"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "frmRelRecebFormaPgto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   7395
   Begin Crystal.CrystalReport crpView 
      Left            =   6900
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Relatório de Recebimentos por Forma de Pagamento"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Data datFiliais 
      Caption         =   "datFiliais"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4530
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   2430
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   510
      Width           =   4545
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   2580
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox mskDataInicio 
         Height          =   315
         Left            =   780
         TabIndex        =   2
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
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
      Begin VB.Label Label3 
         Caption         =   "De"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   12
         Top             =   390
         Width           =   285
      End
      Begin VB.Label Label4 
         Caption         =   "até"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2250
         TabIndex        =   11
         Top             =   390
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   7215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   9
      Top             =   510
      Width           =   2655
      Begin VB.OptionButton optSaidaVideo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton optSaidaImpressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1110
         TabIndex        =   5
         Top             =   420
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdFechar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2070
      Width           =   7215
   End
   Begin VB.TextBox txtNomeFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   5565
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "frmRelRecebFormaPgto.frx":4E95A
      Height          =   315
      Left            =   570
      TabIndex        =   1
      Top             =   90
      Width           =   1155
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
      Columns(0).Width=   3200
      _ExtentX        =   2037
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Filial"
   End
   Begin VB.Label Label1 
      Caption         =   "Filial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   345
   End
End
Attribute VB_Name = "frmRelRecebFormaPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
'07/07/2006 - Andrea
'Criação de relatório Recebimentos por Forma de Pagamento
'para atender cliente específico (BeStar)
'--------------------------------------------------------
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


Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  
  On Error GoTo ErrHandler
  
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
  
  dbTemp.Execute "DELETE * FROM tblRelRecebFormaPgto WHERE Owner = " & gnUserCode, dbFailOnError
   
  '---[ Chamada das funções para geração da tabela temporária ]---'
  
  Call StatusMsg("Gerando as informações do relatório, aguarde...")
  Call GeraRelatorio
  
  Rem  Nome do BD
  'Para o BD QuickStore
  'crpView.DataFiles(0) = gsQuickDBFileName
  
  'Para o BD Temp
  crpView.DataFiles(0) = gsTempDBFileName
  crpView.DataFiles(1) = gsTempDBFileName
  
  Call StatusMsg("Aguarde, imprimindo...")
  
  'Muda o ponteiro do mouse para ampulheta
  MousePointer = vbHourglass
  
  ' Definição da Saída (Vídeo ou impressora)
  If optSaidaVideo = True Then crpView.Destination = 0
  If optSaidaImpressora = True Then crpView.Destination = 1
 
  ' Seta o Nome do arquivo .rpt
  crpView.ReportFileName = gsReportPath & "RecebeFormaPgto.RPT"
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 crpView
 
  ' Seta as variáveis que são passadas para o relatório
  crpView.Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
  crpView.Formulas(1) = "nome_filial = '" & txtNomeFilial.Text & "'"
  crpView.Formulas(2) = "Filtro_DataIni = '" & mskDataInicio.Text & "'"
  crpView.Formulas(3) = "Filtro_DataFim = '" & mskDataFinal.Text & "'"
  
  ' Seta a impressora para relatório
  
  Call SetPrinterName("REL", crpView)
        
  'Filtra os registros por código do usuário
  crpView.SelectionFormula = "{tblRelRecebFormaPgto.Owner} = " & gnUserCode
  
  'Executa o relatório
  crpView.Action = 1
  
  'Para o ponteiro do mouse voltar ao normal
  MousePointer = vbDefault
 
  Call StatusMsg("")
 
  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  MsgBox "Erro ao imprimir relatório: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub GeraRelatorio()
  Dim rstRel As Recordset
  Dim rstSaidas As Recordset
  Dim rstMovCheques As Recordset
  Dim strSQL As String
  Dim intFilial As Integer
  Dim dteDataInicial As Date
  Dim dteDataFinal As Date
  
  Dim dteData As Date
  Dim dblContaCliente As Double
  Dim dblDinheiro As Double
  Dim dblCartao As Double
  Dim dblValesOutros As Double
  Dim dblCheque As Double
  Dim dblChequePre As Double
  Dim dblParcelamento As Double
  Dim dblTotalContaCliente As Double
  Dim dblTotalDinheiro As Double
  Dim dblTotalCartao As Double
  Dim dblTotalValesOutros As Double
  Dim dblTotalCheque As Double
  Dim dblTotalChequePre As Double
  Dim dblTotalParcelamento As Double
  
  
  Call cboFilial_LostFocus
  
  'Filial
  If txtNomeFilial.Text <> "" Then
    intFilial = CInt(cboFilial.Text)
  End If
  
  'Intervalo
  dteDataInicial = CDate(mskDataInicio.Text)
  dteDataFinal = CDate(mskDataFinal.Text)
  
  'Abre a tabela temporária para inserir os registros a serem exibidos
  'no relatório
  Set rstRel = dbTemp.OpenRecordset("tblRelRecebFormaPgto", dbOpenDynaset)
  
  'Obtém as movimentações de saída para cálculo dos recebimentos
  strSQL = "SELECT Filial, Sequência, Data, [Recebe - Conta], [Recebe - Dinheiro], "
  strSQL = strSQL & "[Recebe - Cartão], [Recebe - Vale], [Total Prazo] "
  strSQL = strSQL & "FROM Saídas WHERE "
  If intFilial <> 0 Then
    strSQL = strSQL & "Filial = " & intFilial & " AND "
  End If
  strSQL = strSQL & "Data BETWEEN #" & Format(dteDataInicial, "MM/dd/yyyy") & "# "
  strSQL = strSQL & "AND #" & Format(dteDataFinal, "MM/dd/yyyy") & "# "
  strSQL = strSQL & "AND Efetivada AND Recebimento AND NOT [Movimentação Desfeita] "
  strSQL = strSQL & "ORDER BY Data"
  
  Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstSaidas
    If Not (.BOF And .EOF) Then
      'Atribui o valor inicial de data
      dteData = .Fields("Data").Value
      
      Do Until .EOF
        
        'Se a data for diferente da acumulada é gravado
        'os acumuladores de recebimento
        If dteData <> .Fields("Data").Value Then
          With rstRel
            .AddNew
            .Fields("Owner").Value = gnUserCode
            .Fields("Data").Value = dteData
            .Fields("ContaCliente").Value = dblTotalContaCliente
            .Fields("Dinheiro").Value = dblTotalDinheiro
            .Fields("Cartao").Value = dblTotalCartao
            .Fields("ValesOutros").Value = dblTotalValesOutros
            .Fields("Cheque").Value = dblTotalCheque
            .Fields("ChequePre").Value = dblTotalChequePre
            .Fields("Parcelamento").Value = dblTotalParcelamento
            .Update
          End With
          dteData = .Fields("Data").Value
          dblTotalContaCliente = 0
          dblTotalDinheiro = 0
          dblTotalCartao = 0
          dblTotalValesOutros = 0
          dblTotalCheque = 0
          dblTotalChequePre = 0
          dblTotalParcelamento = 0
        End If
        
        'Acumuladores
        If .Fields("Recebe - Conta").Value Then
          'Conta cliente
          Call IsDataType(dtDouble, .Fields("Total Prazo").Value, dblContaCliente)
          dblTotalContaCliente = dblTotalContaCliente + dblContaCliente
        Else
          'Dinheiro
          Call IsDataType(dtDouble, .Fields("Recebe - Dinheiro").Value, dblDinheiro)
          dblTotalDinheiro = dblTotalDinheiro + dblDinheiro
          
          'Cartão
          Call IsDataType(dtDouble, .Fields("Recebe - Cartão").Value, dblCartao)
          dblTotalCartao = dblTotalCartao + dblCartao
          
          'Vales/Outros
          Call IsDataType(dtDouble, .Fields("Recebe - Vale").Value, dblValesOutros)
          dblTotalValesOutros = dblTotalValesOutros + dblValesOutros
          
          'Parcelamento
          Call IsDataType(dtDouble, .Fields("Total Prazo").Value, dblParcelamento)
          dblTotalParcelamento = dblTotalParcelamento + dblParcelamento
          
          'Cheques
          strSQL = "SELECT Sum(Valor) AS Total FROM [Movimento - Cheques] "
          strSQL = strSQL & "WHERE Filial = " & .Fields("Filial").Value
          strSQL = strSQL & " AND Sequência = " & .Fields("Sequência").Value
          strSQL = strSQL & " AND Bom = #" & .Fields("Data").Value & "#"
          
          Set rstMovCheques = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
          Call IsDataType(dtDouble, rstMovCheques.Fields("Total").Value, dblCheque)
          rstMovCheques.Close
          Set rstMovCheques = Nothing
          dblTotalCheque = dblTotalCheque + dblCheque
          
          'Cheques Pré
          strSQL = "SELECT Sum(Valor) AS Total FROM [Movimento - Cheques] "
          strSQL = strSQL & "WHERE Filial = " & .Fields("Filial").Value
          strSQL = strSQL & " AND Sequência = " & .Fields("Sequência").Value
          strSQL = strSQL & " AND Bom <> #" & .Fields("Data").Value & "#"
          
          Set rstMovCheques = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
          Call IsDataType(dtDouble, rstMovCheques.Fields("Total").Value, dblChequePre)
          rstMovCheques.Close
          Set rstMovCheques = Nothing
          dblTotalChequePre = dblTotalChequePre + dblChequePre
        End If
        
        .MoveNext
      Loop
      
      'Grava último registro
      With rstRel
        .AddNew
        .Fields("Owner").Value = gnUserCode
        .Fields("Data").Value = dteData
        .Fields("ContaCliente").Value = dblTotalContaCliente
        .Fields("Dinheiro").Value = dblTotalDinheiro
        .Fields("Cartao").Value = dblTotalCartao
        .Fields("ValesOutros").Value = dblTotalValesOutros
        .Fields("Cheque").Value = dblTotalCheque
        .Fields("ChequePre").Value = dblTotalChequePre
        .Fields("Parcelamento").Value = dblTotalParcelamento
        .Update
      End With
    End If
    .Close
  End With
  
  Set rstSaidas = Nothing
  
  Set rstRel = Nothing
  
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  datFiliais.DatabaseName = gsQuickDBFileName

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

