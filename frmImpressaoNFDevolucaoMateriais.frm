VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmImpressaoNFDevolucaoMateriais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impress�o de Notas Fiscais (Devolu��o de Materiais)"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpressaoNFDevolucaoMateriais.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2280
   ScaleWidth      =   6060
   Begin VB.Frame Frame4 
      Caption         =   "Intervalo para Notas Fiscais"
      Height          =   900
      Left            =   0
      TabIndex        =   12
      Top             =   820
      Width           =   2980
      Begin VB.TextBox txtNFFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   2
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox txtNFIni 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   1
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fim"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "Im&primir"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   " Per�odo ( Vendas ) "
      Height          =   900
      Left            =   3030
      TabIndex        =   10
      Top             =   820
      Width           =   2980
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDataInicial 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fim"
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   540
         Width           =   90
      End
   End
   Begin VB.Data datFornecedor 
      Caption         =   "datFornecedor"
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
      RecordSource    =   "SELECT C�digo, Nome, Tipo FROM Cli_For WHERE Tipo = 'F' ORDER BY Nome, C�digo"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtNomeFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   3855
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmImpressaoNFDevolucaoMateriais.frx":058A
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "� necess�rio que a Filial de Sa�da esteja cadastrada como Fornecedor"
         Top             =   240
         Width           =   885
         DataFieldList   =   "C�digo"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   1561
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   65535
         DataFieldToDisplay=   "C�digo"
      End
      Begin VB.Label lblFornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmImpressaoNFDevolucaoMateriais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Vari�veis oriundas da table ParamDevoMat
Dim m_bytFilial   As Byte
Dim m_intOperacao As Integer
Dim m_bytCaixa    As Byte
Dim m_strTabela   As String
Dim m_lngSeq      As Long

Private Sub cmdImprimir_Click()
  
  If ValidarCampos Then Exit Sub
  If VerificarTableParamDevoMat Then Exit Sub
  If VerificarNF Then Exit Sub
  
  Call StatusMsg("Aguarde...")
  Screen.MousePointer = vbHourglass
  
  Call CriarSaidas
  If m_lngSeq <> 0 Then Call ImprimindoNF(m_bytFilial, m_lngSeq)
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
End Sub

Private Sub cmdSair_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFornecedor.DatabaseName = gsQuickDBFileName
  
End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(0).Text
  cboFornecedor_LostFocus
End Sub

Private Sub cboFornecedor_LostFocus()
  Dim rstFornecedor As Recordset
  
  txtNomeFornecedor.Text = ""
  
  If Not IsNumeric(cboFornecedor.Text) Then Exit Sub
  
  Set rstFornecedor = db.OpenRecordset("SELECT C�digo, Nome, Tipo FROM Cli_For WHERE C�digo = " & CInt(cboFornecedor.Text) & " AND Tipo ='" & "F" & "'" & " ORDER BY C�digo ", dbOpenDynaset)

  With rstFornecedor
    If Not (.BOF And .EOF) Then
      txtNomeFornecedor.Text = .Fields("Nome") & ""
    End If
  End With

  rstFornecedor.Close
  Set rstFornecedor = Nothing

End Sub

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinal.Text = frmCalendario.gsDateCalender(mskDataFinal.Text)
  End If
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal.Text = Ajusta_Data(mskDataFinal.Text)
End Sub

Private Sub mskDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicial.Text = frmCalendario.gsDateCalender(mskDataInicial.Text)
  End If
End Sub

Private Sub mskDataInicial_LostFocus()
  mskDataInicial.Text = Ajusta_Data(mskDataInicial.Text)
End Sub

Private Function ValidarCampos() As Boolean
  
  If Len(txtNomeFornecedor.Text) <= 0 Then
    ValidarCampos = True
    MsgBox "Fornecedor inv�lido, verifique.", vbExclamation, "Aten��o"
    cboFornecedor.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataInicial.Text) Then
    ValidarCampos = True
    MsgBox "Data Inicial inv�lida, verifique.", vbExclamation, "Aten��o"
    mskDataInicial.SetFocus
    Exit Function
  End If
  
  
  If Not IsDate(mskDataFinal.Text) Then
    ValidarCampos = True
    MsgBox "Data Final inv�lida, verifique.", vbExclamation, "Aten��o"
    mskDataFinal.SetFocus
    Exit Function
  End If
  
  If CDate(mskDataInicial.Text) > CDate(mskDataFinal.Text) Then
    ValidarCampos = True
    MsgBox "Data Inicial maior que a Final, verifique.", vbExclamation, "Aten��o"
    mskDataInicial.SetFocus
    Exit Function
  End If
  
End Function

Private Function VerificarTableParamDevoMat() As Boolean
  Dim rstParamDevoMat As Recordset
  Dim strSQL          As String
  
  strSQL = "SELECT * FROM ParamDevoMat "
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  
  Set rstParamDevoMat = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstParamDevoMat.RecordCount = 0 Then
    VerificarTableParamDevoMat = True
    MsgBox "Configure a Sa�da em: Par�metros / Devolu��o de Materiais.", vbExclamation, "Aten��o"
    Exit Function
  End If
  
  With rstParamDevoMat
    If Not (.BOF And .EOF) Then
      .MoveFirst
      m_bytFilial = .Fields("Filial").Value
      m_intOperacao = .Fields("Operacao").Value
      m_bytCaixa = .Fields("Caixa").Value
      m_strTabela = .Fields("Tabela").Value & ""
    End If
    .Close
  End With

  Set rstParamDevoMat = Nothing

End Function

Private Function VerificarNF() As Boolean
  Dim rstParametros As Recordset
  Dim strSQL        As String
  
  strSQL = "SELECT [Nota Sa�da] FROM [Par�metros Filial] "
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  
  Set rstParametros = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      .MoveFirst
      If Len(.Fields("Nota Sa�da").Value) <= 0 Then VerificarNF = True
    End If
    .Close
  End With
  
  Set rstParametros = Nothing
  
  If VerificarNF Then MsgBox "N�o h� nenhuma NF padr�o cadastrada em Par�metros Filial, verifique.", vbExclamation, "Aten��o"
  
End Function

Private Sub CriarSaidas()
  'Ser� criado nesta procedure:
  'Sa�das, [Sa�das - Produtos] e atualizar o Par�metros
  Dim rstParametros         As Recordset
  Dim rstSaidas             As Recordset
  Dim rstSaidasProdutos     As Recordset
  Dim rstPrestacaoContas    As Recordset
  
  Dim strQuery              As String
  
  Dim nSequencia            As Long
  Dim blnTransaction        As Boolean
  
  Dim nRet                  As Integer

  Dim bytLinha              As Byte
  Dim sngPercentualIPI      As Single
  Dim sngPercentualIcmSaida As Single
  Dim strUnidadeVenda       As String
  Dim dblPreco              As Double
  Dim blnFechar             As Boolean

  On Error GoTo Err_Handlel
  
  '-------------------------------------
  'Abrir a transa��o
  '-------------------------------------
  ws.BeginTrans
  blnTransaction = True
        
      '*** Opera��es com o DB
      
      'Presta��o de Contas
      strQuery = "SELECT * FROM PrestacaoContas "
      strQuery = strQuery & " WHERE Fornecedor = " & CLng(cboFornecedor.Text)
      strQuery = strQuery & " AND Finalizado "
      strQuery = strQuery & " AND NOT ImpressoNF "
      strQuery = strQuery & " AND DatadaFinalizacao >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
      strQuery = strQuery & " AND DatadaFinalizacao <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
      strQuery = strQuery & " AND QtdeDevolvida <> 0 " 'HOUVE DEVOLU��O !!!
      '18/11/2004 - Daniel
      'Adicionado filtro por NF
      If Len(txtNFIni.Text) > 0 And Len(txtNFFin.Text) > 0 Then
        If CLng(txtNFIni.Text) <= CLng(txtNFFin.Text) Then
          strQuery = strQuery & " AND NotaFiscal >= " & CLng(txtNFIni.Text)
          strQuery = strQuery & " AND NotaFiscal <= " & CLng(txtNFFin.Text)
        End If
      End If
      
      strQuery = strQuery & " ORDER BY Filial, Sequencia, Linha, DatadaFinalizacao "
      
      Set rstPrestacaoContas = db.OpenRecordset(strQuery, dbOpenDynaset)
      
      If rstPrestacaoContas.RecordCount = 0 Then
        MsgBox "Nenhuma informa��o encontrada neste intervalo, verifique.", vbInformation, "Aten��o"
        '-------------------------------------
        'Fechar a transa��o
        '-------------------------------------
        ws.CommitTrans
        blnTransaction = False
        Exit Sub
      End If
      
      'Buscar uma pr�xima Sequ�ncia
      nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("�ltima Movimenta��o") + 1
      
      'Abrimos Sa�da e [Sa�das - Produtos]
      Set rstSaidas = db.OpenRecordset("Sa�das", dbOpenDynaset)
      Set rstSaidasProdutos = db.OpenRecordset("Sa�das - Produtos", dbOpenDynaset)
      
      With rstPrestacaoContas
        If Not (.BOF And .EOF) Then
          .MoveFirst
          
          Do Until .EOF
            'PrestacaoContas
            .Edit
            .Fields("ImpressoNF").Value = True
            '18/11/2004 - Daniel
            'Tratamento caso haja uma Devolu��o Parcial
            If .Fields("Resultado").Value = 1 Then 'Devolu��o
              If .Fields("QtdeOriginal").Value > .Fields("QtdeDevolvida").Value Then
                'Valida��o caso o SUM esteja igual a QtdeOriginal n�o daremos UpdateFieldSelecionadoEP
                If Not CompletouDevolucao(.Fields("Filial").Value, .Fields("Sequencia").Value, .Fields("Linha").Value, .Fields("QtdeOriginal").Value) Then
                  'Liberar o campo Selecionado em EP para False
                  'para podermos carregar em uma pr�xima presta��o
                  blnFechar = False
                  Call UpdateFieldSelecionadoEP(.Fields("Filial").Value, .Fields("Sequencia").Value, .Fields("Linha").Value, blnFechar)
                Else 'Fechou
                  blnFechar = True
                  Call UpdateFieldSelecionadoEP(.Fields("Filial").Value, .Fields("Sequencia").Value, .Fields("Linha").Value, blnFechar)
                  Call VerificarConsignacoesDaEntrada(.Fields("Filial").Value, .Fields("Sequencia").Value)
                End If
              End If
            End If
            '------------------------------------------
            .Update
          
          
            bytLinha = bytLinha + 1
            Call BuscarPercentuais(.Fields("Produto").Value, sngPercentualIPI, sngPercentualIcmSaida, strUnidadeVenda)
            dblPreco = dblPreco + Format((.Fields("QtdeDevolvida").Value * .Fields("Custo").Value), FORMAT_VALUE)
          
            'Cria��o das [Sa�das - Produtos]
            rstSaidasProdutos.AddNew
            rstSaidasProdutos.Fields("Filial").Value = gnCodFilial
            rstSaidasProdutos.Fields("Sequ�ncia").Value = nSequencia
            rstSaidasProdutos.Fields("Linha").Value = bytLinha
            rstSaidasProdutos.Fields("C�digo").Value = .Fields("Produto").Value & ""
            rstSaidasProdutos.Fields("Qtde").Value = .Fields("QtdeDevolvida").Value
            rstSaidasProdutos.Fields("Pre�o").Value = .Fields("Custo").Value
            rstSaidasProdutos.Fields("Desconto").Value = 0
            rstSaidasProdutos.Fields("Desconto Valor").Value = 0
            rstSaidasProdutos.Fields("ICM").Value = sngPercentualIPI      'Valor da taxa ICM do produto
            rstSaidasProdutos.Fields("IPI").Value = sngPercentualIcmSaida 'Valor da taxa IPI do produto
            rstSaidasProdutos.Fields("Pre�o Final").Value = Format((.Fields("QtdeDevolvida").Value * .Fields("Custo").Value), FORMAT_VALUE)
            rstSaidasProdutos.Fields("Etiqueta").Value = False
            rstSaidasProdutos.Fields("C�digo sem Grade").Value = .Fields("Produto").Value & ""
            If Len(strUnidadeVenda) > 0 Then
              rstSaidasProdutos.Fields("Unidade Venda").Value = strUnidadeVenda
            Else
              rstSaidasProdutos.Fields("Unidade Venda").Value = "UN"
            End If
            rstSaidasProdutos.Fields("QtdeEntregue").Value = 0
            
            rstSaidasProdutos.Update
            
          .MoveNext
          Loop
          
        End If
        .Close
      End With
      
      Set rstPrestacaoContas = Nothing
      
      rstSaidasProdutos.Close
      Set rstSaidasProdutos = Nothing
      
      'Sa�das
      With rstSaidas
        .AddNew
        .Fields("Filial").Value = gnCodFilial
        .Fields("Data").Value = Data_Atual
        .Fields("Sequ�ncia").Value = nSequencia
        .Fields("Opera��o").Value = m_intOperacao
        .Fields("Caixa").Value = m_bytCaixa
        .Fields("Tabela").Value = m_strTabela
        .Fields("Digitador").Value = gnUserCode
        .Fields("Operador").Value = gnUserCode
        .Fields("Cliente").Value = CLng(cboFornecedor.Text)
        .Fields("Observa��es").Value = "Sa�da criada em " & Now
        .Fields("Produtos").Value = Format(dblPreco, FORMAT_VALUE)
        .Fields("Servi�os").Value = 0
        .Fields("Total").Value = Format(dblPreco, FORMAT_VALUE)
        .Fields("Efetivada").Value = False 'No Efetiva Sa�da ficar� True...
        .Fields("Recebimento").Value = False
        .Fields("Nota Impressa").Value = 0
        .Fields("Valor Recebido").Value = 0
        
        .Update
        .Close
      End With
      
      Set rstSaidas = Nothing
      
      bytLinha = 0
      dblPreco = 0
      

      '-------------------------------------------------------
      'EFETIVA A SA�DA
      '-------------------------------------------------------
      Call StatusMsg("Aguarde, efetivando venda...")
  
      nRet = Efetiva_Sa�da(gnCodFilial, nSequencia)
  
      If nRet <> 0 Then
        Select Case nRet
          Case -1
            'A��o cancelada
            Call StatusMsg("A��o cancelada.")
          Case 5
            Call DisplayMsg("Tabela de pre�os inexistente.")
          Case Else
            Call DisplayMsg("Opera��o N�O efetivada. Erro" & str(nRet))
        End Select
        'Cancelamento da transa��o
        ws.Rollback
        Exit Sub
      End If
      '-------------------------------------------------------
      'FIM DA EFETIVA A SA�DA
      '-------------------------------------------------------

      'Tratamento para Atualiza��o de Par�metros
      Set rstParametros = db.OpenRecordset(" SELECT [�ltima Movimenta��o] FROM [Par�metros Filial] WHERE Filial =" & gnCodFilial, dbOpenDynaset)
  
        rstParametros.Edit
        rstParametros.Fields("�ltima Movimenta��o").Value = nSequencia
        rstParametros.Update
        rstParametros.Close
  
      Set rstParametros = Nothing
      'Fim do Tratamento para Atualiza��o de Par�metros

      '*** Final de Opera��es com o DB
      
      '-------------------------------------------------------
      'POPULAR VARI�VEIS PARA A EMISS�O DA NOTA FISCAL
      '-------------------------------------------------------
      'm_bytFilial (j� est� com valor...)
      m_lngSeq = nSequencia

  
  '-------------------------------------
  'Fechar a transa��o
  '-------------------------------------
  ws.CommitTrans
  blnTransaction = False
  
  Call StatusMsg("")
  
  Exit Sub

Err_Handlel:
  If blnTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Erro na transa��o"

End Sub

Private Sub BuscarPercentuais(ByVal Produto As String, ByRef PercentualIPI As Single, ByRef PercentualIcmSaida As Single, ByRef UnidadeVenda As String)
  Dim rstProdutos As Recordset
  Dim strSQL      As String
  
  strSQL = "SELECT [Percentual IPI], [Percentual Icm Saida], [Unidade Venda] "
  strSQL = strSQL & " FROM Produtos "
  strSQL = strSQL & " WHERE C�digo = '" & Produto & "'"
   
  Set rstProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If IsNumeric(.Fields("Percentual IPI").Value) Then
        PercentualIPI = .Fields("Percentual IPI").Value
      Else
        PercentualIPI = 0
      End If
      
      If IsNumeric(.Fields("Percentual Icm Saida").Value) Then
        PercentualIcmSaida = .Fields("Percentual Icm Saida").Value
      Else
        PercentualIcmSaida = 0
      End If
      
      UnidadeVenda = .Fields("Unidade Venda").Value & ""
      
    End If
    .Close
  End With
  
  Set rstProdutos = Nothing

End Sub

Private Sub ImprimindoNF(ByVal Filial As Byte, ByVal Sequencia As Long)
  'Copiado a Private do mesmo modo que existe em Sa�das
  Dim strSQL                As String
  Dim intX                  As Integer
  Dim strFileNF             As String
  Dim intRet                As Integer
  Dim lngNotaFiscal         As Long
  Dim blnInTransaction      As Boolean
  Dim intRepeatUpdateLocked As Integer
  
  Dim rstSaidas             As Recordset
  Dim rstParametros         As Recordset
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")
  
  'Abrir a tabela Par�metros
  strSQL = ""
  strSQL = "SELECT * FROM [Par�metros Filial]"
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  
  Set rstParametros = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  strSQL = ""
  lngNotaFiscal = 0 'Impressa pela primeira vez...
  
  '--------------------------------------------------------------------------
  'Grava nova NF
  '--------------------------------------------------------------------------
  If lngNotaFiscal = 0 Then
    'Modificado leitura e grava��o do n�mero da �ltima nota fiscal
    'Inclu�do transa��o durante grava��o
    'lngNotaFiscal = rsParametros.Fields("�ltima Nota").Value + 1
    '
    ws.BeginTrans
    blnInTransaction = True
    
    lngNotaFiscal = g_lngNextNotaFiscal(Filial)
    
    strSQL = "SELECT * FROM Sa�das WHERE Filial = " & Filial
    strSQL = strSQL & " AND Sequ�ncia = " & Sequencia
    
    Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)

    With rstSaidas
      .LockEdits = True
      .Edit
      .Fields("Nota Impressa").Value = lngNotaFiscal
      
      .Update
      .LockEdits = False
    End With
    
    'Finaliza transa��o
    ws.CommitTrans
    blnInTransaction = False
  End If
  '--------------------------------------------------------------------------
  
  
  '--------------------------------------------------------------------------
  'Imprime NF
  '--------------------------------------------------------------------------
  strFileNF = gsConfigPath + rstParametros.Fields("Nota Sa�da").Value + ".CNF"
  intRet = Imprime_Nota(strFileNF, rstSaidas.Fields("Filial").Value, rstSaidas.Fields("Sequ�ncia").Value)
  If intRet = 0 Then
    '14/04/2003 - mpdea
    'Atualiza a data da impress�o da nota fiscal
    strSQL = "UPDATE Sa�das SET DataEmissaoNota = #"
    strSQL = strSQL & Format(Date, "mm/dd/yyyy") & "# "
    strSQL = strSQL & "WHERE Filial = " & rstSaidas.Fields("Filial").Value
    strSQL = strSQL & " AND Sequ�ncia = " & rstSaidas.Fields("Sequ�ncia").Value
    db.Execute strSQL, dbFailOnError
    
    DisplayMsg "Nota [" & lngNotaFiscal & "] impressa com sucesso."
  Else
    DisplayMsg "Houve o erro " & intRet & " durante a impress�o da Nota."
  End If
  '--------------------------------------------------------------------------
  
  'Fechar os Recordsets
  rstParametros.Close
  rstSaidas.Close
  Set rstParametros = Nothing
  Set rstSaidas = Nothing
  
  'Limpar as vars modulares
  'm_bytFilial (n�o precisa...)
  m_lngSeq = 0
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3197, 3187, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Resume
        Else
          'Cancelamento da transa��o
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
      
'        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
'          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Sa�das - Imprimir Nota Fiscal") = vbOK Then
'          intRepeatUpdateLocked = 0
'          Resume
'        Else
'          'Cancelamento da transa��o
'          If blnInTransaction Then ws.Rollback
'          Exit Sub
'        End If
      End If
    Case Else
      'Cancelamento da transa��o
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  End Select

End Sub

Private Sub UpdateFieldSelecionadoEP(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte, ByVal Fechar As Boolean)
  '18/11/2004 - Daniel
  Dim rstEntraProdu As Recordset
  Dim strSQL        As String
  
  strSQL = "SELECT Selecionado, ConsignacaoFechada FROM [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequ�ncia = " & Sequencia
  strSQL = strSQL & " AND Linha = " & Linha

  Set rstEntraProdu = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntraProdu
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Edit
      If Not Fechar Then
        '.Fields("ConsignacaoFechada").Value = False
        .Fields("Selecionado").Value = False
      Else
        .Fields("ConsignacaoFechada").Value = True
        .Fields("Selecionado").Value = True
      End If
      .Update
    End If
    .Close
  End With
  
  Set rstEntraProdu = Nothing

End Sub

Private Function CompletouDevolucao(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte, ByVal QtdeOriginal As Double) As Boolean
  Dim rstPrestacao As Recordset
  Dim strSQL       As String
  Dim dblSomas     As Double
  
  strSQL = "SELECT SUM(QtdeDevolvida) AS Devolvida, MAX(QtdeVendida) AS Vendida, SUM(QtdeComprada) AS Comprada "
  strSQL = strSQL & " FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequencia = " & Sequencia
  strSQL = strSQL & " AND Linha = " & Linha

  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      dblSomas = .Fields("Devolvida").Value + .Fields("Vendida").Value + .Fields("Comprada").Value
      
      If dblSomas = QtdeOriginal Then CompletouDevolucao = True
      
    End If
    .Close
  End With

  Set rstPrestacao = Nothing

End Function

Private Sub VerificarConsignacoesDaEntrada(ByVal Filial As Byte, ByVal Sequencia As Long)
  Dim rstEntradas As Recordset
  Dim strSQL      As String
  Dim blnFlag     As Boolean
  
  strSQL = ""
  strSQL = "SELECT [Entradas - Produtos].ConsignacaoFechada "
  strSQL = strSQL & " FROM Entradas, [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Entradas.Filial = " & Filial
  strSQL = strSQL & " AND Entradas.Sequ�ncia = " & Sequencia
  strSQL = strSQL & " AND [Entradas - Produtos].Sequ�ncia = Entradas.Sequ�ncia "

  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        blnFlag = .Fields("ConsignacaoFechada").Value
        
        If Not blnFlag Then Exit Do
        
      .MoveNext
      Loop
    
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing

  If blnFlag Then
  
    strSQL = ""
    strSQL = "SELECT Entradas.ConsignacaoFechada "
    strSQL = strSQL & " FROM Entradas "
    strSQL = strSQL & " WHERE Entradas.Filial = " & Filial
    strSQL = strSQL & " AND Entradas.Sequ�ncia = " & Sequencia
  
    Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)

    With rstEntradas
      If Not (.BOF And .EOF) Then
        .MoveFirst
        .Edit
        .Fields("ConsignacaoFechada").Value = True
        .Update
      End If
      .Close
    End With
  
    Set rstEntradas = Nothing
  
  End If

End Sub

