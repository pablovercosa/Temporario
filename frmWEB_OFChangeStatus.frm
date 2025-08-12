VERSION 5.00
Begin VB.Form frmWEB_OFChangeStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Status do Pedido"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmWEB_OFChangeStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTraceCode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2880
      Width           =   4335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtStatusAdmin 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox txtStatusShopper 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblTraceCode 
      AutoSize        =   -1  'True
      Caption         =   "Informe o código de rastreamento do Pedido Enviado"
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   3765
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Status visível ao comprador"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1980
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Status visível ao administrador"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2160
   End
End
Attribute VB_Name = "frmWEB_OFChangeStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSucess As Boolean
Private mlngID As Long
Private menuStep As enWEB_OrderFormStep
Private menuCurrentStep As enWEB_OrderFormStep
Private mblnChanged As Boolean
Private mblnNewStep As Boolean
Private mbytFilial As Byte
Private mlngSequence As Long
Private Const ERR_EFETIVA_ENTRADA = vbObjectError + 513
Private Const ERR_EFETIVA_ENTRADA_PROD_NC = vbObjectError + 514

Public Function ChangeStatus(ByVal lngID As Long, _
                             ByVal enuStep As enWEB_OrderFormStep, _
                             Optional ByVal blnNewStep As Boolean = False, _
                             Optional ByVal enuCurrentStep As enWEB_OrderFormStep, _
                             Optional ByVal bytFilial As Byte, _
                             Optional ByVal lngSequence As Long) As Boolean
                             
  Dim strStatusShopper As String
  Dim strStatusAdmin As String
  
  mblnNewStep = blnNewStep
  menuCurrentStep = enuCurrentStep
  mbytFilial = bytFilial
  mlngSequence = lngSequence
  
  If blnNewStep Then
    Call GetDataDescPasso(enuStep, strStatusShopper, strStatusAdmin)
  Else
    Call GetOrderFormVStatus(lngID, strStatusShopper, strStatusAdmin)
  End If
  
  txtStatusShopper.Text = strStatusShopper
  txtStatusAdmin.Text = strStatusAdmin
  If mblnNewStep Then
    txtTraceCode.Enabled = enuStep = ofsHasSent
    lblTraceCode.Enabled = enuStep = ofsHasSent
  End If
  
  mlngID = lngID
  menuStep = enuStep
  mblnChanged = False
  mblnSucess = False
  
  Me.Show vbModal
  ChangeStatus = mblnSucess
  Set frmWEB_OFChangeStatus = Nothing
  
End Function

Private Sub cmdCancel_Click()
  mblnSucess = False
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim strAux As String
  Dim blnInTransaction As Boolean
  Dim intRet As Integer
  
  On Error GoTo ErrHandler
  
  If mblnNewStep And menuStep = ofsHasSent Then
    If Trim(txtTraceCode.Text) = "" Then
      MsgBox "Preencher o código de rastreamento do Pedido Enviado.", vbExclamation, "Atenção"
      txtTraceCode.SetFocus
      Exit Sub
    End If
  End If
  
  If mblnChanged Or mblnNewStep Then
    ws.BeginTrans
    blnInTransaction = True
    
    If mblnNewStep And menuStep = ofsHasSent Then
      strAux = "TraceCode = '" & txtTraceCode.Text & "', "
    End If
    
    'Atualiza o Pedido
    Call db.Execute("UPDATE WEB_OrderForms SET " & strAux & _
      "StatusShopper = '" & txtStatusShopper.Text & _
      "', StatusAdmin = '" & txtStatusAdmin.Text & _
      "', Passo = " & menuStep & " WHERE ID = " & mlngID, dbFailOnError)
    
    'Atualiza o Histórico do Pedido
    Call db.Execute("INSERT INTO WEB_OrderStatusHistoric " & _
      "(OrderFormID, Passo, StatusShopper, StatusAdmin, Data, WebSynchronize) " & _
      "VALUES (" & mlngID & ", " & menuStep & ", '" & txtStatusShopper.Text & _
      "', '" & txtStatusAdmin.Text & "', #" & Format(Now, "MM/DD/YYYY HH:MM:SS") & _
      "#, True)", dbFailOnError)
    
    'Em caso de cancelamento verifica a situação
    If mblnNewStep And menuStep = ofsCanceled Then
      Select Case menuCurrentStep
        Case ofsReceived 'Pedido recebido
          
          'Cria movimentação de entrada (repor estoque reservado)
          Call CreateMovementCancelOrderForm(menuCurrentStep)
        
          'Verifica se a venda foi criada
          If mlngSequence > 0 Then
            'Exclui venda
            Call EraseTypeMoviment(tmSaidas, mbytFilial, mlngSequence)
            Call EraseTypeMoviment(tmSaidasProdutos, mbytFilial, mlngSequence)
          End If
        
        'Pagamento confirmado ou Pedido embalado
        Case ofsConfirmedPayment, ofsPacked
          
          'Cria movimentação de entrada (repor estoque reservado)
          Call CreateMovementCancelOrderForm(menuCurrentStep)
          
          'Desfaz a venda e o recebimento
          ws.BeginTrans
          intRet = Desefetiva_Saída(CInt(mbytFilial), mlngSequence)
          If intRet = 0 Then
            ws.CommitTrans
          Else
            ws.Rollback
            MsgBox "Erro [" & intRet & "] ao desefetivar venda.", vbCritical, "Erro"
            Exit Sub
          End If
          
          'Exclui venda
          Call EraseTypeMoviment(tmSaidas, mbytFilial, mlngSequence)
          Call EraseTypeMoviment(tmSaidasProdutos, mbytFilial, mlngSequence)
          
      End Select
    End If
    
    ws.CommitTrans
    blnInTransaction = False
  End If
  Call StatusMsg("")
  
  mblnSucess = True
  Unload Me
  
  Exit Sub
  
ErrHandler:
  If blnInTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub txtStatusAdmin_Change()
  mblnChanged = True
End Sub

Private Sub txtStatusShopper_Change()
  mblnChanged = True
End Sub

Private Sub GetOrderFormVStatus(ByVal lngID As Long, ByRef strStatusShopper As String, ByRef strStatusAdmin As String)
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT StatusShopper, StatusAdmin FROM WEB_OrderForms WHERE ID = " & lngID
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not .BOF And Not .EOF Then
      strStatusShopper = .Fields("StatusShopper").Value
      strStatusAdmin = .Fields("StatusAdmin").Value
    End If
    .Close
  End With
  Set rs = Nothing
  
End Sub

Private Sub CreateMovementCancelOrderForm(ByVal enuStep As enWEB_OrderFormStep)
  Dim intCodOpCancelamento As Integer
  Dim rsEntradas As Recordset
  Dim lngSequence As Long
  Dim strSQL As String
  Dim intRet As Integer
  
  Dim rsItens As Recordset
  Dim rsEntradasProdutos As Recordset
  Dim intX As Integer
  Dim strCodProd As String
  Dim intErro As Integer
  
  'Obtém o código para a operação de venda
  Call GetWEBCod_Op(0, 0, intCodOpCancelamento)
  
  'Nova sequência
  lngSequence = gnGetNextSequencia(CInt(mbytFilial))
  
  Set rsEntradas = db.OpenRecordset("Entradas", dbOpenDynaset)
  With rsEntradas
    .AddNew
    .Fields("Filial").Value = mbytFilial
    .Fields("Data").Value = Date
    .Fields("Sequência").Value = lngSequence
    .Fields("Operação").Value = intCodOpCancelamento
    .Fields("Digitador").Value = 0
    .Fields("Fornecedor").Value = 0
    .Fields("Observações").Value = ""
    .Fields("Nota Fiscal").Value = ""
    .Fields("Data Emissão").Value = Null
    .Fields("Pedido").Value = ""
    .Fields("Forma Pagto").Value = 0
    .Fields("Produtos").Value = 0
    .Fields("Desconto").Value = 0
    .Fields("IPI").Value = 0
    .Fields("Frete").Value = 0
    .Fields("Base ICM").Value = 0
    .Fields("Valor ICM").Value = 0
    .Fields("Base ICM Subs").Value = 0
    .Fields("Valor ICM Subs").Value = 0
    .Fields("Total").Value = 0
    .Fields("Dinheiro Caixa").Value = 0
    .Fields("Cheque Caixa").Value = 0
    .Fields("Caixa").Value = 0
    .Fields("Conta").Value = 0
    .Fields("Num Cheque").Value = ""
    .Fields("Bom para").Value = Null
    .Fields("Valor Cheque").Value = 0
    .Fields("Descrição").Value = "Cancelamento de pedido da Loja Virtual"
    .Fields("Efetivada").Value = False
    .Fields("Nota Impressa").Value = 0
    .Fields("Nota Cancelada").Value = False
    .Fields("Data Acerto Empréstimo").Value = Null
    .Fields("WebOrderFormID").Value = mlngID
    .Update
    .Close
  End With
  Set rsEntradas = Nothing
  
  
  '06/03/2003 - mpdea
  'Adicionado tratamento para cancelamento de pedidos com status igual a recebido
  If enuStep = ofsReceived Then
    'Copia os itens do Pedido para Entradas
    
    'Itens do pedido
    strSQL = "SELECT * FROM WEB_OrderItens WHERE OrderFormID = " & mlngID
    Set rsItens = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    Set rsEntradasProdutos = db.OpenRecordset("Entradas - Produtos", dbOpenDynaset)
    
    intX = 0
    With rsItens
      If Not .BOF And Not .EOF Then
        Do Until .EOF
          Call Acha_Produto(.Fields("sku").Value, strCodProd, 0, 0, 0, 0, intErro)
          If intErro = 0 Then
            'Inclui item na venda
            intX = intX + 1
            With rsEntradasProdutos
              .AddNew
              .Fields("Filial").Value = mbytFilial
              .Fields("Sequência").Value = lngSequence
              .Fields("Linha").Value = intX
              .Fields("Código").Value = rsItens.Fields("sku").Value
              .Fields("Qtde").Value = rsItens.Fields("Quantity").Value
              .Fields("Preço").Value = rsItens.Fields("ListPrice").Value
              'Desconto
              If CCur("0" & rsItens.Fields("Discount").Value) > 0 Then
                .Fields("Desconto").Value = rsItens.Fields("Discount").Value / rsItens.Fields("ListPrice").Value * 100
                .Fields("Desconto Valor").Value = rsItens.Fields("Discount").Value
              End If
              .Fields("ICM").Value = 0
              .Fields("IPI").Value = 0
              .Fields("Preço Final").Value = rsItens.Fields("Total").Value
              .Fields("Etiqueta").Value = False
              .Fields("Código sem Grade").Value = strCodProd
              .Fields("InGeradoViaConsig").Value = False
              .Update
            End With
          Else
            Err.Raise ERR_EFETIVA_ENTRADA_PROD_NC, "WEB - Movimentação de Entrada", _
              "Produto [" & .Fields("sku").Value & "] não cadastrado."
            Exit Sub
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    
    rsEntradasProdutos.Close
    
    Set rsItens = Nothing
    Set rsEntradasProdutos = Nothing
    
  Else
    
    '
    ' ANALISAR SQL ------------ !! mpdea !! --- 06/03/2003
    '
    
    'Copia os itens de Saídas para Entradas
    strSQL = "INSERT INTO [Entradas - Produtos] " & _
             "SELECT SP.Filial, SP.Sequência, SP.Linha, SP.Código, " & _
             "SP.Qtde, SP.Preço, SP.Desconto, SP.ICM, SP.IPI, SP.[Preço Final], " & _
             "SP.Etiqueta, SP.[Código sem Grade], SP.InGeradoViaConsig " & _
             "FROM [Saídas - Produtos] AS SP WHERE SP.Filial = " & mbytFilial & _
             " AND SP.Sequência = " & mlngSequence
    
    db.Execute strSQL, dbFailOnError
  
    'Atualiza para sequência atual
    db.Execute "UPDATE [Entradas - Produtos] SET Sequência = " & lngSequence & _
               " WHERE Filial = " & mbytFilial & " AND Sequência = " & mlngSequence, _
               dbFailOnError
    
  End If
  
  intRet = Efetiva_Entrada(CInt(mbytFilial), lngSequence)
  If intRet <> 0 Then
    Err.Raise ERR_EFETIVA_ENTRADA, "WEB - Movimentação de Entrada", _
      "Erro [" & intRet & "] ao efetivar movimentação de entrada."
  End If
  
End Sub
