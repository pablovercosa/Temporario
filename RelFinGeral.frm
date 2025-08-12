VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelFinGeral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório da Situação Financeira Geral"
   ClientHeight    =   2070
   ClientLeft      =   1965
   ClientTop       =   2040
   ClientWidth     =   5385
   HelpContextID   =   1630
   Icon            =   "RelFinGeral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2070
   ScaleWidth      =   5385
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   525
         Width           =   1095
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   3975
      TabIndex        =   4
      Top             =   1530
      Width           =   1335
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   345
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   4890
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelFinGeral.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   795
      TabIndex        =   0
      Top             =   165
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
      Columns(0).Width=   8811
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
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
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filial:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   585
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1635
      TabIndex        =   1
      Top             =   165
      Width           =   3675
   End
End
Attribute VB_Name = "frmRelFinGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTempo As Recordset
Dim rsParametros As Recordset

Private Sub InsertItemIntoZZZGeral(ByVal sTexto As String, _
  ByVal cValor1 As Currency, ByVal cValor2 As Currency, ByVal nOrdem As Long, _
  Optional ByVal bCheckPosition As Boolean = False)
  
  'Verifica o tamanho do texto
  sTexto = IIf(sTexto = "", Space(1), Left(sTexto, 70))
  If bCheckPosition Then
    'Verifica se a posição do valor está correta, somente informado-se cValor1
    If cValor1 < 0 Then
      cValor2 = Abs(cValor1)
      cValor1 = 0
    End If
  End If
  With rsTempo
    .AddNew
    !Texto = sTexto
    ![Valor 1] = cValor1
    ![Valor 2] = cValor2
    !Ordem = nOrdem
    .Update
  End With
End Sub

Private Sub CheckValuePedidos(ByVal nFilial As Byte, ByRef cTotalPendente As Currency)
  Dim rsPedidos As Recordset
  Dim sSql As String
  
  sSql = "SELECT Sum(Entradas.Total) AS SaldoFinal FROM Entradas " & _
    "INNER JOIN [Operações Entrada] ON Entradas.Operação = " & _
    "[Operações Entrada].Código WHERE Entradas.Filial = " & nFilial & _
    "AND NOT Entradas.Efetivada AND [Operações Entrada].Tipo = 'P';"
  Set rsPedidos = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsPedidos
    cTotalPendente = IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
    .Close
  End With
  Set rsPedidos = Nothing
End Sub

Private Sub CheckValueCartao(ByVal nFilial As Byte, ByRef cTotalReceber As Currency)
  Dim rsCartao As Recordset
  Dim sSql As String
  
  sSql = "SELECT Sum(Valor) AS SaldoFinal FROM [Contas a Receber] " & _
    "WHERE Filial = " & nFilial & " AND Tipo = 'O' AND [Valor Recebido] = 0 " & _
    "AND Vencimento >= #" & Format(Data_Atual, "mm/dd/yyyy") & "# AND " & _
    "Processado = False;"
  Set rsCartao = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsCartao
    cTotalReceber = IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
    .Close
  End With
  Set rsCartao = Nothing
End Sub

Private Sub CheckValueCheques(ByVal nFilial As Byte, ByRef cTotalReceber As Currency, ByRef cTotalDevolvido As Currency)
  Dim rsCheque As Recordset
  Dim sSql As String
  
  sSql = "SELECT Sum(Valor) AS SaldoFinal FROM [Contas a Receber] " & _
    "WHERE Filial = " & nFilial & " AND Tipo = 'C' AND (Vencimento <> [Data Emissão] " & _
    "OR [Data Emissão] is NULL) AND Vencimento >= #" & Format(Data_Atual, "mm/dd/yyyy") & "# AND " & _
    "Processado = False AND Devolvido = False;"
  Set rsCheque = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsCheque
    cTotalReceber = IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
    .Close
  End With
  
  sSql = "SELECT Sum(Valor) AS SaldoFinal FROM [Contas a Receber] " & _
    "WHERE Filial = " & nFilial & " AND Tipo = 'C' AND " & _
    "Devolvido = True;"
  Set rsCheque = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsCheque
    cTotalDevolvido = IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
    .Close
  End With
  Set rsCheque = Nothing
End Sub

Private Sub CheckValueContasPagar(ByVal nFilial As Byte, ByRef cTotalPagar As Currency, ByRef cTotalPagarAtrasado As Currency)
  Dim rsContasPagar As Recordset
  Dim sSql As String
  
  sSql = "SELECT Sum(Valor) AS SaldoFinal FROM [Contas a Pagar] " & _
    "WHERE Filial = " & nFilial & "AND [Valor Pago] = 0 " & _
    "AND Vencimento >= #" & Format(Data_Atual, "mm/dd/yyyy") & "#;"
  Set rsContasPagar = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsContasPagar
    cTotalPagar = IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
    .Close
  End With
  
  sSql = "SELECT Sum(Valor) AS SaldoFinal FROM [Contas a Pagar] " & _
    "WHERE Filial = " & nFilial & " AND [Valor Pago] = 0 " & _
    "AND Vencimento < #" & Format(Data_Atual, "mm/dd/yyyy") & "#;"
  Set rsContasPagar = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsContasPagar
    cTotalPagarAtrasado = IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
    .Close
  End With
  Set rsContasPagar = Nothing
End Sub

Private Sub CheckValueContasReceber(ByVal nFilial As Byte, ByRef cTotalReceber As Currency, ByRef cTotalReceberAtrasado As Currency)
  Dim rsContasReceber As Recordset
  Dim sSql As String
  
  sSql = "SELECT Sum(Valor) AS SaldoFinal FROM [Contas a Receber] " & _
    "WHERE Filial = " & nFilial & " AND Tipo = 'R' AND [Valor Recebido] = 0 " & _
    "AND Vencimento >= #" & Format(Data_Atual, "mm/dd/yyyy") & "#;"
  Set rsContasReceber = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsContasReceber
    cTotalReceber = IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
    .Close
  End With
  
  sSql = "SELECT Sum(Valor) AS SaldoFinal FROM [Contas a Receber] " & _
    "WHERE Filial = " & nFilial & " AND Tipo = 'R' AND [Valor Recebido] = 0 " & _
    "AND Vencimento < #" & Format(Data_Atual, "mm/dd/yyyy") & "#;"
  Set rsContasReceber = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsContasReceber
    cTotalReceberAtrasado = IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
    .Close
  End With
  Set rsContasReceber = Nothing
End Sub

Private Function cGetSaldoAllCC() As Currency
  Dim rsContaCorrente As Recordset
  Dim rsLancamentos As Recordset
  Dim cSaldoTotal As Currency
  Dim sSql As String
  
  sSql = "SELECT Conta, Código FROM [Contas Bancárias] ORDER BY Conta"
  Set rsContaCorrente = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsContaCorrente
    If .RecordCount > 0 Then
      Do Until .EOF
        sSql = "SELECT [Saldo Atual] AS SaldoFinal FROM [Lançamentos Bancários] " & _
          "WHERE Conta = " & !Código & " ORDER BY Data, Ordem"
        Set rsLancamentos = db.OpenRecordset(sSql, dbOpenSnapshot)
        With rsLancamentos
          If .RecordCount > 0 Then
            .MoveLast 'Último Saldo
            cSaldoTotal = cSaldoTotal + IIf(IsNull(!SaldoFinal), 0, !SaldoFinal)
          End If
          .Close
        End With
        .MoveNext
      Loop
    End If
    .Close
  End With
  cGetSaldoAllCC = cSaldoTotal
  Set rsContaCorrente = Nothing
  Set rsContaCorrente = Nothing
End Function

Private Function cSaveInfoCaixas(ByVal nFilial As Byte) As Currency
  Dim rsCaixa As Recordset
  Dim rsLancamentos As Recordset
  Dim sSql As String
  Dim cCaixaValue As Currency
  Dim cTotalCaixaValue As Currency
  
  sSql = "SELECT Caixa, Descrição FROM [Caixas em Uso]"
  Set rsCaixa = db.OpenRecordset(sSql, dbOpenSnapshot)
  
  With rsCaixa
    If .RecordCount > 0 Then
      Do Until .EOF
        sSql = "SELECT Sum(Dinheiro + Cheques + Vales) AS TotalCaixa " & _
          "FROM Caixa WHERE Filial = " & nFilial & "AND Data = #" & _
          Format(Data_Atual, "mm/dd/yyyy") & "# AND Caixa = " & !Caixa
        Set rsLancamentos = db.OpenRecordset(sSql, dbOpenSnapshot)
        cCaixaValue = IIf(IsNull(rsLancamentos!TotalCaixa), 0, rsLancamentos!TotalCaixa)
        cTotalCaixaValue = cTotalCaixaValue + cCaixaValue
        Call InsertItemIntoZZZGeral(rsCaixa!Caixa & "-" & rsCaixa!Descrição & _
          IIf(gbHasMovimentCaixa(!Caixa), "", " (Caixa não inicializado)"), _
          cCaixaValue, 0, 2, True)
        rsLancamentos.Close
        .MoveNext
      Loop
    End If
    .Close
  End With
  cSaveInfoCaixas = cTotalCaixaValue
  Set rsCaixa = Nothing
  Set rsLancamentos = Nothing
End Function

Private Sub B_Imprime_Click()
  Dim cAuxValue1 As Currency
  Dim cAuxValue2 As Currency
  Dim cSubTotal1 As Currency
  Dim cSubTotal2 As Currency
  Dim cSubTotal3 As Currency
  Dim cSubTotal4 As Currency
  Dim cSubTotal5 As Currency
  Dim nFilial As Byte
  
  Call StatusMsg("")

  'Verifica empresa
  If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
    DisplayMsg "Escolha a empresa."
    Combo.SetFocus
    Exit Sub
  End If
  
  nFilial = Val(Combo.Text)
  
  'Verifica o acesso do funcionário
  If Filial_Liberada <> 0 Then
    If nFilial <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If

  'Apaga o arquivo ZZZ
  Call StatusMsg("Preparando o arquivo temporário...")
  Call db.Execute("DELETE * FROM ZZZGeral")
  
  'Inicia transação
  ws.BeginTrans
  
  'Obtém o total em saldo das c/c e grava
  Call StatusMsg("Lendo Saldo em C/C...")
  cAuxValue1 = cGetSaldoAllCC
  Call InsertItemIntoZZZGeral("Saldo Contas Correntes", cAuxValue1, 0, 1, True)
  
  cSubTotal1 = cAuxValue1
  
  'Obtém os valores dos Caixas e grava
  Call StatusMsg("Lendo Saldo(s) no(s) Caixa(s)...")
  
  cAuxValue1 = cSaveInfoCaixas(nFilial)
  
  Call InsertItemIntoZZZGeral("Saldo do Caixa = Dinheiro + Cheque + Vale", 0, 0, 3)
  
  cSubTotal1 = cSubTotal1 + cAuxValue1
  
  'Obtém os valores de Contas a Receber e grava
  Call StatusMsg("Lendo Contas a Receber...")
  Call CheckValueContasReceber(nFilial, cAuxValue1, cAuxValue2)
  Call InsertItemIntoZZZGeral("Contas Futuras a Receber", cAuxValue1, 0, 4)
  Call InsertItemIntoZZZGeral("Contas Atrasadas a Receber", cAuxValue2, 0, 8)
  
  cSubTotal2 = cSubTotal2 + cAuxValue1
  cSubTotal3 = cSubTotal3 + cAuxValue2
  
  Call StatusMsg("Lendo Contas a Pagar...")
  'Obtém os valores de Contas a Pagar e grava
  Call CheckValueContasPagar(nFilial, cAuxValue1, cAuxValue2)
  Call InsertItemIntoZZZGeral("Contas Futuras a Pagar", 0, cAuxValue1, 11)
  Call InsertItemIntoZZZGeral("Contas Atrasadas a Pagar", 0, cAuxValue2, 14)
  
  cSubTotal4 = cSubTotal4 + cAuxValue1
  cSubTotal5 = cSubTotal5 + cAuxValue2
  
  Call StatusMsg("Lendo Cheques...")
  'Obtém os valores de Cheques e grava
  Call CheckValueCheques(nFilial, cAuxValue1, cAuxValue2)
  Call InsertItemIntoZZZGeral("Cheques a Receber", cAuxValue1, 0, 5)
  Call InsertItemIntoZZZGeral("Cheques Devolvidos", cAuxValue2, 0, 9)
  
  cSubTotal2 = cSubTotal2 + cAuxValue1
  cSubTotal3 = cSubTotal3 + cAuxValue2
  
  Call StatusMsg("Lendo Cartões de Crédito...")
  'Obtém os valores de Cheques e grava
  Call CheckValueCartao(nFilial, cAuxValue1)
  Call InsertItemIntoZZZGeral("Cartões de Crédito a Receber", cAuxValue1, 0, 6)
  
  cSubTotal2 = cSubTotal2 + cAuxValue1
  
  Call StatusMsg("Lendo Pedidos Pendentes...")
  'Obtém os valores de Cheques e grava
  Call CheckValuePedidos(nFilial, cAuxValue1)
  Call InsertItemIntoZZZGeral("Pedidos Pendentes junto a Fornecedores", 0, cAuxValue1, 12)
  
  cSubTotal4 = cSubTotal4 + cAuxValue1
  
  Call StatusMsg("Atualizando Valores...")
  
  Call InsertItemIntoZZZGeral("SUBTOTAL 1 - Crédito Disponível ", cSubTotal1, 0, 3, True)
  Call InsertItemIntoZZZGeral("SUBTOTAL 2 - Crédito a Receber ", cSubTotal2, 0, 7)
  Call InsertItemIntoZZZGeral("SUBTOTAL 3 - Crédito Pendente", cSubTotal3, 0, 10)
  Call InsertItemIntoZZZGeral("SUBTOTAL 4 - A Pagar", 0, cSubTotal4, 13)
  Call InsertItemIntoZZZGeral("SUBTOTAL 5 - A Pagar Atrasado", 0, cSubTotal5, 15)
  Call InsertItemIntoZZZGeral("", 0, 0, 16)
  Call InsertItemIntoZZZGeral("TOTAL 1 - Total de Créditos Bons", cSubTotal1 + cSubTotal2, 0, 17, True)
  Call InsertItemIntoZZZGeral("TOTAL 2 - Total Geral de Créditos", cSubTotal1 + cSubTotal2 + cSubTotal3, 0, 18, True)
  Call InsertItemIntoZZZGeral("TOTAL 3 - Total de Débitos", 0, cSubTotal4 + cSubTotal5, 19)
  Call InsertItemIntoZZZGeral("TOTAL 4 - Situação Geral (excluindo créditos pendentes)", _
    cSubTotal1 + cSubTotal2 + cSubTotal3 - cSubTotal4 - cSubTotal5, 0, 20, True)
  
  ws.CommitTrans
  
  With Rel1
    .DataFiles(0) = gsQuickDBFileName
    .Destination = IIf(B_Vídeo.Value, crptToWindow, crptToPrinter)
    .ReportFileName = gsReportPath & "GERAL.RPT"
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
    .Formulas(1) = "nome_filial = '" & Nome_Empresa.Caption & "'"
    Call StatusMsg("Aguarde, imprimindo...")
    Screen.MousePointer = vbHourglass
  
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel1
    
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel1)
  
    
    .Action = 1
  End With
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(1).Text
  Combo_LostFocus
End Sub

Private Sub Combo_KeyPress(KeyAscii As Integer)
  If Len(Combo.Text) >= 2 Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
      Exit Sub
    End If
  End If
End Sub

Private Sub Combo_LostFocus()
  Call StatusMsg("")
 
  Nome_Empresa.Caption = ""
  If IsNull(Combo.Text) Then Exit Sub
  If Combo.Text = "" Then Exit Sub
  If Not IsNumeric(Combo.Text) Then Exit Sub
  If Val(Combo.Text) < 0 Then Exit Sub
  If Val(Combo.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Set rsTempo = db.OpenRecordset("ZZZGeral")
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Combo.Text = gnCodFilial
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsTempo.Close
  rsParametros.Close
  Set rsTempo = Nothing
  Set rsParametros = Nothing
End Sub
