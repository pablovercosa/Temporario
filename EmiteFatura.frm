VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmEmiteFatura 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   3420
   ClientTop       =   900
   ClientWidth     =   6090
   Icon            =   "EmiteFatura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   6090
   Begin VB.Frame Frame2 
      Caption         =   "Base de impressão"
      Height          =   855
      Left            =   1680
      TabIndex        =   25
      Top             =   2760
      Width           =   2295
      Begin VB.OptionButton optTotalParcela 
         Caption         =   "Separar em parcelas"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optTotalNota 
         Caption         =   "Total da Nota"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.TextBox txtPcaPagamento 
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox Texto_Recibo 
      Height          =   285
      Index           =   4
      Left            =   105
      MaxLength       =   60
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox Texto_Recibo 
      Height          =   285
      Index           =   3
      Left            =   105
      MaxLength       =   60
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox Texto_Recibo 
      Height          =   285
      Index           =   2
      Left            =   105
      MaxLength       =   60
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox Texto_Recibo 
      Height          =   285
      Index           =   1
      Left            =   105
      MaxLength       =   60
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox Texto_Recibo 
      Height          =   285
      Index           =   0
      Left            =   105
      MaxLength       =   60
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   105
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   5520
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label lblSequencia 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label_PcaPagamento 
      Caption         =   "Praça pagamento"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblDataEmissao 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCheckValue 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "True"
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label L_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label L_Fatura 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label L_Nota 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label L_Vencimento 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label L_Valor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label L_Encontrar 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label_Recibo 
      AutoSize        =   -1  'True
      Caption         =   "Descrição dos produtos para o recibo"
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Nome_Cliente 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente / Fornecedor :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Tipo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Transf2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Transf1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmEmiteFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsReceber As Recordset
Dim rsCliFor As Recordset
Dim rsParametros As Recordset

'14/01/2004 - Daniel
'Vars para impressão do Ticket
'Case: F. Linhares
Dim m_lngCodCli As Long
Dim m_lngSeq As Long
Dim m_dblValorRec As Long
Dim m_dblJuros As Double
Dim m_dblDesconto As Double
Dim m_dblValor As Double

Private Sub Imprime_Fatura()
  Dim Str1 As String
  Dim Str_Rel As String
  
  Dim Endereço As String
  Dim Cidade As String
  Dim Estado As String
  Dim CEP As String
  Dim Bairro As String
  Dim Comple As String
  
  Endereço = rsCliFor("Endereço") & ""
  Cidade = rsCliFor("Cidade") & ""
  Estado = rsCliFor("Estado") & ""
  CEP = rsCliFor("CEP") & ""
  Bairro = rsCliFor("Bairro") & ""
  Comple = rsCliFor("Complemento") & ""
  
  If Not IsNull(rsCliFor("Endereço Cob")) Then
    If Trim(rsCliFor("Endereço Cob")) <> "" Then
      Endereço = rsCliFor("Endereço Cob") & ""
      Cidade = rsCliFor("Cidade Cob") & ""
      Estado = rsCliFor("Estado Cob") & ""
      CEP = rsCliFor("CEP Cob") & ""
      Bairro = rsCliFor("Bairro Cob") & ""
      Comple = rsCliFor("Complemento Cob") & ""
    End If
  End If
  
  With Rel
    .DataFiles(0) = gsQuickDBFileName
    .Destination = IIf(B_Vídeo.Value, crptToWindow, crptToPrinter)
    .ReportFileName = gsReportPath & "FATURA.RPT"
    .Formulas(0) = "nome_empresa = '" & (rsParametros("Razão Social") & "") & "'"
    .Formulas(1) = "endereço1 = '" & (rsParametros("Endereço") & "") & "'"
    .Formulas(2) = "endereço2 = '" & (rsParametros("Cidade") & "") & "/" & (rsParametros("Estado") & "") & "'"
    .Formulas(3) = "cgc = '" & (rsParametros("CGC") & "") & "'"
    .Formulas(4) = "inscrição = '" & (rsParametros("Inscrição") & "") & "'"
    .Formulas(5) = "nota = '" & L_Nota.Caption & "'"
    .Formulas(6) = "valor = '" & Format(L_Valor.Caption, "###,###,###,##0.00") & "'"
    .Formulas(7) = "duplicata = '" & L_Fatura.Caption & "'"
    .Formulas(8) = "nome = '" & (rsCliFor("Nome") & "") & "'"
    .Formulas(9) = "endereço = '" & (Endereço & "") & "'"
    .Formulas(10) = "complemento = '" & (Comple & "") & "'"
    .Formulas(11) = "cidade = '" & (Cidade & "") & "'"
    .Formulas(12) = "estado = '" & (Estado & "") & "'"
    .Formulas(13) = "cep = '" & (CEP & "") & "'"
    .Formulas(14) = "cgc2 = '" & (rsCliFor("CGC") & "") & "'"
    .Formulas(15) = "inscrição2 = '" & (rsCliFor("Inscrição") & "") & "'"
    .Formulas(16) = "extenso = '" & Extenso(CDbl(L_Valor.Caption)) & "'"
    .Formulas(17) = "vencimento = '" & Format(L_Vencimento.Caption, "dd/mm/yyyy") & "'"
    .Formulas(18) = "texto_final = 'Reconheço(cemos) a exatidão desta DUPLICATA DE VENDA MERCANTIL " & _
      "na importância acima, que pagarei(emos) à " & (rsParametros("Razão Social") & "") & " ou à sua ordem " & _
      "na praça e vencimentos indicados.'"
    .Formulas(19) = "bairro = '" & (Bairro & "") & "'"
    .Formulas(20) = "DataEmissao = '" & Format(lblDataEmissao.Caption, "dd/mm/yyyy") & "'"
    .Formulas(21) = "PcaPagamento = '" & txtPcaPagamento.Text & "'"
    .Formulas(22) = "sequencia = '" & Format(lblSequencia.Caption, "########") & "'"
    
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel)
      
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel)
    
    .Action = 1
  End With
  
End Sub

Private Sub Imprime_Recibo()
  With Rel
    .DataFiles(0) = gsQuickDBFileName
    .Destination = IIf(B_Vídeo.Value, crptToWindow, crptToPrinter)
    .ReportFileName = gsReportPath & "RECIBO.RPT"
    .Formulas(0) = "nome_empresa = '" & (rsParametros("Razão Social") & "") & "'"
    .Formulas(1) = "endereço1 = '" & (rsParametros("Endereço") & "") & "'"
    .Formulas(2) = "endereço2 = '" & (rsParametros("Cidade") & "") & "/" & (rsParametros("Estado") & "") & "'"
    .Formulas(3) = "cgc = '" & (rsParametros("CGC") & "") & "'"
    .Formulas(4) = "inscrição = '" & (rsParametros("Inscrição") & "") & "'"
    .Formulas(5) = "nota = '" & L_Nota.Caption & "'"
    .Formulas(6) = "valor = '" & Format(CDbl(L_Valor.Caption), "###,###,###,##0.00") & "'"
    .Formulas(7) = "duplicata = '" & (L_Fatura.Caption) & "'"
    .Formulas(8) = "nome = """ & (rsCliFor("Nome") & "") & """"
    .Formulas(9) = "endereço = '" & (rsCliFor("Endereço") & "") & "'"
    .Formulas(10) = "complemento = '" & (rsCliFor("Complemento") & "") & "'"
    .Formulas(11) = "cidade = '" & (rsCliFor("Cidade") & "") & "'"
    .Formulas(12) = "estado = '" & (rsCliFor("Estado") & "") & "'"
    .Formulas(13) = "extenso = '" & Extenso(CDbl(L_Valor.Caption)) & "'"
    .Formulas(14) = "vencimento = '" & Format(L_Vencimento.Caption, "dd/mm/yyyy") & "'"
    .Formulas(15) = "texto_final1= '" & Texto_Recibo(0).Text & "'"
    .Formulas(16) = "texto_final2= '" & Texto_Recibo(1).Text & "'"
    .Formulas(17) = "texto_final3= '" & Texto_Recibo(2).Text & "'"
    .Formulas(18) = "texto_final4= '" & Texto_Recibo(3).Text & "'"
    .Formulas(19) = "texto_final5= '" & Texto_Recibo(4).Text & "'"
    .Formulas(20) = ""
      
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel)
    
    .Action = 1
  End With
End Sub

Private Sub B_Imprime_Click()
  On Error Resume Next
  
  Call StatusMsg("Aguarde...")
  If Tipo.Caption = "F" Then
    If optTotalNota.Value Then
      Call Imprime_Fatura
    ElseIf optTotalParcela.Value Then
      Call Imprime_FaturaParcelados
    End If
  Else
    Call Imprime_Recibo
  End If
  Call StatusMsg("")

End Sub

Private Sub Form_Activate()
  Dim bVisible As Boolean
  Dim nX As Byte
  Dim nValor As Currency
  Dim nAcrescimo As Currency
  Dim nDesconto As Currency
  
  nValor = nReciboVALOR           'frmManContasReceber.Valor.Text
  nAcrescimo = nReciboACRESCIMO   'frmManContasReceber.Acréscimo.Text
  nDesconto = nReciboDESCONTO     'frmManContasReceber.Desconto.Text
  
  
  If IsNull(Transf1.Caption) Then Exit Sub
  
  If L_Encontrar.Caption = "SIM" Then
    rsReceber.Index = "Vencimento"
    rsReceber.Seek "=", "R", gnCodFilial, Transf1.Caption, Transf2.Caption
    If rsReceber.NoMatch Then
      Exit Sub
    Else
      L_Nota.Caption = rsReceber("Nota") & ""
      L_Fatura.Caption = rsReceber("Fatura") & ""
      If CBool(lblCheckValue.Caption) Then
'        L_Valor.Caption = rsReceber("Valor") + rsReceber("Acréscimo") - rsReceber("Desconto")
        'L_Valor.Caption = frmManContasReceber.Valor.Text + frmManContasReceber.Acréscimo.Text - frmManContasReceber.Desconto.Text
         L_Valor.Caption = nValor + nAcrescimo - nDesconto
         
      End If
      L_Vencimento.Caption = rsReceber("Vencimento")
      lblDataEmissao.Caption = rsReceber("Data Emissão")
      L_Cliente.Caption = rsReceber("Cliente")
      lblSequencia.Caption = rsReceber("Sequência")
      
      '14/01/2004 - Daniel
      'Populando vars para impressão do ticket
      m_lngSeq = Format(rsReceber("Sequência"), FORMAT_VALUE)
      m_dblValorRec = Format(rsReceber("Valor Recebido"), FORMAT_VALUE)
      m_dblJuros = Format(rsReceber("Acréscimo"), FORMAT_VALUE)
      m_dblDesconto = Format(rsReceber("Desconto"), FORMAT_VALUE)
      m_dblValor = Format(rsReceber("Valor"), FORMAT_VALUE)

    End If
  End If
  
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", L_Cliente.Caption  'rsReceber("Cliente")
  If Not rsCliFor.NoMatch Then
     Nome_Cliente.Caption = str(rsCliFor("Código")) + " - " + rsCliFor("Nome")
  End If
  
  '14/01/2004 - Daniel
  'Populando var para impressão do ticket
  m_lngCodCli = Format(rsCliFor("Código"), FORMAT_VALUE)
  '---------------------------------------
    
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  
  If Tipo.Caption = "F" Then
    bVisible = False
  ElseIf Tipo.Caption = "R" Then
    bVisible = True
  End If
  Label_Recibo.Visible = bVisible
  For nX = 0 To 4
    Texto_Recibo(nX).Visible = bVisible
  Next nX
  
  Label_PcaPagamento.Visible = Not bVisible
  txtPcaPagamento.Visible = Not bVisible
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Set rsReceber = db.OpenRecordset("Contas a Receber", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsReceber.Close
  rsCliFor.Close
  rsParametros.Close
  Set rsReceber = Nothing
  Set rsCliFor = Nothing
  Set rsParametros = Nothing
End Sub

Private Sub Imprime_FaturaParcelados()
  Dim Str1     As String
  Dim Str_Rel  As String
  
  Dim Endereço As String
  Dim Cidade   As String
  Dim Estado   As String
  Dim CEP      As String
  Dim Bairro   As String
  Dim Comple   As String
  
  Dim intParcelas As Integer
  Dim bolFlag     As Boolean
  
  '17/06/2005 - Daniel
  '
  'Alterações para a impressão de Duplicatas (Faturas) a partir da beta 6.52.0.55
  '
  'A partir da emissão em VR, o Quick não imprime às Duplicatas (Faturas) 'depois' que o
  'usuário finaliza o aplicativo. Por esta razão estamos otimizando rotinas que proporcionem
  'ao usuário imprimir às Duplicatas (Faturas) a partir da tela de Saídas buscando às parcelas
  'através da Sequência gerada na gravação da venda.
  '
  'Nota: O código para VR permanece intacto
  Dim intAuxi As Integer
  Dim strSQL  As String
  Dim rstCR   As Recordset
  Dim intQtdeParce As Integer
  Dim dblValor(1 To 500) As Double
  Dim datVencimento(1 To 500) As Date
  Dim strFatura(1 To 500) As String
  
  
  '---[Fim das definições de variáveis]---
  
  If IsNumeric(lblSequencia.Caption) Then  'Chamada vinda a partir de Saídas
  
    intAuxi = 0
  
    strSQL = ""
    strSQL = "SELECT * FROM [Contas a Receber] "
    strSQL = strSQL & " WHERE Filial = " & gnCodFilial
    strSQL = strSQL & " AND Sequência = " & CLng(lblSequencia.Caption)
    strSQL = strSQL & " AND Cliente = " & rsCliFor("Código").Value
    strSQL = strSQL & " AND [Valor Recebido] = 0 "
    strSQL = strSQL & " ORDER BY Contador "
  
    Set rstCR = db.OpenRecordset(strSQL, dbOpenDynaset)
  
    If rstCR.RecordCount = 0 Then
      MsgBox "Não há parcelas para serem exibidas.", vbExclamation, "Quick Store"
      Exit Sub
    End If
    
    rstCR.MoveFirst
    rstCR.MoveLast
    rstCR.MoveFirst
    
    intQtdeParce = rstCR.RecordCount
  
    With rstCR
      If Not (.BOF And .EOF) Then
        '.MoveFirst
        
        Do Until .EOF
          intAuxi = intAuxi + 1
          
          '''strFatura(intAuxi) = .Fields("Fatura").Value & ""
          strFatura(intAuxi) = .Fields("Parcela").Value & ""
          dblValor(intAuxi) = Format(.Fields("Valor").Value, FORMAT_VALUE)
          datVencimento(intAuxi) = Format(.Fields("Vencimento").Value, "DD/MM/YYYY")
        
         .MoveNext
        Loop
        
      End If
      .Close
    End With
  
    Set rstCR = Nothing
    
    'Montando às duplicatas (faturas)
    
    'For intAuxi = 1 To intQtdeParce
          
        Endereço = rsCliFor("Endereço") & ""
        Cidade = rsCliFor("Cidade") & ""
        Estado = rsCliFor("Estado") & ""
        CEP = rsCliFor("CEP") & ""
        Bairro = rsCliFor("Bairro") & ""
        Comple = rsCliFor("Complemento") & ""
        
        If Not IsNull(rsCliFor("Endereço Cob")) Then
          If Trim(rsCliFor("Endereço Cob")) <> "" Then
            Endereço = rsCliFor("Endereço Cob") & ""
            Cidade = rsCliFor("Cidade Cob") & ""
            Estado = rsCliFor("Estado Cob") & ""
            CEP = rsCliFor("CEP Cob") & ""
            Bairro = rsCliFor("Bairro Cob") & ""
            Comple = rsCliFor("Complemento Cob") & ""
          End If
        End If
        
        With Rel
          .Reset
          .DataFiles(0) = gsQuickDBFileName
          .Destination = IIf(B_Vídeo.Value, crptToWindow, crptToPrinter)
          .ReportFileName = gsReportPath & "FATURA.RPT"
          .Formulas(0) = "nome_empresa = '" & (rsParametros("Razão Social") & "") & "'"
          .Formulas(1) = "endereço1 = '" & (rsParametros("Endereço") & "") & "'"
          .Formulas(2) = "endereço2 = '" & (rsParametros("Cidade") & "") & "/" & (rsParametros("Estado") & "") & "'"
          .Formulas(3) = "cgc = '" & (rsParametros("CGC") & "") & "'"
          .Formulas(4) = "inscrição = '" & (rsParametros("Inscrição") & "") & "'"
          .Formulas(5) = "nota = '" & L_Nota.Caption & "'"
          '.Formulas(6) = "valor = '" & Format(dblValor(intAuxi), "###,###,###,##0.00") & "'"
          '.Formulas(7) = "duplicata = '" & strFatura(intAuxi) & "'"
          .Formulas(8) = "nome = '" & (rsCliFor("Nome") & "") & "'"
          .Formulas(9) = "endereço = '" & (Endereço & "") & "'"
          .Formulas(10) = "complemento = '" & (Comple & "") & "'"
          .Formulas(11) = "cidade = '" & (Cidade & "") & "'"
          .Formulas(12) = "estado = '" & (Estado & "") & "'"
          .Formulas(13) = "cep = '" & (CEP & "") & "'"
          .Formulas(14) = "cgc2 = '" & (rsCliFor("CGC") & "") & "'"
          .Formulas(15) = "inscrição2 = '" & (rsCliFor("Inscrição") & "") & "'"
          '.Formulas(16) = "extenso = '" & Extenso(CDbl(dblValor(intAuxi))) & "'"
          '.Formulas(17) = "vencimento = '" & datVencimento(intAuxi) & "'"
          '.Formulas(18) = "texto_final = 'Reconheço(cemos) a exatidão desta DUPLICATA DE VENDA MERCANTIL " & _
          '  "na importância acima, que pagarei(emos) à " & (rsParametros("Razão Social") & "") & " ou à sua ordem " & _
          '  "na praça e vencimentos indicados.'"
          .Formulas(19) = "bairro = '" & (Bairro & "") & "'"
          .Formulas(20) = "DataEmissao = '" & Format(lblDataEmissao.Caption, "DD/MM/YYYY") & "'"
          .Formulas(21) = "PcaPagamento = '" & txtPcaPagamento.Text & "'"
          .Formulas(22) = "sequencia = '" & Format(lblSequencia.Caption, "########") & "'"
          .ParameterFields(0) = "pFilial;" & gnCodFilial & ";true"
          .ParameterFields(1) = "pSequencia;" & CLng(lblSequencia.Caption) & ";true"
          .WindowState = crptMaximized
          
          '25/07/2003 - mpdea
          'Seta a impressora para relatório
          Call SetPrinterName("REL", Rel)
  
          .Action = 1
        End With
      
    'Next
  
  
  Else 'Chamada vinda a partir de VR
  
      '20/06/2005 - Daniel
      'Tratamento para o contador de parcelas com a finalidade de exibir duplicata NF/Parcela
      intQtdeParce = 0
  
      For intParcelas = 0 To UBound(pfParcelasFatura)
        If IsDate(pfParcelasFatura(intParcelas).pfDataVencimento) Then
          '-----------------------------------------------------------
            Endereço = rsCliFor("Endereço") & ""
            Cidade = rsCliFor("Cidade") & ""
            Estado = rsCliFor("Estado") & ""
            CEP = rsCliFor("CEP") & ""
            Bairro = rsCliFor("Bairro") & ""
            Comple = rsCliFor("Complemento") & ""
            
            If Not IsNull(rsCliFor("Endereço Cob")) Then
              If Trim(rsCliFor("Endereço Cob")) <> "" Then
                Endereço = rsCliFor("Endereço Cob") & ""
                Cidade = rsCliFor("Cidade Cob") & ""
                Estado = rsCliFor("Estado Cob") & ""
                CEP = rsCliFor("CEP Cob") & ""
                Bairro = rsCliFor("Bairro Cob") & ""
                Comple = rsCliFor("Complemento Cob") & ""
              End If
            End If
            
            '20/06/2005 - Daniel
            'Tratamento para o contador de parcelas com a finalidade de exibir duplicata NF/Parcela
            intQtdeParce = intQtdeParce + 1
            
            With Rel
              .DataFiles(0) = gsQuickDBFileName
              .Destination = IIf(B_Vídeo.Value, crptToWindow, crptToPrinter)
              .ReportFileName = gsReportPath & "FATURA.RPT"
              .Formulas(0) = "nome_empresa = '" & (rsParametros("Razão Social") & "") & "'"
              .Formulas(1) = "endereço1 = '" & (rsParametros("Endereço") & "") & "'"
              .Formulas(2) = "endereço2 = '" & (rsParametros("Cidade") & "") & "/" & (rsParametros("Estado") & "") & "'"
              .Formulas(3) = "cgc = '" & (rsParametros("CGC") & "") & "'"
              .Formulas(4) = "inscrição = '" & (rsParametros("Inscrição") & "") & "'"
              .Formulas(5) = "nota = '" & L_Nota.Caption & "'"
              .Formulas(6) = "valor = '" & Format(CDbl(pfParcelasFatura(intParcelas).pfValor), "###,###,###,##0.00") & "'"
              '20/06/2005 - Daniel
              'Alterado para exibir o seguinte valor: NF/Parcela
              '.Formulas(7) = "duplicata = '" & L_Fatura.Caption & "'"
              If Len(L_Nota.Caption) > 0 Then
                .Formulas(7) = "duplicata = '" & L_Nota.Caption & "/" & intQtdeParce & "'"
              Else 'Antes da 6.52.0.55 exibia:
                .Formulas(7) = "duplicata = '" & L_Fatura.Caption & "'"
              End If
              .Formulas(8) = "nome = '" & (rsCliFor("Nome") & "") & "'"
              .Formulas(9) = "endereço = '" & (Endereço & "") & "'"
              .Formulas(10) = "complemento = '" & (Comple & "") & "'"
              .Formulas(11) = "cidade = '" & (Cidade & "") & "'"
              .Formulas(12) = "estado = '" & (Estado & "") & "'"
              .Formulas(13) = "cep = '" & (CEP & "") & "'"
              .Formulas(14) = "cgc2 = '" & (rsCliFor("CGC") & "") & "'"
              .Formulas(15) = "inscrição2 = '" & (rsCliFor("Inscrição") & "") & "'"
              .Formulas(16) = "extenso = '" & Extenso(CDbl(pfParcelasFatura(intParcelas).pfValor)) & "'"
              .Formulas(17) = "vencimento = '" & Format(pfParcelasFatura(intParcelas).pfDataVencimento, "dd/mm/yyyy") & "'"
              .Formulas(18) = "texto_final = 'Reconheço(cemos) a exatidão desta DUPLICATA DE VENDA MERCANTIL " & _
                "na importância acima, que pagarei(emos) à " & (rsParametros("Razão Social") & "") & " ou à sua ordem " & _
                "na praça e vencimentos indicados.'"
              .Formulas(19) = "bairro = '" & (Bairro & "") & "'"
              .Formulas(20) = "DataEmissao = '" & Format(lblDataEmissao.Caption, "dd/mm/yyyy") & "'"
              .Formulas(21) = "PcaPagamento = '" & txtPcaPagamento.Text & "'"
              .Formulas(22) = "sequencia = '" & Format(lblSequencia.Caption, "########") & "'"
      
              '25/07/2003 - mpdea
              'Seta a impressora para relatório
              Call SetPrinterName("REL", Rel)
      
              .Action = 1
            End With
          '---------------------------------------------------------------
        End If
      Next

  End If

End Sub
