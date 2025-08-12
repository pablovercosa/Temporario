VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmVerificaPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Controle de entregas"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerificaPedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   15765
   Begin VB.CommandButton cmd_visualizaEntregasDaSequencia 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Visualiza as Entregas parciais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3900
      Width           =   3105
   End
   Begin VB.CommandButton cmd_abreSequencia 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Detalhar Sequência"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7470
      Width           =   3105
   End
   Begin VB.CommandButton cmdZerarSaldo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Zerar Saldo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3900
      Width           =   3105
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Relatório de Controle de Entregas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   3135
   End
   Begin VB.CommandButton cmdProcurar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   3135
   End
   Begin VB.CommandButton cmdFechar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14310
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   1395
   End
   Begin VB.CommandButton cmdGerarSaida 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3900
      Width           =   3105
   End
   Begin VB.TextBox txtSequencia 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1140
      TabIndex        =   0
      Top             =   90
      Width           =   1845
   End
   Begin SSDataWidgets_B.SSDBGrid grdPedidos 
      Height          =   3255
      Left            =   30
      TabIndex        =   2
      Top             =   570
      Width           =   15675
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   10
      BackColorOdd    =   12648447
      RowHeight       =   503
      ExtraHeight     =   132
      Columns.Count   =   10
      Columns(0).Width=   2143
      Columns(0).Caption=   "Sequência"
      Columns(0).Name =   "Seq"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   2778
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Produto"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   6456
      Columns(2).Caption=   "Nome"
      Columns(2).Name =   "Nome"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1826
      Columns(3).Caption=   "Tam"
      Columns(3).Name =   "Tamanho"
      Columns(3).Alignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1799
      Columns(4).Caption=   "Cor"
      Columns(4).Name =   "Cor"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   2858
      Columns(5).Caption=   "Q. Vendido"
      Columns(5).Name =   "QtdePedido"
      Columns(5).Alignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   2831
      Columns(6).Caption=   "Q. Lanç."
      Columns(6).Name =   "QtdeLancamento"
      Columns(6).Alignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   3175
      Columns(7).Caption=   "Saldo"
      Columns(7).Name =   "Saldo"
      Columns(7).Alignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   2778
      Columns(8).Caption=   "Digitar"
      Columns(8).Name =   "QtdeComCliente"
      Columns(8).Alignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "Linha"
      Columns(9).Name =   "Linha"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   27649
      _ExtentY        =   5741
      _StockProps     =   79
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid gridOperacoesDetalhe 
      Height          =   2985
      Left            =   30
      TabIndex        =   9
      Top             =   4410
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   5265
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   12648384
      BackColorSel    =   12648384
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16250871
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Sequência"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   135
      Width           =   855
   End
End
Attribute VB_Name = "frmVerificaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Public rsSaidas As Recordset
  Public rsSaidasProdutos  As Recordset
  Public rsProdutos As Recordset

  Public rsParametros      As Recordset
  Public rsOpSaidas        As Recordset

Private Sub cmd_abreSequencia_Click()
On Error GoTo Erro

  Dim sNumSeq As String

  If gridOperacoesDetalhe.RowSel > 0 Then
      sNumSeq = gridOperacoesDetalhe.TextMatrix(gridOperacoesDetalhe.RowSel, 2)
      
      Dim objSaidas As frmSaidas
      Set objSaidas = New frmSaidas
      
      objSaidas.txtSeq = gsHandleNull(sNumSeq)
      objSaidas.SearchRecord_peloNumSeq gsHandleNull(sNumSeq)
      objSaidas.Show

      Set objSaidas = Nothing
  Else
      MsgBox "selecione uma Sequência na grade", vbInformation, "Atenção"
      Exit Sub
  End If

  Exit Sub
Erro:
  MsgBox "Erro ao detalhar a sequência " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cmd_visualizaEntregasDaSequencia_Click()
On Error GoTo Erro

  Dim nRow As Integer
  Dim bm As Variant
  Dim sNumSeq As String
  Dim sSQL As String
  Dim rsOper As Recordset
  Dim sEfetAux As String
  Dim sDesfAux As String

  If grdPedidos.Row >= 0 Then
      bm = grdPedidos.AddItemBookmark(grdPedidos.Row)
      sNumSeq = gsHandleNull(grdPedidos.Columns("Seq").CellValue(bm))
      
      ' Buscar as sequencias das entregas parciais da SEQUENCIA SELECIONADA
      sSQL = ""
      sSQL = sSQL & " Select * from Saídas "
      sSQL = sSQL & " Where Filial = " & gnCodFilial & " and "
      sSQL = sSQL & " SequênciaPai = " & sNumSeq
      'sSQL = sSQL & " Observações like '*Entrega referente a movimentação " & sNumSeq & "*'"
      sSQL = sSQL & " Order by Data"
      
      gridOperacoesDetalhe.Rows = 1
      
      ' Listar na grade
      Set rsOper = db.OpenRecordset(sSQL, dbOpenDynaset)
      If Not (rsOper.EOF And rsOper.BOF) Then
          While Not rsOper.EOF
          
                sEfetAux = ""
                sDesfAux = ""
                If rsOper.Fields("Efetivada").Value = True Then
                    sEfetAux = "SIM"
                Else
                    sEfetAux = "NÃO"
                End If
                
                If rsOper.Fields("Movimentação Desfeita").Value = True Then
                    sDesfAux = "SIM"
                Else
                    sDesfAux = "NÃO"
                End If
          
                gridOperacoesDetalhe.AddItem vbTab & rsOper.Fields("Data").Value & vbTab & _
                      rsOper.Fields("Sequência").Value & vbTab & _
                      rsOper.Fields("SequênciaPai").Value & vbTab & _
                      sEfetAux & vbTab & _
                      sDesfAux & vbTab & _
                      FormataValorTexto(rsOper.Fields("total").Value, 2) & "" & vbTab & _
                      rsOper.Fields("Digitador").Value & vbTab & _
                      rsOper.Fields("Operador").Value & vbTab & _
                      rsOper.Fields("Observações").Value

              rsOper.MoveNext
          Wend
      End If
      rsOper.Close
      Set rsOper = Nothing
      
  Else
      MsgBox "selecione uma Sequência na grade", vbInformation, "Atenção"
      Exit Sub
  End If

  Exit Sub
Erro:
  MsgBox "Erro ao detalhar entregas parciais da venda original " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
  
  If lngCasasDecimais = 2 Then
      If Len(FormataValorTexto) = 7 Then  ' 9999.99     = 9.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 6)
      ElseIf Len(FormataValorTexto) = 8 Then ' 99999.99    = 99.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 6)
      ElseIf Len(FormataValorTexto) = 9 Then ' 999999.99   = 999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 6)
      ElseIf Len(FormataValorTexto) = 10 Then ' 9999999.99   = 9.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 3) + "." + Mid(FormataValorTexto, 5, 6)
      ElseIf Len(FormataValorTexto) = 11 Then ' 99999999.99   = 99.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 3) + "." + Mid(FormataValorTexto, 6, 6)
      ElseIf Len(FormataValorTexto) = 12 Then ' 999999999.99   = 999.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 3) + "." + Mid(FormataValorTexto, 7, 6)
      End If
  End If
End Function

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdGerarSaida_Click()
  Dim nSequencia        As Long
  Dim nOperacao         As Long
  Dim nCliente          As Long
  Dim nDigitador        As Long
  Dim nVendedor         As Long
  Dim sTabela           As String
  Dim nX                As Integer
  
  '22/01/2003 - mpdea
  'Verifica se há movimentação
  If grdPedidos.Rows = 0 Then
    DisplayMsg "Encontre uma movimentação antes."
    Exit Sub
  End If
  
  'Teste todos os produtos
    Dim bQtdeIncorreta As Boolean
    bQtdeIncorreta = False
    grdPedidos.MoveFirst
    For nX = 0 To grdPedidos.Rows - 1
      If Len(grdPedidos.Columns(8).Text) <= 0 Then
        grdPedidos.Columns(8).Text = 0
      Else
        If CDbl(grdPedidos.Columns(8).Text) > CDbl(grdPedidos.Columns(7).Text) Then
          bQtdeIncorreta = True
        End If
      End If
      If bQtdeIncorreta Then
        MsgBox "Qtde incorreta na linha " & nX + 1, vbCritical, "Sistema de entregas"
        Exit Sub
      End If
      grdPedidos.MoveNext
    Next nX
    
  '04/08/2007 - Anderson
  'Implementação de mensagem de confirmação de entregas
  'Solicitado pela Technomax
  If MsgBox("Tem certeza que deseja efetuar as entregas selecionadas?", vbYesNo + vbQuestion, "Controle de Entregas") = vbNo Then
    Exit Sub
  End If
  
  grdPedidos.MoveFirst
  Set rsParametros = db.OpenRecordset("SELECT Filial, [Última Movimentação] FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset)
  Set rsSaidas = db.OpenRecordset("SELECT * FROM Saídas WHERE Sequência = " & grdPedidos.Columns(0).Text & " AND Filial = " & gnCodFilial, dbOpenDynaset)
  '16/04/2007 - Anderson
  'Retirada para fazer parte do looping, resolvendo o problema de entrega quando existe o mesmo produto registrado duas vezes na mesma venda
  'Set rsSaidasProdutos = db.OpenRecordset("SELECT * FROM [Saídas - Produtos] WHERE Filial = " & gnCodFilial & " AND Sequência = " & grdPedidos.Columns(0).Text, dbOpenDynaset)
  
  If rsSaidas.EOF Then Exit Sub
  
  rsSaidas.MoveFirst
  nOperacao = rsSaidas.Fields("Operação")
  nCliente = rsSaidas.Fields("Cliente")
  nDigitador = rsSaidas.Fields("Digitador")
  nVendedor = rsSaidas.Fields("Operador")
  sTabela = rsSaidas.Fields("Tabela")
  
  Set rsOpSaidas = db.OpenRecordset("SELECT * FROM [Operações Saída] WHERE Código = " & nOperacao, dbOpenSnapshot)
  If (rsOpSaidas.EOF And rsOpSaidas.BOF) Then
    MsgBox "A operação selecionada para essa movimentação não suporta entregas !!", vbCritical, "ERRO "
    '16/04/2007 - Anderson
    'Não estava finalizando a rotina quando a operação não suporta entregas
    Exit Sub
  Else
    rsOpSaidas.MoveFirst
    nOperacao = rsOpSaidas.Fields("OpEntrega")
  End If
  
  With rsSaidas
    .AddNew
    !Filial = gnCodFilial
    !Data = Date
    !SequênciaPai = grdPedidos.Columns(0).Text
    nSequencia = gnGetNextSequencia(gnCodFilial)  'rsParametros![Última Movimentação] + 1
    !Sequência = nSequencia
    !Operação = nOperacao  'txtOperacao.Text
    !Cliente = nCliente
    !Digitador = nDigitador
    !Operador = nVendedor
    !Tabela = sTabela
    !Observações = "Entrega referente a movimentação " & !SequênciaPai
    
    .Update

    rsParametros.Edit
    rsParametros![Última Movimentação] = nSequencia
    rsParametros.Update
  End With
  
  grdPedidos.MoveFirst
  For nX = 0 To grdPedidos.Rows - 1
    If grdPedidos.Columns(8).Text > 0 Then
      Set rsSaidasProdutos = db.OpenRecordset("SELECT * FROM [Saídas - Produtos] WHERE Filial = " & gnCodFilial & " AND Sequência = " & grdPedidos.Columns(0).Text & " AND [Código Sem grade] = '" & grdPedidos.Columns(1).Text & "' AND Linha=" & grdPedidos.Columns(9).Text, dbOpenDynaset)
      With rsSaidasProdutos
  
        '16/04/2007 - Anderson
        'Retirada para resolver o problema de entrega quando existe o mesmo produto registrado duas vezes na mesma venda
        'rsSaidasProdutos.FindFirst (" Filial = " & gnCodFilial & _
        '                            " AND Sequência = " & grdPedidos.Columns(0).Text & _
        '                            " AND [Código Sem grade] = '" & grdPedidos.Columns(1).Text & "'")
        Do Until .EOF
        'If Not .NoMatch Then
          .Edit
          !QtdeEntregue = !QtdeEntregue + grdPedidos.Columns(8).Text
          Dim nPreco As Double
          nPreco = !Preço
          .Update
          .MoveNext
        'End If
        Loop
        
        .AddNew
        !Filial = gnCodFilial
        !Sequência = nSequencia
        !Linha = nX + 1
        !Código = grdPedidos.Columns(1).Text
        !Qtde = grdPedidos.Columns(8).Text
        !QtdeEntregue = grdPedidos.Columns(8).Text
        !Preço = nPreco
        .Update
        
        
        .Close
      End With
    End If
    grdPedidos.MoveNext
  Next nX

  MsgBox "Movimentação de Entrega (" & nSequencia & ") gerada com sucesso !!", vbInformation, "Movimentação de Entrega"
  
  'rsSaidasProdutos.Close
  rsParametros.Close
  rsOpSaidas.Close
  
  Set rsSaidasProdutos = Nothing
  Set rsParametros = Nothing
  Set rsOpSaidas = Nothing
  
  cmdGerarSaida.Enabled = False
End Sub

Private Sub cmdProcurar_Click()
  If Len(txtSequencia.Text) <= 0 Then Exit Sub
  grdPedidos.RemoveAll
  
  Dim sSQL     As String
  
  Dim nTamanho As Integer
  Dim nCor     As Integer
  Dim nQtde    As Double
  Dim nQtdeE   As Double
  Dim sNomProd As String
  
  sSQL = "SELECT * FROM [Saídas - Produtos] WHERE Filial = " & gnCodFilial
  
  If (Len(txtSequencia.Text) > 0) Or (IsNumeric(txtSequencia.Text)) Then sSQL = sSQL & " AND Sequência = " & Val(txtSequencia.Text)
  
  'Verifica se a operação utilizada na saida suporta entrega
    Dim rsOperacaoTEMP As Recordset
    Dim rsSaidasTEMP As Recordset
    
    gridOperacoesDetalhe.Rows = 1
                
    Set rsSaidasTEMP = db.OpenRecordset("SELECT * FROM Saídas WHERE Filial = " & gnCodFilial & " AND Sequência = " & txtSequencia.Text, dbOpenDynaset)
    With rsSaidasTEMP
      If Not (.BOF And .EOF) Then
        .MoveFirst
        Set rsOperacaoTEMP = db.OpenRecordset("SELECT * FROM [Operações Saída] WHERE Código = " & !Operação)
        With rsOperacaoTEMP
          If Not (.BOF And .EOF) Then
            .MoveFirst
            If !ControleEntregas <> -1 Then
                MsgBox "A operação utilizada para esta movimentação, não suporta entregas ... " & _
                     "verifique se a sequência digitada está correta !!", vbCritical, "Operação não suportada"
             
                Exit Sub
            End If
          End If
        End With
      Else
        MsgBox "Não existe movimentação com esse número de sequência !!", vbInformation, "Nenhum registro"
        Exit Sub
      End If
    End With
    
    rsOperacaoTEMP.Close
    rsSaidasTEMP.Close
  
    Set rsOperacaoTEMP = Nothing
    Set rsSaidasTEMP = Nothing
  '---------------------------------------------------------
  Set rsSaidas = db.OpenRecordset(sSQL, dbOpenSnapshot)
  
  With rsSaidas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        
        nTamanho = 0
        nCor = 0
        
        If (!Código <> ![Código Sem grade]) Then
          If IsNumeric(Left(Right(!Código, 6), 3)) Then
             nTamanho = Left(Right(!Código, 6), 3)
          End If
          If IsNumeric(Right(!Código, 3)) Then
             nCor = Right(!Código, 3)
          End If
        End If
        
        Set rsProdutos = db.OpenRecordset("SELECT Código, Nome FROM Produtos WHERE Código = '" & !Código & "'", dbOpenSnapshot)
        
        sNomProd = IIf(rsProdutos.EOF, "-", rsProdutos!Nome)
        '12/02/2007 - Anderson
        'Correção do problema de quantidades fracionadas entregues. Ex: o valor 2,4 estava apresentando como 2,40000009536743 (Var Double)
        'nQtde = IIf(Len(!Qtde) > 0, !Qtde, 0)
        'nQtdeE = IIf(Len(!QtdeEntregue) > 0, !QtdeEntregue, 0)
        nQtde = IIf(Len(!Qtde) > 0, Format(!Qtde, "#,##0.00"), 0)
        nQtdeE = IIf(Len(!QtdeEntregue) > 0, Format(!QtdeEntregue, "#,##0.00"), 0)
        
        If nQtde - nQtdeE > 0 Then
          grdPedidos.AddItem !Sequência & vbTab & _
                             !Código & vbTab & _
                             sNomProd & vbTab & _
                             nTamanho & vbTab & _
                             nCor & vbTab & _
                             nQtde & vbTab & _
                             nQtdeE & vbTab & _
                             nQtde - nQtdeE & vbTab & _
                             "" & vbTab & _
                             !Linha
        End If
        .MoveNext
      Loop
    End If
  End With
  
  
  rsProdutos.Close
  rsSaidas.Close
  
  Set rsSaidas = Nothing
  Set rsProdutos = Nothing
End Sub

Private Sub cmdZerarSaldo_Click()
  Dim nX                As Integer

  grdPedidos.MoveFirst
  For nX = 0 To grdPedidos.Rows - 1
    If grdPedidos.Columns(8).Text = "" Then
      grdPedidos.Columns(8).Text = grdPedidos.Columns(5).Text - grdPedidos.Columns(6).Text
    End If
    grdPedidos.MoveNext
  Next nX

End Sub

Private Sub Command1_Click()
  frmRelPedidosPendentes.Show
End Sub

Private Sub Form_Load()
  Dim sSeq As String
  
  Call CenterForm(Me)
  sSeq = frmSaidas.txtSeq.Text
  If Len(sSeq) > 0 Then
    If IsNumeric(sSeq) Then
      txtSequencia.Text = sSeq
      cmdProcurar_Click
    End If
  End If
  
  gridOperacoesDetalhe.ColWidth(0) = 0
  gridOperacoesDetalhe.ColWidth(1) = 1150
  gridOperacoesDetalhe.ColWidth(2) = 1200
  gridOperacoesDetalhe.ColWidth(3) = 1200
  gridOperacoesDetalhe.ColWidth(4) = 1300
  gridOperacoesDetalhe.ColWidth(5) = 1300
  gridOperacoesDetalhe.ColWidth(6) = 1300
  gridOperacoesDetalhe.ColWidth(7) = 1000
  gridOperacoesDetalhe.ColWidth(8) = 1000
  gridOperacoesDetalhe.ColWidth(9) = 8500
  
  gridOperacoesDetalhe.Row = 0
  gridOperacoesDetalhe.TextMatrix(0, 0) = ""
  gridOperacoesDetalhe.TextMatrix(0, 1) = "Data"
  gridOperacoesDetalhe.TextMatrix(0, 2) = "Sequência"
  gridOperacoesDetalhe.TextMatrix(0, 3) = "SequênciaPai"
  gridOperacoesDetalhe.TextMatrix(0, 4) = "Efetivada"
  gridOperacoesDetalhe.TextMatrix(0, 5) = "Desfeita"
  gridOperacoesDetalhe.TextMatrix(0, 6) = "Valor"
  gridOperacoesDetalhe.TextMatrix(0, 7) = "Digitador"
  gridOperacoesDetalhe.TextMatrix(0, 8) = "Operador"
  gridOperacoesDetalhe.TextMatrix(0, 9) = "Observações"

  
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  rsSaidas.Close
'  rsProdutos.Close
  

End Sub

Private Sub txtSequencia_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub
