VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMovCaixa_PosicaoCrediario 
   Caption         =   " Movimentação Diária - Posição Crediário"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovCaixa_PosicaoCrediario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   16950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_imprimirOperDet 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   2115
   End
   Begin VB.CommandButton cmd_abreSequencia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Detalhar Sequência"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14790
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   16860
      Begin VB.Label lbl_data 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   1
         Top             =   180
         Width           =   14145
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridLancamentos 
      Height          =   6990
      Left            =   60
      TabIndex        =   2
      Top             =   630
      Width           =   16860
      _ExtentX        =   29739
      _ExtentY        =   12330
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454143
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16250871
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_totalRegistros 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Total de registros: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   7680
      Width           =   2385
   End
End
Attribute VB_Name = "frmMovCaixa_PosicaoCrediario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DataPosicao As String
'Public CaixaPosicao As String
'Public NomeCaixaPosicao As String

Private Sub cmd_abreSequencia_Click()
On Error GoTo Erro

  Dim sTipoOperacao As String

  If gridLancamentos.RowSel > 0 Then
      Dim objSaidas As frmSaidas
      Set objSaidas = New frmSaidas
      
      objSaidas.txtSeq = gridLancamentos.TextMatrix(gridLancamentos.RowSel, 12)
      objSaidas.SearchRecord_peloNumSeq (gridLancamentos.TextMatrix(gridLancamentos.RowSel, 12))
      objSaidas.Show
      
      Set objSaidas = Nothing
  Else
      MsgBox "selecione uma Sequência na grade", vbInformation, "Atenção"
      Exit Sub
  End If
    
  Exit Sub
  
Erro:
  MsgBox "Erro no detalhamento da Sequência" & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_imprimirOperDet_Click()
  On Error GoTo Erro
  
  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome As String
  Dim strNomeLPT As String
  Dim strPortaLPT As String
  Dim intX As Integer
  Dim i As Integer
  
  strNome = "REL"
  strNomeLPT = "NOME IMPRESSORA REL"
  strPortaLPT = "PORTA IMPRESSORA REL"

  strImpressora = GetSetting("QuickStore", "ConfigLPT", strNomeLPT, "")
  strPorta = GetSetting("QuickStore", "ConfigLPT", strPortaLPT, "")
      
  If Len(Trim(strImpressora)) > 0 And Len(Trim(strPorta)) > 0 Then
      For Each objPrinter In Printers
        If objPrinter.DeviceName = strImpressora And objPrinter.Port = strPorta Then
            Set Printer = objPrinter
            Exit For
        End If
      Next objPrinter
  End If

  Dim nRow As Long
  Dim sLinha As String
  
  Printer.Font = "LUCIDA CONSOLE"
  
  Printer.Print ""
  sLinha = "                                                                   Quick Store 10 - Soluções Comerciais inteligentes"
  
  Printer.Print ""

  sLinha = "                           Movimentação Diária - Posição Crediário do Período: " & DataPosicao
  Printer.Print sLinha

  Printer.Print ""

  sLinha = "   Filial : " & gnCodFilial
  Printer.Print sLinha
  sLinha = "   Caixas : TODOS OS CAIXAS"
  Printer.Print sLinha

  Printer.Print ""

  sLinha = "   Código   Nome                                                          Crediário  Descrição"
  Printer.Print sLinha
  sLinha = "   Vencimento      R$ Valor      R$ Recebido          Parcela      Nota          Cartão           Fatura    Sequência   "
  Printer.Print sLinha
  sLinha = "   ID             Tem Pendência?   "
  Printer.Print sLinha

  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""

  Dim sCodigo         As String
  Dim sNome           As String
  Dim sCrediario      As String
  Dim sDescricao      As String
  Dim sVencimento     As String
  Dim sValor          As String
  Dim sRecebido       As String
  Dim sParcela        As String
  Dim sNota           As String
  Dim sCartao         As String
  Dim sFatura         As String
  Dim sSequencia      As String
  Dim sContador       As String

  With gridLancamentos
      For nRow = 1 To .Rows - 1
          ' ************************** ATENÇÃO ***********************************
          ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
          ' De preferência com o mesmo nome da impressora !!!

          sCodigo = gridLancamentos.TextMatrix(nRow, 1)
          If Len(sCodigo) < 10 Then
            For i = Len(sCodigo) To 9
                sCodigo = sCodigo & " "
            Next
          End If

          sNome = gridLancamentos.TextMatrix(nRow, 2)
          If Len(sNome) < 50 Then
              For i = Len(sNome) To 49
                sNome = sNome & " "
              Next
          Else
              sNome = Mid(sNome, 1, 50)
          End If

          sCrediario = gridLancamentos.TextMatrix(nRow, 3)
          If Len(sCrediario) < 16 Then
            For i = Len(sCrediario) To 15
                sCrediario = " " & sCrediario
            Next
          End If

          sDescricao = gridLancamentos.TextMatrix(nRow, 4)

          sLinha = sCodigo
          sLinha = sLinha & "  " & sNome
          sLinha = sLinha & "  " & sCrediario
          sLinha = sLinha & "  " & sDescricao
          Printer.Print "   " & sLinha


          sVencimento = gridLancamentos.TextMatrix(nRow, 5)

          sValor = gridLancamentos.TextMatrix(nRow, 6)
          If Len(sValor) < 13 Then
            For i = Len(sValor) To 12
                sValor = " " & sValor
            Next
          End If
          
          sRecebido = gridLancamentos.TextMatrix(nRow, 7)
          If Len(sRecebido) < 13 Then
            For i = Len(sRecebido) To 12
                sRecebido = " " & sRecebido
            Next
          End If
          
          sParcela = gridLancamentos.TextMatrix(nRow, 8)
          If Len(sParcela) < 13 Then
            For i = Len(sParcela) To 12
                sParcela = " " & sParcela
            Next
          End If
          
          sNota = gridLancamentos.TextMatrix(nRow, 9)
          If Len(sNota) < 13 Then
            For i = Len(sNota) To 12
                sNota = " " & sNota
            Next
          End If
          
          sCartao = gridLancamentos.TextMatrix(nRow, 10)
          If Len(sCartao) < 13 Then
            For i = Len(sCartao) To 12
                sCartao = " " & sCartao
            Next
          End If
          
          sFatura = gridLancamentos.TextMatrix(nRow, 11)
          If Len(sFatura) < 13 Then
            For i = Len(sFatura) To 12
                sFatura = " " & sFatura
            Next
          End If
          
          sSequencia = gridLancamentos.TextMatrix(nRow, 12)
          If Len(sSequencia) < 10 Then
            For i = Len(sSequencia) To 9
                sSequencia = " " & sSequencia
            Next
          End If
  
          sLinha = sVencimento
          sLinha = sLinha & "  " & sValor
          sLinha = sLinha & "  " & sRecebido
          sLinha = sLinha & "  " & sParcela
          sLinha = sLinha & "  " & sNota
          sLinha = sLinha & "  " & sCartao
          sLinha = sLinha & "  " & sFatura
          sLinha = sLinha & "  " & sSequencia
              
          Printer.Print "   " & sLinha
          
          sContador = gridLancamentos.TextMatrix(nRow, 13)
          If Len(sContador) < 10 Then
            For i = Len(sContador) To 9
                sContador = " " & sContador
            Next
          End If
          
          sLinha = sContador
          sLinha = sLinha & "  " & gridLancamentos.TextMatrix(nRow, 14)
          
          If gridLancamentos.TextMatrix(nRow, 14) = "TEM" Then
              sLinha = sLinha & "     ****************************"
          End If
          Printer.Print "   " & sLinha
          
          
          Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      Next nRow
  End With
      
  Printer.Print ""
    
  Printer.Print "   " & lbl_totalRegistros.Caption
  Printer.Print "   -----------------------------------------------------------------------------------------------------------------"

  Printer.EndDoc

  Exit Sub
Erro:
    MsgBox "Erro na impressão da grade " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Form_Load()
On Error GoTo Erro
  Dim sSQL As String
  Dim rsCaixaPosicao As Recordset
  Dim lngContadorRegGrid As Long
  Dim sFatura As String
  Dim sParcela As String
  Dim sCartao As String
  Dim sNota As String
  Dim sDescricao As String
  Dim sSituacaoCrediario As String
  Dim iCol As Integer
  Dim sCodigoAnterior As String
  Dim sCorAnterior As String
  Dim sPendencia As String
  
  lbl_data.Caption = "Data Posição: " & DataPosicao & "         Caixa: TODOS OS CAIXAS" '& "         Caixa : " & CaixaPosicao & " - " & NomeCaixaPosicao
  
  gridLancamentos.ColWidth(0) = 0
  gridLancamentos.ColWidth(1) = 900
  gridLancamentos.ColWidth(2) = 4100
  gridLancamentos.ColWidth(3) = 1000
  gridLancamentos.ColWidth(4) = 2600
  gridLancamentos.ColWidth(5) = 1100
  gridLancamentos.ColWidth(6) = 1200
  gridLancamentos.ColWidth(7) = 1200
  gridLancamentos.ColWidth(8) = 700
  gridLancamentos.ColWidth(9) = 700
  gridLancamentos.ColWidth(10) = 1200
  gridLancamentos.ColWidth(11) = 900
  gridLancamentos.ColWidth(12) = 900
  gridLancamentos.ColWidth(13) = 900
  gridLancamentos.ColWidth(14) = 900

  gridLancamentos.Row = 0
  gridLancamentos.TextMatrix(0, 1) = "Código"
  gridLancamentos.TextMatrix(0, 2) = "Nome"
  gridLancamentos.TextMatrix(0, 3) = "Crediário"
  gridLancamentos.TextMatrix(0, 4) = "Descrição"
  gridLancamentos.TextMatrix(0, 5) = "DtVencimento"
  gridLancamentos.TextMatrix(0, 6) = "R$ Valor"
  gridLancamentos.TextMatrix(0, 7) = "R$ Recebido"
  gridLancamentos.TextMatrix(0, 8) = "Parcela"
  gridLancamentos.TextMatrix(0, 9) = "Nota"
  gridLancamentos.TextMatrix(0, 10) = "Cartão"
  gridLancamentos.TextMatrix(0, 11) = "Fatura"
  gridLancamentos.TextMatrix(0, 12) = "Sequência"
  gridLancamentos.TextMatrix(0, 13) = "ID"
  gridLancamentos.TextMatrix(0, 14) = "Pendência"


  sSQL = "Select R.Cliente, C.Nome, R.Valor, R.[Valor Recebido], R.Fatura, R.Parcela, "
  sSQL = sSQL & " C.Faturado, R.Descrição, R.Nota, R.Cartão, R.Sequência, R.Vencimento, R.Contador, R.Pendencia "
  sSQL = sSQL & " From [Contas a Receber] R, Cli_For C "
  sSQL = sSQL & " Where R.[Data Recebimento] >= CDATE('" & DataPosicao & " 00:00:00') and "
  sSQL = sSQL & " R.[Data Recebimento] <= CDATE('" & DataPosicao & " 23:59:59') and "
  sSQL = sSQL & " R.Filial = " & gnCodFilial & " And "
  sSQL = sSQL & " R.Cliente = C.Código "
  sSQL = sSQL & " Order by C.Nome, R.Parcela "
  
  Screen.MousePointer = vbHourglass
  
  Set rsCaixaPosicao = db.OpenRecordset(sSQL, dbOpenDynaset, dbReadOnly)
  
  lngContadorRegGrid = 0
  
  If Not (rsCaixaPosicao.EOF And rsCaixaPosicao.BOF) Then
    rsCaixaPosicao.MoveFirst
  End If
  
  sCorAnterior = vbWhite
  
  While Not rsCaixaPosicao.EOF
  
      sFatura = ""
      sParcela = ""
      sCartao = ""
      sNota = ""
      sDescricao = ""
      
      If Not IsNull(rsCaixaPosicao.Fields("Fatura").Value) Then
          sFatura = rsCaixaPosicao.Fields("Fatura").Value
      End If
  
      If Not IsNull(rsCaixaPosicao.Fields("Parcela").Value) Then
          sParcela = rsCaixaPosicao.Fields("Parcela").Value
      End If
  
      If Not IsNull(rsCaixaPosicao.Fields("Cartão").Value) Then
          sCartao = rsCaixaPosicao.Fields("Cartão").Value
      End If
  
      If Not IsNull(rsCaixaPosicao.Fields("Descrição").Value) Then
          sDescricao = rsCaixaPosicao.Fields("Descrição").Value
      End If
  
      If Not IsNull(rsCaixaPosicao.Fields("Nota").Value) Then
          sNota = rsCaixaPosicao.Fields("Nota").Value
      End If
  
  
      sSituacaoCrediario = ""
      If rsCaixaPosicao.Fields(6).Value = True Then
          sSituacaoCrediario = "Habilitado"
      Else
          sSituacaoCrediario = "Suspenso"
      End If
      
      If rsCaixaPosicao.Fields("Pendencia").Value = -1 Then
        sPendencia = "TEM"
      Else
        sPendencia = ""
      End If
      
      gridLancamentos.AddItem vbTab & rsCaixaPosicao.Fields("Cliente").Value & vbTab & _
                          rsCaixaPosicao.Fields("Nome").Value & vbTab & _
                          sSituacaoCrediario & vbTab & _
                          sDescricao & vbTab & _
                          rsCaixaPosicao.Fields("Vencimento").Value & vbTab & _
                          FormataValorTexto(rsCaixaPosicao.Fields("Valor").Value, 2) & vbTab & _
                          FormataValorTexto(rsCaixaPosicao.Fields(3).Value, 2) & vbTab & _
                          sParcela & vbTab & _
                          sNota & vbTab & _
                          sCartao & vbTab & _
                          sFatura & vbTab & _
                          rsCaixaPosicao.Fields("Sequência").Value & vbTab & _
                          rsCaixaPosicao.Fields("Contador").Value & vbTab & _
                          sPendencia
                          
      lngContadorRegGrid = lngContadorRegGrid + 1

      ' tratar cores da grid
      If sCodigoAnterior = rsCaixaPosicao.Fields("Cliente").Value Then
          For iCol = 0 To gridLancamentos.Cols - 1
              gridLancamentos.Col = iCol
              gridLancamentos.Row = lngContadorRegGrid
              'alterna a cor das linhas do grid
              gridLancamentos.CellBackColor = sCorAnterior
          Next
      Else
          If sCorAnterior = vbWhite Then
              sCorAnterior = vbButtonFace
          Else
              sCorAnterior = vbWhite
          End If
      
          For iCol = 0 To gridLancamentos.Cols - 1
              gridLancamentos.Col = iCol
              gridLancamentos.Row = lngContadorRegGrid
              'alterna a cor das linhas do grid
              gridLancamentos.CellBackColor = sCorAnterior
          Next
      End If
      sCodigoAnterior = rsCaixaPosicao.Fields("Cliente").Value
      ' fim cores

      rsCaixaPosicao.MoveNext
  Wend
  rsCaixaPosicao.Close
  Set rsCaixaPosicao = Nothing
  
  lbl_totalRegistros.Caption = "Total de registros: " & lngContadorRegGrid
  
  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If
  
  MsgBox "Erro na abertura da tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
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

