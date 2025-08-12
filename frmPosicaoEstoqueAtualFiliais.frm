VERSION 5.00
Begin VB.Form frmPosicaoEstoqueAtualFiliais 
   Caption         =   " Estoque atual nas Filiais"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12810
   Icon            =   "frmPosicaoEstoqueAtualFiliais.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12765
   End
End
Attribute VB_Name = "frmPosicaoEstoqueAtualFiliais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lCodigoTransferencia As Long

Dim arrayTamanhos() As Variant
Dim arrayCores() As Variant
Dim contador_arrayTamanhos As Integer
Dim contador_arrayCores As Integer

Dim rsParametros As Recordset
Dim rsEstoque As Recordset
Dim rsGrade As Recordset
Dim rsEdicoes As Recordset
Dim Filial As Integer

Private Function AchaTamanho(pTamanho As Integer) As String
  Dim i As Integer
  AchaTamanho = ""

  For i = 0 To contador_arrayTamanhos - 1
      If arrayTamanhos(i, 0) = pTamanho Then
          AchaTamanho = arrayTamanhos(i, 1)
          Exit For
      End If
  Next
End Function

Private Function AchaCor(pCor As Integer) As String
  Dim i As Integer
  AchaCor = ""
  For i = 0 To contador_arrayCores - 1
      If arrayCores(i, 0) = pCor Then
          AchaCor = arrayCores(i, 1)
          Exit For
      End If
  Next
End Function

Sub Mostra_Estoque_Edição()
  Dim Prod As String
  Dim Prod1 As String
  Dim Cod_Completo As String
  Dim Fim As Integer
  Dim Edição As Long
  Dim Cor As Integer
  Dim Erro As Integer
  Dim Tipo As Integer
  Dim Edição_Str As String
  Dim Cor_Str As String
  Dim Estoque As Single
  Dim Aux_Str As String
  Dim Est_Str As String
 
  rsEdicoes.Index = "Produto"
  
  Edição = 0
  Fim = False
  Do
    rsEdicoes.Seek ">", gsCodProduto, Edição
    If rsEdicoes.NoMatch Then Fim = True
    If Fim = False Then If rsEdicoes("Produto") <> gsCodProduto Then Fim = True
    If Fim = False Then
      Edição = rsEdicoes("Código")
      Aux_Str = Right(String(5, "0") & Edição, 5)
      Cod_Completo = gsCodProduto & Aux_Str
      Acha_Produto Cod_Completo, Prod1, 0, 0, Edição, Tipo, Erro
      If Erro = 0 Then
        Estoque = Acha_Estoque(Filial, gsCodProduto, 0, 0, Edição, Erro)
        Edição_Str = Right(Space(5) & Edição, 5)
        Est_Str = Right(Space(10) & Estoque, 10)
        Lista.AddItem " " & Edição_Str & Space(2) & Est_Str & "  " & rsParametros("Nome")
      End If
    End If
  Loop While Fim = False
End Sub

Sub Mostra_Estoque_Grade()
  Dim Prod As String
  Dim Prod1 As String
  Dim Cod_Completo As String
  Dim Fim As Integer
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Erro As Integer
  Dim Tipo As Integer
  Dim Tam_Str As String
  Dim Cor_Str As String
  Dim Estoque As Single
  Dim Est_Str As String
  Dim iCont As Integer
  
  rsGrade.Index = "Original"
  
  Cod_Completo = ""
  Fim = False
  Do
    rsGrade.Seek ">", gsCodProduto, Cod_Completo
    If rsGrade.NoMatch Then Fim = True
    If Fim = False Then
      If rsGrade("Código Original") <> gsCodProduto Then
        Fim = True
      End If
    End If
    If Fim = False Then
      Acha_Produto rsGrade("Código"), Prod1, Tamanho, Cor, 0, Tipo, Erro
      If Erro = 0 Then
        Cod_Completo = rsGrade("Código")
        Estoque = Acha_Estoque(Filial, gsCodProduto, Tamanho, Cor, 0, Erro)
        
        Tam_Str = Right(Space(5) & Tamanho, 5)
        Tam_Str = AchaTamanho(CInt(Tam_Str))
        For iCont = Len(Tam_Str) To 20
            Tam_Str = Tam_Str & " "
        Next
        
        Cor_Str = Right(Space(3) & Cor, 3)
        Cor_Str = AchaCor(CInt(Cor_Str))
        For iCont = Len(Cor_Str) To 20
            Cor_Str = Cor_Str & " "
        Next
        
        Est_Str = Right(Space(10) & Estoque, 10)
        Est_Str = Trim(Est_Str)
        For iCont = Len(Est_Str) To 9
            Est_Str = " " & Est_Str
        Next
        
        If Trim(Est_Str) <> "0" Then
            Lista.AddItem Tam_Str & Cor_Str & Est_Str & "   " & rsParametros("Nome")
        End If
      End If
    End If
  Loop While Fim = False
End Sub

Sub Mostra_Estoque_Normal()
  Dim Estoque As Double
  Dim Est_Str As String
  Dim nQtdeCasaDec As Integer
  
  Estoque = Acha_Estoque(Filial, gsCodProduto, 0, 0, 0, 0)
  
  If gbIsFrac(gsCodProduto, nQtdeCasaDec) Then
    Est_Str = Right(Space(10) & Round(Estoque, nQtdeCasaDec), 10)
  Else
    Est_Str = Right(Space(10) & Format(Estoque, "#0"), 10)
  End If

  If Trim(Est_Str) <> "0" Then
      Lista.AddItem " " & Est_Str & "  " & rsParametros("Nome")
  End If
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Prod As String
  Dim Prod1 As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Aux As String
  Dim Cod_Completo As String
  Dim Fim As Integer
  Dim Tam_Str As String
  Dim Cor_Str As String
  Dim Erro As Integer
  Dim Tipo As Integer
  Dim iContador As Integer
  Dim rsTamanho As Recordset
  Dim rsCor As Recordset
  
  Filial = 0

  Call CenterForm(Me)

  On Error GoTo Processa_Erro

  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsEdicoes = db.OpenRecordset("Edições", , dbReadOnly)

  Dim strSQL As String
  Dim rstTransfDet As Recordset
  
  For iContador = 0 To 1
  
      If iContador = 0 Then
          'Lista.AddItem " Quantidade  Filial/Empresa"
          
          '-----------------------------------------------------------------------------
          ' Selecionar os produtos NORMAIS da Transferência
          strSQL = "SELECT T.codigoProduto, T.Quantidade, T.NomeProduto, P.Tipo "
          strSQL = strSQL & " From TransferenciaProdutos T, Produtos P "
          strSQL = strSQL & " Where T.CodigoTransf = " & lCodigoTransferencia
          strSQL = strSQL & " And T.codigoProduto = P.Código "
      
      Else
          'Lista.AddItem "Tamanho              Cor                  Quantidade   Filial/Empresa"
      
          '-----------------------------------------------------------------------------
          ' Selecionar os produtos COM GRADE da Transferência
          rstTransfDet.Close
          Set rstTransfDet = Nothing
    
          strSQL = "SELECT T.codigoProduto, T.Quantidade, T.NomeProduto, P.Tipo, G.[Código Original] "
          strSQL = strSQL & " From TransferenciaProdutos T, Produtos P, [Códigos da Grade] G "
          strSQL = strSQL & " Where T.CodigoTransf = " & lCodigoTransferencia
          strSQL = strSQL & " And T.codigoProduto = G.Código "
          strSQL = strSQL & " And G.[Código Original] = P.Código "
      End If

      Set rstTransfDet = db.OpenRecordset(strSQL, dbOpenDynaset)
    
      '-----------------------------------------------------------------------------
      'Abaixo, para cada produto da transferência, mostrar o estoque em todos as filiais
      If Not (rstTransfDet.EOF And rstTransfDet.BOF) Then
          rstTransfDet.MoveFirst
          
          If iContador = 0 Then
              Lista.AddItem " Quantidade  Filial/Empresa"
          Else
              Lista.AddItem "Tamanho              Cor                  Quantidade   Filial/Empresa"
          End If
          
          While Not rstTransfDet.EOF
          
              If iContador = 0 Then
                  gsCodProduto = rstTransfDet.Fields(0).Value
                  gsTipoProduto = rstTransfDet.Fields(3).Value
              Else
                  gsCodProduto = rstTransfDet.Fields(4).Value
                  gsTipoProduto = rstTransfDet.Fields(3).Value
              End If
              Filial = 0
              
              Select Case gsTipoProduto
                Case "N"
                  Lista.AddItem "---------------------------------------------------------------------------------------------------------"
                  Lista.AddItem rstTransfDet.Fields(0).Value & " " & rstTransfDet.Fields(2).Value
                  'Lista.AddItem ""
                Case "G"
                  Lista.AddItem "---------------------------------------------------------------------------------------------------------"
                  Lista.AddItem rstTransfDet.Fields(0).Value & " " & rstTransfDet.Fields(2).Value
                  'Lista.AddItem ""
                  
                  iContador = 0
                  Set rsTamanho = db.OpenRecordset("select Código, Nome from Tamanhos ", dbOpenDynaset)
                  If Not (rsTamanho.EOF And rsTamanho.BOF) Then
                      rsTamanho.MoveLast
                      rsTamanho.MoveFirst
                      
                      ReDim arrayTamanhos(rsTamanho.RecordCount, 2)
                      contador_arrayTamanhos = rsTamanho.RecordCount
                      While Not rsTamanho.EOF
                          arrayTamanhos(iContador, 0) = rsTamanho.Fields(0).Value
                          arrayTamanhos(iContador, 1) = rsTamanho.Fields(1).Value
                          iContador = iContador + 1
                          rsTamanho.MoveNext
                      Wend
                  End If
                  rsTamanho.Close
                  Set rsTamanho = Nothing
                  
                  iContador = 0
                  Set rsCor = db.OpenRecordset("select Código, Nome from Cores ", dbOpenDynaset)
                  If Not (rsCor.EOF And rsCor.BOF) Then
                      rsCor.MoveLast
                      rsCor.MoveFirst
                      
                      ReDim arrayCores(rsCor.RecordCount, 2)
                      contador_arrayCores = rsCor.RecordCount
                      While Not rsCor.EOF
                          arrayCores(iContador, 0) = rsCor.Fields(0).Value
                          arrayCores(iContador, 1) = rsCor.Fields(1).Value
                          iContador = iContador + 1
                          rsCor.MoveNext
                      Wend
                  End If
                  rsCor.Close
                  Set rsCor = Nothing
            
                Case "E"
                  Lista.AddItem "---------------------------------------------------------------------------------------------------------"
                  Lista.AddItem rstTransfDet.Fields(0).Value & " " & rstTransfDet.Fields(2).Value
                  Lista.AddItem ""
                  'Lista.AddItem "Edição        Qtde  Filial/Empresa"
              End Select
          
              rsParametros.Index = "Filial"
Lp1:
              rsParametros.Seek ">", Filial
              If Not rsParametros.NoMatch Then
                Filial = rsParametros("Filial")
              '  Linha = rsParametros("Nome")
                Select Case gsTipoProduto
                  Case "N"
                    Call Mostra_Estoque_Normal
                  Case "G"
                    Call Mostra_Estoque_Grade
                  Case "E"
                    Call Mostra_Estoque_Edição
                End Select
                GoTo Lp1
              End If
          
              Lista.AddItem "---------------------------------------------------------------------------------------------------------"
              Lista.AddItem ""
              
              rstTransfDet.MoveNext
          Wend
      End If
  Next
  
  rsParametros.Close
  rsEstoque.Close
  rsGrade.Close
  rsEdicoes.Close
  Set rsParametros = Nothing
  Set rsEstoque = Nothing
  Set rsGrade = Nothing
  Set rsEdicoes = Nothing
  
  rstTransfDet.Close
  Set rstTransfDet = Nothing
  
  Exit Sub

Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apresentar registros em Consulta Estoque."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)

End Sub

