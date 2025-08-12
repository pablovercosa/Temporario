VERSION 5.00
Begin VB.Form frmConsultaEstoque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Consulta Estoque"
   ClientHeight    =   5115
   ClientLeft      =   1275
   ClientTop       =   2010
   ClientWidth     =   12825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConsultaEstoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5115
   ScaleWidth      =   12825
   Begin VB.CommandButton cmdFechar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4650
      Width           =   12735
   End
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
      Height          =   4515
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12765
   End
End
Attribute VB_Name = "frmConsultaEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Sub Mostra_Estoque_Edi��o()
  Dim Prod As String
  Dim Prod1 As String
  Dim Cod_Completo As String
  Dim Fim As Integer
  Dim Edi��o As Long
  Dim Cor As Integer
  Dim Erro As Integer
  Dim Tipo As Integer
  Dim Edi��o_Str As String
  Dim Cor_Str As String
  Dim Estoque As Single
  Dim Aux_Str As String
  Dim Est_Str As String
 
  rsEdicoes.Index = "Produto"
  
  Edi��o = 0
  Fim = False
  Do
    rsEdicoes.Seek ">", gsCodProduto, Edi��o
    If rsEdicoes.NoMatch Then Fim = True
    If Fim = False Then If rsEdicoes("Produto") <> gsCodProduto Then Fim = True
    If Fim = False Then
      Edi��o = rsEdicoes("C�digo")
      Aux_Str = Right(String(5, "0") & Edi��o, 5)
      Cod_Completo = gsCodProduto & Aux_Str
      Acha_Produto Cod_Completo, Prod1, 0, 0, Edi��o, Tipo, Erro
      If Erro = 0 Then
        Estoque = Acha_Estoque(Filial, gsCodProduto, 0, 0, Edi��o, Erro)
        Edi��o_Str = Right(Space(5) & Edi��o, 5)
        Est_Str = Right(Space(10) & Estoque, 10)
        Lista.AddItem " " & Edi��o_Str & Space(2) & Est_Str & "  " & rsParametros("Nome")
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
      If rsGrade("C�digo Original") <> gsCodProduto Then
        Fim = True
      End If
    End If
    If Fim = False Then
      Acha_Produto rsGrade("C�digo"), Prod1, Tamanho, Cor, 0, Tipo, Erro
      If Erro = 0 Then
        Cod_Completo = rsGrade("C�digo")
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
        
        Lista.AddItem Tam_Str & Cor_Str & Est_Str & "   " & rsParametros("Nome")
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
  
  
'  If InStr(Estoque, gsCurrencyDecimal) > 0 Then 'Qtde Fracionada
'    If Len(Estoque) > 10 Then
'      Est_Str = Right(Space(10) & Format(Estoque, "#0.00000"), 10)
'    Else
'      Est_Str = Right(Space(10) & Estoque, 10)
'    End If
'  Else
'    Est_Str = Right(Space(10) & Estoque, 10)
'  End If
'
  
  
  
  Lista.AddItem " " & Est_Str & "  " & rsParametros("Nome")
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Prod As String
  Dim Prod1 As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edi��o As Long
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

  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("C�digos da Grade", , dbReadOnly)
  Set rsEdicoes = db.OpenRecordset("Edi��es", , dbReadOnly)

  Select Case gsTipoProduto
    Case "N"
      Lista.AddItem " Quantidade  Filial/Empresa"
      Lista.AddItem "---------------------------------------------------------------------------------------------------------"
      Lista.AddItem ""
    Case "G"
      Lista.AddItem "Tamanho              Cor                  Quantidade   Filial/Empresa"
      Lista.AddItem "---------------------------------------------------------------------------------------------------------"
      Lista.AddItem ""
      
      iContador = 0
      Set rsTamanho = db.OpenRecordset("select C�digo, Nome from Tamanhos ", dbOpenDynaset)
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
      Set rsCor = db.OpenRecordset("select C�digo, Nome from Cores ", dbOpenDynaset)
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
      Lista.AddItem "Edi��o        Qtde  Filial/Empresa"
      Lista.AddItem "---------------------------------------------------------------------------------------------------------"
      Lista.AddItem ""
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
        Call Mostra_Estoque_Edi��o
    End Select
    GoTo Lp1
  End If
  Exit Sub

Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apresentar registros em Consulta Estoque."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)

End Sub
