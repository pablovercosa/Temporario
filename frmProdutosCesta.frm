VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProdutosCesta 
   Caption         =   " Cesta de Produtos"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdutosCesta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   14310
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00F7F7F7&
      Caption         =   "Sugestão de preço da Cesta"
      Height          =   3405
      Left            =   10770
      TabIndex        =   20
      Top             =   3450
      Width           =   3495
      Begin VB.ComboBox cmb_tabPreco 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   630
         Width           =   3075
      End
      Begin VB.CommandButton cmd_sugestaoPreco 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sugestão de Preço"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1200
         Width           =   3075
      End
      Begin VB.TextBox txt_precoSugerido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   1965
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Tabela de Preço"
         Height          =   225
         Left            =   240
         TabIndex        =   25
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Preço Sugerido"
         Height          =   225
         Left            =   240
         TabIndex        =   24
         Top             =   2250
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E5E5E5&
      Caption         =   "Movimentar ESTOQUE da CESTA"
      Height          =   3405
      Left            =   10770
      TabIndex        =   15
      Top             =   -30
      Width           =   3495
      Begin VB.CommandButton cmd_entradaEstoqueCesta 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Adicionar unidades de Cesta no Estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1410
         Width           =   3075
      End
      Begin VB.CommandButton cmd_retirarEstoqueCesta 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Baixar unidades de Cesta do Estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2190
         Width           =   3075
      End
      Begin VB.TextBox txt_estoqueAtual 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   390
         Left            =   240
         TabIndex        =   16
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Estoque ATUAL"
         Height          =   225
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1185
      End
   End
   Begin VB.TextBox txt_codProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   90
      TabIndex        =   14
      Top             =   5370
      Width           =   2235
   End
   Begin VB.TextBox txt_nomeProd 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   2340
      TabIndex        =   13
      Top             =   5370
      Width           =   6105
   End
   Begin VB.CommandButton cmd_acharProduto 
      BackColor       =   &H00C0FFFF&
      Height          =   435
      Left            =   8490
      Picture         =   "frmProdutosCesta.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5325
      Width           =   705
   End
   Begin VB.TextBox txt_codigoCesta 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   1440
      TabIndex        =   7
      Top             =   90
      Width           =   2355
   End
   Begin VB.TextBox txt_nomeCesta 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   3810
      TabIndex        =   6
      Top             =   90
      Width           =   6825
   End
   Begin VB.CommandButton cmd_novo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Incluir Item na Cesta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4650
      Width           =   5175
   End
   Begin VB.CommandButton cmd_excluir 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Excluir Item da Cesta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4650
      Width           =   5175
   End
   Begin VB.TextBox txt_qtde 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   9480
      MaxLength       =   4
      TabIndex        =   2
      Top             =   5385
      Width           =   1125
   End
   Begin VB.CommandButton cmd_salvar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salvar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5940
      Width           =   10545
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6420
      Width           =   10545
   End
   Begin MSFlexGridLib.MSFlexGrid grade_produtos 
      Height          =   3855
      Left            =   90
      TabIndex        =   5
      Top             =   750
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColor       =   12648447
      BackColorSel    =   12648384
      ForeColorSel    =   0
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
   Begin VB.Label Label5 
      Caption         =   "Lista de produtos que compõe a Cesta"
      Height          =   225
      Left            =   90
      TabIndex        =   11
      Top             =   510
      Width           =   2985
   End
   Begin VB.Label Label9 
      Caption         =   "Código da Cesta"
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   150
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Qtde"
      Height          =   195
      Left            =   9480
      TabIndex        =   9
      Top             =   5160
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "Produto"
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   5130
      Width           =   645
   End
End
Attribute VB_Name = "frmProdutosCesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CodigoProduto As String
Public NomeProduto As String
Public s_Funcao As String

Private Function SensibilizaEstoqueProduto(iFuncao As Integer, sCodProduto As String, dQtde As Double, sCodProdutoCesta As String)
On Error GoTo Erro:
  Dim strSQL As String
  Dim rsEstoque As Recordset
  Dim rsProdAux As Recordset
  Dim Classe As Integer
  Dim SubClasse As Integer
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim Estoque_Final As Double
  Dim Estoque_Final_CESTA As Double
  
  '***************************************
  'Achar estoque atual da CESTA
  Call Acha_Produto(sCodProdutoCesta, sCodProdutoCesta, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)

  strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & gnCodFilial & _
             " AND Produto = '" & sCodProdutoCesta & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " ORDER BY Data"

  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

  With rsEstoque
    If Not (.BOF And .EOF) Then
      '.MoveFirst
      .MoveLast
      Estoque_Final_CESTA = .Fields("Estoque Final")
    Else
      Estoque_Final_CESTA = 0
    End If

    .Close
  End With

  Set rsEstoque = Nothing
  '***************************************
  

  '***************************************
  'Achar estoque atual do produto item da CESTA
  
  'iFuncao:     1 - Incluir na CESTA (então DIMINUIR O ESTOQUE DO PRODUTO)
  '             2 - Excluir da CESTA (então SOMAR O ESTOQUE DO PRODUTO)
  
  strSQL = "SELECT * FROM Produtos WHERE Código='" & sCodProduto & "'"
  Set rsProdAux = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  Classe = rsProdAux.Fields("Classe").Value
  SubClasse = rsProdAux.Fields("Sub Classe").Value
  rsProdAux.Close
  Set rsProdAux = Nothing

  Call Acha_Produto(sCodProduto, sCodProduto, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)

  strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & gnCodFilial & _
             " AND Produto = '" & sCodProduto & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " ORDER BY Data"

  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

  With rsEstoque
    If Not (.BOF And .EOF) Then
      '.MoveFirst
      .MoveLast
      Estoque_Final = .Fields("Estoque Final")
    Else
      Estoque_Final = 0
    End If

    .Close
  End With

  Set rsEstoque = Nothing

  strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & gnCodFilial & _
             " AND Produto = '" & sCodProduto & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " AND Data = #" & Format(Now, "mm/dd/yyyy") & "#"

  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rsEstoque
    If Not (.BOF And .EOF) Then
      .LockEdits = True
      .Edit
    Else
      .AddNew
      .Fields("Filial").Value = gnCodFilial
      .Fields("Data").Value = Format(Now, "dd/mm/yyyy")
      .Fields("Produto").Value = sCodProduto
      .Fields("Tamanho").Value = Tamanho
      .Fields("Cor").Value = Cor
      .Fields("Edição").Value = Edição
      .Fields("Classe").Value = Classe
      .Fields("Sub Classe").Value = SubClasse
      .Fields("Estoque Anterior").Value = Estoque_Final
      '.Update
      '.Requery
    End If
  End With
  
  If iFuncao = 1 Then
    rsEstoque.Fields("Ajuste Saída") = CDbl(Estoque_Final_CESTA * dQtde)
    Estoque_Final = Estoque_Final - (Estoque_Final_CESTA * dQtde)
  Else
    rsEstoque.Fields("Ajuste Entra") = CDbl(Estoque_Final_CESTA * dQtde)
    Estoque_Final = Estoque_Final + (Estoque_Final_CESTA * dQtde)
  End If
  
  rsEstoque("Estoque Final") = Estoque_Final
  
  rsEstoque.Update
  
  rsEstoque.LockEdits = False
  
  rsEstoque.Close
  
  Rem Arruma Estoque Final
  Grava_Estoque_Final gnCodFilial, sCodProduto, Tamanho, Cor, Edição, CSng(Estoque_Final), Format(Now, "dd/mm/yyyy")

  Exit Function

Erro:
  MsgBox "Erro ao tentar sensibilar o estoque...Detalhes do Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

Private Sub cmd_acharProduto_Click()
  CodigoProdutoCestaPesq = ""
  nChamaConsulta = 5
  
  frmPesquisaProduto.Show
End Sub

Private Sub cmd_cancelar_Click()
  grade_produtos.Enabled = True
  cmd_novo.Enabled = True
  cmd_excluir.Enabled = True
  
  txt_codProduto.Enabled = False
  txt_qtde.Enabled = False
  cmd_salvar.Enabled = False

  txt_codProduto.Text = ""
  txt_nomeProd.Text = ""
  txt_qtde.Text = ""
  s_Funcao = ""
End Sub

Private Sub cmd_excluir_Click()
On Error GoTo Erro:

  Dim sCodItem As String
  Dim lnResponse As Long
  Dim dQtde As Double

  If grade_produtos.RowSel < 1 Then
    MsgBox "Selecione um registro na grade.", vbInformation
    Exit Sub
  End If

  lnResponse = MsgBox("Deseja realmente excluir o registro?", vbYesNo, "Atenção")
  If lnResponse = vbNo Then
    Exit Sub
  End If

  sCodItem = grade_produtos.TextMatrix(grade_produtos.RowSel, 1)
  dQtde = CDbl(grade_produtos.TextMatrix(grade_produtos.RowSel, 3))
  
  'SENSIBILIZAR ESTOQUE
  SensibilizaEstoqueProduto 2, sCodItem, dQtde, txt_codigoCesta.Text
  
  db.Execute "Delete from ProdutoCesta where CodigoCesta= '" & txt_codigoCesta.Text & "' and CodigoItem = '" & sCodItem & "' "

  cmd_cancelar_Click
  carregarGrade

  MsgBox "Registro desvinculado com sucesso.", vbInformation, "Sucesso"
  
  Exit Sub
Erro:
  MsgBox "Erro ao tentar desvincular o registro...Detalhes do Erro: " & Err.Number & " " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmd_entradaEstoqueCesta_Click()
On Error GoTo Erro:

  Dim strSQL As String
  Dim sEstoque As String
  Dim rsEstoque As Recordset
  Dim rsProdAux As Recordset
  Dim rs_ProdutoCesta As Recordset
  Dim Classe As Integer
  Dim SubClasse As Integer
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim Estoque_Final As Double
  Dim Estoque_Final_CESTA As Double
  
  sEstoque = InputBox("Adicionar quantas unidades de Cesta no Estoque? ", "Cesta de Produtos", "1")
  
  If LTrim(RTrim(sEstoque)) = "" Then
    MsgBox "Informe o número de unidades de Cestas para adicionar ao Estoque. Ex: '1' ou '3.5' ou '4' etc", vbInformation
    Exit Sub
  End If
  
  ws.BeginTrans
  
  
  ' **********
  ' Estoque para a CESTA
  strSQL = "SELECT * FROM Produtos WHERE Código='" & CodigoProduto & "'"
  Set rsProdAux = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  Classe = rsProdAux.Fields("Classe").Value
  SubClasse = rsProdAux.Fields("Sub Classe").Value
  rsProdAux.Close
  Set rsProdAux = Nothing

  Call Acha_Produto(CodigoProduto, CodigoProduto, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)

  strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & gnCodFilial & _
             " AND Produto = '" & CodigoProduto & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " ORDER BY Data"

  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

  With rsEstoque
    If Not (.BOF And .EOF) Then
      '.MoveFirst
      .MoveLast
      Estoque_Final = .Fields("Estoque Final")
    Else
      Estoque_Final = 0
    End If

    .Close
  End With

  Set rsEstoque = Nothing

  strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & gnCodFilial & _
             " AND Produto = '" & CodigoProduto & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " AND Data = #" & Format(Now, "mm/dd/yyyy") & "#"

  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rsEstoque
    If Not (.BOF And .EOF) Then
      .LockEdits = True
      .Edit
    Else
      .AddNew
      .Fields("Filial").Value = gnCodFilial
      .Fields("Data").Value = Format(Now, "dd/mm/yyyy")
      .Fields("Produto").Value = CodigoProduto
      .Fields("Tamanho").Value = Tamanho
      .Fields("Cor").Value = Cor
      .Fields("Edição").Value = Edição
      .Fields("Classe").Value = Classe
      .Fields("Sub Classe").Value = SubClasse
      .Fields("Estoque Anterior").Value = Estoque_Final
      '.Update
      '.Requery
    End If
  End With
  
  Estoque_Final = Estoque_Final + CDbl(sEstoque)
  
  rsEstoque.Fields("Ajuste Entra") = CDbl(sEstoque)
  rsEstoque("Estoque Final") = Estoque_Final
  
  rsEstoque.Update
  
  rsEstoque.LockEdits = False
  
  rsEstoque.Close
  
  Rem Arruma Estoque Final
  Grava_Estoque_Final gnCodFilial, CodigoProduto, Tamanho, Cor, Edição, CSng(Estoque_Final), Format(Now, "dd/mm/yyyy")

  
  
  '****************************
  ' Agora Estoque para cada Item da CESTA, neste caso: Diminuir o estoque unitários dos itens pois foram para a CESTA
  
  strSQL = "SELECT C.CodigoCesta, C.CodigoItem, C.QuantidadeItem, P.Nome "
  strSQL = strSQL & " from [ProdutoCesta] C, Produtos P "
  strSQL = strSQL & " Where C.CodigoCesta = '" & CodigoProduto & "' and "
  strSQL = strSQL & " C.CodigoItem = P.Código order by P.Nome "
  Set rs_ProdutoCesta = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rs_ProdutoCesta.EOF And rs_ProdutoCesta.BOF) Then
    rs_ProdutoCesta.MoveFirst
  End If
  While Not rs_ProdutoCesta.EOF
      
      'Para cada item/produto da CESTA
      strSQL = "SELECT * FROM Produtos WHERE Código='" & rs_ProdutoCesta.Fields(1).Value & "'"
      Set rsProdAux = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      Classe = rsProdAux.Fields("Classe").Value
      SubClasse = rsProdAux.Fields("Sub Classe").Value
      rsProdAux.Close
      Set rsProdAux = Nothing
    
      Call Acha_Produto(rs_ProdutoCesta.Fields(1).Value, rs_ProdutoCesta.Fields(1).Value, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)
    
      strSQL = "SELECT * FROM Estoque WHERE " & _
                 " Filial = " & gnCodFilial & _
                 " AND Produto = '" & rs_ProdutoCesta.Fields(1).Value & "'" & _
                 " AND Tamanho = " & Tamanho & _
                 " AND Cor = " & Cor & _
                 " AND Edição = " & Edição & _
                 " ORDER BY Data"
    
      Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    
      With rsEstoque
        If Not (.BOF And .EOF) Then
          '.MoveFirst
          .MoveLast
          Estoque_Final = .Fields("Estoque Final")
        Else
          Estoque_Final = 0
        End If
    
        .Close
      End With
    
      Set rsEstoque = Nothing
    
      strSQL = "SELECT * FROM Estoque WHERE " & _
                 " Filial = " & gnCodFilial & _
                 " AND Produto = '" & rs_ProdutoCesta.Fields(1).Value & "'" & _
                 " AND Tamanho = " & Tamanho & _
                 " AND Cor = " & Cor & _
                 " AND Edição = " & Edição & _
                 " AND Data = #" & Format(Now, "mm/dd/yyyy") & "#"
    
      Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
      With rsEstoque
        If Not (.BOF And .EOF) Then
          .LockEdits = True
          .Edit
        Else
          .AddNew
          .Fields("Filial").Value = gnCodFilial
          .Fields("Data").Value = Format(Now, "dd/mm/yyyy")
          .Fields("Produto").Value = rs_ProdutoCesta.Fields(1).Value
          .Fields("Tamanho").Value = Tamanho
          .Fields("Cor").Value = Cor
          .Fields("Edição").Value = Edição
          .Fields("Classe").Value = Classe
          .Fields("Sub Classe").Value = SubClasse
          .Fields("Estoque Anterior").Value = Estoque_Final
          '.Update
          '.Requery
        End If
      End With
      
      Estoque_Final = Estoque_Final - (CDbl(sEstoque) * CDbl(rs_ProdutoCesta.Fields(2).Value))
      
      rsEstoque.Fields("Ajuste Saída") = CDbl(sEstoque) * CDbl(rs_ProdutoCesta.Fields(2).Value)
      rsEstoque("Estoque Final") = Estoque_Final
      
      rsEstoque.Update
      
      rsEstoque.LockEdits = False
      
      rsEstoque.Close
      
      Rem Arruma Estoque Final
      Grava_Estoque_Final gnCodFilial, rs_ProdutoCesta.Fields(1).Value, Tamanho, Cor, Edição, CSng(Estoque_Final), Format(Now, "dd/mm/yyyy")
      '
      
      rs_ProdutoCesta.MoveNext
  Wend
  rs_ProdutoCesta.Close
  Set rs_ProdutoCesta = Nothing
  
  ws.CommitTrans
  
  If txt_estoqueAtual.Text = "" Then
    txt_estoqueAtual.Text = sEstoque
  Else
    txt_estoqueAtual.Text = CDbl(txt_estoqueAtual.Text) + CDbl(sEstoque)
  End If
  MsgBox "Adicionado " & sEstoque & " unidade(s) de Cesta no Estoque.", vbInformation, "Sucesso"

  Exit Sub
Erro:
  MsgBox "Erro ao tentar salvar o registro...Detalhes do Erro: " & Err.Number & " " & Err.Description, vbCritical, "Erro"

  ws.Rollback

End Sub

Private Sub cmd_novo_Click()
  cmd_acharProduto.Enabled = True
  txt_codProduto.Enabled = True
  grade_produtos.Enabled = False
  cmd_excluir.Enabled = False
  txt_qtde.Enabled = True
  cmd_salvar.Enabled = True

  txt_qtde.Text = ""
  txt_codProduto.Text = ""
  txt_nomeProd.Text = ""

  s_Funcao = "INCLUIR"
End Sub

Private Sub cmd_retirarEstoqueCesta_Click()
On Error GoTo Erro:

  Dim strSQL As String
  Dim sEstoque As String
  Dim rsEstoque As Recordset
  Dim rsProdAux As Recordset
  Dim rs_ProdutoCesta As Recordset
  Dim Classe As Integer
  Dim SubClasse As Integer
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim Estoque_Final As Double
  Dim Estoque_Final_CESTA As Double
  
  sEstoque = InputBox("Baixar quantas unidades de Cesta do Estoque? ", "Cesta de Produtos", "1")
  
  If LTrim(RTrim(sEstoque)) = "" Then
    MsgBox "Informe o número de unidades de Cestas para diminuir do Estoque. Ex: '1' ou '3.5' ou '4' etc", vbInformation
    Exit Sub
  End If
  
  ws.BeginTrans
  
  
  ' **********
  ' Estoque para a CESTA
  strSQL = "SELECT * FROM Produtos WHERE Código='" & CodigoProduto & "'"
  Set rsProdAux = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  Classe = rsProdAux.Fields("Classe").Value
  SubClasse = rsProdAux.Fields("Sub Classe").Value
  rsProdAux.Close
  Set rsProdAux = Nothing

  Call Acha_Produto(CodigoProduto, CodigoProduto, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)

  strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & gnCodFilial & _
             " AND Produto = '" & CodigoProduto & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " ORDER BY Data"

  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

  With rsEstoque
    If Not (.BOF And .EOF) Then
      '.MoveFirst
      .MoveLast
      Estoque_Final = .Fields("Estoque Final")
    Else
      Estoque_Final = 0
    End If

    .Close
  End With

  Set rsEstoque = Nothing

  strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & gnCodFilial & _
             " AND Produto = '" & CodigoProduto & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " AND Data = #" & Format(Now, "mm/dd/yyyy") & "#"

  Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rsEstoque
    If Not (.BOF And .EOF) Then
      .LockEdits = True
      .Edit
    Else
      .AddNew
      .Fields("Filial").Value = gnCodFilial
      .Fields("Data").Value = Format(Now, "dd/mm/yyyy")
      .Fields("Produto").Value = CodigoProduto
      .Fields("Tamanho").Value = Tamanho
      .Fields("Cor").Value = Cor
      .Fields("Edição").Value = Edição
      .Fields("Classe").Value = Classe
      .Fields("Sub Classe").Value = SubClasse
      .Fields("Estoque Anterior").Value = Estoque_Final
      '.Update
      '.Requery
    End If
  End With
  
  Estoque_Final = Estoque_Final - CDbl(sEstoque)
  
  rsEstoque.Fields("Ajuste Saída") = CDbl(sEstoque)
  rsEstoque("Estoque Final") = Estoque_Final
  
  rsEstoque.Update
  
  rsEstoque.LockEdits = False
  
  rsEstoque.Close
  
  Rem Arruma Estoque Final
  Grava_Estoque_Final gnCodFilial, CodigoProduto, Tamanho, Cor, Edição, CSng(Estoque_Final), Format(Now, "dd/mm/yyyy")

  
  
  '****************************
  ' Agora Estoque para cada Item da CESTA, neste caso: Diminuir o estoque unitários dos itens pois foram para a CESTA
  
  strSQL = "SELECT C.CodigoCesta, C.CodigoItem, C.QuantidadeItem, P.Nome "
  strSQL = strSQL & " from [ProdutoCesta] C, Produtos P "
  strSQL = strSQL & " Where C.CodigoCesta = '" & CodigoProduto & "' and "
  strSQL = strSQL & " C.CodigoItem = P.Código order by P.Nome "
  Set rs_ProdutoCesta = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rs_ProdutoCesta.EOF And rs_ProdutoCesta.BOF) Then
    rs_ProdutoCesta.MoveFirst
  End If
  While Not rs_ProdutoCesta.EOF
      
      'Para cada item/produto da CESTA
      strSQL = "SELECT * FROM Produtos WHERE Código='" & rs_ProdutoCesta.Fields(1).Value & "'"
      Set rsProdAux = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      Classe = rsProdAux.Fields("Classe").Value
      SubClasse = rsProdAux.Fields("Sub Classe").Value
      rsProdAux.Close
      Set rsProdAux = Nothing
    
      Call Acha_Produto(rs_ProdutoCesta.Fields(1).Value, rs_ProdutoCesta.Fields(1).Value, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)
    
      strSQL = "SELECT * FROM Estoque WHERE " & _
                 " Filial = " & gnCodFilial & _
                 " AND Produto = '" & rs_ProdutoCesta.Fields(1).Value & "'" & _
                 " AND Tamanho = " & Tamanho & _
                 " AND Cor = " & Cor & _
                 " AND Edição = " & Edição & _
                 " ORDER BY Data"
    
      Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    
      With rsEstoque
        If Not (.BOF And .EOF) Then
          '.MoveFirst
          .MoveLast
          Estoque_Final = .Fields("Estoque Final")
        Else
          Estoque_Final = 0
        End If
    
        .Close
      End With
    
      Set rsEstoque = Nothing
    
      strSQL = "SELECT * FROM Estoque WHERE " & _
                 " Filial = " & gnCodFilial & _
                 " AND Produto = '" & rs_ProdutoCesta.Fields(1).Value & "'" & _
                 " AND Tamanho = " & Tamanho & _
                 " AND Cor = " & Cor & _
                 " AND Edição = " & Edição & _
                 " AND Data = #" & Format(Now, "mm/dd/yyyy") & "#"
    
      Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
      With rsEstoque
        If Not (.BOF And .EOF) Then
          .LockEdits = True
          .Edit
        Else
          .AddNew
          .Fields("Filial").Value = gnCodFilial
          .Fields("Data").Value = Format(Now, "dd/mm/yyyy")
          .Fields("Produto").Value = rs_ProdutoCesta.Fields(1).Value
          .Fields("Tamanho").Value = Tamanho
          .Fields("Cor").Value = Cor
          .Fields("Edição").Value = Edição
          .Fields("Classe").Value = Classe
          .Fields("Sub Classe").Value = SubClasse
          .Fields("Estoque Anterior").Value = Estoque_Final
          '.Update
          '.Requery
        End If
      End With
      
      Estoque_Final = Estoque_Final + (CDbl(sEstoque) * CDbl(rs_ProdutoCesta.Fields(2).Value))
      
      rsEstoque.Fields("Ajuste Entra") = CDbl(sEstoque) * CDbl(rs_ProdutoCesta.Fields(2).Value)
      rsEstoque("Estoque Final") = Estoque_Final
      
      rsEstoque.Update
      
      rsEstoque.LockEdits = False
      
      rsEstoque.Close
      
      Rem Arruma Estoque Final
      Grava_Estoque_Final gnCodFilial, rs_ProdutoCesta.Fields(1).Value, Tamanho, Cor, Edição, CSng(Estoque_Final), Format(Now, "dd/mm/yyyy")
      '
      
      rs_ProdutoCesta.MoveNext
  Wend
  rs_ProdutoCesta.Close
  Set rs_ProdutoCesta = Nothing
  
  ws.CommitTrans
  
  If txt_estoqueAtual.Text = "" Then
    txt_estoqueAtual.Text = sEstoque
  Else
    txt_estoqueAtual.Text = CDbl(txt_estoqueAtual.Text) - CDbl(sEstoque)
  End If
  MsgBox "Baixado " & sEstoque & " unidade(s) de Cesta do Estoque.", vbInformation, "Sucesso"

  Exit Sub
Erro:
  MsgBox "Erro ao tentar salvar o registro...Detalhes do Erro: " & Err.Number & " " & Err.Description, vbCritical, "Erro"

  ws.Rollback
End Sub

Private Sub cmd_salvar_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim sDado As String
  Dim sDadoArray() As String

  If s_Funcao = "INCLUIR" Then
    If txt_codProduto.Text = "" Then
      MsgBox "Informe um Produto que irá compor a Cesta", vbInformation, "Atenção"
      Exit Sub
    End If

    If txt_qtde.Text = "" Then
      MsgBox "Informe a quantidade de unidades do Produto que irá compor a Cesta", vbInformation, "Atenção"
      Exit Sub
    End If
    
    If txt_nomeProd.Text = "" Then
      Dim rsProdPesqExiste As Recordset
      sSql = "SELECT * FROM Produtos WHERE Código='" & txt_codProduto.Text & "' "
    
      Set rsProdPesqExiste = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
    
      With rsProdPesqExiste
        If .BOF And .EOF Then
          MsgBox "Informe um Produto válido.", vbInformation, "Atenção"
          .Close
          Set rsProdPesqExiste = Nothing
          Exit Sub
        End If
    
        .Close
      End With
    
      Set rsProdPesqExiste = Nothing
    End If
    
    

    sSql = "Insert into ProdutoCesta (CodigoCesta, CodigoItem, QuantidadeItem) VALUES ('" & CodigoProduto & "','" & LTrim(RTrim(txt_codProduto.Text)) & "',"
    sSql = sSql & Replace(txt_qtde.Text, ",", ".") & ") "

    db.Execute sSql

    If db.RecordsAffected > 0 Then

      'Sensibilidar ESTOQUE
      SensibilizaEstoqueProduto 1, txt_codProduto.Text, CDbl(txt_qtde.Text), txt_codigoCesta.Text
    Else
      MsgBox "Registro já existe na base ou algum parâmetro que vc digitou esta inconsistente", vbInformation, "Atenção"
    End If
    cmd_cancelar_Click

    carregarGrade
  End If

  Exit Sub
Erro:
  MsgBox "Erro ao tentar salvar o registro...Detalhes do Erro: " & Err.Number & " " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmd_sugestaoPreco_Click()
On Error GoTo Erro
  Dim dPrecoSugerido As Double
  Dim sMensagem As String
  
  If cmb_tabPreco.Text = "" Then
    MsgBox "Selecione uma tabela de preços.", vbInformation, "Atenção"
    cmb_tabPreco.SetFocus
    Exit Sub
  End If
  
  dPrecoSugerido = 0
  sMensagem = ""
  
  Dim rs_ProdutoCesta As Recordset
  Dim strSQL As String

  strSQL = "SELECT C.CodigoCesta, C.CodigoItem, C.QuantidadeItem, [Preços].Preço "
  strSQL = strSQL & " from [ProdutoCesta] C "
  strSQL = strSQL & " LEFT JOIN [Preços] ON "
  strSQL = strSQL & " (C.[CodigoItem] = [Preços].produto and [Preços].Tabela = '" & cmb_tabPreco.Text & "') "
  strSQL = strSQL & " Where C.CodigoCesta = '" & CodigoProduto & "' " ' and "
  Set rs_ProdutoCesta = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rs_ProdutoCesta.EOF And rs_ProdutoCesta.BOF) Then
    rs_ProdutoCesta.MoveFirst
  End If
  While Not rs_ProdutoCesta.EOF
      If rs_ProdutoCesta.Fields(3).Value = "" Or rs_ProdutoCesta.Fields(3).Value = "0" Then
        sMensagem = sMensagem & rs_ProdutoCesta.Fields(1).Value & vbCrLf
      Else
        dPrecoSugerido = dPrecoSugerido + CDbl(rs_ProdutoCesta.Fields(3).Value) * CDbl(rs_ProdutoCesta.Fields(2).Value)
      End If
      rs_ProdutoCesta.MoveNext
  Wend
  rs_ProdutoCesta.Close
  Set rs_ProdutoCesta = Nothing
  
  If sMensagem <> "" Then
    MsgBox "Este(s) produto(s) não possuem preço (R$) na tabela selecionada: " & vbCrLf & sMensagem, vbInformation, "Atenção"
    txt_precoSugerido.Text = "Incompleto"
  Else
    txt_precoSugerido.Text = Format(dPrecoSugerido, "###,###,##0.00")
  End If

  Exit Sub
Erro:
  
  If Not (rs_ProdutoCesta Is Nothing) Then
      rs_ProdutoCesta.Close
      Set rs_ProdutoCesta = Nothing
  End If
  
  MsgBox "Erro ao realizar carga da grade...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_Activate()
  If CodigoProdutoCestaPesq <> "" Then
    txt_codProduto.Text = CodigoProdutoCestaPesq
    txt_nomeProd.Text = NomeProdutoCestaPesq
  Else
    txt_codProduto.Text = ""
    txt_nomeProd.Text = ""
  End If
End Sub

Private Sub Form_Load()
On Error GoTo Erro:

  grade_produtos.ColWidth(0) = 10
  grade_produtos.ColWidth(1) = 2000
  grade_produtos.ColAlignment(1) = flexAlignRightCenter
  grade_produtos.ColWidth(2) = 6200
  grade_produtos.ColWidth(3) = 2000
  
  grade_produtos.Row = 0
  grade_produtos.TextMatrix(0, 1) = "Cód.Produto"
  grade_produtos.TextMatrix(0, 2) = "Nome Produto"
  grade_produtos.TextMatrix(0, 3) = "Qtde"
  
  carregarGrade
  
  Dim rs_TabPrecos As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Tabela from [Tabela de Preços] "
  Set rs_TabPrecos = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rs_TabPrecos.EOF And rs_TabPrecos.BOF) Then
    rs_TabPrecos.MoveFirst
  End If
  While Not rs_TabPrecos.EOF
      cmb_tabPreco.AddItem rs_TabPrecos.Fields(0).Value
      rs_TabPrecos.MoveNext
  Wend
  rs_TabPrecos.Close
  Set rs_TabPrecos = Nothing
  
  If Not IsNull(CodigoProduto) And CodigoProduto <> "" Then
    txt_codigoCesta.Text = CodigoProduto
    txt_nomeCesta.Text = NomeProduto
  End If
  
  'Carregar ESTOQUE ATUAL
  Dim rs_Estoque As Recordset
  strSQL = "SELECT [Estoque Atual] from [Estoque Final] where Produto='" & CodigoProduto & "' and filial=" & gnCodFilial
  Set rs_Estoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If rs_Estoque.EOF And rs_Estoque.BOF Then
    'lbl_aviso.Caption = "Inicie Estoque com pelo Menos uma UNIDADE para somente depois poder montar a Cesta."
    txt_estoqueAtual.Text = ""
  Else
    txt_estoqueAtual.Text = rs_Estoque.Fields(0).Value
  End If
  rs_Estoque.Close
  Set rs_Estoque = Nothing
  
  Exit Sub
Erro:
  If Not (rs_TabPrecos Is Nothing) Then
      rs_TabPrecos.Close
      Set rs_TabPrecos = Nothing
  End If

  MsgBox "Erro ao realizar carga da tela...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub carregarGrade()
On Error GoTo Erro
  Dim iStatus As Integer
  Dim sStatus As String
  
  grade_produtos.Rows = 1
  
  Dim rs_ProdutoCesta As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT C.CodigoCesta, C.CodigoItem, C.QuantidadeItem, P.Nome "
  strSQL = strSQL & " from [ProdutoCesta] C, Produtos P "
  strSQL = strSQL & " Where C.CodigoCesta = '" & CodigoProduto & "' and "
  strSQL = strSQL & " C.CodigoItem = P.Código order by P.Nome "
  Set rs_ProdutoCesta = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rs_ProdutoCesta.EOF And rs_ProdutoCesta.BOF) Then
    rs_ProdutoCesta.MoveFirst
  End If
  While Not rs_ProdutoCesta.EOF
      grade_produtos.AddItem 0 & vbTab & rs_ProdutoCesta.Fields(1).Value & vbTab & rs_ProdutoCesta.Fields(3).Value & vbTab & rs_ProdutoCesta.Fields(2).Value
      rs_ProdutoCesta.MoveNext
  Wend
  rs_ProdutoCesta.Close
  Set rs_ProdutoCesta = Nothing

  Exit Sub
Erro:
  
  If Not (rs_ProdutoCesta Is Nothing) Then
      rs_ProdutoCesta.Close
      Set rs_ProdutoCesta = Nothing
  End If
  
  MsgBox "Erro ao realizar carga da grade...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub
