VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProdutosPareamentoProdutoForn 
   BackColor       =   &H00FFA324&
   Caption         =   " Pareamento entre Códigos de Produto com o Fornecedor"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdutosPareamentoProdutoForn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_codigoProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1890
      TabIndex        =   10
      Top             =   60
      Width           =   2355
   End
   Begin VB.CommandButton cmd_consultarProduto 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultar produto"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5220
      Width           =   1815
   End
   Begin VB.CommandButton cmd_atualizar 
      BackColor       =   &H00FFA324&
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5220
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid FlxGd 
      Height          =   3735
      Left            =   60
      TabIndex        =   0
      Top             =   1410
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16250871
      BackColorSel    =   -2147483633
      BackColorBkg    =   16777215
      AllowUserResizing=   1
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
   Begin VB.Label lbl_tipoProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1890
      TabIndex        =   7
      Top             =   495
      Width           =   2355
   End
   Begin VB.Label lbl_nomeProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   4290
      TabIndex        =   6
      Top             =   60
      Width           =   5955
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFA324&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   540
      Width           =   450
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFA324&
      Caption         =   "Código do Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   1830
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFA324&
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Label lbl_fornCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1890
      TabIndex        =   2
      Top             =   1035
      Width           =   2355
   End
   Begin VB.Label lbl_fornNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4290
      TabIndex        =   1
      Top             =   1035
      Width           =   5955
   End
End
Attribute VB_Name = "frmProdutosPareamentoProdutoForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrayTamanhos() As Variant
Dim arrayCores() As Variant
Dim contador_arrayTamanhos As Integer
Dim contador_arrayCores As Integer

Public pCodigoProduto As String
Public pNomeProduto As String
Public pTipoProduto As String
Public pCodigoFornecedor As String
Public pNomeFornecedor As String

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

Private Sub cmd_atualizar_Click()
On Error GoTo Erro
  Dim sSql As String
  Dim iContador As Integer
  

  If lbl_tipoProduto.Caption = "NORMAL" Then
    ws.BeginTrans
    
    sSql = "Delete From ProdutoPareamentoFornecedor "
    sSql = sSql & " where Fornecedor = " & lbl_fornCodigo.Caption
    sSql = sSql & " AND Produto = '" & txt_codigoProduto.Text & "' "

    db.Execute sSql
    
    sSql = "Insert into ProdutoPareamentoFornecedor (Produto, Tipo, ProdutoForn, Fornecedor) "
    sSql = sSql & " values('" & txt_codigoProduto.Text & "', 'N', '" & FlxGd.TextMatrix(1, 4) & "', "
    sSql = sSql & lbl_fornCodigo.Caption & ")"

    db.Execute sSql

    ws.CommitTrans
  Else
    ' Grade
    ws.BeginTrans
    
    For iContador = 1 To FlxGd.Rows - 1
      sSql = "Delete From ProdutoPareamentoFornecedor "
      sSql = sSql & " where Fornecedor = " & lbl_fornCodigo.Caption
      sSql = sSql & " AND Produto = '" & FlxGd.TextMatrix(iContador, 1) & "' "
      
      db.Execute sSql
    Next
    
    For iContador = 1 To FlxGd.Rows - 1
      sSql = "Insert into ProdutoPareamentoFornecedor (Produto, Tipo, ProdutoForn, Fornecedor) "
      sSql = sSql & " values('" & FlxGd.TextMatrix(iContador, 1) & "', 'G', '" & FlxGd.TextMatrix(iContador, 4) & "', "
      sSql = sSql & lbl_fornCodigo.Caption & ")"

      db.Execute sSql
    Next

    ws.CommitTrans
  End If
  
  MsgBox "Pareamento realizado com sucesso.", vbInformation, "Sucesso"
  
  Exit Sub
Erro:
  ws.Rollback
  MsgBox "Erro na função de pareamento " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub cmd_consultarProduto_Click()
  Dim F As Form
  Set F = New frmPesquisaProduto
  F.Show
End Sub

Private Sub Form_Load()
On Error GoTo Erro
  Dim sSql As String
  Dim rsTamanho As Recordset
  Dim rsCor As Recordset
  Dim iContador As Integer
  Dim rsProdutos As Recordset
  Dim rsProdutos2 As Recordset
  Dim sTamanho As String
  Dim sCor As String
  Dim sTamanhoCorAux As String
      
  If pTipoProduto = "GRADE" Or pCodigoProduto = "" Then
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
  End If
  
  If pCodigoProduto = "" Then
      ' Tela de consulta de todos os produtos pareados do Fornecedor
      
      cmd_atualizar.Visible = False
      cmd_consultarProduto.Visible = False
      Label1.Visible = False
      Label3.Visible = False
      txt_codigoProduto.Visible = False
      lbl_nomeProduto.Visible = False
      lbl_tipoProduto.Visible = False
      Label2.Top = 150
      lbl_fornCodigo.Top = 120
      lbl_fornNome.Top = 120
      lbl_fornNome.Width = 10400
      FlxGd.Top = 540
      FlxGd.Height = 8000
      FlxGd.Width = 16650
      
      Me.Width = 16900
      Me.Height = 9100
      FlxGd.Cols = 6
      
      FlxGd.ColWidth(0) = 0
      FlxGd.ColWidth(1) = 2400
      FlxGd.ColWidth(2) = 2150
      FlxGd.ColWidth(3) = 2150
      FlxGd.ColWidth(4) = 7200
      FlxGd.ColWidth(5) = 2400
      
      FlxGd.Row = 0
      FlxGd.TextMatrix(0, 0) = ""
      FlxGd.TextMatrix(0, 1) = "Produto"
      FlxGd.TextMatrix(0, 2) = "Tamanho"
      FlxGd.TextMatrix(0, 3) = "Cor"
      FlxGd.TextMatrix(0, 4) = "Nome"
      FlxGd.TextMatrix(0, 5) = "Produto Fornecedor"
      
      lbl_fornCodigo.Caption = pCodigoFornecedor
      lbl_fornNome.Caption = pNomeFornecedor
      
      FlxGd.Rows = 1
      
      sSql = "Select P.Produto, P.Tipo, P.ProdutoForn, Produtos.Nome "
      sSql = sSql & " From Produtos, ProdutoPareamentoFornecedor P, [Códigos da Grade] "
      sSql = sSql & " Where P.Fornecedor = " & lbl_fornCodigo.Caption & " AND "
      sSql = sSql & " P.Tipo = 'G' AND "
      sSql = sSql & " P.Produto = [Códigos da Grade].[Código] AND "
      sSql = sSql & " [Códigos da Grade].[Código Original] = Produtos.Código "
      sSql = sSql & " Order by P.Produto "
      
      Set rsProdutos = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
      
      With rsProdutos
        If Not (.BOF And .EOF) Then
          Do Until .EOF

              sTamanhoCorAux = .Fields(0).Value
              sTamanhoCorAux = Mid(sTamanhoCorAux, Len(sTamanhoCorAux) - 5, 3)
              sTamanho = AchaTamanho(CInt(sTamanhoCorAux))
              
              sTamanhoCorAux = .Fields(0).Value
              sTamanhoCorAux = Mid(sTamanhoCorAux, Len(sTamanhoCorAux) - 2, 3)
              sCor = AchaCor(CInt(sTamanhoCorAux))
              
              'Adiciona registro
              FlxGd.AddItem vbTab & .Fields(0).Value & vbTab & _
                      sTamanho & vbTab & _
                      sCor & vbTab & _
                      .Fields(3).Value & vbTab & _
                      .Fields(2).Value
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsProdutos = Nothing
      
      sSql = "Select P.Produto, P.Tipo, P.ProdutoForn, Produtos.Nome "
      sSql = sSql & " From Produtos, ProdutoPareamentoFornecedor P "
      sSql = sSql & " Where P.Fornecedor = " & lbl_fornCodigo.Caption & " AND "
      sSql = sSql & " P.Tipo = 'N' AND "
      sSql = sSql & " P.Produto = Produtos.Código "
      sSql = sSql & " Order by P.Produto "
      
      Set rsProdutos2 = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
      
      With rsProdutos2
        If Not (.BOF And .EOF) Then
          Do Until .EOF
              
              sTamanho = ""
              sCor = ""
              'Adiciona registro
              FlxGd.AddItem vbTab & .Fields(0).Value & vbTab & _
                      sTamanho & vbTab & _
                      sCor & vbTab & _
                      .Fields(3).Value & vbTab & _
                      .Fields(2).Value
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsProdutos2 = Nothing
      
  Else
      ' Tela de manutenção de pareamento

      FlxGd.ColWidth(0) = 0
      FlxGd.ColWidth(1) = 2400
      
      If pTipoProduto = "GRADE" Then
        FlxGd.ColWidth(2) = 2500
        FlxGd.ColWidth(3) = 2500
      Else
        FlxGd.ColWidth(2) = 0
        FlxGd.ColWidth(3) = 0
        FlxGd.Height = 900
        cmd_atualizar.Top = 2400
        cmd_consultarProduto.Top = 2400
        Me.Height = 3270
      End If
      FlxGd.ColWidth(4) = 2400
    
      FlxGd.Row = 0
      FlxGd.TextMatrix(0, 0) = ""
      FlxGd.TextMatrix(0, 1) = "Produto"
      FlxGd.TextMatrix(0, 2) = "Tamanho"
      FlxGd.TextMatrix(0, 3) = "Cor"
      FlxGd.TextMatrix(0, 4) = "Produto Fornecedor"
      
      txt_codigoProduto.Text = pCodigoProduto
      lbl_nomeProduto.Caption = pNomeProduto
      lbl_tipoProduto.Caption = pTipoProduto
      lbl_fornCodigo.Caption = pCodigoFornecedor
      lbl_fornNome.Caption = pNomeFornecedor
    
      FlxGd.Rows = 1
    
      If pTipoProduto = "GRADE" Then
          sSql = "Select [Códigos da Grade].Código, ProdutoPareamentoFornecedor.ProdutoForn "
          sSql = sSql & " From [Códigos da Grade] LEFT Join ProdutoPareamentoFornecedor ON "
          sSql = sSql & " ([Códigos da Grade].[Código] = ProdutoPareamentoFornecedor.Produto AND "
          sSql = sSql & " ProdutoPareamentoFornecedor.Fornecedor = " & lbl_fornCodigo.Caption & ") "
          sSql = sSql & " where [Códigos da Grade].[Código Original] = '" & txt_codigoProduto.Text & "'"
          sSql = sSql & " Order by [Códigos da Grade].Código "
    
          Set rsProdutos = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
          With rsProdutos
            If Not (.BOF And .EOF) Then
              Do Until .EOF
    
                  sTamanhoCorAux = .Fields(0).Value
                  sTamanhoCorAux = Mid(sTamanhoCorAux, Len(sTamanhoCorAux) - 5, 3)
                  sTamanho = AchaTamanho(CInt(sTamanhoCorAux))
                  
                  sTamanhoCorAux = .Fields(0).Value
                  sTamanhoCorAux = Mid(sTamanhoCorAux, Len(sTamanhoCorAux) - 2, 3)
                  sCor = AchaCor(CInt(sTamanhoCorAux))
                  
                  'Adiciona registro
                  FlxGd.AddItem vbTab & .Fields(0).Value & vbTab & _
                          sTamanho & vbTab & _
                          sCor & vbTab & _
                          .Fields(1).Value
                    
    '''              FlxGd.Row = lLinhas
    '''              FlxGd.Col = 0
    '''              FlxGd.CellBackColor = &HC0FFFF
    '''              FlxGd.Col = 1
    '''              FlxGd.CellBackColor = &HC0FFFF
    '''              FlxGd.Col = 2
    '''              FlxGd.CellBackColor = &HC0FFFF
    '''              lLinhas = lLinhas + 1
                          
                .MoveNext
              Loop
            End If
            .Close
          End With
          Set rsProdutos = Nothing
      Else
          'Produto sem grade
          sSql = "Select ProdutoForn From ProdutoPareamentoFornecedor "
          sSql = sSql & " Where Produto = '" & txt_codigoProduto.Text & "'"
          sSql = sSql & " AND Fornecedor = " & lbl_fornCodigo.Caption
    
          Set rsProdutos = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
          With rsProdutos
            If Not (.BOF And .EOF) Then
              Do Until .EOF
                  'Adiciona registro
                  FlxGd.AddItem vbTab & txt_codigoProduto.Text & vbTab & _
                          vbTab & _
                          vbTab & _
                          .Fields(0).Value
                .MoveNext
              Loop
            Else
                  FlxGd.AddItem vbTab & txt_codigoProduto.Text & vbTab & _
                          vbTab & _
                          vbTab & _
                          ""
            End If
            .Close
          End With
          Set rsProdutos = Nothing
      End If
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro na abertura da tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub FlxGd_EnterCell()
    FlxGd.CellBackColor = &HC0FFFF    'Amarelo claro
    FlxGd.Tag = ""
End Sub
 
Private Sub FlxGd_LeaveCell()
    FlxGd.CellBackColor = &H80000005
End Sub
 
Private Sub FlxGd_KeyDown(KeyCode As Integer, Shift As Integer)

  '----------------------------------------------
  ' Não poderá digitar valor nas colunas < que 4
  If FlxGd.MouseCol < 4 Then
      Exit Sub
  End If
  '----------------------------------------------

  Select Case KeyCode
    Case 46                 '<Del>, apaga celula
        FlxGd.Tag = FlxGd
        FlxGd = ""
  End Select
End Sub
 
Private Sub FlxGd_KeyPress(KeyAscii As Integer)

    '----------------------------------------------
    ' Não poderá digitar valor nas colunas < que 4
    If FlxGd.MouseCol < 4 Then
        Exit Sub
    End If
    '----------------------------------------------

    Select Case KeyAscii
        Case 13            'Tecla ENTER
            Advance_Cell   'Avancar para nova cellula
        Case 8             'Backspace
            If Len(FlxGd) Then
              FlxGd = Left$(FlxGd, Len(FlxGd) - 1)
            End If
        Case 27                      'ESC
            If FlxGd.Tag > "" Then   'Se todos forem NULL
              FlxGd = FlxGd.Tag      'Retorna ao texto original
            End If
        Case Else
            FlxGd = FlxGd + Chr(KeyAscii)
    End Select
End Sub

Private Sub FlxGd_Click()
    FlxGd.CellBackColor = &H80000005
End Sub

Private Sub FlxGd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Row As Integer, Col As Integer
    Row = FlxGd.MouseRow
    Col = FlxGd.MouseCol
    If Button = 2 And (Col = 0 Or Row = 0) Then
      FlxGd.Col = IIf(Col = 0, 1, Col)
      FlxGd.Row = IIf(Row = 0, 1, Row)
      'PopupMenu MnuFGridRows
    End If
End Sub
 
Private Sub Advance_Cell()                  'Avancar para proxima cellula
    With FlxGd
        .HighLight = flexHighlightNever
        If .Col < .Cols - 1 Then
          .Col = .Col + 1
        Else
          If .Row < .Rows - 1 Then
            .Row = .Row + 1                 'Desce uma linha
            .Col = 1
          Else
            .Row = 1
            .Col = 1
          End If
        End If
        If .CellTop + .CellHeight > .Top + .Height Then
          .TopRow = .TopRow + 1
        End If
        .HighLight = flexHighlightAlways
    End With
End Sub

