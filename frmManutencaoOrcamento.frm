VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManutencaoOrcamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção de orçamentos"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManutencaoOrcamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   9015
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   -120
      TabIndex        =   4
      Top             =   -120
      Width           =   9375
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmManutencaoOrcamento.frx":058A
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   1440
         TabIndex        =   6
         Top             =   480
         Width           =   7575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manutenção dos orçamentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   270
         Left            =   360
         Picture         =   "frmManutencaoOrcamento.frx":068A
         Top             =   480
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdAlterarOrcamento 
      BackColor       =   &H0000C0C0&
      Caption         =   "Atualizar Orçamentos"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin SSDataWidgets_B.SSDBGrid grdProdutos 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   8775
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
      Col.Count       =   4
      BevelColorFrame =   -2147483632
      BevelColorShadow=   -2147483633
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   6271
      Columns(0).Caption=   "Produto"
      Columns(0).Name =   "Produto"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1773
      Columns(1).Caption=   "Qtde"
      Columns(1).Name =   "Qtde"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2302
      Columns(2).Caption=   "Preço"
      Columns(2).Name =   "PrecoUnitario"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2646
      Columns(3).Caption=   "Total"
      Columns(3).Name =   "PrecoTotal"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      _ExtentX        =   15478
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "Produtos"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid grdOrcamentos 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8775
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
      Col.Count       =   4
      BevelColorFrame =   -2147483632
      BevelColorShadow=   -2147483633
      CheckBox3D      =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Caption=   "Filial"
      Columns(0).Name =   "Filial"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3200
      Columns(1).Caption=   "Sequencia"
      Columns(1).Name =   "Sequencia"
      Columns(1).Alignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Caption=   "Total"
      Columns(2).Name =   "Total"
      Columns(2).Alignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3200
      Columns(3).Caption=   "Aprovado"
      Columns(3).Name =   "Aprovado"
      Columns(3).Alignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      _ExtentX        =   15478
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "Orçamentos"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmManutencaoOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadOrcamentos()
  Dim rstOrcamentos As Recordset
  Dim strSQL        As String
  
  strSQL = " SELECT Saídas.*, [Operações Saída].* "
  strSQL = strSQL & " FROM Saídas, [Operações Saída] "
  strSQL = strSQL & " WHERE Saídas.Operação = [Operações Saída].Código "
  strSQL = strSQL & " AND [Operações Saída].Tipo = 'O' AND Saídas.OrcamentoAprovado = FALSE "
  
  Set rstOrcamentos = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  grdOrcamentos.Redraw = False
  grdOrcamentos.RemoveAll
  grdProdutos.RemoveAll
  
  With rstOrcamentos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        grdOrcamentos.AddNew
        
        grdOrcamentos.Columns("Filial").Text = .Fields("Filial").Value
        grdOrcamentos.Columns("Sequencia").Text = .Fields("Sequência").Value
        grdOrcamentos.Columns("Total").Text = Format(.Fields("Total").Value, FORMAT_VALUE)
                
        grdOrcamentos.Update
        
        .MoveNext
      Loop
    End If
  
    .Close
    Set rstOrcamentos = Nothing
  End With
  
  grdOrcamentos.MoveFirst
  grdOrcamentos.Redraw = True
End Sub

Private Sub cmdAlterarOrcamento_Click()
  Dim rstSaidas As Recordset
  Dim strSQL    As String
  Dim intX      As Integer
  
  grdOrcamentos.MoveFirst
  For intX = 0 To grdOrcamentos.Rows - 1
    strSQL = " SELECT * FROM Saídas "
    strSQL = strSQL & " WHERE Filial = " & grdOrcamentos.Columns("Filial").Text
    strSQL = strSQL & " AND Sequência = " & grdOrcamentos.Columns("Sequencia").Text
    
    Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
    With rstSaidas
      .Edit
      
      If IsDataType(dtBoolean, grdOrcamentos.Columns("Aprovado").Value) Then
        If CBool(.Fields("OrcamentoAprovado").Value) <> CBool(grdOrcamentos.Columns("Aprovado").Value) Then
          .Fields("ComentariosSobreOrcamento").Value = InputBox("Insira as observações sobre a aprovação do orçamento " & _
                  grdOrcamentos.Columns("Sequencia").Text, "Quick Store")
        End If
        .Fields("OrcamentoAprovado").Value = grdOrcamentos.Columns("Aprovado").Value
        .Update
      End If
      .Close
      Set rstSaidas = Nothing
    End With
    
    grdOrcamentos.MoveNext
  Next intX
  
  LoadOrcamentos
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  LoadOrcamentos
End Sub

Private Sub grdOrcamentos_Click()
  Dim rstProdutos As Recordset
  Dim strSQL      As String
  
  '14/06/2005 - Daniel
  'Adicionado validação para o usuário não clicar sem ter nada
  'carregado na grid e dar o Run-time error '3075'
  If Len(grdOrcamentos.Columns("Filial").Text) <= 0 Then Exit Sub
  '-----------------------------------------------------------
  
  strSQL = " SELECT * FROM [Saídas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & grdOrcamentos.Columns("Filial").Text
  strSQL = strSQL & " AND Sequência = " & grdOrcamentos.Columns("Sequencia").Text
  
  Set rstProdutos = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  grdProdutos.Redraw = False
  grdProdutos.RemoveAll
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        grdProdutos.AddNew
        grdProdutos.Columns("Produto").Text = .Fields("Código") & " - " & getNomeProduto(.Fields("Código"))
        grdProdutos.Columns("Qtde").Text = .Fields("Qtde")
        grdProdutos.Columns("PrecoUnitario").Text = Format(.Fields("Preço"), FORMAT_VALUE)
        grdProdutos.Columns("PrecoTotal").Text = Format(.Fields("Preço Final"), FORMAT_VALUE)
        grdProdutos.Update
        
        .MoveNext
      Loop
    End If
    
    .Close
    Set rstProdutos = Nothing
  End With
  
  grdProdutos.MoveFirst
  grdProdutos.Redraw = True
End Sub

Private Sub grdOrcamentos_DblClick()
  With grdOrcamentos
    If IsNumeric(.Columns("Sequencia").Text) Then
      frmSaidas.txtSeq.Text = .Columns("Sequencia").Text
      frmSaidas.SearchRecord
      frmSaidas.Show
    End If
  End With
End Sub

Private Function getNomeProduto(strCodigoProduto As String) As String
  Dim rstProdutos As Recordset
  
  Set rstProdutos = db.OpenRecordset("SELECT Nome FROM Produtos WHERE Código = '" & strCodigoProduto & "'")
  
  With rstProdutos
    If (.BOF And .EOF) Then
      getNomeProduto = "<Produto_não_cadastrado>"
    Else
      getNomeProduto = .Fields("Nome").Value & ""
    End If
    
    .Close
    Set rstProdutos = Nothing
  End With
End Function
