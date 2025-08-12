VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmProdutosFavoritos 
   Caption         =   " Produtos Favoritos"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdutosFavoritos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_salvar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3840
      Width           =   9405
   End
   Begin VB.Data datProduto 
      Caption         =   "Produto"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7530
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   90
      Visible         =   0   'False
      Width           =   1830
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto1 
      Bindings        =   "frmProdutosFavoritos.frx":4E95A
      DataSource      =   "datProduto"
      Height          =   315
      Left            =   915
      TabIndex        =   1
      Top             =   90
      Width           =   2265
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3651
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto2 
      Bindings        =   "frmProdutosFavoritos.frx":4E973
      DataSource      =   "datProduto"
      Height          =   315
      Left            =   915
      TabIndex        =   4
      Top             =   570
      Width           =   2265
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3651
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto3 
      Bindings        =   "frmProdutosFavoritos.frx":4E98C
      DataSource      =   "datProduto"
      Height          =   315
      Left            =   915
      TabIndex        =   7
      Top             =   1050
      Width           =   2265
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3651
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto4 
      Bindings        =   "frmProdutosFavoritos.frx":4E9A5
      DataSource      =   "datProduto"
      Height          =   315
      Left            =   915
      TabIndex        =   10
      Top             =   1530
      Width           =   2265
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3651
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto5 
      Bindings        =   "frmProdutosFavoritos.frx":4E9BE
      DataSource      =   "datProduto"
      Height          =   315
      Left            =   915
      TabIndex        =   13
      Top             =   1980
      Width           =   2265
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3651
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto6 
      Bindings        =   "frmProdutosFavoritos.frx":4E9D7
      DataSource      =   "datProduto"
      Height          =   315
      Left            =   915
      TabIndex        =   16
      Top             =   2460
      Width           =   2265
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3651
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto7 
      Bindings        =   "frmProdutosFavoritos.frx":4E9F0
      DataSource      =   "datProduto"
      Height          =   315
      Left            =   915
      TabIndex        =   19
      Top             =   2940
      Width           =   2265
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3651
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboProduto8 
      Bindings        =   "frmProdutosFavoritos.frx":4EA09
      DataSource      =   "datProduto"
      Height          =   315
      Left            =   915
      TabIndex        =   22
      Top             =   3420
      Width           =   2265
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3651
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label lblNomeProduto7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3225
      TabIndex        =   20
      Top             =   2940
      Width           =   6240
   End
   Begin VB.Label Label13 
      Caption         =   "Produto  7"
      Height          =   225
      Left            =   105
      TabIndex        =   18
      Top             =   2985
      Width           =   795
   End
   Begin VB.Label Label12 
      Caption         =   "Produto  8"
      Height          =   225
      Left            =   105
      TabIndex        =   21
      Top             =   3465
      Width           =   795
   End
   Begin VB.Label lblNomeProduto8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3225
      TabIndex        =   23
      Top             =   3420
      Width           =   6240
   End
   Begin VB.Label lblNomeProduto5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3225
      TabIndex        =   14
      Top             =   1980
      Width           =   6240
   End
   Begin VB.Label Label9 
      Caption         =   "Produto  5"
      Height          =   225
      Left            =   105
      TabIndex        =   12
      Top             =   2025
      Width           =   795
   End
   Begin VB.Label Label8 
      Caption         =   "Produto  6"
      Height          =   225
      Left            =   105
      TabIndex        =   15
      Top             =   2505
      Width           =   795
   End
   Begin VB.Label lblNomeProduto6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3225
      TabIndex        =   17
      Top             =   2460
      Width           =   6240
   End
   Begin VB.Label lblNomeProduto3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3225
      TabIndex        =   8
      Top             =   1050
      Width           =   6240
   End
   Begin VB.Label Label5 
      Caption         =   "Produto  3"
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   1095
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Produto  4"
      Height          =   225
      Left            =   105
      TabIndex        =   9
      Top             =   1575
      Width           =   795
   End
   Begin VB.Label lblNomeProduto4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3210
      TabIndex        =   11
      Top             =   1530
      Width           =   6240
   End
   Begin VB.Label lblNomeProduto2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3225
      TabIndex        =   5
      Top             =   570
      Width           =   6240
   End
   Begin VB.Label Label2 
      Caption         =   "Produto  2"
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   615
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Produto  1"
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   795
   End
   Begin VB.Label lblNomeProduto1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3225
      TabIndex        =   2
      Top             =   90
      Width           =   6240
   End
End
Attribute VB_Name = "frmProdutosFavoritos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboProduto1_CloseUp()
  cboProduto1.Text = cboProduto1.Columns("Codigo").Text
  cboProduto1_LostFocus
End Sub

Private Sub cboProduto1_LostFocus()
  Call StatusMsg("")
  If cboProduto1.Text <> "" Then
    lblNomeProduto1.Caption = gsGetNameProduto(cboProduto1.Text)
  Else
    lblNomeProduto1.Caption = ""
  End If
End Sub

Private Sub cboProduto2_CloseUp()
  cboProduto2.Text = cboProduto2.Columns("Codigo").Text
  cboProduto2_LostFocus
End Sub

Private Sub cboProduto2_LostFocus()
  Call StatusMsg("")
  If cboProduto2.Text <> "" Then
    lblNomeProduto2.Caption = gsGetNameProduto(cboProduto2.Text)
  Else
    lblNomeProduto2.Caption = ""
  End If
End Sub

Private Sub cmd_salvar_Click()
On Error GoTo Erro

  ws.BeginTrans
  db.Execute "Delete from ProdutoFavoritos where Filial = " & gnCodFilial
  ws.CommitTrans
  
  ws.BeginTrans
  
  If Trim(cboProduto1.Text) <> "" Then
      db.Execute "Insert into ProdutoFavoritos(Filial, Produto) values (" & gnCodFilial & ",'" & Trim(cboProduto1.Text) & "') "
  End If
    
  If Trim(cboProduto2.Text) <> "" Then
      db.Execute "Insert into ProdutoFavoritos(Filial, Produto) values (" & gnCodFilial & ",'" & Trim(cboProduto2.Text) & "') "
  End If
    
  If Trim(cboProduto3.Text) <> "" Then
      db.Execute "Insert into ProdutoFavoritos(Filial, Produto) values (" & gnCodFilial & ",'" & Trim(cboProduto3.Text) & "') "
  End If
  
  If Trim(cboProduto4.Text) <> "" Then
      db.Execute "Insert into ProdutoFavoritos(Filial, Produto) values (" & gnCodFilial & ",'" & Trim(cboProduto4.Text) & "') "
  End If
  
  If Trim(cboProduto5.Text) <> "" Then
      db.Execute "Insert into ProdutoFavoritos(Filial, Produto) values (" & gnCodFilial & ",'" & Trim(cboProduto5.Text) & "') "
  End If
  
  If Trim(cboProduto6.Text) <> "" Then
      db.Execute "Insert into ProdutoFavoritos(Filial, Produto) values (" & gnCodFilial & ",'" & Trim(cboProduto6.Text) & "') "
  End If
  
  If Trim(cboProduto7.Text) <> "" Then
      db.Execute "Insert into ProdutoFavoritos(Filial, Produto) values (" & gnCodFilial & ",'" & Trim(cboProduto7.Text) & "') "
  End If
  
  If Trim(cboProduto8.Text) <> "" Then
      db.Execute "Insert into ProdutoFavoritos(Filial, Produto) values (" & gnCodFilial & ",'" & Trim(cboProduto8.Text) & "') "
  End If
  
  ws.CommitTrans

  Exit Sub
Erro:
  MsgBox "Erro ao salvar produtos favoritos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub Form_Load()
On Error GoTo Erro

  datProduto.DatabaseName = gsQuickDBFileName

  Dim rsProdutosFavoritos As Recordset
  Dim sSql As String
  
  sSql = "Select F.Produto, P.Nome from ProdutoFavoritos F, Produtos P "
  sSql = sSql & " where F.Filial = " & gnCodFilial & " and F.Produto = P.Código "
  
  Set rsProdutosFavoritos = db.OpenRecordset(sSql, dbOpenDynaset)
  If Not (rsProdutosFavoritos.EOF And rsProdutosFavoritos.BOF) Then
      rsProdutosFavoritos.MoveLast
      rsProdutosFavoritos.MoveFirst
      
      If rsProdutosFavoritos.RecordCount > 0 Then
          If rsProdutosFavoritos.RecordCount = 1 Then
              cboProduto1.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto1.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 2 Then
              cboProduto1.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto2.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto2.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 3 Then
              cboProduto1.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto2.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto3.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto3.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 4 Then
              cboProduto1.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto2.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto3.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto4.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto4.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 5 Then
              cboProduto1.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto2.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto3.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto4.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto4.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto5.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto5.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 6 Then
              cboProduto1.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto2.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto3.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto4.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto4.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto5.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto5.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto6.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto6.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 7 Then
              cboProduto1.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto2.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto3.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto4.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto4.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto5.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto5.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto6.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto6.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto7.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto7.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 8 Then
              cboProduto1.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto2.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto3.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto4.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto4.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto5.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto5.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto6.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto6.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto7.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto7.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              cboProduto8.Text = rsProdutosFavoritos.Fields(0).Value
              lblNomeProduto8.Caption = rsProdutosFavoritos.Fields(1).Value
          End If
      
      End If
  End If
  rsProdutosFavoritos.Close
  Set rsProdutosFavoritos = Nothing

  Exit Sub
Erro:
  MsgBox "Erro ao salvar produtos favoritos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cboProduto3_CloseUp()
  cboProduto3.Text = cboProduto3.Columns("Codigo").Text
  cboProduto3_LostFocus
End Sub

Private Sub cboProduto3_LostFocus()
  Call StatusMsg("")
  If cboProduto3.Text <> "" Then
    lblNomeProduto3.Caption = gsGetNameProduto(cboProduto3.Text)
  Else
    lblNomeProduto3.Caption = ""
  End If
End Sub

Private Sub cboProduto4_CloseUp()
  cboProduto4.Text = cboProduto4.Columns("Codigo").Text
  cboProduto4_LostFocus
End Sub

Private Sub cboProduto4_LostFocus()
  Call StatusMsg("")
  If cboProduto4.Text <> "" Then
    lblNomeProduto4.Caption = gsGetNameProduto(cboProduto4.Text)
  Else
    lblNomeProduto4.Caption = ""
  End If
End Sub

Private Sub cboProduto5_CloseUp()
  cboProduto5.Text = cboProduto5.Columns("Codigo").Text
  cboProduto5_LostFocus
End Sub

Private Sub cboProduto5_LostFocus()
  Call StatusMsg("")
  If cboProduto5.Text <> "" Then
    lblNomeProduto5.Caption = gsGetNameProduto(cboProduto5.Text)
  Else
    lblNomeProduto5.Caption = ""
  End If
End Sub

Private Sub cboProduto6_CloseUp()
  cboProduto6.Text = cboProduto6.Columns("Codigo").Text
  cboProduto6_LostFocus
End Sub

Private Sub cboProduto6_LostFocus()
  Call StatusMsg("")
  If cboProduto6.Text <> "" Then
    lblNomeProduto6.Caption = gsGetNameProduto(cboProduto6.Text)
  Else
    lblNomeProduto6.Caption = ""
  End If
End Sub

Private Sub cboProduto7_CloseUp()
  cboProduto7.Text = cboProduto7.Columns("Codigo").Text
  cboProduto7_LostFocus
End Sub

Private Sub cboProduto7_LostFocus()
  Call StatusMsg("")
  If cboProduto7.Text <> "" Then
    lblNomeProduto7.Caption = gsGetNameProduto(cboProduto7.Text)
  Else
    lblNomeProduto7.Caption = ""
  End If
End Sub

Private Sub cboProduto8_CloseUp()
  cboProduto8.Text = cboProduto8.Columns("Codigo").Text
  cboProduto8_LostFocus
End Sub

Private Sub cboProduto8_LostFocus()
  Call StatusMsg("")
  If cboProduto8.Text <> "" Then
    lblNomeProduto8.Caption = gsGetNameProduto(cboProduto8.Text)
  Else
    lblNomeProduto8.Caption = ""
  End If
End Sub
