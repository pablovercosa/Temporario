VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmResumoConsignacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo das movimentações de produtos consignados"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResumoConsignacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   11010
   Begin VB.Data datClientes 
      Appearance      =   0  'Flat
      Caption         =   "datClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'C'"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdFechar 
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
      Height          =   375
      Left            =   9720
      TabIndex        =   34
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtrato 
      Caption         =   "Extrato"
      Height          =   375
      Left            =   5400
      TabIndex        =   33
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdGerarSaida 
      Caption         =   "Gerar Saída"
      Height          =   375
      Left            =   4200
      TabIndex        =   32
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdGerarEntrada 
      Caption         =   "Gerar Entrada"
      Height          =   375
      Left            =   2880
      TabIndex        =   31
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdFechamento 
      Caption         =   "Fechamento"
      Height          =   375
      Left            =   1680
      TabIndex        =   30
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdContarProdutos 
      Caption         =   "Contar produtos"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0,00"
      Top             =   5115
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4680
      TabIndex        =   21
      Top             =   5055
      Width           =   3375
      Begin VB.TextBox txtTotalHojeDevolvido 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtTotalHojeSaldo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Devolvido"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Saldo"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Histórico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   5055
      Width           =   3855
      Begin VB.TextBox txtTotalHistoricoConsignado 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtTotalHistoricoDevolvido 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Consignado"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Devolvido"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   40
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.TextBox txtDataPrevisaoAcerto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtDataConsignacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtStatusConsignacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtCreditoDisponivel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtLimite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtConsignacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtClientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   3375
      End
      Begin SSDataWidgets_B.SSDBCombo cboClientes 
         Bindings        =   "frmResumoConsignacao.frx":058A
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         DataFieldList   =   "Código"
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
         Columns(0).Width=   3200
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin VB.Label Label7 
         Caption         =   "Previsão do acerto"
         Height          =   255
         Left            =   8160
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Data"
         Height          =   255
         Left            =   6120
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3720
         TabIndex        =   11
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Crédito Disponível"
         Height          =   255
         Left            =   8160
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Limite"
         Height          =   255
         Left            =   6120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Consignação"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdProdutos_ 
      Height          =   3975
      Left            =   0
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   10935
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   9
      UseGroups       =   -1  'True
      AllowAddNew     =   -1  'True
      RowHeight       =   423
      ExtraHeight     =   79
      Groups.Count    =   4
      Groups(0).Width =   7594
      Groups(0).Columns.Count=   3
      Groups(0).Columns(0).Width=   2064
      Groups(0).Columns(0).Caption=   "Codigo"
      Groups(0).Columns(0).Name=   "Codigo"
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   8
      Groups(0).Columns(0).FieldLen=   256
      Groups(0).Columns(1).Width=   3519
      Groups(0).Columns(1).Caption=   "Nome"
      Groups(0).Columns(1).Name=   "Nome"
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   256
      Groups(0).Columns(2).Width=   2011
      Groups(0).Columns(2).Caption=   "Sequencia"
      Groups(0).Columns(2).Name=   "Sequencia"
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   8
      Groups(0).Columns(2).FieldLen=   256
      Groups(1).Width =   3281
      Groups(1).Caption=   "Historico"
      Groups(1).Columns.Count=   2
      Groups(1).Columns(0).Width=   1773
      Groups(1).Columns(0).Caption=   "Consignado"
      Groups(1).Columns(0).Name=   "Hist_Consignado"
      Groups(1).Columns(0).Alignment=   2
      Groups(1).Columns(0).DataField=   "Column 3"
      Groups(1).Columns(0).DataType=   8
      Groups(1).Columns(0).FieldLen=   256
      Groups(1).Columns(1).Width=   1508
      Groups(1).Columns(1).Caption=   "Devolvido"
      Groups(1).Columns(1).Name=   "Hist_Devolvido"
      Groups(1).Columns(1).Alignment=   2
      Groups(1).Columns(1).DataField=   "Column 4"
      Groups(1).Columns(1).DataType=   8
      Groups(1).Columns(1).FieldLen=   256
      Groups(2).Width =   2990
      Groups(2).Caption=   "Hoje"
      Groups(2).Columns.Count=   2
      Groups(2).Columns(0).Width=   1482
      Groups(2).Columns(0).Caption=   "Devolvido"
      Groups(2).Columns(0).Name=   "Devolvido"
      Groups(2).Columns(0).Alignment=   2
      Groups(2).Columns(0).DataField=   "Column 5"
      Groups(2).Columns(0).DataType=   8
      Groups(2).Columns(0).FieldLen=   256
      Groups(2).Columns(1).Width=   1508
      Groups(2).Columns(1).Caption=   "Saldo"
      Groups(2).Columns(1).Name=   "Saldo"
      Groups(2).Columns(1).Alignment=   2
      Groups(2).Columns(1).DataField=   "Column 6"
      Groups(2).Columns(1).DataType=   8
      Groups(2).Columns(1).FieldLen=   256
      Groups(3).Width =   4339
      Groups(3).Caption=   "R$"
      Groups(3).Columns.Count=   2
      Groups(3).Columns(0).Width=   1931
      Groups(3).Columns(0).Caption=   "Unitário"
      Groups(3).Columns(0).Name=   "Unitario"
      Groups(3).Columns(0).Alignment=   2
      Groups(3).Columns(0).DataField=   "Column 7"
      Groups(3).Columns(0).DataType=   8
      Groups(3).Columns(0).FieldLen=   256
      Groups(3).Columns(1).Width=   2408
      Groups(3).Columns(1).Caption=   "Total"
      Groups(3).Columns(1).Name=   "Total"
      Groups(3).Columns(1).Alignment=   2
      Groups(3).Columns(1).DataField=   "Column 8"
      Groups(3).Columns(1).DataType=   8
      Groups(3).Columns(1).FieldLen=   256
      _ExtentX        =   19288
      _ExtentY        =   7011
      _StockProps     =   79
      Caption         =   "Produtos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   10920
      X2              =   45
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label13 
      Caption         =   "R$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   27
      Top             =   5220
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Totais"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5290
      Width           =   615
   End
End
Attribute VB_Name = "frmResumoConsignacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'  Dim rsClientes As Recordset
'
'Private Sub cboClientes_CloseUp()
'  If IsNumeric(cboClientes.Columns(0).Text) Then
'    cboClientes.Text = cboClientes.Columns(0).Text
'  End If
'End Sub
'
'Private Sub cboClientes_LostFocus()
'  txtClientes.Text = ""
'
'  If Not IsNumeric(cboClientes.Text) Then Exit Sub
'  Set rsClientes = db.OpenRecordset("SELECT Código, Nome, DataProxAcertoConsignacao, UltimaConsignacao, [Limite Crédito] FROM Cli_For WHERE Tipo = 'C' AND Código = " & cboClientes.Text, dbOpenSnapshot)
'
'  With rsClientes
'    If Not (.BOF And .EOF) Then
'      txtClientes.Text = .Fields("Nome") & ""
'      GetDataHeader
'    End If
'
'    If Not rsClientes Is Nothing Then rsClientes.Close
'  End With
'
'  Set rsClientes = Nothing
'End Sub
'
'Private Sub cboMovimentacoes_InitColumnProps()
'
'End Sub
'
'Private Sub cmdContarProdutos_Click()
'  frmContagemProdutos.Show
'End Sub
'
'Private Sub cmdExtrato_Click()
'  Dim lngBuffer As Long
'
'  Open "LPT1" For Output As #lngBuffer
'
'
'
'  Close #lngBuffer
'End Sub
'
'Private Sub cmdFechar_Click()
'  Unload Me
'End Sub
'
'Private Sub cmdGerarSaida_Click()
''  Dim intX As Integer
''
''  With frmSaidas
''    .Show
''
''    .Combo_Caixa = 1
''    .Combo_Operador = 1
''    .Combo_Preço = "TABELA1"
''    .cboCliente.Text = cboClientes.Text
''  End With
''
''  With frmSaidas.Grade1
''    .Redraw = False
''    .MoveFirst
''
''    For intX = 0 To .Rows - 1
''      .Columns("Código").Text = grdProdutos.Columns("Codigo").Text
''      .Columns("Qtde").Text = grdProdutos.Columns("Qtde").Text
''
''      .MoveNext
''    Next intX
''
''    .Redraw = True
''  End With
''
''  With rsClientes
''    .Edit
''
''    .Fields("UltimaConsignacao") = gnGetNextSequencia(gnCodFilial)
''    .Fields("DataProximoAcerto") = txtDataPrevisaoAcerto.Text & "/" & Month(Data_Atual) + 1 & "/" & Year(Data_Atual)
''
''    .Update
''  End With
'End Sub
'
'Private Sub Form_Load()
'  Call CenterForm(Me)
'  datClientes.DatabaseName = gsQuickDBFileName
'End Sub
'
'Private Sub GetDataHeader()
'  Dim rsSaidas          As Recordset
'  Dim rsSaidasProdutos  As Recordset
'  Dim rsCR              As Recordset
'
'  Dim dblQtde           As Double
'  Dim dblQtdeEntregue   As Double
'
'  With rsClientes
'    txtLimite.Text = Format(.Fields("Limite Crédito").Value & "", FORMAT_VALUE)
'    txtDataPrevisaoAcerto.Text = .Fields("DataProxAcertoConsignacao") & ""
'    txtConsignacao.Text = .Fields("UltimaConsignacao") & ""
'  End With
'
'  If (Not IsNumeric(txtConsignacao.Text)) Or (txtConsignacao.Text = "0") Then
'    MsgBox "Sem movimentação a exibir !", vbCritical, "Quick Store"
'    Exit Sub
'  End If
'
'  Set rsCR = db.OpenRecordset("SELECT SUM(Valor) AS SOMA FROM [Contas a Receber] WHERE Valor > [Valor Recebido]", dbOpenSnapshot)
'
'  txtCreditoDisponivel.Text = Format((rsClientes.Fields("Limite Crédito").Value - rsCR.Fields("Soma").Value), FORMAT_VALUE)
'
'  Set rsSaidas = db.OpenRecordset(" SELECT Data FROM Saídas WHERE Filial = " & gnCodFilial & _
'                                  " Sequência = " & txtConsignacao.Text, dbOpenSnapshot)
'
'  With rsSaidas
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'      txtDataConsignacao.Text = rsSaidas.Fields("Data") & ""
'    End If
'  End With
'
'  Set rsSaidasProdutos = db.OpenRecordset(" SELECT [Saídas - Produtos].Filial, [Saídas - Produtos].Sequência, [Saídas - Produtos].Código, [Saídas - Produtos].Qtde, [Saídas - Produtos].QtdeEntregue, Produtos.Nome " & _
'                                          " FROM [Saídas - Produtos] INNER JOIN Produtos ON [Saídas - Produtos].Código = Produtos.Código " & _
'                                          " WHERE Filial = " & gnCodFilial & _
'                                          " AND   Sequência = " & txtConsignacao.Text, dbOpenSnapshot)
'
'  With rsSaidasProdutos
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'
'      grdProdutos.Redraw = False
'
'      Do While Not .EOF
'        dblQtde = IIf(IsNumeric(.Fields("Qtde")), .Fields("Qtde"), 0)
'        dblQtdeEntregue = IIf(IsNumeric(.Fields("QtdeEntregue")), .Fields("QtdeEntregue"), 0)
'
'        grdProdutos.AddItem .Fields("Código") & vbTab & _
'                            .Fields("Nome") & vbTab & _
'                            .Fields("Qtde") & vbTab & _
'                            dblQtde - dblQtdeEntregue
'
'        .MoveNext
'      Loop
'
'      grdProdutos.Redraw = True
'    End If
'  End With
'End Sub
