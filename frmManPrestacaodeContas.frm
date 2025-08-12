VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManPrestacaodeContas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção da Prestação de Contas com o Fornecedor"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManPrestacaodeContas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   11535
   Begin VB.TextBox txtTotParaAcerto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0,00"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdGerarEntradaCompras 
      Caption         =   "Gerar Entrada p/ &Compras"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdGerarEntradaPrestacao 
      Caption         =   "Gerar Entrada p/ &Prestações"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox txtTotalVendido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0,00"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdGerarNF 
      Caption         =   "Gerar &NF para Devoluções"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfirmar 
      BackColor       =   &H0080C0FF&
      Caption         =   "Confirmar o &Acerto"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Acerto com todos os ítens da Grid."
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Frame fraPesquisa 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "Intervalo para Notas Fiscais"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   6050
         TabIndex        =   23
         Top             =   120
         Width           =   2500
         Begin VB.TextBox txtNFIni 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            MaxLength       =   8
            TabIndex        =   1
            Top             =   480
            Width           =   1005
         End
         Begin VB.TextBox txtNFFin 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   2
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   1200
            TabIndex        =   24
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.Frame fraProdutos 
         Appearance      =   0  'Flat
         Caption         =   " Produtos "
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   6050
         TabIndex        =   22
         Top             =   1080
         Width           =   2500
         Begin VB.OptionButton optNaoVendidos 
            Appearance      =   0  'Flat
            Caption         =   "Não Vendidos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1080
            TabIndex        =   6
            Top             =   250
            Width           =   1335
         End
         Begin VB.OptionButton optVendidos 
            Appearance      =   0  'Flat
            Caption         =   "Vendidos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   250
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Data datFornecedores 
         Caption         =   "datFornecedores"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Cli_FOR WHERE Tipo = 'F'"
         Top             =   720
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.TextBox txtNomeFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1250
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   4755
      End
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H0000C0C0&
         Caption         =   "Pesquisar"
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Frame fraPeriodo 
         Appearance      =   0  'Flat
         Caption         =   " Período ( Vendas ) "
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   8640
         TabIndex        =   13
         Top             =   120
         Width           =   2535
         Begin MSMask.MaskEdBox mskDataFinal 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDataInicial 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   1320
            TabIndex        =   15
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   375
         End
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmManPrestacaodeContas.frx":058A
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   975
         DataFieldList   =   "Nome"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   7805
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).DataField=   "Nome"
         Columns(0).FieldLen=   256
         Columns(1).Width=   3731
         Columns(1).Caption=   "Codigo"
         Columns(1).Name =   "Codigo"
         Columns(1).DataField=   "Código"
         Columns(1).FieldLen=   256
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label lblAcao 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   825
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdGeral 
      Height          =   3495
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   11295
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   16
      AllowDelete     =   -1  'True
      RowHeight       =   423
      Columns.Count   =   16
      Columns(0).Width=   503
      Columns(0).Caption=   "Fil"
      Columns(0).Name =   "Filial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1720
      Columns(1).Caption=   "Geração"
      Columns(1).Name =   "Data"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3122
      Columns(2).Caption=   "Fornecedor"
      Columns(2).Name =   "Fornecedor"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1508
      Columns(3).Caption=   "Nota"
      Columns(3).Name =   "Nota"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1111
      Columns(4).Caption=   "Seq"
      Columns(4).Name =   "Sequencia"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   900
      Columns(5).Caption=   "Linha"
      Columns(5).Name =   "Linha"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   1826
      Columns(6).Caption=   "Codigo"
      Columns(6).Name =   "Codigo"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   3731
      Columns(7).Caption=   "Nome"
      Columns(7).Name =   "Nome"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   873
      Columns(8).Caption=   "Qtde"
      Columns(8).Name =   "Qtde"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   1773
      Columns(9).Caption=   "Preco Custo"
      Columns(9).Name =   "Preco"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1270
      Columns(10).Caption=   "Vendido"
      Columns(10).Name=   "Vendido"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(10).Locked=   -1  'True
      Columns(11).Width=   1482
      Columns(11).Caption=   "Devolver"
      Columns(11).Name=   "Devolvido"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   1482
      Columns(12).Caption=   "Comprar"
      Columns(12).Name=   "Comprar"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   1296
      Columns(13).Caption=   "Estoque"
      Columns(13).Name=   "Estoque"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(13).Locked=   -1  'True
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "Acertados"
      Columns(14).Name=   "QtdeAcertada"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(14).Locked=   -1  'True
      Columns(14).HasForeColor=   -1  'True
      Columns(14).ForeColor=   255
      Columns(15).Width=   4842
      Columns(15).Caption=   "Decisão"
      Columns(15).Name=   "Resultado"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(15).Style=   3
      _ExtentX        =   19923
      _ExtentY        =   6165
      _StockProps     =   79
      Caption         =   "Ítens para Prestação de Contas com o Fornecedor"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTotParaAcerto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total p/ Acerto (R$)"
      Height          =   195
      Left            =   8415
      TabIndex        =   27
      Top             =   5460
      Width           =   1440
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Vendido (R$)"
      Height          =   195
      Left            =   5400
      TabIndex        =   20
      Top             =   5460
      Width           =   1335
   End
   Begin VB.Label lblMensagem 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   5460
      Width           =   5085
   End
End
Attribute VB_Name = "frmManPrestacaodeContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strCaption As String

Private Sub cmdConfirmar_Click()
  Dim intAuxi              As Integer
  Dim sngDiminuirDoEstoque As Single
  Dim sngX                 As Single
  Dim sngY                 As Single
  
  If grdGeral.Rows <= 0 Then
    MsgBox "Não há nenhum ítem carregado, verifique.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  '-------------------------------------------------------------------------------
  'Verificar se a coluna Resultado está vazia para alguém
  '-------------------------------------------------------------------------------
  If ValidarColunaResultado Then Exit Sub
  If ValidarLinhas Then Exit Sub

  If MsgBox("Deseja confirmar o Acerto? ", vbQuestion + vbYesNo, "Atenção") = vbYes Then
    grdGeral.MoveFirst
  
    For intAuxi = 0 To (grdGeral.Rows - 1)
    
      'Verificar o case baseado na coluna Resultado
      Select Case (grdGeral.Columns("Resultado").Text)
        Case "1 - Devolver"
          Call AtualizarPrestacaoContas(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text, grdGeral.Columns("Linha").Text, 1)
          Call VerificarConsignacoesDaEntrada(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text)
          'Caso a Devolução seja parcial não fecharemos a consignação, atualizaremos a prestação e
          'em seguida o estoque
          If CDbl(grdGeral.Columns("Devolvido").Value) < CDbl(grdGeral.Columns("Qtde").Value) Then
            Call AtualizarEstoqueFinal(gnCodFilial, Trim(grdGeral.Columns("Codigo").Value), CSng(grdGeral.Columns("Devolvido").Value))
          End If
        
        Case "2 - Comprar"
          Call AtualizarPrestacaoContas(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text, grdGeral.Columns("Linha").Text, 2)
          Call VerificarConsignacoesDaEntrada(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text)
          
          'Validação: Caso a Compra seja inferior a Qtde deveremos
          'atualizar o Estoque Final
          '23/11/2004 - Não passará mais por aqui...a decisão 4 tratará a situação de
          'forma mais coerente...
          If (CDbl(grdGeral.Columns("Qtde").Text) > CDbl(grdGeral.Columns("Comprar").Text)) Then
            'sngX = Qtde - Vendido
            'sngY = sngX - Comprado
            '
            'sngDiminuirDoEstoque = sngY
            
            sngX = CSng(grdGeral.Columns("Qtde").Text) - CSng(grdGeral.Columns("Comprar").Text)
            sngY = sngX
            
            sngDiminuirDoEstoque = sngY
            
            'Filial / CodProduto / Qtde que diminuíremos do Estoque
            Call AtualizarEstoqueFinal(CByte(grdGeral.Columns("Filial").Text), Trim(grdGeral.Columns("Codigo").Text), sngDiminuirDoEstoque)
          End If
        
        Case "3 - Prestar Contas"
          Call AtualizarPrestacaoContas(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text, grdGeral.Columns("Linha").Text, 3)
          Call VerificarConsignacoesDaEntrada(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text)
        
        Case "4 - Devolver e Comprar"
          Call AtualizarPrestacaoContas(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text, grdGeral.Columns("Linha").Text, 4)
          Call VerificarConsignacoesDaEntrada(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text)
          
        Case "5 - Devolver e Prestar Contas"
          Call AtualizarPrestacaoContas(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text, grdGeral.Columns("Linha").Text, 5)
          Call VerificarConsignacoesDaEntrada(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text)
          
        Case "6 - Comprar e Prestar Contas"
          Call AtualizarPrestacaoContas(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text, grdGeral.Columns("Linha").Text, 6)
          Call VerificarConsignacoesDaEntrada(grdGeral.Columns("Filial").Text, grdGeral.Columns("Sequencia").Text)
          
          'Validação: Caso a Compra seja inferior a Qtde deveremos
          'atualizar o Estoque Final
          If (CDbl(grdGeral.Columns("Vendido").Text) + CDbl(grdGeral.Columns("Comprar").Text)) < CDbl(grdGeral.Columns("Qtde").Text) Then
            'sngX = Qtde - Vendido
            'sngY = sngX - Comprado
            '
            'sngDiminuirDoEstoque = sngY
            
            sngX = CSng(grdGeral.Columns("Qtde").Text) - CSng(grdGeral.Columns("Vendido").Text)
            sngY = sngX - CSng(grdGeral.Columns("Comprar").Text)
            
            sngDiminuirDoEstoque = sngY
            
            'Filial / CodProduto / Qtde que diminuíremos do Estoque
            Call AtualizarEstoqueFinal(CByte(grdGeral.Columns("Filial").Text), Trim(grdGeral.Columns("Codigo").Text), sngDiminuirDoEstoque)
          End If
          
      End Select
    
      grdGeral.MoveNext
    Next intAuxi
    
  End If
  
  grdGeral.Redraw = False
  grdGeral.RemoveAll
  grdGeral.Refresh
  grdGeral.Redraw = True
  
  txtTotalVendido.Text = "0,00"
  txtTotParaAcerto.Text = "0,00"
  lblMensagem.Caption = ""

  MsgBox "Acerto finalizado com sucesso.", vbInformation, "Quick Store"

End Sub

Private Sub cmdGerarEntradaCompras_Click()
  strCaption = "Compras"
  frmImpressaoNFPrestacao.Show
End Sub

Private Sub cmdGerarEntradaPrestacao_Click()
  strCaption = "Prestação de Contas"
  frmImpressaoNFPrestacao.Show
End Sub

Private Sub cmdGerarNF_Click()
  frmImpressaoNFDevolucaoMateriais.Show
End Sub


Private Sub cmdPesquisar_Click()
  Dim rstPrestacaoContas As Recordset
  Dim strSQL             As String
  Dim dblVendido         As Double
  Dim sngEstoqueAtual    As Single
  Dim dblQtdeAcertada    As Double
  Dim dblTotParaAcertar  As Double
  
  If Len(txtNomeFornecedor.Text) <= 0 Then
    MsgBox "Selecione um Fornecedor.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If ValidarDatas Then Exit Sub
  
  Call StatusMsg("Verificando as quantidades vendidas...")
  Screen.MousePointer = vbHourglass
  '15/12/2004 - Daniel
  'Call AtualizarPrestacaoContasQtdeVendida 'Informamos a QtdeVendida Atualizada...
  Screen.MousePointer = vbDefault
  Call StatusMsg("")

  strSQL = "SELECT * FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  strSQL = strSQL & " AND Fornecedor = " & CLng(cboFornecedor.Text)
  strSQL = strSQL & " AND PeriodoVenda >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND PeriodoVenda <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND NOT Finalizado "
  
  '08/10/2004 - Daniel
  'Adicionado filtro de notas fiscais
  If Len(txtNFIni.Text) > 0 And Len(txtNFFin.Text) > 0 Then
    If CLng(txtNFIni.Text) <= CLng(txtNFFin.Text) Then
      strSQL = strSQL & " AND NotaFiscal >= " & CLng(txtNFIni.Text)
      strSQL = strSQL & " AND NotaFiscal <= " & CLng(txtNFFin.Text)
    End If
  End If
  
  '07/10/2004 - Daniel
  If optVendidos.Value Then
    strSQL = strSQL & " AND QtdeVendida <> 0 "
  Else
    strSQL = strSQL & " AND QtdeVendida = 0 "
  End If
  
  strSQL = strSQL & " ORDER BY Sequencia, Linha, DatadaGeracao "
  
  Set rstPrestacaoContas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  grdGeral.Redraw = False
  grdGeral.RemoveAll
  
  With rstPrestacaoContas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        grdGeral.AddNew
        grdGeral.Columns("Filial").Value = .Fields("Filial").Value
        grdGeral.Columns("Data").Value = .Fields("DatadaGeracao").Value
        grdGeral.Columns("Fornecedor").Value = .Fields("Fornecedor").Value & " - " & FindFornecedor(CLng(.Fields("Fornecedor").Value))
        grdGeral.Columns("Nota").Value = .Fields("NotaFiscal").Value
        grdGeral.Columns("Sequencia").Value = .Fields("Sequencia").Value
        grdGeral.Columns("Linha").Value = .Fields("Linha").Value
        grdGeral.Columns("Codigo").Value = .Fields("Produto").Value
        grdGeral.Columns("Nome").Value = FindProduto(CStr(.Fields("Produto").Value))
        grdGeral.Columns("Qtde").Value = .Fields("QtdeOriginal").Value
        grdGeral.Columns("Preco").Value = Format(.Fields("Custo").Value, FORMAT_VALUE)
        grdGeral.Columns("Vendido").Value = .Fields("QtdeVendida").Value
        grdGeral.Columns("Devolvido").Value = .Fields("QtdeDevolvida").Value
        grdGeral.Columns("Comprar").Value = .Fields("QtdeComprada").Value
        
        Call BuscarEstoqueAtual(.Fields("Filial").Value, .Fields("Produto").Value, sngEstoqueAtual)
        grdGeral.Columns("Estoque").Value = sngEstoqueAtual
        
        '15/12/2004 - Daniel
        'Call BuscarQtdeAcertada(.Fields("Filial").Value, .Fields("Sequencia").Value, .Fields("Linha").Value, dblQtdeAcertada)
        'grdGeral.Columns("QtdeAcertada").Value = dblQtdeAcertada
        'não irá acertar n vezes...
        grdGeral.Columns("QtdeAcertada").Value = 0
        
        'If optVendidos.Value Then grdGeral.Columns("Resultado").Value = "3 - Prestar Contas"
        
        grdGeral.Update
      
        If grdGeral.Columns("Vendido").Text <> "0" Then dblVendido = dblVendido + (CDbl(grdGeral.Columns("Preco").Text) * (CDbl(grdGeral.Columns("Vendido").Value)))
        If grdGeral.Columns("Vendido").Text <> "0" Then dblTotParaAcertar = dblTotParaAcertar + (CDbl(grdGeral.Columns("Preco").Text) * ((CDbl(grdGeral.Columns("Vendido").Value) - (CDbl(grdGeral.Columns("QtdeAcertada").Value)))))
        
      .MoveNext
      Loop
      
    End If
    
    txtTotalVendido.Text = Format(dblVendido, FORMAT_VALUE)
    txtTotParaAcerto.Text = Format(dblTotParaAcertar, FORMAT_VALUE)
    lblMensagem.Caption = "Pesquisa concluída, " & .RecordCount & " registros encontrados."
    .Close
  End With
  
  grdGeral.MoveFirst
  grdGeral.Redraw = True
  
  
  'Estava perdendo o nome do Fornecedor...
  cboFornecedor_LostFocus
  
  Set rstPrestacaoContas = Nothing

End Sub


Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFornecedores.DatabaseName = gsQuickDBFileName
  
'  mskDataInicial.Text = "01/01/" & Year(Data_Atual)
'  mskDataFinal.Text = CDate(Data_Atual + 30)

'  grdGeral.Columns("Devolvido").Visible = False
'  grdGeral.Columns("Comprar").Visible = False
  
  lblTotParaAcerto.Visible = True
  txtTotParaAcerto.Visible = True
  
  lblAcao.Caption = "  Prestação de Contas  "
  
  grdGeral.Columns("Resultado").AddItem "3 - Prestar Contas"
  'grdGeral.Columns("Resultado").AddItem "5 - Devolver e Prestar Contas"
  'grdGeral.Columns("Resultado").AddItem "6 - Comprar e Prestar Contas"

End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(1).Text
  txtNomeFornecedor.Text = FindFornecedor
End Sub

Private Sub cboFornecedor_LostFocus()
  txtNomeFornecedor.Text = FindFornecedor
End Sub

Private Function FindFornecedor(Optional lngFornecedor As Long)
  Dim strSQL    As String
  Dim rstForn   As Recordset
  Dim lng_Fornecedor As Long
  
  txtNomeFornecedor.Text = ""
  
  If lngFornecedor <= 0 Then
    If Len(Trim(cboFornecedor.Text)) <= 0 Then Exit Function
    If Not IsNumeric(Trim(cboFornecedor.Text)) Then Exit Function
    
    lng_Fornecedor = CLng(cboFornecedor.Text)
  Else
    lng_Fornecedor = lngFornecedor
  End If
  
  strSQL = " SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' AND "
  strSQL = strSQL & " Código = " & lng_Fornecedor
  
  Set rstForn = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstForn
    If Not (.BOF And .EOF) Then
      FindFornecedor = .Fields("Nome").Value & ""
    Else
      FindFornecedor = ""
    End If
    .Close
  End With
  
  Set rstForn = Nothing
End Function

Private Function FindProduto(strCodigo As String)
  Dim strSQL    As String
  Dim rstProd   As Recordset
  
  strSQL = " SELECT Nome FROM Produtos WHERE Código = '" & strCodigo & "'"
  
  Set rstProd = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstProd
    If Not (.BOF And .EOF) Then
      FindProduto = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  
  Set rstProd = Nothing
End Function

Private Sub grdGeral_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  If gbPodeApagar = False Then
    Beep
    Cancel = True
    Exit Sub
  End If
  
  If bGridBeforeDelete Then
    Call StatusMsg("Seleção de itens apagada.")
    Cancel = False
  Else
    Cancel = True
  End If

End Sub

Public Function bGridBeforeDelete() As Boolean
  Dim intI    As Integer
  Dim varBook As Variant
  
  gsTitle = LoadResString(201)
  gsMsg = "Apagar seleção atual?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  If gnResponse = vbNo Then
    bGridBeforeDelete = False
  Else
    bGridBeforeDelete = True
    
    For intI = 0 To (grdGeral.SelBookmarks.Count - 1)
      varBook = grdGeral.SelBookmarks(intI)
      grdGeral.Bookmark = varBook
    
      Call UpdateFieldSelecionadoEP(grdGeral.Columns("Filial").CellValue(varBook), grdGeral.Columns("Sequencia").CellValue(varBook), grdGeral.Columns("Linha").CellValue(varBook))
      Call DeletePrestacaoContas(grdGeral.Columns("Filial").CellValue(varBook), grdGeral.Columns("Sequencia").CellValue(varBook), grdGeral.Columns("Linha").CellValue(varBook))
      
    Next intI
      
    Call UpdateTxtTotalVendido
      
  End If
End Function

Private Sub UpdateFieldSelecionadoEP(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte)
  Dim rstEntraProdu As Recordset
  Dim strSQL        As String
  
  strSQL = "SELECT Selecionado, ConsignacaoFechada FROM [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequência = " & Sequencia
  strSQL = strSQL & " AND Linha = " & Linha

  Set rstEntraProdu = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntraProdu
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Edit
      '.Fields("ConsignacaoFechada").Value = False
      .Fields("Selecionado").Value = False
      .Update
    End If
    .Close
  End With
  
  Set rstEntraProdu = Nothing

End Sub

Private Sub DeletePrestacaoContas(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte)
  Dim rstPrestacao As Recordset
  Dim strSQL       As String
  
  strSQL = "DELETE * FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequencia = " & Sequencia
  strSQL = strSQL & " AND Linha = " & Linha

  db.Execute (strSQL)
  
End Sub

Private Sub UpdateTxtTotalVendido()
  Dim intAuxi           As Integer
  Dim dblVendido        As Double
  Dim dblTotParaAcertar As Double

  If grdGeral.Rows <= 0 Then Exit Sub

  For intAuxi = 0 To (grdGeral.Rows - 1)
    If grdGeral.Columns("Vendido").Text <> "0" Then dblVendido = dblVendido + (CDbl(grdGeral.Columns("Preco").Text) * (CDbl(grdGeral.Columns("Vendido").Value)))
    
    If optVendidos.Value Then
      If grdGeral.Columns("Vendido").Text <> "0" Then dblTotParaAcertar = dblTotParaAcertar + (CDbl(grdGeral.Columns("Preco").Text) * ((CDbl(grdGeral.Columns("Vendido").Value) - (CDbl(grdGeral.Columns("QtdeAcertada").Value)))))
    End If
  Next intAuxi
  
  txtTotalVendido.Text = Format(dblVendido, FORMAT_VALUE)
  
  If optVendidos.Value Then txtTotParaAcerto.Text = Format(dblTotParaAcertar, FORMAT_VALUE)

End Sub

Private Function ValidarColunaResultado() As Boolean
  Dim intAuxi As Integer
  
  Call StatusMsg("Validando a coluna Decisão...")
  
    grdGeral.MoveFirst
    
    For intAuxi = 0 To (grdGeral.Rows - 1)
      If Len(grdGeral.Columns("Resultado").Text) <= 0 Then ValidarColunaResultado = True
      
      If ValidarColunaResultado Then
        Call StatusMsg("")
        MsgBox "Coluna Decisão na linha " & intAuxi + 1 & " está inválida, verifique.", vbExclamation, "Atenção"
        Exit Function
      End If
      
      grdGeral.MoveNext
    Next intAuxi
  
  Call StatusMsg("")
  
End Function

Private Function ValidarLinhas() As Boolean
  Dim intAuxi          As Integer
  Dim intLinhaGrid     As Integer
  Dim dblQtdeBaixada   As Double
  
    grdGeral.MoveFirst
    
    For intAuxi = 0 To (grdGeral.Rows - 1)
    
      intLinhaGrid = intAuxi + 1
      
      'Validação geral
      'Vendido + Devolvido + Comprar > Qtde
      If (CDbl(grdGeral.Columns("Vendido").Text) + CDbl(grdGeral.Columns("Devolvido").Text) + CDbl(grdGeral.Columns("Comprar").Text)) > CDbl(grdGeral.Columns("Qtde").Text) Then
        ValidarLinhas = True
        MsgBox "A Soma de Vendido, Devolvido e Comprado é maior que a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
        Exit Function
      End If
    
      'Verificar o case baseado na coluna Resultado
      Select Case (grdGeral.Columns("Resultado").Text)
        Case "1 - Devolver"
          
          'Devolvido > Qtde
          If CDbl(grdGeral.Columns("Devolvido").Text) > CDbl(grdGeral.Columns("Qtde").Text) Then
            ValidarLinhas = True
            MsgBox "Devolução maior que a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          'Devolvido < Qtde
          'Foi Solicitado pela Resultado para devolver x unidades
          'Exemplo: Qtde 5 mas vou devolver 3 e ficar com 2
          'If CDbl(grdGeral.Columns("Devolvido").Text) < CDbl(grdGeral.Columns("Qtde").Text) Then
          '  ValidarLinhas = True
          '  MsgBox "Devolução menor que a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
          '  Exit Function
          'End If
          'Vendido + Devolvido + Comprar > Qtde
          If (CDbl(grdGeral.Columns("Vendido").Text) + CDbl(grdGeral.Columns("Devolvido").Text) + CDbl(grdGeral.Columns("Comprar").Text)) > CDbl(grdGeral.Columns("Qtde").Text) Then
            ValidarLinhas = True
            MsgBox "A Soma de Vendido, Devolvido e Comprado é maior que a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          'Devolução nula
          If CDbl(grdGeral.Columns("Devolvido").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Devolução ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
                    
          If CDbl(grdGeral.Columns("Vendido").Text) = 0 And CDbl(grdGeral.Columns("Devolvido").Text) = 0 And CDbl(grdGeral.Columns("Comprar").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Devolução ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
        
        Case "2 - Comprar"
        
          'Compra > Qtde
          If CDbl(grdGeral.Columns("Comprar").Text) > CDbl(grdGeral.Columns("Qtde").Text) Then
            ValidarLinhas = True
            MsgBox "Compra maior que a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          'Vendido + Devolvido + Comprar > Qtde
          If (CDbl(grdGeral.Columns("Vendido").Text) + CDbl(grdGeral.Columns("Devolvido").Text) + CDbl(grdGeral.Columns("Comprar").Text)) > CDbl(grdGeral.Columns("Qtde").Text) Then
            ValidarLinhas = True
            MsgBox "A Soma de Vendido, Devolvido e Comprado é maior que a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          'Compra nula
          If CDbl(grdGeral.Columns("Comprar").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Compra ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          'Compra menor que a Qtde deverá preencher a QtdeDevolvida
          'Compra < Qtde
          If CDbl(grdGeral.Columns("Comprar").Text) < CDbl(grdGeral.Columns("Qtde").Text) Then
            ValidarLinhas = True
            MsgBox "Compra menor que a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            MsgBox "Você deverá escolher a Decisão '4 - Devolver e Comprar', verifique", vbExclamation, "Atenção"
            Exit Function
          End If
        
        Case "3 - Prestar Contas"
          'Venda nula
          If CDbl(grdGeral.Columns("Vendido").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Venda ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          
          'Validação para não acertar duas vezes
          If CDbl(grdGeral.Columns("Vendido").Text) = CDbl(grdGeral.Columns("QtdeAcertada").Text) Then
            ValidarLinhas = True
            MsgBox "Vendido igual Acertados, na linha " & intLinhaGrid & ", verifique.", vbExclamation, "Acertar Duas vezes"
            Exit Function
          End If

        Case "4 - Devolver e Comprar"
          'Devolução nula
          If CDbl(grdGeral.Columns("Devolvido").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Devolução ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          
          'Compra nula
          If CDbl(grdGeral.Columns("Comprar").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Compra ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
        
          'A Soma da Devolução com a Compra deve dar a Qtde
          If (CDbl(grdGeral.Columns("Devolvido").Text) + CDbl(grdGeral.Columns("Comprar").Text)) <> CDbl(grdGeral.Columns("Qtde").Text) Then
            ValidarLinhas = True
            MsgBox "Soma da Devolução com a Compra menor que a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
        
        
        Case "5 - Devolver e Prestar Contas"
          'Devolução nula
          If CDbl(grdGeral.Columns("Devolvido").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Devolução ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
        
          'Venda nula
          If CDbl(grdGeral.Columns("Vendido").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Venda ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          
          'Vendido + Devolvido > Qtde
          If (CDbl(grdGeral.Columns("Vendido").Text) + CDbl(grdGeral.Columns("Devolvido").Text)) > CDbl(grdGeral.Columns("Qtde").Text) Then
            ValidarLinhas = True
            MsgBox "Qtde Vendida + Devolvida é Superior a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          
          'Vendido + Devolvido + Comprado < Qtde
          'If (CDbl(grdGeral.Columns("Vendido").Text) + CDbl(grdGeral.Columns("Devolvido").Text)) < CDbl(grdGeral.Columns("Qtde").Text) Then
          '  ValidarLinhas = True
          '  MsgBox "Qtde Devolvida é Inferior a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
          '  MsgBox "Você deverá Devolver o restante do que não vendeu integralmente.", vbExclamation, "Atenção"
          '  Exit Function
          'End If
          Call BuscarQtdeADevolver(CByte(grdGeral.Columns("Filial").Text), CLng(grdGeral.Columns("Sequencia").Text), CByte(grdGeral.Columns("Linha").Text), dblQtdeBaixada)
          If CDbl(grdGeral.Columns("Qtde").Text) - dblQtdeBaixada <> CDbl(grdGeral.Columns("Devolvido").Text) Then
            MsgBox "Qtde Devolvida deverá ser " & (CDbl(grdGeral.Columns("Qtde").Text) - dblQtdeBaixada) & " na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            MsgBox "Você deverá Devolver o restante do que não vendeu integralmente.", vbExclamation, "Atenção"
            Exit Function
          End If
        
        Case "6 - Comprar e Prestar Contas"
          'Compra nula
          If CDbl(grdGeral.Columns("Comprar").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Compra ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If

          'Venda nula
          If CDbl(grdGeral.Columns("Vendido").Text) = 0 Then
            ValidarLinhas = True
            MsgBox "Venda ZERADA na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          
          'Compra + Venda > Qtde
          If (CDbl(grdGeral.Columns("Vendido").Text) + CDbl(grdGeral.Columns("Comprar").Text)) > CDbl(grdGeral.Columns("Qtde").Text) Then
            ValidarLinhas = True
            MsgBox "Qtde Vendida + Comprada é Superior a Qtde na linha, " & intLinhaGrid & ", verifique.", vbExclamation, "Atenção"
            Exit Function
          End If
          
          
      End Select
    
      grdGeral.MoveNext
    Next intAuxi

End Function

Private Sub AtualizarPrestacaoContas(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte, ByVal Resultado As Byte)
  Dim rstEntraProdu    As Recordset
  Dim rstPrestacao     As Recordset
  Dim strSQL           As String
  Dim intAuxi          As Integer
  Dim blnFecharConsig  As Boolean
  Dim blnDevolverXUnid As Boolean

  '------------------------------------------------------------
  'Atualizar os campos em PrestacaoContas:
  ' PrestacaoContas.QtdeDevolvida
  ' PrestacaoContas.QtdeComprada
  ' PrestacaoContas.QtdeVendida
  ' PrestacaoContas.Finalizado para True
  ' PrestacaoContas.DatadaFinalizacao para Data_Atual
  ' PrestacaoContas.Resultado
  ' PrestacaoContas.QtdeAcertada
  '------------------------------------------------------------
  
  'Verificar se a devolução é menor que a Qtde deixando
  'assim x unidades no estoque
  If Resultado = 1 Then
    If CDbl(grdGeral.Columns("Devolvido").Value) < CDbl(grdGeral.Columns("Qtde").Value) Then blnDevolverXUnid = True
  End If
  
  strSQL = ""
  strSQL = "SELECT Filial, Sequencia, Linha, Finalizado, QtdeDevolvida, QtdeVendida, QtdeComprada, DatadaFinalizacao, Resultado, QtdeAcertada, QtdeOriginal, Custo FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequencia = " & Sequencia
  strSQL = strSQL & " AND Linha = " & Linha
  strSQL = strSQL & " AND NOT Finalizado "
  
  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Edit
      .Fields("QtdeDevolvida").Value = CDbl(grdGeral.Columns("Devolvido").Value)
      .Fields("QtdeVendida").Value = CDbl(grdGeral.Columns("Vendido").Value)
      .Fields("QtdeComprada").Value = CDbl(grdGeral.Columns("Comprar").Value)
      .Fields("Custo").Value = Format((CDbl(grdGeral.Columns("Preco").Value)), FORMAT_VALUE)
      .Fields("Finalizado").Value = True
      .Fields("DatadaFinalizacao").Value = Data_Atual
      .Fields("Resultado").Value = Resultado
      .Fields("QtdeAcertada").Value = CDbl(grdGeral.Columns("Vendido").Value) - CDbl(grdGeral.Columns("QtdeAcertada").Value)
      
      If blnDevolverXUnid Then .Fields("QtdeOriginal").Value = CDbl(grdGeral.Columns("Qtde").Value) 'menos CDbl(grdGeral.Columns("Devolvido").Value)
      
      'Acendemos o flag para fechar em [Entradas - Produtos] este registro
      If .Fields("QtdeOriginal").Value = .Fields("QtdeVendida").Value Then blnFecharConsig = True
      '18/11/2004 - Daniel
      'Verificar se Devolveu X unidades e depois houve venda
      'caso feche então blnFecharConsig será True
      If Resultado = 3 Then
        If Completou(.Fields("Filial").Value, .Fields("Sequencia").Value, .Fields("Linha").Value, .Fields("QtdeOriginal").Value) Then blnFecharConsig = True
      End If
      
      .Update
    End If
    .Close
  End With
  
  Set rstPrestacao = Nothing

  '-------------------------------------------------------------------
  ' Atualizar os campos em [Entradas - Produtos]:
  '
  ' [Entradas - Produtos].Acertado para True
  ' [Entradas - Produtos].ConsignacaoFechada para True
  '
  ' Onde na grid a Qtde seja Igual a QtdeVendida (Prestação de Contas)
  '-------------------------------------------------------------------
  
  If Resultado = 3 And blnFecharConsig Then 'Prestação de Contas
  
    'grdGeral.MoveFirst
  
    'For intAuxi = 0 To (grdGeral.Rows - 1)

      'If grdGeral.Columns("Qtde").Value = grdGeral.Columns("Vendido").Value Then
      
        strSQL = ""
        strSQL = "SELECT Acertado, ConsignacaoFechada FROM [Entradas - Produtos] "
        strSQL = strSQL & " WHERE Filial = " & Filial     'CByte(grdGeral.Columns("Filial").Value)
        strSQL = strSQL & " AND Sequência = " & Sequencia 'CLng(grdGeral.Columns("Sequencia").Value)
        strSQL = strSQL & " AND Linha = " & Linha         'CByte(grdGeral.Columns("Linha").Value)
        
        Set rstEntraProdu = db.OpenRecordset(strSQL, dbOpenDynaset)
        
        With rstEntraProdu
          If Not (.BOF And .EOF) Then
            .MoveFirst
            .Edit
            .Fields("Acertado").Value = True
            .Fields("ConsignacaoFechada").Value = True 'Pelo efetiva saída caso feche este campo já deverá estar True...
            .Update
          End If
          .Close
        End With
        
        Set rstEntraProdu = Nothing
      
      'End If
    
      'grdGeral.MoveNext
    'Next intAuxi
  
  Else 'Para as demais situações (1,2,4,5)
  
    'grdGeral.MoveFirst
  
    'For intAuxi = 0 To (grdGeral.Rows - 1)

      'If grdGeral.Columns("Resultado").Value <> "3 - Prestar Contas" Then
      
      If Not blnDevolverXUnid Then
      
        If Resultado = 3 Then Exit Sub '17/11/2004 - Daniel
      
        strSQL = ""
        strSQL = "SELECT Acertado, ConsignacaoFechada FROM [Entradas - Produtos] "
        strSQL = strSQL & " WHERE Filial = " & Filial     'CByte(grdGeral.Columns("Filial").Value)
        strSQL = strSQL & " AND Sequência = " & Sequencia 'CLng(grdGeral.Columns("Sequencia").Value)
        strSQL = strSQL & " AND Linha = " & Linha         'CByte(grdGeral.Columns("Linha").Value)
        
        Set rstEntraProdu = db.OpenRecordset(strSQL, dbOpenDynaset)
        
        With rstEntraProdu
          If Not (.BOF And .EOF) Then
            .MoveFirst
            .Edit
            .Fields("Acertado").Value = True
            .Fields("ConsignacaoFechada").Value = True 'Pelo efetiva saída caso feche este campo já deverá estar True...
            .Update
          End If
          .Close
        End With
        
        Set rstEntraProdu = Nothing
      
      End If
      
      'End If
    
      'grdGeral.MoveNext
    'Next intAuxi
  
  
  End If


End Sub

Private Sub VerificarConsignacoesDaEntrada(ByVal Filial As Byte, ByVal Sequencia As Long)
  Dim rstEntradas As Recordset
  Dim strSQL      As String
  Dim blnFlag     As Boolean
  
  strSQL = ""
  strSQL = "SELECT [Entradas - Produtos].ConsignacaoFechada "
  strSQL = strSQL & " FROM Entradas, [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Entradas.Filial = " & Filial
  strSQL = strSQL & " AND Entradas.Sequência = " & Sequencia
  strSQL = strSQL & " AND [Entradas - Produtos].Sequência = Entradas.Sequência "

  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        blnFlag = .Fields("ConsignacaoFechada").Value
        
        If Not blnFlag Then Exit Do
        
      .MoveNext
      Loop
    
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing

  If blnFlag Then
  
    strSQL = ""
    strSQL = "SELECT Entradas.ConsignacaoFechada "
    strSQL = strSQL & " FROM Entradas "
    strSQL = strSQL & " WHERE Entradas.Filial = " & Filial
    strSQL = strSQL & " AND Entradas.Sequência = " & Sequencia
  
    Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)

    With rstEntradas
      If Not (.BOF And .EOF) Then
        .MoveFirst
        .Edit
        .Fields("ConsignacaoFechada").Value = True
        .Update
      End If
      .Close
    End With
  
    Set rstEntradas = Nothing
  
  End If

End Sub

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinal.Text = frmCalendario.gsDateCalender(mskDataFinal.Text)
  End If
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal.Text = Ajusta_Data(mskDataFinal.Text)
End Sub

Private Sub mskDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicial.Text = frmCalendario.gsDateCalender(mskDataInicial.Text)
  End If
End Sub

Private Sub mskDataInicial_LostFocus()
  mskDataInicial.Text = Ajusta_Data(mskDataInicial.Text)
End Sub

Private Sub BuscarEstoqueAtual(ByVal Filial As Byte, ByVal CodProduto As String, ByRef EstoqueAtual As Single)
  Dim rstEstoqueFinal As Recordset
  Dim strSQL          As String

  strSQL = "SELECT [Estoque Atual] FROM [Estoque Final] "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Produto = '" & CodProduto & "'"
  
  Set rstEstoqueFinal = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEstoqueFinal
    If Not (.BOF And .EOF) Then
      .MoveFirst
      EstoqueAtual = .Fields("Estoque Atual").Value
    End If
    .Close
  End With
  
  Set rstEstoqueFinal = Nothing

End Sub

Private Function ValidarDatas() As Boolean

  If Not IsDate(mskDataInicial.Text) Then
    ValidarDatas = True
    MsgBox "Data Inicial inválida, verifique.", vbInformation, "Atenção"
    mskDataInicial.SetFocus
    Exit Function
  End If

  If Not IsDate(mskDataFinal.Text) Then
    ValidarDatas = True
    MsgBox "Data Final inválida, verifique.", vbInformation, "Atenção"
    mskDataFinal.SetFocus
    Exit Function
  End If

  If CDate(mskDataInicial.Text) > CDate(mskDataFinal.Text) Then
    ValidarDatas = True
    MsgBox "Data Final menor que a Inicial, verifique.", vbInformation, "Atenção"
    mskDataFinal.SetFocus
    Exit Function
  End If

End Function

Private Sub optNaoVendidos_Click()
'  grdGeral.Columns("Devolvido").Visible = True
'  grdGeral.Columns("Comprar").Visible = True
  
  lblTotParaAcerto.Visible = False
  txtTotParaAcerto.Visible = False
  
  lblAcao.Caption = "  Compra ou Devolução  "
  
  grdGeral.Columns("Resultado").RemoveAll
  grdGeral.Columns("Resultado").AddItem "1 - Devolver"
  grdGeral.Columns("Resultado").AddItem "2 - Comprar"
  grdGeral.Columns("Resultado").AddItem "4 - Devolver e Comprar"
  
End Sub

Private Sub optVendidos_Click()
'  grdGeral.Columns("Devolvido").Visible = False
'  grdGeral.Columns("Comprar").Visible = False
  
  lblTotParaAcerto.Visible = True
  txtTotParaAcerto.Visible = True
  
  lblAcao.Caption = "  Prestação de Contas  "
  
  grdGeral.Columns("Resultado").RemoveAll
  grdGeral.Columns("Resultado").AddItem "3 - Prestar Contas"
  grdGeral.Columns("Resultado").AddItem "5 - Devolver e Prestar Contas"
  'grdGeral.Columns("Resultado").AddItem "6 - Comprar e Prestar Contas"
  
End Sub

Private Sub AtualizarPrestacaoContasQtdeVendida()
  Dim rstPrestacaoContas As Recordset
  Dim strSQL             As String
  
  strSQL = "SELECT * FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  strSQL = strSQL & " AND Fornecedor = " & CLng(cboFornecedor.Text)
  strSQL = strSQL & " AND PeriodoVenda >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND PeriodoVenda <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND NOT Finalizado "
  
  '08/10/2004 - Daniel
  'Adicionado filtro de notas fiscais
  If Len(txtNFIni.Text) > 0 And Len(txtNFFin.Text) > 0 Then
    If CLng(txtNFIni.Text) <= CLng(txtNFFin.Text) Then
      strSQL = strSQL & " AND NotaFiscal >= " & CLng(txtNFIni.Text)
      strSQL = strSQL & " AND NotaFiscal <= " & CLng(txtNFFin.Text)
    End If
  End If
  
  '07/10/2004 - Daniel
  If optVendidos.Value Then
    strSQL = strSQL & " AND QtdeVendida <> 0 "
  Else
    strSQL = strSQL & " AND QtdeVendida = 0 "
  End If
  
  strSQL = strSQL & " ORDER BY Sequencia, Linha, DatadaGeracao "
  
  Set rstPrestacaoContas = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstPrestacaoContas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        .Edit
        .Fields("QtdeVendida").Value = getAcertosConsignacao(.Fields("Filial").Value, .Fields("Sequencia").Value, .Fields("Produto").Value)
        .Update
        
      .MoveNext
      Loop
      
    End If
    .Close
  End With

End Sub

Private Function getAcertosConsignacao(bytFilial As Byte, lngSequencia As Long, ByVal strProduto As String) As Double
  Dim strSQL        As String
  Dim rstAcerConsig As Recordset
  Dim dblVend       As Double
  
  strSQL = " SELECT * FROM AcertoConsignacaoEntrada WHERE "
  strSQL = strSQL & " Filial = " & bytFilial & " AND Sequencia = " & lngSequencia
  '22/09/2004 - Daniel - Adicionado AND abaixo
  strSQL = strSQL & " AND CodigoProduto = '" & strProduto & "'"
  
  '06/10/2004 - Daniel - Adicionado AND abaixo
  If optVendidos.Value Then
    strSQL = strSQL & " AND DataAcerto >= #" & Format(mskDataInicial.Text, "MM/DD/YYYY") & "#"
    strSQL = strSQL & " AND DataAcerto <= #" & Format(mskDataFinal.Text, "MM/DD/YYYY") & "#"
  End If
  
  dblVend = 0
  
  Set rstAcerConsig = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstAcerConsig
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do Until .EOF
      dblVend = dblVend + .Fields("QtdeVendida").Value
      .MoveNext
    Loop
    .Close
  End With
  
  getAcertosConsignacao = dblVend
  
  Set rstAcerConsig = Nothing
End Function

Private Sub BuscarQtdeAcertada(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte, ByRef QtdeAcertada As Double)
  Dim rstPrestacaoContas As Recordset
  Dim strQuery           As String

  QtdeAcertada = 0

  strQuery = "SELECT QtdeAcertada FROM PrestacaoContas "
  strQuery = strQuery & " WHERE Filial = " & Filial
  strQuery = strQuery & " AND Sequencia = " & Sequencia
  strQuery = strQuery & " AND Linha = " & Linha

  Set rstPrestacaoContas = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstPrestacaoContas.RecordCount = 0 Then
    rstPrestacaoContas.Close
    Set rstPrestacaoContas = Nothing
    Exit Sub
  End If

  With rstPrestacaoContas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        QtdeAcertada = QtdeAcertada + .Fields("QtdeAcertada").Value
        
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstPrestacaoContas = Nothing

End Sub

Private Sub AtualizarEstoqueFinal(ByVal Filial As Byte, ByVal Produto As String, ByVal DiminuirDoEstoque As Single)
  Dim rstEstoqueFinal As Recordset
  Dim strSQL          As String
  Dim sngEstoqueAtual As Single

  strSQL = "SELECT [Estoque Atual], [Última Data] FROM [Estoque Final] "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Produto = '" & Produto & "'"
  
  Set rstEstoqueFinal = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEstoqueFinal
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      sngEstoqueAtual = .Fields("Estoque Atual").Value
      
      .Edit
      .Fields("Estoque Atual").Value = sngEstoqueAtual - DiminuirDoEstoque
      .Fields("Última Data").Value = Data_Atual
      .Update
    End If
    .Close
  End With
  
  Set rstEstoqueFinal = Nothing

End Sub

Private Function Completou(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte, ByVal QtdeOriginal As Double) As Boolean
  Dim rstPrestacao As Recordset
  Dim strSQL       As String
  Dim dblSomas     As Double
  
  strSQL = "SELECT SUM(QtdeDevolvida) AS Devolvida, MAX(QtdeVendida) AS Vendida, SUM(QtdeComprada) AS Comprada "
  strSQL = strSQL & " FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequencia = " & Sequencia
  strSQL = strSQL & " AND Linha = " & Linha

  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      dblSomas = .Fields("Devolvida").Value + .Fields("Vendida").Value + .Fields("Comprada").Value
      
      If dblSomas = QtdeOriginal Then Completou = True
      
    End If
    .Close
  End With

  Set rstPrestacao = Nothing

End Function

Private Sub BuscarQtdeADevolver(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Linha As Byte, ByRef QtdeBaixada As Double)
  Dim rstPrestacao As Recordset
  Dim strSQL       As String
  
  strSQL = "SELECT SUM(QtdeDevolvida) AS Devolvida, MAX(QtdeVendida) AS Vendida, SUM(QtdeComprada) AS Comprada "
  strSQL = strSQL & " FROM PrestacaoContas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequencia = " & Sequencia
  strSQL = strSQL & " AND Linha = " & Linha

  Set rstPrestacao = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstPrestacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      QtdeBaixada = .Fields("Devolvida").Value + .Fields("Vendida").Value + .Fields("Comprada").Value
      
    End If
    .Close
  End With

  Set rstPrestacao = Nothing

End Sub
