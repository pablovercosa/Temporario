VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmLeTransfe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leitura do Arquivo de Transfer�ncia entre Filiais"
   ClientHeight    =   6525
   ClientLeft      =   945
   ClientTop       =   1410
   ClientWidth     =   11070
   Icon            =   "LeTransfe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11070
   Begin VB.Frame fraX 
      Caption         =   "Para preencher os registros transferidos de maneira adequada, informe a seguir a Opera��o de Entrada e a Empresa de Origem  "
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   10815
      Begin VB.TextBox txtTabelaPrecos 
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
         Height          =   315
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   720
         Width           =   2055
      End
      Begin VB.Data datOperEntrada 
         Caption         =   "datOperEntrada"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT C�digo, Nome FROM [Opera��es Entrada] ORDER BY C�digo"
         Top             =   240
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Data datFornecedor 
         Caption         =   "datFornecedor"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT C�digo, Nome, Tipo FROM Cli_For WHERE Tipo = 'F' ORDER BY C�digo"
         Top             =   720
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.TextBox txtFornecedor 
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
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   885
         Width           =   4575
      End
      Begin VB.TextBox txtOperEntrada 
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
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   4575
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "LeTransfe.frx":058A
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   885
         Width           =   885
         DataFieldList   =   "C�digo"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   1561
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   65535
         DataFieldToDisplay=   "C�digo"
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperEntrada 
         Bindings        =   "LeTransfe.frx":05A6
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   885
         DataFieldList   =   "C�digo"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   1561
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   65535
         DataFieldToDisplay=   "C�digo"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tabela de Pre�os"
         Height          =   195
         Left            =   8400
         TabIndex        =   15
         Top             =   420
         Width           =   1260
      End
      Begin VB.Label lblOpEntrada 
         AutoSize        =   -1  'True
         Caption         =   "Opera��o Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   1305
      End
      Begin VB.Label lblFornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Empresa de Origem"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   945
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton B_Procurar 
         Caption         =   "&Procurar..."
         Height          =   400
         Left            =   9360
         TabIndex        =   0
         Top             =   197
         Width           =   1335
      End
      Begin VB.TextBox Nome_Arq 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1635
         TabIndex        =   1
         Top             =   240
         Width           =   7590
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Arquivo :"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Fechar"
      Height          =   400
      Left            =   9645
      TabIndex        =   7
      Top             =   3555
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   9600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.TSC"
   End
   Begin VB.CommandButton B_Recebe 
      Caption         =   "&Receber"
      Enabled         =   0   'False
      Height          =   400
      Left            =   9645
      TabIndex        =   5
      ToolTipText     =   "Receber Produtos"
      Top             =   3060
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Height          =   3900
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   9465
      _Version        =   196617
      DataMode        =   2
      Rows            =   500
      Col.Count       =   3
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "C�digo"
      Columns(0).Name =   "C�digo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   9260
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Qtde"
      Columns(2).Name =   "Qtde"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   16695
      _ExtentY        =   6879
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
   Begin VB.CommandButton B_Le 
      Caption         =   "&Ler"
      Height          =   400
      Left            =   9645
      TabIndex        =   4
      Top             =   2550
      Width           =   1335
   End
End
Attribute VB_Name = "frmLeTransfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsProdutos    As Recordset
Dim rsCores       As Recordset
Dim rsTamanhos    As Recordset
Dim rsEstoque     As Recordset
'30/03/2004 - Daniel
'Atualizador da Ultima Movimenta��o de Par�metros
Dim m_nSequencia    As Long
'Var de controle para a Sub CriarRegistros
'Case: Casagrande
Dim m_intContador   As Integer

'24/10/2005 - mpdea
'Modificado o campo separador do grid
Private Sub B_Le_Click()
  Dim Texto As String
  Dim Pos As Integer
  Dim Produto As String
  Dim Qtde As Single
  Dim Tamanho As Integer
  Dim C�digo As String
  Dim Cor As Integer
  Dim Edi��o As Long
  Dim Tipo As Integer
  Dim Erro As Integer
  Dim Texto2 As String
  Dim Resp As Integer
  
  '25/10/2005 - mpdea
  'Prefixo da configura��o da Tabela de Pre�os
  Const PREF_TABELA As String = "TABELA:"
  

  If IsNull(Nome_Arq.Text) Then Nome_Arq.Text = ""
  If Nome_Arq.Text = "" Then
     DisplayMsg "Digite o nome do arquivo ou pressione Procurar."
     Exit Sub
  End If
  
  
  On Error GoTo Erro_Arq
  Open Nome_Arq.Text For Input As #1
  On Error GoTo 0
  
  rsProdutos.Index = "C�digo"
  rsTamanhos.Index = "C�digo"
  rsCores.Index = "C�digo"
  
  
  Input #1, Texto
  If Left(Texto, 20) <> "***TRANSFEREQUICK***" Then GoTo Arquivo_Inv
  
  
  '25/10/2005 - mpdea
  'L� a tabela de pre�os
  Input #1, Texto
  If Left(Texto, Len(PREF_TABELA)) = PREF_TABELA Then
    txtTabelaPrecos.Text = Right(Texto, Len(Texto) - Len(PREF_TABELA))
    'Verifica tabela de pre�os
    If Not gbCheckTabPreco(txtTabelaPrecos.Text) Then
      Close #1
      MsgBox "Tabela de Pre�os [" & txtTabelaPrecos.Text & "] n�o existe.", vbExclamation, "Aviso"
      Exit Sub
    End If
  Else
    Close #1
    MsgBox "Informa��o da Tabela de Pre�os n�o localizada no local padr�o.", vbExclamation, "Aviso"
    Exit Sub
  End If
  
  
  Do Until EOF(1)
    Input #1, Texto
    If Left(Texto, 15) = "***FIMTRANSFERE" Then GoTo Fim_Transfere
    Pos = InStr(1, Texto, "#")
    If Pos = 0 Then GoTo Arquivo_Inv
    Tamanho = Len(Texto)
    Produto = Left(Texto, (Pos - 1))
    Qtde = Val(Right(Texto, (Tamanho - Pos)))
    
    Acha_Produto Produto, C�digo, Tamanho, Cor, Edi��o, Tipo, Erro
    If Erro <> 0 Then
      Texto2 = "N�o foi poss�vel encontrar o produto : "
      Texto2 = Texto2 + Produto
      Texto2 = Texto2 & vbCrLf & "Deseja continuar o processo ? "
      Resp = MsgBox(Texto2, vbYesNo + vbQuestion)
      If Resp = vbNo Then
        Grade1.RemoveAll
        GoTo Fim_Transfere
      End If
    End If
    
    If Erro = 0 Then
      rsProdutos.Seek "=", C�digo
      Texto2 = Produto + vbTab + rsProdutos("Nome")
      
      If rsProdutos("Tipo") = "G" Then
         rsTamanhos.Seek "=", Tamanho
         If Not rsTamanhos.NoMatch Then Texto2 = Texto2 + " " + rsTamanhos("Nome")
         rsCores.Seek "=", Cor
         If Not rsCores.NoMatch Then Texto2 = Texto2 + " " + rsCores("Nome")
      End If
      
      Texto2 = Texto2 + vbTab
      Texto2 = Texto2 + str(Qtde)
      Grade1.AddItem Texto2
    End If
  Loop
  
Fim_Transfere:
  Close #1
  B_Recebe.Enabled = True
  Exit Sub
 
Arquivo_Inv:
  '25/10/2005 - mpdea
  'Fecha o arquivo aberto
  Close #1

  DisplayMsg "Este arquivo n�o � um arquivo de transfer�ncia deste aplicativo."
  Exit Sub
 
Erro_Arq:
  DisplayMsg "N�o foi poss�vel abrir o arquivo. Verifique o nome e tente novamente."

End Sub

Private Sub B_Procurar_Click()
 
  On Error Resume Next
  
  Call StatusMsg("")
  
  With Cmd1
    .CancelError = True
    .DialogTitle = "Escolha o arquivo de transfer�ncia"
    .DefaultExt = "TSC"
    .InitDir = gsDefaultPath
    .Filter = "Arquivo de Transfer�ncia | *.TSC"
    .Flags = cdlOFNPathMustExist & cdlOFNHideReadOnly
    .ShowOpen
  End With
  
  If Err.Number = 0 Then
    Nome_Arq = Cmd1.FileName
  Else
    Nome_Arq = ""
  End If
  
  On Error GoTo 0
  
  If Cmd1.FileName = "*.TSC" Then Exit Sub
 
End Sub

Private Sub B_Recebe_Click()
  Dim i As Integer
  Dim J As Integer
  Dim Criar_Registro As Integer
  Dim Estoque_Final As Single
  Dim Produto As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edi��o As Long
  Dim Aux_Str As String
  Dim Tipo As Integer
  Dim Erro As Integer
  Dim bRecOK As Boolean
  
  '21/10/2005 - mpdea
  'C�digo completo do produto
  Dim strCodigoCompleto As String
  'Controle de transa��o
  Dim blnInTransaction As Boolean
  
  
  '21/10/2005 - mpdea
  'Inclu�do tratamento de erro
  On Error GoTo ErrHandler
  

  '29/03/2004 - Daniel
  'Valida��o dos campos Opera��o de Entrada e Fornecedor
  If Len(txtOperEntrada.Text) <= 0 Then
    MsgBox "Opera��o de Entrada inv�lida.", vbExclamation, "Quick Sore"
    cboOperEntrada.SetFocus
    Exit Sub
  End If
  
  If Len(txtFornecedor.Text) <= 0 Then
    MsgBox "Empresa de Origem inv�lida.", vbExclamation, "Quick Sore"
    cboFornecedor.SetFocus
    Exit Sub
  End If
  '---------------------------------------------------------------------
  
  
  '21/10/2005 - mpdea
  'In�cio de transa��o
  ws.BeginTrans: blnInTransaction = True
  

  Grade1.MoveFirst
  For i = 0 To Grade1.Rows - 1
    If Grade1.Columns(0).Text <> "" And Val(Grade1.Columns(2).Text) <> 0 Then
      
      '21/10/2005 - mpdea
      'C�digo completo do produto
      strCodigoCompleto = Grade1.Columns(0).Text
    
      Produto = ""
      Tamanho = 0
      Cor = 0
      Edi��o = 0
      
      Call Acha_Produto(strCodigoCompleto, Produto, Tamanho, Cor, Edi��o, Tipo, Erro)
      If Erro <> 0 Then
        '21/10/2005 - mpdea
        'Desfaz transa��o
        If blnInTransaction Then ws.Rollback
        'Exibe mensagem
        DisplayMsg "Produto " & strCodigoCompleto & " n�o encontrado. Transfer�ncia cancelada."
        Exit Sub
      End If
      
      bRecOK = True
      rsProdutos.Seek "=", Produto
      'Ajusta estoque de ENTRADA
      'Encontra a posi��o do estoque
      Criar_Registro = False
      Estoque_Final = 0
      rsEstoque.Index = "Produto"
      rsEstoque.Seek "=", gnCodFilial, Data_Atual, rsProdutos("C�digo"), Tamanho, Cor, Edi��o

      If Not rsEstoque.NoMatch Then
        Estoque_Final = rsEstoque("Estoque Final")
      End If
      If rsEstoque.NoMatch Then
        rsEstoque.Index = "Data"
        rsEstoque.Seek "<", gnCodFilial, Produto, Tamanho, Cor, Edi��o, Data_Atual
        If rsEstoque.NoMatch Then
          Criar_Registro = True
        End If
        If Not rsEstoque.NoMatch Then
          If rsEstoque("Filial") = gnCodFilial And rsEstoque("Produto") = Produto And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edi��o") = Edi��o Then
            Criar_Registro = True
            Estoque_Final = rsEstoque("Estoque Final")
          Else
            Criar_Registro = True
            Estoque_Final = 0
          End If
        End If

        If Criar_Registro = True Then
          rsEstoque.AddNew
           rsEstoque("Filial") = gnCodFilial
           rsEstoque("Data") = Data_Atual
           rsEstoque("Produto") = Produto
           rsEstoque("Tamanho") = Tamanho
           rsEstoque("Cor") = Cor
           rsEstoque("Edi��o") = Edi��o
           rsEstoque("Classe") = rsProdutos("Classe")
           rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
           rsEstoque("Estoque Anterior") = Estoque_Final
          rsEstoque.Update
        End If

        rsEstoque.Index = "Produto"
        rsEstoque.Seek "=", gnCodFilial, Data_Atual, Produto, Tamanho, Cor, Edi��o
      End If

      'neste ponto esta com o registro de estoque
      'no buffer, agora soma com os valores da movimenta��o
      rsEstoque.Edit
      rsEstoque("Transf Entra") = rsEstoque("Transf Entra") + Val(Grade1.Columns(2).Text)
      Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
      Estoque_Final = Estoque_Final - rsEstoque("Transf Sa�da") + rsEstoque("Transf Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Ajuste Sa�da") + rsEstoque("Ajuste Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Gr�tis Sa�da") + rsEstoque("Gr�tis Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Empre Sa�da") + rsEstoque("Empre Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolu��o")
      If rsProdutos("Estoque") = False Then
        Estoque_Final = 0
      End If
      rsEstoque("Estoque Final") = Estoque_Final
      rsEstoque.Update
      'Arruma Estoque Final
      Grava_Estoque_Final gnCodFilial, Produto, Tamanho, Cor, Edi��o, Estoque_Final, CDate(Data_Atual)
    End If
    
    '21/10/2005 - mpdea
    'Corrigido inser��o do c�digo completo na cria��o dos registros
    '
    '25/03/2004 - Daniel
    'Case: Casagrande
    'Implementa��o para cria��o de registros nas
    'seguintes Tabelas: Entradas e [Entradas - Produtos], al�m de Atualizar a
    '�ltima movimenta��o em Par�metros Filial
    Call CriarRegistros(strCodigoCompleto, Produto, CSng(Grade1.Columns(2).Text))
    '------------------------------------------------------------------------------
    
    Grade1.MoveNext 'Agora podemos dar o MoveNext
    
  Next i
  
  '30/03/2004 - Daniel
  'Case: Casagrande
  'Zerando os contadores
  m_nSequencia = 0
  m_intContador = 0
  '------------------------

  
  '21/10/2005 - mpdea
  'Fim de transa��o
  ws.CommitTrans: blnInTransaction = False


  If bRecOK Then
    DisplayMsg "Transfer�ncia efetuada com Sucesso."
  End If
  B_Recebe.Enabled = False
  
  Exit Sub
  
ErrHandler:
  '21/10/2005 - mpdea
  'Desfaz transa��o
  If blnInTransaction Then ws.Rollback
  'Exibe mensagem de erro
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(0).Text
  cboFornecedor_LostFocus
End Sub

'05/12/2005 - mpdea
'Inclu�do tratamento de erro
Private Sub cboFornecedor_LostFocus()
  Dim rstFornecedor As Recordset
  Dim lngRet As Long

  On Error GoTo ErrHandler
  
  txtFornecedor.Text = ""
  
  If Not IsNumeric(cboFornecedor.Text) Then Exit Sub
  
  '05/12/2005 - mpdea
  'Tratamento de overflow
  Call IsDataType(dtLong, cboFornecedor.Text, lngRet)
  
  Set rstFornecedor = db.OpenRecordset("SELECT C�digo, Nome, Tipo FROM Cli_For WHERE C�digo = " & lngRet & " AND Tipo = 'F' ORDER BY C�digo", dbOpenDynaset, dbReadOnly)

  With rstFornecedor
    If Not (.BOF And .EOF) Then
      txtFornecedor.Text = .Fields("Nome") & ""
    End If
  End With

  rstFornecedor.Close
  Set rstFornecedor = Nothing
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboOperEntrada_CloseUp()
  cboOperEntrada.Text = cboOperEntrada.Columns(0).Text
  cboOperEntrada_LostFocus
End Sub

'05/12/2005 - mpdea
'Inclu�do tratamento de erro
Private Sub cboOperEntrada_LostFocus()
  Dim rstOperEntrada As Recordset
  Dim intRet As Integer
  
  On Error GoTo ErrHandler
  
  txtOperEntrada.Text = ""
  
  If Not IsNumeric(cboOperEntrada.Text) Then Exit Sub
  
  '05/12/2005 - mpdea
  'Tratamento de overflow
  Call IsDataType(dtInteger, cboOperEntrada.Text, intRet)
  
  Set rstOperEntrada = db.OpenRecordset("SELECT C�digo, Nome FROM [Opera��es Entrada] WHERE C�digo = " & intRet & " ORDER BY C�digo ", dbOpenDynaset, dbReadOnly)

  With rstOperEntrada
    If Not (.BOF And .EOF) Then
      txtOperEntrada.Text = .Fields("Nome") & ""
    End If
  End With

  rstOperEntrada.Close
  Set rstOperEntrada = Nothing

  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()

 Call CenterForm(Me)

 '29/03/2004 - Daniel
 datOperEntrada.DatabaseName = gsQuickDBFileName
 datFornecedor.DatabaseName = gsQuickDBFileName
 '----------------------------------------------

 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
 Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
 Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
 Set rsEstoque = db.OpenRecordset("Estoque")
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsProdutos.Close
  rsCores.Close
  rsTamanhos.Close
  rsEstoque.Close
  Set rsProdutos = Nothing
  Set rsCores = Nothing
  Set rsTamanhos = Nothing
  Set rsEstoque = Nothing
End Sub

'21/10/2005 - mpdea
'Corrigido inser��o do c�digo completo na cria��o dos registros
Private Sub CriarRegistros(ByVal strCodigoCompleto As String, ByVal strProduto As String, ByVal sngQtde As Single)
  Dim rstEntradas         As Recordset
  Dim rstEntradasProdutos As Recordset
  Dim rstParametros       As Recordset
  Dim strObs              As String
    
  m_intContador = m_intContador + 1
  
  If m_intContador = 1 Then

    m_nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("�ltima Movimenta��o") + 1

    'Entradas
    Set rstEntradas = db.OpenRecordset("SELECT * FROM Entradas", dbOpenDynaset)
    
    With rstEntradas
      .AddNew
      
      .Fields("Filial").Value = gnCodFilial
      .Fields("Data").Value = Data_Atual
      .Fields("Sequ�ncia").Value = m_nSequencia
      .Fields("Opera��o").Value = CInt(cboOperEntrada.Text)
      .Fields("Fornecedor").Value = CLng(cboFornecedor.Text)
      .Fields("Digitador").Value = gnUserCode
      '26/01/2005 - Daniel
      'Tratamento para o campo Observa��es
      'Para alguns clientes ocorria o Erro: 3163
      strObs = "Importado da Empresa " & (cboFornecedor.Text) & " - " & (txtFornecedor.Text) & " em " & Data_Atual
      If Len(strObs) <= 70 Then
        .Fields("Observa��es").Value = strObs & ""
      Else
        .Fields("Observa��es").Value = Left(strObs, 70) & ""
      End If
      '27/04/2004 - Daniel
      'Adicionado o field Data Emiss�o
      .Fields("Data Emiss�o").Value = Data_Atual
      .Fields("Produtos").Value = 0
      .Fields("Caixa").Value = 1
      .Fields("Efetivada").Value = True
        
      .Update
      .Close
    End With

    Set rstEntradas = Nothing
    'Fim Entradas
  
    'Abrindo o Par�metros Filial
    Set rstParametros = db.OpenRecordset(" SELECT [�ltima Movimenta��o] FROM [Par�metros Filial] WHERE Filial =" & gnCodFilial, dbOpenDynaset)
    
    rstParametros.Edit
    rstParametros.Fields("�ltima Movimenta��o").Value = m_nSequencia
    rstParametros.Update
    rstParametros.Close

    Set rstParametros = Nothing
    'Fim 'Par�metros

  End If 'If m_intContador = 1

  If m_intContador >= 1 Then

    '[Entradas - Produtos]
    Set rstEntradasProdutos = db.OpenRecordset("SELECT * FROM [Entradas - Produtos]", dbOpenDynaset)
  
    With rstEntradasProdutos
      .AddNew
      
      .Fields("Filial").Value = gnCodFilial
      .Fields("Sequ�ncia").Value = m_nSequencia
      .Fields("Linha").Value = m_intContador
      
      '21/10/2005 - mpdea
      'Corrigido inser��o do c�digo completo na cria��o dos registros
      .Fields("C�digo").Value = strCodigoCompleto
      
      .Fields("Qtde").Value = sngQtde
      
      
      '25/10/2005 - mpdea
      'Inclu�do o pre�o do produto de acordo com a tabela de pre�os informada
      .Fields("Pre�o").Value = Format(gcGetPrecoProduto(strProduto, txtTabelaPrecos.Text), FORMAT_VALUE)
      
      
      .Fields("Desconto").Value = 0
      .Fields("ICM").Value = 0
      .Fields("IPI").Value = 0
      
      
      '25/10/2005 - mpdea
      'C�lculo do Pre�o Final (simplificdo)
      .Fields("Pre�o Final").Value = Format(.Fields("Qtde").Value * .Fields("Pre�o").Value, FORMAT_VALUE)
      
      
      '27/04/2004 - Daniel
      'Adicionado o field C�digo sem Grade
      .Fields("C�digo sem Grade").Value = strProduto
      
      .Update
      .Close
    End With
    
    Set rstEntradasProdutos = Nothing
    'Fim [Entradas - Produtos]

  End If 'If m_intContador >= 1

End Sub
