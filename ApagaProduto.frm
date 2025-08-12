VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmApagaProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apagar Produtos"
   ClientHeight    =   4245
   ClientLeft      =   2190
   ClientTop       =   1920
   ClientWidth     =   7665
   HelpContextID   =   1840
   Icon            =   "ApagaProduto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4245
   ScaleWidth      =   7665
   Begin VB.TextBox txt_alvoEnderecoWEBAPI 
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3825
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   2625
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton B_Cancelar 
      Caption         =   "Cancelar"
      Height          =   400
      Left            =   4005
      TabIndex        =   10
      Top             =   3645
      Width           =   1335
   End
   Begin VB.CommandButton B_Apagar 
      Caption         =   "Apagar"
      Height          =   400
      Left            =   2490
      TabIndex        =   9
      Top             =   3645
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Produto2 
      Bindings        =   "ApagaProduto.frx":058A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
      DataFieldList   =   "Nome"
      MaxDropDownItems=   16
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8705
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3281
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Produto1 
      Bindings        =   "ApagaProduto.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   2700
      Width           =   1815
      DataFieldList   =   "Nome"
      MaxDropDownItems=   16
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8387
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3572
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Produto2 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3360
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Nome_Produto1 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3360
      TabIndex        =   7
      Top             =   2715
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Produto Final :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Produto Inicial :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "É altamente desaconselhável que você apague produtos. Considere a opção de deixá-lo inativo ao invés de apagá-lo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   7455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ao apagar produtos você poderá ter vários relatórios com valores, quantidades e totais incorretos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   795
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ATENÇÃO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   7455
   End
End
Attribute VB_Name = "frmApagaProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsComissoes As Recordset
Private rsConta_Cliente As Recordset
Private rsEmprestimos As Recordset
Private rsEntradas_Prod As Recordset
Private rsEtiquetas As Recordset
Private rsResumo_Clientes As Recordset
Private rsResumo_Produtos As Recordset
Private rsSaidas_Prod As Recordset


Private Cod_Ini As String
Private Cod_Fim As String

' Variaveis para tratamento do Componente WebApi
Dim http As MSXML2.XMLHTTP
'Dim sRETORNO_ERRO_WEBAPI As String
'Dim alvoEnderecoWEBAPI As String
'Dim tentativasWEBAPI As Integer
'Dim sCodigoProdutoDELETE As String
'Dim cod_produtoArray(2000) As String
'Dim iContProdutoArray As Integer
'Private rsProdutosGradeDELETE As Recordset

Sub Apaga_Comissões()
 Dim Aux_Data     As Variant
 Dim Aux_Vendedor As Long
 Dim Aux_Produto  As String
 Dim Aux_Tamanho  As Integer
 Dim Aux_Cor      As Integer
 Dim Aux_Edição   As Long
 Dim Aux_Contador As Long

  
  Call StatusMsg("Alterando comissões...")
  Aux_Data = CDate("01/01/1980")
  Aux_Vendedor = 0
  Aux_Produto = ""
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edição = 0
  Aux_Contador = 0
  rsComissoes.Index = "Vendedor"
Lp1:
  rsComissoes.Seek ">", Aux_Data, Aux_Vendedor, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Contador
  If rsComissoes.NoMatch Then Exit Sub
  Aux_Data = rsComissoes("Data")
  Aux_Vendedor = rsComissoes("Vendedor")
  Aux_Produto = rsComissoes("Produto")
  Aux_Tamanho = rsComissoes("Tamanho")
  Aux_Cor = rsComissoes("Cor")
  Aux_Edição = rsComissoes("Edição")
  Aux_Contador = rsComissoes("Contador")
  If rsComissoes("Produto") = Cod_Ini Then
    rsComissoes.Edit
      rsComissoes("Produto") = 0
    rsComissoes.Update
  End If
  GoTo Lp1
  
  

End Sub

Sub Apaga_Entradas_Prod()
  Dim Aux_Filial    As Integer
  Dim Aux_Sequência As Long
  Dim Aux_Linha     As Integer
  
  
  Call StatusMsg("Alterando entradas...")
  
  Aux_Filial = 0
  Aux_Sequência = 0
  Aux_Linha = 0
  rsEntradas_Prod.Index = "Sequência"
Lp1:
  rsEntradas_Prod.Seek ">", Aux_Filial, Aux_Sequência, Aux_Linha
  If rsEntradas_Prod.NoMatch Then Exit Sub
  
  Aux_Filial = rsEntradas_Prod("Filial")
  Aux_Sequência = rsEntradas_Prod("Sequência")
  Aux_Linha = rsEntradas_Prod("Linha")
  
  If rsEntradas_Prod("Código") = Cod_Ini Then
    rsEntradas_Prod.Edit
      rsEntradas_Prod("Código") = 0
    rsEntradas_Prod.Update
  End If
  
  GoTo Lp1
  
End Sub

Sub Apaga_Saídas_Prod()
  Dim Aux_Filial    As Integer
  Dim Aux_Sequência As Long
  Dim Aux_Linha     As Integer
  
  
  Call StatusMsg("Alterando saídas...")
  
  Aux_Filial = 0
  Aux_Sequência = 0
  Aux_Linha = 0
  rsSaidas_Prod.Index = "Sequência"
Lp1:
  rsSaidas_Prod.Seek ">", Aux_Filial, Aux_Sequência, Aux_Linha
  If rsSaidas_Prod.NoMatch Then Exit Sub
  
  Aux_Filial = rsSaidas_Prod("Filial")
  Aux_Sequência = rsSaidas_Prod("Sequência")
  Aux_Linha = rsSaidas_Prod("Linha")
  
  If rsSaidas_Prod("Código") = Cod_Ini Then
    rsSaidas_Prod.Edit
      rsSaidas_Prod("Código") = 0
    rsSaidas_Prod.Update
  End If
  
  
  GoTo Lp1
  

End Sub

Sub Apaga_Conta_Cliente()
  Dim Aux_Filial As Integer
  Dim Aux_Sequência As Long
  Dim Aux_Contador  As Long
  
  
  Call StatusMsg("Alterando conta do cliente...")
  
  Aux_Filial = 0
  Aux_Sequência = 0
  Aux_Contador = 0
  rsConta_Cliente.Index = "Sequência"
Lp1:
  rsConta_Cliente.Seek ">", Aux_Filial, Aux_Sequência, Aux_Contador
  If rsConta_Cliente.NoMatch Then Exit Sub
  Aux_Filial = rsConta_Cliente("Filial")
  Aux_Sequência = rsConta_Cliente("Sequência")
  Aux_Contador = rsConta_Cliente("Contador")
  
  If rsConta_Cliente("Produto") = Cod_Ini Then
    rsConta_Cliente.Edit
    rsConta_Cliente("Produto") = "0"
    rsConta_Cliente("Data Alteração") = Format(Date, "dd/mm/yyyy")
    rsConta_Cliente.Update
  End If
  GoTo Lp1
  
  
End Sub

Sub Apaga_Empréstimos()
  Dim Aux_Cliente As Long
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Saldo_Ant As Double
  Dim Data_Ant As Variant
  
  Call StatusMsg("Alterando empréstimos...")
  Aux_Cliente = 0
  Aux_Tamanho = 0
  Aux_Cor = 0
  rsEmprestimos.Index = "Produto"
Lp1:
  rsEmprestimos.Seek ">", Cod_Ini, Aux_Tamanho, Aux_Cor, Aux_Cliente
  If rsEmprestimos.NoMatch Then Exit Sub
  If rsEmprestimos("Produto") <> Cod_Ini Then Exit Sub
  Aux_Tamanho = rsEmprestimos("Tamanho")
  Aux_Cor = rsEmprestimos("Cor")
  Aux_Cliente = rsEmprestimos("Cliente")
  
  Data_Ant = rsEmprestimos("Última Data")
  Saldo_Ant = rsEmprestimos("Saldo Emprestado")
  
  rsEmprestimos.Delete
  
  
  rsEmprestimos.Seek "=", 0, 0, 0, Aux_Cliente
  If rsEmprestimos.NoMatch Then
    rsEmprestimos.AddNew
      rsEmprestimos("Cliente") = Aux_Cliente
      rsEmprestimos("Produto") = Cod_Ini
      rsEmprestimos("Tamanho") = Aux_Tamanho
      rsEmprestimos("Cor") = Aux_Cor
      rsEmprestimos("Saldo Emprestado") = Saldo_Ant
      rsEmprestimos("Última Data") = Data_Ant
      rsEmprestimos("Data Alteração") = Format(Date, "dd/mm/yyyy")
    rsEmprestimos.Update
  Else
    rsEmprestimos.Edit
       rsEmprestimos("Saldo Emprestado") = rsEmprestimos("Saldo Emprestado") + Saldo_Ant
       rsEmprestimos("Data Alteração") = Format(Date, "dd/mm/yyyy")
    rsEmprestimos.Update
  End If
  

  GoTo Lp1
  
  
End Sub

Sub Apaga_Resumo_Clientes()
 Dim Aux_Cliente As Long
 Dim Aux_Dia As Variant
 Dim Aux_Produto As String
 Dim Aux_Tamanho As Integer
 Dim Aux_Cor As Integer
 Dim Aux_Edição As Long
 Dim Aux_Sequência As Long
 Dim Qtde As Double
 Dim Valor_Total As Double
 
 
  Call StatusMsg("Alterando resumo de clientes...")
  
  Aux_Cliente = 0
  Aux_Dia = CDate("01/01/1980")
  Aux_Produto = ""
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edição = 0
  Aux_Sequência = 0
  
  rsResumo_Clientes.Index = "Cliente"
Lp1:
  rsResumo_Clientes.Seek ">", Aux_Cliente, Aux_Dia, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Sequência
  If rsResumo_Clientes.NoMatch Then Exit Sub
  Aux_Cliente = rsResumo_Clientes("Cliente")
  Aux_Dia = rsResumo_Clientes("Dia")
  Aux_Produto = rsResumo_Clientes("Produto")
  Aux_Tamanho = rsResumo_Clientes("Tamanho")
  Aux_Cor = rsResumo_Clientes("Cor")
  Aux_Edição = rsResumo_Clientes("Edição")
  Aux_Sequência = rsResumo_Clientes("Sequência")
  
  If rsResumo_Clientes("Produto") <> Cod_Ini Then GoTo Lp1
  
  Qtde = rsResumo_Clientes("Qtde")
  Valor_Total = rsResumo_Clientes("Valor Total")
  rsResumo_Clientes.Delete
  
  rsResumo_Clientes.Seek "=", Aux_Cliente, Aux_Dia, 0, 0, 0, 0, Aux_Sequência
  If rsResumo_Clientes.NoMatch Then
    rsResumo_Clientes.AddNew
      rsResumo_Clientes("Cliente") = Aux_Cliente
      rsResumo_Clientes("Dia") = Aux_Dia
      rsResumo_Clientes("Produto") = 0
      rsResumo_Clientes("Tamanho") = 0
      rsResumo_Clientes("Cor") = 0
      rsResumo_Clientes("Edição") = 0
      rsResumo_Clientes("Sequência") = Aux_Sequência
      rsResumo_Clientes("Qtde") = Qtde
      rsResumo_Clientes("Valor Total") = Valor_Total
      rsResumo_Clientes("Filial") = 1
    rsResumo_Clientes.Update
  Else
    rsResumo_Clientes.Edit
      rsResumo_Clientes("Qtde") = rsResumo_Clientes("Qtde") + Qtde
      rsResumo_Clientes("Valor Total") = rsResumo_Clientes("Valor Total") + Valor_Total
    rsResumo_Clientes.Update
  End If
  
  GoTo Lp1
  
End Sub

Sub Apaga_Resumo_Produtos()
 Dim Aux_Filial As Integer
 Dim Aux_Mes As Integer
 Dim Aux_Ano As Integer
 Dim Aux_Produto As Double
 Dim Aux_Classe As Integer
 Dim Aux_Sub_Classe As Integer
 Dim Aux_Tamanho As Integer
 Dim Aux_Cor As Integer

  Call StatusMsg("Apagando resumo de produtos...")
  
  Aux_Filial = 0
  Aux_Mes = 0
  Aux_Ano = 0
  Aux_Produto = 0
  Aux_Classe = 0
  Aux_Sub_Classe = 0
  Aux_Tamanho = 0
  Aux_Cor = 0
  rsResumo_Produtos.Index = "Produto"
Lp1:
  rsResumo_Produtos.Seek ">", Aux_Filial, Aux_Mes, Aux_Ano, Aux_Classe, Aux_Sub_Classe, Aux_Produto, Aux_Tamanho, Aux_Cor
  If rsResumo_Produtos.NoMatch Then Exit Sub
  
  Aux_Filial = rsResumo_Produtos("Filial")
  Aux_Mes = rsResumo_Produtos("Mes")
  Aux_Ano = rsResumo_Produtos("Ano")
  Aux_Produto = rsResumo_Produtos("Produto")
  Aux_Classe = rsResumo_Produtos("Classe")
  Aux_Sub_Classe = rsResumo_Produtos("Sub Classe")
  Aux_Tamanho = rsResumo_Produtos("Tamanho")
  Aux_Cor = rsResumo_Produtos("Cor")
  
  If rsResumo_Produtos("Produto") = Cod_Ini Then rsResumo_Produtos.Delete
  
  GoTo Lp1
  
End Sub


'Public Function WebRequest(url As String) As Boolean
'On Error GoTo trata_WebApiErro
'    ' Função que chama o componente WebApi
'    Dim retWebApiAux As String
'
'    http.Open "GET", url, False
'    http.Send
'
'    retWebApiAux = http.statusText
'
'    If retWebApiAux = "OK" Then
'        sRETORNO_ERRO_WEBAPI = "Cadastrado com sucesso!"
'        WebRequest = True
'    Else
'        sRETORNO_ERRO_WEBAPI = http.responseText
'        WebRequest = False
'    End If
'
'    Exit Function
'
'trata_WebApiErro:
'    Dim iInStrRet As Integer
'    iInStrRet = InStr(1, Err.Description, "timed out")
'    If iInStrRet > 0 And tentativasWEBAPI < 4 Then
'        tentativasWEBAPI = tentativasWEBAPI + 1
'        WebRequest url
'    End If
'    sRETORNO_ERRO_WEBAPI = Err.Description
'End Function

'Private Sub ChamadaWebApiWiseDELETE()
'  On Error GoTo ErrChamadaWebApiWiseDELETE
'
'    Dim cod_produto As String
'    Dim retWebApi As Boolean
'    Dim cnpj As String
'    Dim msgRetWebApi As String
'    Dim rsCNPJ As Recordset
'
'    alvoEnderecoWEBAPI = txt_alvoEnderecoWEBAPI.Text
'    If alvoEnderecoWEBAPI = "INTEGRACAO_WISE=NAO" Then
'      bolINTEGRACAO_WISE = False
'    Else
'      bolINTEGRACAO_WISE = True
'    End If
'
'    If bolINTEGRACAO_WISE = True Then
'
'      Set rsCNPJ = db.OpenRecordset("Select CGC From [Parâmetros Filial]")
'      Set http = CreateObject("MSXML2.ServerXMLHTTP")
'
'      Do While Not rsCNPJ.EOF
'          cnpj = rsCNPJ.Fields(0)
'          cnpj = Trim(cnpj)
'          cnpj = Replace(cnpj, ".", "")
'          cnpj = Replace(cnpj, "/", "")
'          cnpj = Replace(cnpj, "-", "")
'
'          If cnpj <> "" Then
'            tentativasWEBAPI = 0
'
'            cod_produto = sCodigoProdutoDELETE
'
'            ' Excluir produto no wise
'            retWebApi = WebRequest(alvoEnderecoWEBAPI + "api/product/DeleteProduct?merchantId=" + cnpj + "&code=" + cod_produto)
'
'            If retWebApi = False Then
'              MsgBox "Atenção: Erro de integração deste produto junto ao Sistema Wise! Erro: " + sRETORNO_ERRO_WEBAPI, vbInformation
'            Else
'
'              Dim i As Integer
'              If iContProdutoArray > 0 Then
'                For i = 0 To iContProdutoArray - 1
'                  cod_produto = cod_produtoArray(i)
'
'                  ' Excluir produto no wise
'                  retWebApi = WebRequest(alvoEnderecoWEBAPI + "api/product/DeleteProduct?merchantId=" + cnpj + "&code=" + cod_produto)
'                Next
'              End If
'            End If
'          End If
'
'          rsCNPJ.MoveNext
'      Loop
'
'      rsCNPJ.Close
'      Set rsCNPJ = Nothing
'      Set http = Nothing
'    End If
'
'    Exit Sub
'
'ErrChamadaWebApiWiseDELETE:
'  MsgBox "Atenção: Erro de integração deste produto junto ao Sistema Wise! Erro: (ChamadaWebApiWiseDELETE) " + sRETORNO_ERRO_WEBAPI, vbInformation
'End Sub


'-----------------------------------------------------------------------------------
'08/07/2002 - mpdea
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'(produtos a excluir)
'-----------------------------------------------------------------------------------
Private Sub B_Apagar_Click()
  Dim sSql As String
  Dim sCodProd As String
  Dim sCriteria As String
  
  On Error GoTo ErrHandler
  
  Cod_Ini = Combo_Produto1.Text
  Cod_Fim = Combo_Produto2.Text
  
  If Gera_Ordenação(Cod_Ini) > Gera_Ordenação(Cod_Fim) Then
    gsTitle = LoadResString(201)
    gsMsg = "Intervalo de Códigos não é crescente. Reentre."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Esta operação não poderá ser desfeita e pode gerar resultados indesejáveis. Deseja realmente prosseguir?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    Exit Sub
  End If
  
  Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
  Set rsComissoes = db.OpenRecordset("Comissão")
  Set rsConta_Cliente = db.OpenRecordset("Conta Cliente")
  Set rsEntradas_Prod = db.OpenRecordset("Entradas - Produtos")
  Set rsSaidas_Prod = db.OpenRecordset("Saídas - Produtos")
  
  Call ws.BeginTrans
  
  Screen.MousePointer = vbHourglass
  
  If Cod_Ini <> Cod_Fim Then
    sCriteria = "[Código Ordenação] >= '" & Gera_Ordenação(Cod_Ini) & "'"
    sCriteria = sCriteria & " AND [Código Ordenação] <= '" & Gera_Ordenação(Cod_Fim) & "'"
  Else
    sCriteria = "Código = '" & Cod_Ini & "'"
  End If
  rsProdutos.FindFirst sCriteria
  
  Do While Not rsProdutos.NoMatch
    
'    DoEvents
    
    sCodProd = rsProdutos("Código").Value
    
    ' Verificar se tem produtos da grade...
'    Dim rsProdutosGradeDELETE As Recordset
'    iContProdutoArray = 0
'    Set rsProdutosGradeDELETE = db.OpenRecordset("Select Código FROM [Códigos da Grade] WHERE [Código Original] = '" & sCodProd & "'")
'    Do While Not rsProdutosGradeDELETE.EOF
'      cod_produtoArray(iContProdutoArray) = rsProdutosGradeDELETE.Fields(0)
'      rsProdutosGradeDELETE.MoveNext
'      iContProdutoArray = iContProdutoArray + 1
'    Loop
'    rsProdutosGradeDELETE.Close
'    Set rsProdutosGradeDELETE = Nothing
    
    sSql = "DELETE * FROM [Códigos da Grade] WHERE [Código Original] = '" & sCodProd & "'"
    Call db.Execute(sSql, dbFailOnError)
    sSql = "DELETE * FROM [Estoque] WHERE Produto = '" & sCodProd & "'"
    Call db.Execute(sSql, dbFailOnError)
    sSql = "DELETE * FROM [Etiquetas] WHERE Produto = '" & sCodProd & "'"
    Call db.Execute(sSql, dbFailOnError)
    sSql = "DELETE * FROM [Preços] WHERE Produto = '" & sCodProd & "'"
    Call db.Execute(sSql, dbFailOnError)
    sSql = "DELETE * FROM [Estoque Final] WHERE Produto = '" & sCodProd & "'"
    Call db.Execute(sSql, dbFailOnError)
    
    Call Apaga_Resumo_Clientes
    Call Apaga_Comissões
    Call Apaga_Conta_Cliente
    Call Apaga_Entradas_Prod
    Call Apaga_Saídas_Prod
    
    'Verifica se o produto já foi do tipo WEB
    Select Case UCase(rsProdutos.Fields("WebLastOp").Value & "")
      Case "U", "I" 'U = update, I = Insert
        'Caso tenha sido, inclui o mesmo na lista de produtos a excluir
        'da Loja Virtual
        Call db.Execute("INSERT INTO WEB_ProdutosExcluir (Codigo) VALUES ('" & _
                        rsProdutos.Fields("Código").Value & "')", dbFailOnError)
    End Select
         
    ' Apaga_Empréstimos
    ' Apaga_Resumo_Produtos
    rsProdutos.Delete
    rsProdutos.FindNext sCriteria
    
  Loop
  
  'sCodigoProdutoDELETE = sCodProd
  
  Call ws.CommitTrans
  
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  DisplayMsg "Operação concluída com sucesso."
  Unload Me
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao apagar referências de produto. Código=" & sCodProd
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub B_Cancelar_Click()
  Unload Me
End Sub

Private Sub Combo_Produto1_CloseUp()
  Combo_Produto1.Text = Combo_Produto1.Columns(1).Text
  Nome_Produto1.Caption = Combo_Produto1.Columns(0).Text
End Sub

Private Sub Combo_Produto1_LostFocus()
  Data1.Recordset.FindFirst "Código = '" & Combo_Produto1.Text & "'"
  If Not Data1.Recordset.NoMatch Then
    Nome_Produto1.Caption = Data1.Recordset.Fields("Nome").Value
  Else
    Nome_Produto1.Caption = ""
  End If
End Sub

'25/01/2006 - mpdea
'Corrigido referências de objetos ao Produto Inicial
'e não ao Produto Final que é o correto
Private Sub Combo_Produto2_CloseUp()
  Combo_Produto2.Text = Combo_Produto2.Columns(1).Text
  Nome_Produto2.Caption = Combo_Produto2.Columns(0).Text
End Sub

Private Sub Combo_Produto2_LostFocus()
  Data2.Recordset.FindFirst "Código = '" & Combo_Produto2.Text & "'"
  If Not Data2.Recordset.NoMatch Then
    Nome_Produto2.Caption = Data2.Recordset.Fields("Nome").Value
  Else
    Nome_Produto2.Caption = ""
  End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Screen.MousePointer = vbHourglass
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Screen.MousePointer = vbDefault
End Sub


'04/11/2002 - mpdea
'Incluído verificação dos recordsets abertos (Run-time 91 ao fechar tela - cancelar)
Private Sub Form_Unload(Cancel As Integer)

  If Not rsResumo_Clientes Is Nothing Then rsResumo_Clientes.Close
  Set rsResumo_Clientes = Nothing
  
  If Not rsComissoes Is Nothing Then rsComissoes.Close
  Set rsComissoes = Nothing
  
  If Not rsConta_Cliente Is Nothing Then rsConta_Cliente.Close
  Set rsConta_Cliente = Nothing
  
  If Not rsEntradas_Prod Is Nothing Then rsEntradas_Prod.Close
  Set rsEntradas_Prod = Nothing
  
  If Not rsSaidas_Prod Is Nothing Then rsSaidas_Prod.Close
  Set rsSaidas_Prod = Nothing

End Sub
