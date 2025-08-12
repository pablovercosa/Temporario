VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Private cData As Date
Private cVendedor As Integer
Private cProduto As String
Private cTamanho As Integer
Private cCor As Integer
Private cEdição As Long
Private cQtde As Integer
Private cValor As Currency
Private cComissão As Currency
Private cSequência As Long
Private cCliente As Long
Private cTabela As String
Private cFilial As Integer
Private cVlPagoEmCartao As Double
Private cVlPagoEmCartaoComRetencao As Double
Private cTaxaRetencao As Integer

Private Const sqlInsert As String = "INSERT INTO Comissão ( Data, Vendedor, Produto, Tamanho, Cor, Edição, Qtde, Valor, Comissão, Sequência, Cliente, Tabela, Filial, VlPagoEmCartao, VlPagoEmCartaoComRetencao, TaxaRetencao ) VALUES ( #@Data#, @Vendedor, '@Produto', @Tamanho, @Cor, @Edição, @Qtde, @Valor, @Comissão, @Sequência, @Cliente, '@Tabela', @Filial, @VlPagoEmCartao, @VlPagoEmCartaoComRetencao, @TaxaRetencao )"

Public Property Get Data() As Date
  Data = cData
End Property
Public Property Let Data(ByVal pData As Date)
  cData = pData
End Property

Public Property Get Vendedor() As Integer
  Vendedor = cVendedor
End Property
Public Property Let Vendedor(ByVal pVendedor As Integer)
  cVendedor = pVendedor
End Property

Public Property Get Produto() As String
  Produto = cProduto
End Property
Public Property Let Produto(pProduto As String)
  cProduto = pProduto
End Property

Public Property Get Tamanho() As Integer
  Tamanho = cTamanho
End Property
Public Property Let Tamanho(pTamanho As Integer)
  cTamanho = pTamanho
End Property

Public Property Get Cor() As Integer
  Cor = cCor
End Property
Public Property Let Cor(pCor As Integer)
  cCor = pCor
End Property

Public Property Get Edição() As Long
  Edição = cEdição
End Property
Public Property Let Edição(pEdição As Long)
  cEdição = pEdição
End Property

Public Property Get Qtde() As Integer
  Qtde = cQtde
End Property
Public Property Let Qtde(pQtde As Integer)
  cQtde = pQtde
End Property

Public Property Get Valor() As Currency
  Valor = cValor
End Property
Public Property Let Valor(pValor As Currency)
  cValor = pValor
End Property

Public Property Get Comissão() As Currency
  Comissão = cComissão
End Property
Public Property Let Comissão(pComissão As Currency)
  cComissão = pComissão
End Property

Public Property Get Sequência() As Long
  Sequência = cSequência
End Property
Public Property Let Sequência(pSequência As Long)
  cSequência = pSequência
End Property

Public Property Get Cliente() As Long
  Cliente = cCliente
End Property
Public Property Let Cliente(pCliente As Long)
  cCliente = pCliente
End Property

Public Property Get Tabela() As String
  Tabela = cTabela
End Property
Public Property Let Tabela(pTabela As String)
  cTabela = pTabela
End Property

Public Property Get Filial() As Integer
  Filial = cFilial
End Property
Public Property Let Filial(pFilial As Integer)
  cFilial = pFilial
End Property

Public Property Get VlPagoEmCartao() As Double
  VlPagoEmCartao = cVlPagoEmCartao
End Property
Public Property Let VlPagoEmCartao(pVlPagoEmCartao As Double)
  cVlPagoEmCartao = pVlPagoEmCartao
End Property

Public Property Get VlPagoEmCartaoComRetencao() As Double
  VlPagoEmCartaoComRetencao = cVlPagoEmCartaoComRetencao
End Property
Public Property Let VlPagoEmCartaoComRetencao(pVlPagoEmCartaoComRetencao As Double)
  cVlPagoEmCartaoComRetencao = pVlPagoEmCartaoComRetencao
End Property

Public Property Get TaxaRetencao() As Integer
  TaxaRetencao = cTaxaRetencao
End Property
Public Property Let TaxaRetencao(pTaxaRetencao As Integer)
  cTaxaRetencao = pTaxaRetencao
End Property

Public Sub Insert()
  Dim sql As String
  sql = sqlInsert
  sql = Replace(sql, "@Data", Format(Data, "mm/dd/yyyy"), 1, 1)
  sql = Replace(sql, "@Vendedor", Replace(CStr(Vendedor), ",", "."), 1, 1)
  sql = Replace(sql, "@Produto", Produto, 1, 1)
  sql = Replace(sql, "@Tamanho", Replace(CStr(Tamanho), ",", "."), 1, 1)
  sql = Replace(sql, "@Cor", Replace(CStr(Cor), ",", "."), 1, 1)
  sql = Replace(sql, "@Edição", Replace(CStr(Edição), ",", "."), 1, 1)
  sql = Replace(sql, "@Qtde", Replace(CStr(Qtde), ",", "."), 1, 1)
  sql = Replace(sql, "@Valor", Replace(CStr(Valor), ",", "."), 1, 1)
  sql = Replace(sql, "@Comissão", Replace(CStr(Comissão), ",", "."), 1, 1)
  sql = Replace(sql, "@Sequência", Replace(CStr(Sequência), ",", "."), 1, 1)
  sql = Replace(sql, "@Cliente", Replace(CStr(Cliente), ",", "."), 1, 1)
  sql = Replace(sql, "@Tabela", Tabela, 1, 1)
  sql = Replace(sql, "@Filial", Replace(CStr(Filial), ",", "."), 1, 1)
  sql = Replace(sql, "@VlPagoEmCartao", Replace(CStr(VlPagoEmCartao), ",", "."), 1, 1)
  sql = Replace(sql, "@VlPagoEmCartaoComRetencao", Replace(CStr(VlPagoEmCartaoComRetencao), ",", "."), 1, 1)
  sql = Replace(sql, "@TaxaRetencao", Replace(CStr(TaxaRetencao), ",", "."), 1, 1)
  
  db.Execute sql
End Sub

