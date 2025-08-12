VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagemNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsMensagensNotaFiscal"
Option Explicit

'27/01/2006 - mpdea
'Classe para a utilização de mensagens para nota fiscal

Public Enum TipoFiltroProduto
  tfpTodos
  tfpGrupoFiscal
  tfpClasseSubClasse
  tfpEspecifico
End Enum

Public Enum TipoFiltroOpSaida
  tfoTodas
  tfoGrupoFiscal
  tfoEspecifica
End Enum

Public Enum TipoFiltroUF
  tfuTodos
  tfuEspecifico
End Enum

'Propriedades
Private m_lngCodigo As Long
Private m_intOrdem As Integer
Private m_enuTipoFiltroProduto As TipoFiltroProduto
Private m_enuTipoFiltroOpSaida As TipoFiltroOpSaida
Private m_enuTipoFiltroUF As TipoFiltroUF
Private m_strFiltroProduto As String
Private m_strFiltroOpSaida As String
Private m_strFiltroUF As String
Private m_strMensagem As String

'------------------------------------------------------------------------------
'Propriedade: Mensagem
'------------------------------------------------------------------------------
Public Property Let Mensagem(ByVal vData As String)
  m_strMensagem = vData
End Property

Public Property Get Mensagem() As String
  Mensagem = m_strMensagem
End Property

'------------------------------------------------------------------------------
'Propriedade: FiltroUF
'------------------------------------------------------------------------------
Public Property Let FiltroUF(ByVal vData As String)
  m_strFiltroUF = vData
End Property

Public Property Get FiltroUF() As String
  FiltroUF = m_strFiltroUF
End Property

'------------------------------------------------------------------------------
'Propriedade: FiltroOpSaida
'------------------------------------------------------------------------------
Public Property Let FiltroOpSaida(ByVal vData As String)
  m_strFiltroOpSaida = vData
End Property

Public Property Get FiltroOpSaida() As String
  FiltroOpSaida = m_strFiltroOpSaida
End Property

'------------------------------------------------------------------------------
'Propriedade: FiltroProduto
'------------------------------------------------------------------------------
Public Property Let FiltroProduto(ByVal vData As String)
  m_strFiltroProduto = vData
End Property

Public Property Get FiltroProduto() As String
  FiltroProduto = m_strFiltroProduto
End Property

'------------------------------------------------------------------------------
'Propriedade: TipoFiltroUF
'------------------------------------------------------------------------------
Public Property Let TipoFiltroUF(ByVal vData As TipoFiltroUF)
  m_enuTipoFiltroUF = vData
End Property

Public Property Get TipoFiltroUF() As TipoFiltroUF
  TipoFiltroUF = m_enuTipoFiltroUF
End Property

'------------------------------------------------------------------------------
'Propriedade: TipoFiltroOpSaida
'------------------------------------------------------------------------------
Public Property Let TipoFiltroOpSaida(ByVal vData As TipoFiltroOpSaida)
  m_enuTipoFiltroOpSaida = vData
End Property

Public Property Get TipoFiltroOpSaida() As TipoFiltroOpSaida
  TipoFiltroOpSaida = m_enuTipoFiltroOpSaida
End Property

'------------------------------------------------------------------------------
'Propriedade: TipoFiltroProduto
'------------------------------------------------------------------------------
Public Property Let TipoFiltroProduto(ByVal vData As TipoFiltroProduto)
  m_enuTipoFiltroProduto = vData
End Property

Public Property Get TipoFiltroProduto() As TipoFiltroProduto
  TipoFiltroProduto = m_enuTipoFiltroProduto
End Property

'------------------------------------------------------------------------------
'Propriedade: Ordem
'Descrição  : Ordem de execução da regra
'------------------------------------------------------------------------------
Public Property Let Ordem(ByVal vData As Integer)
  m_intOrdem = vData
End Property

Public Property Get Ordem() As Integer
  Ordem = m_intOrdem
End Property

'------------------------------------------------------------------------------
'Propriedade: Codigo
'Descrição  : Identificador do registro (Auto numeração no BD)
'------------------------------------------------------------------------------
Public Property Let Codigo(ByVal vData As Long)
  m_lngCodigo = Codigo
End Property

Public Property Get Codigo() As Long
  Codigo = m_lngCodigo
End Property

'------------------------------------------------------------------------------
'Método     : AddNew
'Descrição  : Adiciona o registro ao BD
'------------------------------------------------------------------------------
Public Sub AddNew()
  Dim rstMensagens As Recordset
  
  
  On Error GoTo ErrHandler
  
  
  Set rstMensagens = db.OpenRecordset("MensagensNotaFiscal", dbOpenDynaset)
  With rstMensagens
    .AddNew
    m_lngCodigo = .Fields("Codigo").Value
    .Fields("Ordem").Value = m_intOrdem
    .Fields("TipoFiltroProduto").Value = m_enuTipoFiltroProduto
    .Fields("TipoFiltroOpSaida").Value = m_enuTipoFiltroOpSaida
    .Fields("TipoFiltroUF").Value = m_enuTipoFiltroUF
    .Fields("FiltroProduto").Value = m_strFiltroProduto
    .Fields("FiltroOpSaida").Value = m_strFiltroOpSaida
    .Fields("FiltroUF").Value = m_strFiltroUF
    .Fields("Mensagem").Value = m_strMensagem
    .Update
    .Close
  End With
  Set rstMensagens = Nothing
  
  Exit Sub
  
ErrHandler:
  'Fecha tabela
  If Not rstMensagens Is Nothing Then
    rstMensagens.Close
    Set rstMensagens = Nothing
  End If
  'Repassa erro
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

'------------------------------------------------------------------------------
'Método     : Load
'Descrição  : Adiciona o registro ao BD
'Parâmetros : Codigo -> Código do registro
'Retorno    : True   -> Registro encontrado
'             False  -> Registro não encontrado
'------------------------------------------------------------------------------
Public Function Load(ByVal Codigo As Long) As Boolean
  Dim rstMensagens As Recordset
  Dim strSQL As String
  
  
  On Error GoTo ErrHandler
  
  
  strSQL = "SELECT * FROM MensagensNotaFiscal "
  strSQL = strSQL & "WHERE Codigo = " & Codigo
  Set rstMensagens = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstMensagens
    If Not (.BOF And .EOF) Then
      Call IsDataType(dtLong, .Fields("Codigo").Value, m_lngCodigo)
      Call IsDataType(dtInteger, .Fields("Ordem").Value, m_intOrdem)
      Call IsDataType(dtByte, .Fields("TipoFiltroProduto").Value, m_enuTipoFiltroProduto)
      Call IsDataType(dtByte, .Fields("TipoFiltroOpSaida").Value, m_enuTipoFiltroOpSaida)
      Call IsDataType(dtByte, .Fields("TipoFiltroUF").Value, m_enuTipoFiltroUF)
      Call IsDataType(dtString, .Fields("FiltroProduto").Value, m_strFiltroProduto)
      Call IsDataType(dtString, .Fields("FiltroOpSaida").Value, m_strFiltroOpSaida)
      Call IsDataType(dtString, .Fields("FiltroUF").Value, m_strFiltroUF)
      Call IsDataType(dtString, .Fields("Mensagem").Value, m_strMensagem)
    End If
    .Close
  End With
  Set rstMensagens = Nothing
  
  Exit Function
  
ErrHandler:
  'Fecha tabela
  If Not rstMensagens Is Nothing Then
    rstMensagens.Close
    Set rstMensagens = Nothing
  End If
  'Repassa erro
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Function
