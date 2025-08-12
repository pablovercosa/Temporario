VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagensNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsMensagemNotaFiscal"
Attribute VB_Ext_KEY = "Member0" ,"clsMensagemNotaFiscal"
Option Explicit

'31/01/2006 - mpdea
'Coleção da classe clsMensagemNotaFiscal

Private m_clcMsgNF As Collection

Public Sub Add(ByRef MensagemNotaFiscal As clsMensagemNotaFiscal, Optional Key As String)
  If Len(Key) = 0 Then
    m_clcMsgNF.Add MensagemNotaFiscal
  Else
    m_clcMsgNF.Add MensagemNotaFiscal, Key
  End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As clsMensagemNotaFiscal
Attribute Item.VB_UserMemId = 0
  Set Item = m_clcMsgNF(vntIndexKey)
End Property

Public Property Get Count() As Long
  Count = m_clcMsgNF.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
  m_clcMsgNF.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = m_clcMsgNF.[_NewEnum]
End Property

Public Sub Load()
  Dim objMensagemNotaFiscal As clsMensagemNotaFiscal
  Dim rstMensagens As Recordset
  Dim strSQL As String
  Dim varRet As Variant
  
  
  On Error GoTo ErrHandler
  
  
  'Nova Collection
  Set m_clcMsgNF = New Collection
  
  'SQL
  strSQL = "SELECT * FROM MensagensNotaFiscal ORDER BY Ordem;"
  
  'Retorna recordset e preenche collection
  Set rstMensagens = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstMensagens
    Do Until .EOF
    
      'Novo objeto
      Set objMensagemNotaFiscal = New clsMensagemNotaFiscal
      
      'Atributos do objeto
      Call IsDataType(dtLong, .Fields("Codigo").Value, varRet)
      objMensagemNotaFiscal.Codigo = CLng(varRet)
      Call IsDataType(dtInteger, .Fields("Ordem").Value, varRet)
      objMensagemNotaFiscal.Ordem = CInt(varRet)
      Call IsDataType(dtByte, .Fields("TipoFiltroProduto").Value, varRet)
      objMensagemNotaFiscal.TipoFiltroProduto = CByte(varRet)
      Call IsDataType(dtByte, .Fields("TipoFiltroOpSaida").Value, varRet)
      objMensagemNotaFiscal.TipoFiltroOpSaida = CByte(varRet)
      Call IsDataType(dtByte, .Fields("TipoFiltroUF").Value, varRet)
      objMensagemNotaFiscal.TipoFiltroUF = CByte(varRet)
      Call IsDataType(dtString, .Fields("FiltroProduto").Value, varRet)
      objMensagemNotaFiscal.FiltroProduto = CStr(varRet)
      Call IsDataType(dtString, .Fields("FiltroOpSaida").Value, varRet)
      objMensagemNotaFiscal.FiltroOpSaida = CStr(varRet)
      Call IsDataType(dtString, .Fields("FiltroUF").Value, varRet)
      objMensagemNotaFiscal.FiltroUF = CStr(varRet)
      Call IsDataType(dtString, .Fields("Mensagem").Value, varRet)
      objMensagemNotaFiscal.Mensagem = CStr(varRet)
      
      'Adiciona a coleção
      m_clcMsgNF.Add objMensagemNotaFiscal, CStr(.Fields("Codigo").Value)
      .MoveNext
    Loop
    .Close
  End With
  Set rstMensagens = Nothing
  Set objMensagemNotaFiscal = Nothing
  
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

Private Sub Class_Initialize()
  Set m_clcMsgNF = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_clcMsgNF = Nothing
End Sub
