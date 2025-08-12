VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'12/02/2010 - mpdea
'Classe intermediária no desvio da impressão direta para impressora, a qual salva em arquivo
'para impressão posterior por aplicativo auxiliar

'Buffer das linhas de impressão
Private m_clc_buffer As Collection

'Buffer da linha de impressão
Public Sub BufferLine(ByVal strBuffer As String)
  m_clc_buffer.Add strBuffer
End Sub

'Salva o buffer das linhas de impressão em arquivo
Public Sub SaveBufferToFile(ByVal id As String)
  Dim intFileNum As Integer
  Dim intLinha As Integer
  Dim strPath As String
  Dim strFile As String
  Dim strFilenameTemp As String
  Dim strFilenameFinal As String
  

  On Error GoTo ErrHandler
  
  
  If m_clc_buffer.Count > 0 Then
    'Pasta de armazenamento
    strPath = GetSetting("QuickStore", "ConfigLPT", "PastaSaidaModoRemoto", "")
    strPath = IIf(Right(strPath, 1) = "\", strPath, strPath & "\")
    
    'Nome do arquivo principal
    strFile = Format(Now, "yyyy-MM-dd HH-mm-ss.R") & GetRandomNumber
    
    'Nome do arquivo temporário
    strFilenameTemp = strPath & strFile & ".tmp"
    
    'Nome do arquivo final
    strFilenameFinal = strPath & strFile & ".p" & id
    
    'Exclui arquivos caso existam
    Kill strFilenameTemp
    Kill strFilenameFinal
    
    'Salva arquivo temporário
    intFileNum = FreeFile
    Open strFilenameTemp For Output As #intFileNum
    For intLinha = 1 To m_clc_buffer.Count
      Print #intFileNum, m_clc_buffer.Item(intLinha)
    Next
    Close #intFileNum
    
    'Renomeia arquivo para final
    Name strFilenameTemp As strFilenameFinal
  End If
  
  Exit Sub
  
ErrHandler:
  Close
  Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'Gera número aleatório com 3 dígitos
Private Function GetRandomNumber() As String
  Randomize
  GetRandomNumber = Format(Int(999 * Rnd + 1), "000")
End Function

Private Sub Class_Initialize()
  'Inicializa coleção
  Set m_clc_buffer = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_clc_buffer = Nothing
End Sub
