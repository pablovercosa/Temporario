Attribute VB_Name = "modDBFunctions"
Option Explicit

'25/08/2009 - mpdea
'Incluído tratamento de erro opcional nas funções (blnUseErrHandler)
'Útil quando deseja centralizar o tratamento de erro
'Preservado funcionamento já aplicado com Optional e valor padrão como verdadeiro

'17/09/2002 - mpdea
'Estrutura de índice (base de dados)
Private Type IndexField
  sName As String
  sType As String * 1
End Type

Public Function gbCreateFieldZeroLenght(ByVal sTableName As String, ByVal sFieldName As String, _
                                        ByVal nType As DataTypeEnum, Optional ByVal nSize As Integer = 0, _
                                        Optional ByVal blnUseErrHandler As Boolean = True) As Boolean
  
  '25/08/2009 - mpdea
  If blnUseErrHandler Then
    On Error GoTo ErrHandler
  End If
  
  Dim td As TableDef
  Dim fd As Field
  
  'nType = dbBoolean, dbByte, dbInteger, dbSingle, dbDouble, dbCurrency, ...
  'nSize ignored for fixed-size fields and numeric fields...
  Set td = db.TableDefs(sTableName)
  If nType <> dbText Then
    Set fd = td.CreateField(sFieldName, nType)
  Else
    Set fd = td.CreateField(sFieldName, nType, nSize)
  End If
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = Nothing
  Set td = Nothing
  
  gbCreateFieldZeroLenght = True
  Exit Function
  
ErrHandler:
  gbCreateFieldZeroLenght = False
  
End Function

Public Function gbDeleteField(ByVal sTableName As String, _
                              ByVal sFieldName As String, _
                              Optional ByVal blnUseErrHandler As Boolean = True) As Boolean

  If blnUseErrHandler Then
    On Error GoTo ErrHandler
  End If
  
  Dim td As TableDef
  Dim fd As Field
  
  Set td = db.TableDefs(sTableName)
  td.Fields.Delete sFieldName

  Set td = Nothing
  
  gbDeleteField = True
  
  Exit Function

ErrHandler:
  gbDeleteField = False
  MsgBox Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Function


Public Function gbCreateField(ByVal sTableName As String, ByVal sFieldName As String, _
                              ByVal nType As DataTypeEnum, _
                              Optional ByVal nSize As Integer = 0, _
                              Optional ByVal blnAllowZeroLength As Boolean = True, _
                              Optional ByVal blnRequired As Boolean = False, _
                              Optional ByVal blnUseErrHandler As Boolean = True, _
                              Optional ByVal varDefault As String = "") As Boolean
  
  '25/08/2009 - mpdea
  If blnUseErrHandler Then
    On Error GoTo ErrHandler
  End If
  
  Dim td As TableDef
  Dim fd As Field
  
  'nType = dbBoolean, dbByte, dbInteger, dbSingle, dbDouble, dbCurrency, ...
  'nSize ignored for fixed-size fields and numeric fields...
  
  Set td = db.TableDefs(sTableName)
  
  Select Case nType
    Case dbText
      Set fd = td.CreateField(sFieldName, nType, nSize)
      fd.AllowZeroLength = blnAllowZeroLength
      fd.DefaultValue = varDefault
      
    Case dbMemo
      Set fd = td.CreateField(sFieldName, nType)
      fd.AllowZeroLength = blnAllowZeroLength
      fd.DefaultValue = ""
      
    Case dbByte, dbInteger, dbSingle, dbDouble, dbCurrency, dbLong
      Set fd = td.CreateField(sFieldName, nType)
      If (varDefault <> "NÃO PONHA ZERO") Then fd.DefaultValue = 0
    
    '22/09/2005 - mpdea
    'Incluído o tipo dbBoolean para tratamento do valor padrão
    Case dbBoolean
      Set fd = td.CreateField(sFieldName, nType)
      fd.DefaultValue = False
    
    Case Else
      Set fd = td.CreateField(sFieldName, nType)
      
  End Select
  
  fd.Required = blnRequired
  
  td.Fields.Append fd
  Set fd = Nothing
  Set td = Nothing
  
  gbCreateField = True
  Exit Function
  
ErrHandler:
  gbCreateField = False
  
End Function

Public Function gbGetTable(ByVal sTableName As String) As Boolean
  Dim nI As Integer
  For nI = 0 To db.TableDefs.Count - 1
    If UCase(db.TableDefs(nI).Name) = UCase(sTableName) Then
      gbGetTable = True
      Exit Function
    End If
  Next nI
  gbGetTable = False
End Function
Public Function gbGetTableTemp(ByVal sTableName As String) As Boolean
  Dim nI As Integer
  For nI = 0 To dbTemp.TableDefs.Count - 1
    If UCase(dbTemp.TableDefs(nI).Name) = UCase(sTableName) Then
      gbGetTableTemp = True
      Exit Function
    End If
  Next nI
  gbGetTableTemp = False
End Function

Public Function gbGetField(ByVal sTableName As String, ByVal sFieldName As String) As Boolean
  Dim nI As Integer
  Dim td As TableDef
  Set td = db.TableDefs(sTableName)
  For nI = 0 To td.Fields.Count - 1
    If UCase(td.Fields(nI).Name) = UCase(sFieldName) Then
      gbGetField = True
      gnNum = nI
      Exit Function
    End If
  Next nI
  gbGetField = False
End Function


'24/09/2003 - mpdea
'Verifica a existência do índice em determinada tabela
Public Function g_blnGetIndex(ByVal strTableName As String, ByVal strIndexName As String) As Boolean
  Dim nI As Integer
  Dim td As TableDef
  
  Set td = db.TableDefs(strTableName)
  For nI = 0 To td.Indexes.Count - 1
    If UCase(td.Indexes(nI).Name) = UCase(strIndexName) Then
      g_blnGetIndex = True
      Exit Function
    End If
  Next nI
  g_blnGetIndex = False
End Function

Public Function gbAlteraTamanhoCampoIndex(ByVal sTable As String, ByVal sCampo As String, ByVal sTipo As String, _
                                          ByVal nTamanho As Integer, ByVal sIndex As String, ByVal sCampo1 As String, _
                                          ByVal sCampo2 As String, ByVal sPrimary As Boolean, ByVal sUnique As Boolean, _
                                          Optional ByVal blnUseErrHandler As Boolean = True) As Boolean
  
  Dim rs As Recordset
  Dim td As TableDef
  Dim fd As Field
  Dim bGotValor As Boolean
  Dim sField As String
  Dim iX As Index
  
  '25/08/2009 - mpdea
  '17/09/2002 - mpdea
  'Corrigido ordem da verificação da função
  'Implementado rotina de erro
  If blnUseErrHandler Then
    On Error GoTo ErrHandler
  End If
  
  Set td = db.TableDefs(sTable)

  If td(sCampo).Size >= nTamanho Then
    gbAlteraTamanhoCampoIndex = True
    Set td = Nothing
    Exit Function
  End If


  On Error Resume Next
  td.Indexes.Delete sIndex
  td.Indexes.Refresh
  Err.Clear
  
  On Error GoTo ErrHandler
  
  Set fd = td.CreateField("Campo2", sTipo, nTamanho)
  fd.AllowZeroLength = True

  td.Fields.Append fd


  Set td = Nothing

  
  Set rs = db.OpenRecordset(sTable)
  If Not rs.EOF Then
    Do While Not rs.EOF
      rs.Edit
      rs("Campo2").Value = rs(sCampo).Value & ""
      rs.Update
      rs.MoveNext
    Loop
  End If
  rs.Close
  Set rs = Nothing
  
 Set td = db.TableDefs(sTable)
 
 td.Fields.Delete sCampo
 td.Fields("Campo2").Name = sCampo
' Set td = Nothing
 
  Set iX = td.CreateIndex
  With iX
    .Name = sIndex
    .Fields.Append .CreateField(sCampo1)
    .Fields.Append .CreateField(sCampo2)
    .Primary = sPrimary
    .Unique = sUnique
  End With
  td.Indexes.Append iX
  
  
  Set td = Nothing
  
  gbAlteraTamanhoCampoIndex = True
  
  Exit Function
  
ErrHandler:
  gbAlteraTamanhoCampoIndex = False
  MsgBox Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'19/09/2002 - mpdea
'Função para alteração do tamanho do campo
'Parâmetro opcional para o nome de um índice
Public Function gbAlteraTamanhoCampo2(ByVal strTable As String, _
                                      ByVal strCampo As String, ByVal enuTipo As DataTypeEnum, _
                                      ByVal intTamanho As Integer, Optional ByVal strIndexName As String = "", _
                                      Optional ByVal blnUseErrHandler As Boolean = True) As Boolean
  
  '25/08/2009 - mpdea
  If blnUseErrHandler Then
    On Error GoTo ErrHandler
  End If
  
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  Dim typFieldIndex() As IndexField
  Dim blnIdxFound As Boolean
  Dim blnIdxPrimary As Boolean
  Dim blnIdxUnique As Boolean
  Dim intX As Integer
  
  Set td = db.TableDefs(strTable)
  
  'Validação
  If td(strCampo).Size >= intTamanho Then
    gbAlteraTamanhoCampo2 = True
    Set td = Nothing
    Exit Function
  End If
  
  'Índice
  If strIndexName <> "" Then
    On Error Resume Next
    blnIdxFound = td.Indexes(strIndexName).Name = strIndexName
    Err.Clear
    On Error GoTo ErrHandler
    
    If blnIdxFound Then
      'Armazena dados do índice
      With td.Indexes(strIndexName)
        Call SeparateIndexFields(.Fields, typFieldIndex())
        blnIdxPrimary = .Primary
        blnIdxUnique = .Unique
      End With
      'Exclui índice
      td.Indexes.Delete strIndexName
      td.Indexes.Refresh
    End If
  End If
  
  'Cria novo campo
  Set fd = td.CreateField("NewField", enuTipo, intTamanho)
  With td(strCampo)
    fd.AllowZeroLength = .AllowZeroLength
    fd.Required = .Required
  End With
  td.Fields.Append fd
  
  'Atualiza valores
  Set td = Nothing
  db.Execute "UPDATE [" & strTable & "] SET NewField = [" & strCampo & "]"
  Set td = db.TableDefs(strTable)
  
  'Remove campo temporário
  td.Fields.Delete strCampo
  td.Fields("NewField").Name = strCampo
  
  If strIndexName <> "" Then
    'Recria índice
    If blnIdxFound Then
      Set iX = td.CreateIndex
      With iX
        .Name = strIndexName
        For intX = LBound(typFieldIndex) To UBound(typFieldIndex)
          .Fields.Append .CreateField(typFieldIndex(intX).sName)
        Next intX
        .Primary = blnIdxPrimary
        .Unique = blnIdxUnique
      End With
      td.Indexes.Append iX
    End If
  End If
  
  Set iX = Nothing
  Set fd = Nothing
  Set td = Nothing
  
  gbAlteraTamanhoCampo2 = True
  
  Exit Function
  
ErrHandler:
  gbAlteraTamanhoCampo2 = False
  MsgBox Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

Public Sub SeparateIndexFields(ByVal sTexto As String, ByRef sList() As IndexField)
  Dim nX As Integer
  Dim sFieldIndex() As String
  
  sFieldIndex = Split(sTexto, ";", -1, vbTextCompare)
  For nX = LBound(sFieldIndex) To UBound(sFieldIndex)
    ReDim Preserve sList(nX)
    sList(nX).sName = Right(sFieldIndex(nX), Len(sFieldIndex(nX)) - 1)
    sList(nX).sType = Left(sFieldIndex(nX), 1)
  Next nX
End Sub
Public Function gbAlteraTipoCampo(ByVal sTable As String, ByVal sCampo As String, ByVal sTipo As String, _
                                     Optional ByVal nTamanho As Integer = -1, _
                                     Optional ByVal blnUseErrHandler As Boolean = True) As Boolean
                                     
'Dim banco As Database
Dim executar As String

If nTamanho <= 0 Then
  executar = "ALTER TABLE " & sTable & " ADD COLUMN " & sCampo & "1 " & sTipo & ";"
  db.Execute executar
  executar = "UPDATE " & sTable & " SET " & sCampo & "1 = " & sCampo & ";"
  db.Execute executar
  executar = "ALTER TABLE " & sTable & " DROP " & sCampo
  db.Execute executar
  executar = "ALTER TABLE " & sTable & " ADD COLUMN " & sCampo & " " & sTipo & ";"
  db.Execute executar
  executar = "UPDATE " & sTable & " SET " & sCampo & " = " & sCampo & "1;"
  db.Execute executar
  executar = "ALTER TABLE " & sTable & " DROP " & sCampo & "1;"
  db.Execute executar
  executar = ""
Else
  executar = "ALTER TABLE " & sTable & " ALTER COLUMN " & sCampo & " " & sTipo & "(" & nTamanho & ");"
  db.Execute executar
  executar = "UPDATE " & sTable & " SET " & sCampo & "1 = " & sCampo & ";"
  db.Execute executar
  executar = "ALTER TABLE " & sTable & " DROP " & sCampo
  db.Execute executar
  executar = "ALTER TABLE " & sTable & " ADD COLUMN " & sCampo & " " & sTipo & ";"
  db.Execute executar
  executar = "UPDATE " & sTable & " SET " & sCampo & " = " & sCampo & "1;"
  db.Execute executar
  executar = "ALTER TABLE " & sTable & " DROP " & sCampo & "1;"
  db.Execute executar
  executar = ""
End If
gbAlteraTipoCampo = True
End Function

Public Function gbAlteraTamanhoCampo(ByVal sTable As String, ByVal sCampo As String, ByVal sTipo As String, _
                                     ByVal nTamanho As Integer, _
                                     Optional ByVal blnUseErrHandler As Boolean = True) As Boolean
  
  Dim rs As Recordset
  Dim td As TableDef
  Dim fd As Field
  Dim fdx As Field
  Dim fdx2 As Field
  Dim sField As String
  Dim sAUX As String
  Dim idx As Index
  Dim sIdxNames() As String
  Dim sIdxCols() As String
  Dim sIdxColsTmp() As String
  Dim n As Integer
  Dim i As Integer
    
  '25/08/2009 - mpdea
  If blnUseErrHandler Then
    On Error GoTo ErrHandler
  End If
  
  Set td = db.TableDefs(sTable)
  If td(sCampo).Size >= nTamanho Then
    gbAlteraTamanhoCampo = True
    Set td = Nothing
    Exit Function
  End If
  
  Set fd = td.CreateField("Campo2", sTipo, nTamanho)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set td = Nothing
  
  Set rs = db.OpenRecordset(sTable)
  If Not rs.EOF Then
    Do While Not rs.EOF
      rs.Edit
      rs("Campo2").Value = rs(sCampo).Value & ""
      rs.Update
      rs.MoveNext
    Loop
  End If
  rs.Close
  Set rs = Nothing
  
  Set td = db.TableDefs(sTable)
  
  For Each idx In td.Indexes
    For Each fdx In idx.Fields
      If fdx.Name = sCampo Then
        ReDim Preserve sIdxNames(n)
        ReDim Preserve sIdxCols(n)
        sIdxNames(n) = idx.Name
        For Each fdx2 In idx.Fields
          If sIdxCols(n) = "" Then
            sIdxCols(n) = fdx2.Name
          Else
            sIdxCols(n) = sIdxCols(n) & "#" & fdx2.Name
          End If
        Next fdx2
        n = n + 1
        Exit For
      End If
    Next fdx
  Next idx
  
  For n = 0 To UBound(sIdxNames)
    td.Indexes.Delete (sIdxNames(n))
  Next n
  
  td.Fields.Delete sCampo
  td.Fields("Campo2").Name = sCampo
  
  For n = 0 To UBound(sIdxNames)
    Set idx = td.CreateIndex(sIdxNames(n))
    sIdxColsTmp = Split(sIdxCols(n), "#")
    For i = 0 To UBound(sIdxColsTmp)
      idx.Fields.Append idx.CreateField(sIdxColsTmp(i))
    Next i
    td.Indexes.Append idx
  Next n
  
  Set td = Nothing
  gbAlteraTamanhoCampo = True
  Exit Function
  
ErrHandler:
'''''  If Err.Number = 3280 Then
'''''    DoEvents
'''''    td.Indexes.Delete ("Código Fiscal")
'''''    Resume
'''''  Else
'''''    Screen.MousePointer = vbDefault
'''''    Select Case frmErro.gnShowErr(Err.Number, "Alterar Código Fiscal")
'''''      Case 0 'Repetir
'''''        Resume
'''''      Case 1 'Prosseguir
'''''        Resume Next
'''''      Case 2 'Sair
'''''        Exit Function
'''''      Case 3 'Encerrar
'''''        End
'''''    End Select
'''''  End If
  gbAlteraTamanhoCampo = False

End Function

'17/11/2009 - mpdea
'Adiciona permissão de usuário
'
'strPrograma = Nome do programa/módulo/tela
'strDescricao = Descrição do programa
'intCodigo = Código do programa (exclusivo)
'lngToolId = Identificador da permissão para o menu
Public Sub AddUserPermission(ByVal strPrograma As String, ByVal strDescricao As String, _
  ByVal intCodigo As Integer, ByVal lngToolId As Long)
  
  Dim rstZZZProgramas As Recordset
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  With rstZZZProgramas
    .Index = "Nome"
    .Seek "=", strPrograma
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = strPrograma
      .Fields("Descrição").Value = strDescricao
      .Fields("Número").Value = intCodigo
      .Fields("ToolID").Value = lngToolId
      .Update
    End If
    .Close
  End With
  
  Set rstZZZProgramas = Nothing
    
End Sub
