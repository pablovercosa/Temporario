Attribute VB_Name = "modExcel"
Option Explicit

Public gsSaveExcelEntrada    As String
Public gsSaveExcelSaida      As String
Public gsArquivoExcelEntrada As String
Public gsArquivoExcelSaida   As String

Public Sub ExcelSubstituirCampo(ByVal Campo As String, ByVal Valor As String, ByVal Arquivo As String, ByRef Aplicativo As Excel.Application, Optional Celulas As String)

  Dim exCelula As Range
  Dim strRange As String
  Dim strValor As String
  Dim intPosicaoCampo As Integer
  
  If Celulas = "" Then
    strRange = "A1:IV65536"
  Else
    strRange = Celulas
  End If
  
  With Aplicativo.Workbooks(Arquivo).Worksheets(1).Range(strRange)
      Set exCelula = .Find(Campo, lookin:=xlValues)
      If Not exCelula Is Nothing Then
          Do
              intPosicaoCampo = InStr(1, exCelula.FormulaR1C1, Campo)
              If intPosicaoCampo > 1 Then
                strValor = Mid(exCelula.FormulaR1C1, 1, intPosicaoCampo - 1) & Valor
                strValor = strValor & Mid(exCelula.FormulaR1C1, intPosicaoCampo + Len(Campo))
              Else
                strValor = Valor & Mid(exCelula.FormulaR1C1, Len(Campo) + 1)
              End If
              exCelula.FormulaR1C1 = strValor
              Set exCelula = .Find(Campo)
          Loop While Not exCelula Is Nothing
      End If
  End With

End Sub

Public Function ExcelLocalizarCampo(ByVal Campo As String, ByVal Arquivo As String, ByRef Aplicativo As Excel.Application) As String

  Dim exCelula As Range
  
  ExcelLocalizarCampo = ""

  With Aplicativo.Workbooks(Arquivo).Worksheets(1).Range("A1:IV65536")
      Set exCelula = .Find(Campo, lookin:=xlValues)
      If Not exCelula Is Nothing Then
        ExcelLocalizarCampo = exCelula.Address
      Else
        ExcelLocalizarCampo = ""
      End If
  End With

End Function

