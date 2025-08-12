Attribute VB_Name = "modPrintNotaMN"
Option Explicit

'19-20/02/2004 - mpdea
'Módulo auxiliar para a impressão da Nota Fiscal
'
'26/02/2004 - mpdea
'Tratamento da impressão de asteriscos
'
'01/03/2004 - mpdea
'Reformulado módulo para atender primeiro modelo multinota recebido

'Flag que indica o uso de impressão multinota
'
'True  = a página é impressa com asteriscos no lugar dos totalizadores
'        e em branco em determinados campos conforme a função Pega_Campo
'
'False = impressão normal
Public g_blnPrintMNF As Boolean

'02/06/2004 - mpdea
'Contador de linhas durante a impressão
Private m_intCountPrintLine As Integer
'Flag para o case Embalavi
Private m_blnCaseEmbalavi As Boolean

'Impressão da Nota Fiscal através da collection
Public Sub PrintNotaFiscalByColl(ByVal clcLayoutFile As Collection)
  Dim intQtdeProdLayout As Integer
  Dim intQtdeProdLinhas As Integer
  Dim intQtdeServLayout As Integer
  Dim intQtdeServLinhas As Integer
  Dim intProdPages As Integer
  Dim intServPages As Integer
  Dim intPages As Integer
  Dim intPage As Integer

    
  On Error GoTo ErrHandler
  
  
  '02/06/2004 - mpdea
  'Case Embalavi
  m_blnCaseEmbalavi = CheckSerialCaseMod("QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162")
  
  
  '02/06/2004 - mpdea
  'Zera o contador de linhas durante a impressão
  m_intCountPrintLine = 0
  
  '----------------------------------------------------------------
  'Número de páginas
  '----------------------------------------------------------------
  '
  'Obtém a quantidade de produtos e serviços do layout
  Call GetQtdeProdServLayout(clcLayoutFile, intQtdeProdLayout, intQtdeServLayout)
  '
  'Obtém a quantidade de produtos e serviços da movimentação
  '(aproveitado variáveis globais com esta finalidade)
  intQtdeProdLinhas = gnCtItemProd
  intQtdeServLinhas = gnCtItemServ
  '
  'Produtos
  If intQtdeProdLayout > 0 Then
    intProdPages = (intQtdeProdLinhas \ intQtdeProdLayout)
    'Verifica página adicional
    If intQtdeProdLinhas Mod intQtdeProdLayout <> 0 Then
      intProdPages = intProdPages + 1
    End If
  End If
  '
  'Serviços
  If intQtdeServLayout > 0 Then
    intServPages = (intQtdeServLinhas \ intQtdeServLayout)
    'Verifica página adicional
    If intQtdeServLinhas Mod intQtdeServLayout <> 0 Then
      intServPages = intServPages + 1
    End If
  End If
  '
  'Quantidade máxima de páginas
  intPages = intProdPages
  If intServPages > intProdPages Then intPages = intServPages
  '
  '23/04/2004 - mpdea
  'Validação para a quantidade mínima de uma página
  If intPages = 0 Then intPages = 1
  '----------------------------------------------------------------
  
  
  'Impressão multinota
  If intPages > 1 Then g_blnPrintMNF = True
  
  
  'Impressão de acordo com a quantidade de páginas
  For intPage = 1 To intPages
    'Verifica última página
    If intPage = intPages Then
      g_blnPrintMNF = False
    End If
    'Imprime estrutura
    Call PrintCollection(clcLayoutFile)
  Next intPage
  
  
  'Finaliza impressão
  Printer.Print
  Printer.EndDoc
  
  
  Exit Sub
  
ErrHandler:
  With Err
    .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
  End With
  
End Sub

'Imprime as informações de layout da collection
Private Sub PrintCollection(ByVal clcPrint As Collection)
  Dim intLayoutLinha As Integer
  Dim strLayoutLinha As String
  Dim intLinha As Integer
  Dim strPrintLinha As String
  
  
  For intLayoutLinha = 1 To clcPrint.Count
    '
    '02/06/2004 - mpdea
    'Incrementa o contador de linhas durante a impressão
    m_intCountPrintLine = m_intCountPrintLine + 1
    '
    'Linha do layout
    strLayoutLinha = clcPrint.Item(intLayoutLinha)
    'Verifica avanço de linhas em branco
    If Left(strLayoutLinha, 13) = "[LINHA_BRANCO" Then
      For intLinha = 1 To Val(Mid(strLayoutLinha, 15))
        Printer.Print
      Next intLinha
    Else
    'Conteúdo
      strPrintLinha = Retorna_Texto(strLayoutLinha)
      '
      'Início da formatação em negrito
      If InStr(strLayoutLinha, "LINHA_EM_NEGRITO") > 0 Then
        Printer.FontBold = True
      End If
      'Imprime
      'Printer.Print CStr(m_intCountPrintLine) 'Teste para imprimir o número da linha
      Printer.Print strPrintLinha
      '
      'Término da formatação em negrito
      If InStr(strLayoutLinha, "LINHA_EM_NEGRITO") > 0 Then
        Printer.FontBold = False
      End If
    End If
    '
    '02/06/2004 - mpdea
    'Tratamento de exceções durante a impressão
    'Omissão de impressão das linhas
    Select Case m_intCountPrintLine
      Case 65, 130
        If m_blnCaseEmbalavi And Not IsWindowsNT Then Printer.Print
    End Select
    '
  Next intLayoutLinha
  
End Sub

'Obtém a quantidade de produtos e serviços na collection
Private Sub GetQtdeProdServLayout(ByVal clcLayout As Collection, _
  ByRef intQtdeProd As Integer, ByRef intQtdeServ As Integer)
  
  Dim intX As Integer
  
  intQtdeProd = 0: intQtdeServ = 0
  
  For intX = 1 To clcLayout.Count
    If InStr(clcLayout.Item(intX), "[PROXIMO_PRODUTO,1]") Then
      intQtdeProd = intQtdeProd + 1
    End If
    If InStr(clcLayout.Item(intX), "[PROXIMO_SERVIÇO,1]") Then
      intQtdeServ = intQtdeServ + 1
    End If
  Next intX

End Sub
