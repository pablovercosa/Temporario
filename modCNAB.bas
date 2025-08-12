Attribute VB_Name = "modCNAB"
Option Explicit

Public Enum tpCarteiraCobranca
  Bradesco = 9
End Enum


Public Function GetDigitoVerificador_NossoNumero(ByVal NossoNumero As String, ByVal CarteiraCobranca As tpCarteiraCobranca) As String

  Dim intContador As Integer   'Auxilia em estrutura de repeti��o
  Dim strPeso As String        'Utilizado para calcular o d�gito verificador
  Dim intSomaTotal As Integer  'Utilizado para efetuar a soma dos valores do nosso numero
  Dim intResto As Integer      'Utilizado para obter o resto da divis�o para composi��o do digito verificador
  Dim intDigitoVerificador     'Utilizado para gerar o d�gito verificador

  Select Case CarteiraCobranca
  
    Case Bradesco
    
      'Acrescenta zero a esquerda para gera��o do nosso n�mero
      NossoNumero = Right(String(11, "0") & NossoNumero, 11)
      'Acrescenta o n�mero da carteira ao nosso n�mero
      NossoNumero = "09" & Right(String(11, "0") & NossoNumero, 11)
                              
      strPeso = "2765432765432"   'PESO p/ calcular o d�gito verificador

      'Realiza a somat�ria
      For intContador = 1 To Len(strPeso)
        intSomaTotal = intSomaTotal + (Mid(NossoNumero, intContador, 1) * Mid(strPeso, intContador, 1))
      Next intContador
      
      'Obtem o resto da divis�o
      intResto = intSomaTotal Mod 11
      
      'Se o resto da divis�o for igual a 1, considerar o d�gito verificador como "P"
      If intResto = 1 Then
        GetDigitoVerificador_NossoNumero = "P"
      'Se o resto da divis�o for igual a 0, considerar o digito verificador como "0"
      ElseIf intResto = 0 Then
        GetDigitoVerificador_NossoNumero = "0"
      'Caso contr�rio, realiza a subtra��o do divendo com o resto
      Else
        GetDigitoVerificador_NossoNumero = 11 - intResto
      End If

  End Select
  

End Function
