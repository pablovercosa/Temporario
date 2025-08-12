Attribute VB_Name = "modCNAB"
Option Explicit

Public Enum tpCarteiraCobranca
  Bradesco = 9
End Enum


Public Function GetDigitoVerificador_NossoNumero(ByVal NossoNumero As String, ByVal CarteiraCobranca As tpCarteiraCobranca) As String

  Dim intContador As Integer   'Auxilia em estrutura de repetição
  Dim strPeso As String        'Utilizado para calcular o dígito verificador
  Dim intSomaTotal As Integer  'Utilizado para efetuar a soma dos valores do nosso numero
  Dim intResto As Integer      'Utilizado para obter o resto da divisão para composição do digito verificador
  Dim intDigitoVerificador     'Utilizado para gerar o dígito verificador

  Select Case CarteiraCobranca
  
    Case Bradesco
    
      'Acrescenta zero a esquerda para geração do nosso número
      NossoNumero = Right(String(11, "0") & NossoNumero, 11)
      'Acrescenta o número da carteira ao nosso número
      NossoNumero = "09" & Right(String(11, "0") & NossoNumero, 11)
                              
      strPeso = "2765432765432"   'PESO p/ calcular o dígito verificador

      'Realiza a somatória
      For intContador = 1 To Len(strPeso)
        intSomaTotal = intSomaTotal + (Mid(NossoNumero, intContador, 1) * Mid(strPeso, intContador, 1))
      Next intContador
      
      'Obtem o resto da divisão
      intResto = intSomaTotal Mod 11
      
      'Se o resto da divisão for igual a 1, considerar o dígito verificador como "P"
      If intResto = 1 Then
        GetDigitoVerificador_NossoNumero = "P"
      'Se o resto da divisão for igual a 0, considerar o digito verificador como "0"
      ElseIf intResto = 0 Then
        GetDigitoVerificador_NossoNumero = "0"
      'Caso contrário, realiza a subtração do divendo com o resto
      Else
        GetDigitoVerificador_NossoNumero = 11 - intResto
      End If

  End Select
  

End Function
