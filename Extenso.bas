Attribute VB_Name = "modExtenso"
Option Explicit


Function Extenso(Valor As Double) As String
  Dim Valor2 As Double
  Dim Str_Extenso, Str_Ret As String
  Dim Str2, Str3, Str4, Str5 As String
  Dim Parte1, Parte2, Parte3 As Long
  Dim Parte_Agora As Integer
  
  Str_Ret = ""
  Str_Extenso = ""
  
  Valor2 = Valor
 
  '01/03/2007 - Anderson
  'Alteração realizada para evitar que o valor 999,90 seja igual a um mil como estava sendo impresso na nota fiscal.
  'Parte1 = Valor
  Parte1 = Int(Valor)
  Parte2 = ((Valor - Int(Parte1)) * 100)
  
  
 ' If Parte1 = 0 Then GoTo Ver_Parte2
  
  If Parte1 > 999999999 Then
     Parte_Agora = Ret_Bilhões(Parte1)
     Str_Ret = Nome_Num(Parte_Agora)
     If Parte_Agora = 1 Then Str_Extenso = Str_Ret + " bilhão "
     If Parte_Agora <> 1 Then Str_Extenso = Str_Ret + " bilhões "
    ' Parte3 = CInt(Parte1 / 1000)
     Parte1 = Parte1 - (CLng(Parte_Agora) * 1000000000)
  End If
  
  If Parte1 > 999999 Then
     Parte_Agora = Ret_Milhões(Parte1)
     Str_Ret = Nome_Num(Parte_Agora)
     If Parte_Agora = 1 Then Str_Extenso = Str_Extenso + Str_Ret + " milhão "
     If Parte_Agora <> 1 Then Str_Extenso = Str_Extenso + Str_Ret + " milhões "
  '   Parte3 = CInt(Parte1 / 1000)
     Parte1 = Parte1 - (CLng(Parte_Agora) * 1000000)
  End If
 
  If Parte1 > 999 Then
     Parte_Agora = Ret_Mil(Parte1)
     Str_Ret = Nome_Num(Parte_Agora)
     If Parte_Agora = 1 Then Str_Extenso = Str_Extenso + Str_Ret + " mil "
     If Parte_Agora <> 1 Then Str_Extenso = Str_Extenso + Str_Ret + " mil "
   '  Parte3 = CInt(Parte1 / 1000)
     Parte1 = Parte1 - (CLng(Parte_Agora) * 1000)
  End If
 
  Parte_Agora = Int(Parte1)
  Str_Ret = Nome_Num(Parte_Agora)
  Str_Extenso = Str_Extenso + Str_Ret
  
  If Int(Valor) = 1 Then Str_Extenso = Str_Extenso + " real"
  If Int(Valor) <> 1 Then Str_Extenso = Str_Extenso + " reais"
  
  
  Str_Ret = Nome_Num(CInt(Parte2))
  
  Extenso = Str_Extenso
  
  If Str_Ret <> "" Then
    Extenso = Extenso + " e " + Str_Ret
    If Parte2 = 1 Then Extenso = Extenso + " centavo"
    If Parte2 <> 1 Then Extenso = Extenso + " centavos"
  End If
  
End Function

Function Nome_Num(Número As Integer) As String
 Dim Num As Integer
 Dim Parte1, Parte2 As Integer
 Dim Str1, Str2, Str3 As String
 Dim Nome_Final As String
 
 Str1 = ""
 Str2 = ""
 Str3 = ""
 
 Num = Número
 
 If Num < 100 Then
    Parte2 = Num
    GoTo Ver_Parte2
 End If
 
 Parte1 = Int(Num / 100)
 Parte2 = Num - (Parte1 * 100)
 
 If Parte1 = 1 Then
     Str1 = "cento"
     If Parte2 = 0 Then Str1 = "cem"
 End If
 
 If Parte1 = 2 Then Str1 = "duzentos"
 If Parte1 = 3 Then Str1 = "trezentos"
 If Parte1 = 4 Then Str1 = "quatrocentos"
 If Parte1 = 5 Then Str1 = "quinhentos"
 If Parte1 = 6 Then Str1 = "seiscentos"
 If Parte1 = 7 Then Str1 = "setecentos"
 If Parte1 = 8 Then Str1 = "oitocentos"
 If Parte1 = 9 Then Str1 = "novecentos"
 
Ver_Parte2:
 Num = Parte2
 Parte1 = Parte2
 Parte1 = Int(Parte1 / 10)
 Parte2 = Parte2 - (Parte1 * 10)
 
 If Parte1 = 1 Then
  If Num = 10 Then Str2 = "dez"
  If Num = 11 Then Str2 = "onze"
  If Num = 12 Then Str2 = "doze"
  If Num = 13 Then Str2 = "treze"
  If Num = 14 Then Str2 = "catorze"
  If Num = 15 Then Str2 = "quinze"
  If Num = 16 Then Str2 = "dezesseis"
  If Num = 17 Then Str2 = "dezessete"
  If Num = 18 Then Str2 = "dezoito"
  If Num = 19 Then Str2 = "dezenove"
 End If
 
 If Parte1 = 2 Then Str2 = "vinte"
 If Parte1 = 3 Then Str2 = "trinta"
 If Parte1 = 4 Then Str2 = "quarenta"
 If Parte1 = 5 Then Str2 = "cinqüenta"
 If Parte1 = 6 Then Str2 = "sessenta"
 If Parte1 = 7 Then Str2 = "setenta"
 If Parte1 = 8 Then Str2 = "oitenta"
 If Parte1 = 9 Then Str2 = "noventa"
 
 If Num < 10 Or Num > 19 Then
  If Parte2 = 1 Then Str3 = "um"
  If Parte2 = 2 Then Str3 = "dois"
  If Parte2 = 3 Then Str3 = "três"
  If Parte2 = 4 Then Str3 = "quatro"
  If Parte2 = 5 Then Str3 = "cinco"
  If Parte2 = 6 Then Str3 = "seis"
  If Parte2 = 7 Then Str3 = "sete"
  If Parte2 = 8 Then Str3 = "oito"
  If Parte2 = 9 Then Str3 = "nove"
 End If
  
  Nome_Final = Str1
  If Str2 <> "" Then
    If Nome_Final <> "" Then Nome_Final = Nome_Final + " e " + Str2
    If Nome_Final = "" Then Nome_Final = Str2
  End If
  
  If Str3 <> "" Then
    If Nome_Final <> "" Then Nome_Final = Nome_Final + " e " + Str3
    If Nome_Final = "" Then Nome_Final = Str3
  End If
 
  Nome_Num = Nome_Final
  

End Function

Function Ret_Bilhões(ByVal Parte As Long) As Integer

 Dim i As Integer
 i = Int(Parte / 1000000000)
 Ret_Bilhões = i
 

End Function


Function Ret_Mil(ByVal Parte As Long) As Integer

 Dim i As Integer
 i = Int(Parte / 1000)
 Ret_Mil = i


End Function

Function Ret_Milhões(ByVal Parte As Long) As Integer

 Dim i As Integer
 i = Int(Parte / 1000000)
 Ret_Milhões = i

End Function


