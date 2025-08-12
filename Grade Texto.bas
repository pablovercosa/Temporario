Attribute VB_Name = "modGradeTexto"
Option Explicit

Function Conta_Campos(Texto As String) As Integer
 Dim i, J As Integer
 Dim Campos As Integer
 
 i = Len(Texto)
 If i < 3 Then
   Conta_Campos = 0
   Exit Function
 End If
 
 Campos = 0
 For J = 1 To i
  If Mid(Texto, J, 1) = "[" Or Mid(Texto, J, 1) = "{" Then Campos = Campos + 1
 Next J
 
 Conta_Campos = Campos

End Function

Function Separa_Campos(Texto As String, Campo As Integer, Tipo As String) As String
 Dim i, J As Integer
 Dim Campos As Integer
 Dim Somar As Integer
 Dim Nome_Campo As String
 Dim Letra As String
 
  Campos = 0
  Nome_Campo = ""
  Somar = False
  i = Len(Texto)

  For J = 1 To i
    Letra = Mid(Texto, J, 1)
    If Letra = "[" Or Letra = "{" Then
       Campos = Campos + 1
       If Campos = Campo Then
          Somar = True
          If Letra = "[" Then Tipo = "CAMPO"
          If Letra = "{" Then Tipo = "TEXTO"
       End If
    End If
      
    If Letra = "]" Or Letra = "}" Then
       If Somar = True Then
         Separa_Campos = Nome_Campo
         Exit Function
       End If
    End If

    If Letra <> "{" And Letra <> "[" And Letra <> "}" And Letra <> "]" Then
      If Somar = True Then Nome_Campo = Nome_Campo + Letra
    End If
  Next J
    
End Function

Function Separa_Tamanho(Texto1 As String) As Integer

  Dim Texto2 As String
  Dim Tamanho As String
  Dim Somar As Integer
  Dim Letra As String
  Dim i As Integer
  
  Texto2 = Texto1
  Texto1 = ""
  Somar = False
  
  
  For i = 1 To Len(Texto2)
    Letra = Mid(Texto2, i, 1)
    If Letra = "," Then Somar = True
    If Letra <> "," Then
      If Somar = False Then Texto1 = Texto1 + Letra
      If Somar = True Then Tamanho = Tamanho + Letra
    End If
  Next i
  
  Separa_Tamanho = Val(Tamanho)

End Function
