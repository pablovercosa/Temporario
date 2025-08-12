Attribute VB_Name = "modCGC_CPF"
Option Explicit

Public Function gsRetiraSpecialChars(ByVal sCGCCPF As String) As String
  Dim nI As Integer
  Dim sNum As String
  Dim sCh As String
  sNum = ""
  For nI = 1 To Len(sCGCCPF)
    sCh = Mid(sCGCCPF, nI, 1)
    If InStr("./- ", sCh) = 0 Then
      sNum = sNum & sCh
    End If
  Next nI
  gsRetiraSpecialChars = sNum
End Function

Public Function gnCGCCPFKey(ByVal sCGCCPF As String) As Integer
  Dim nI As Integer
  Dim sV As String
  Dim nSum As Long
  sV = "13576201249181"
  nSum = 0#
  For nI = 1 To Len(sCGCCPF)
    nSum = nSum + CInt(Mid(sCGCCPF, nI, 1)) * CInt(Mid(sV, nI, 1))
  Next nI
  If nSum > 999# Then
    nSum = nSum Mod 999
  End If
  gnCGCCPFKey = CInt(nSum)
End Function

Function Valid_CGC(CGC As String) As Boolean
    Dim i, J, df1, df2, df3, resto1, pridig, segdig, dg(18), Resto
    Dim Letra As String
    
    J = 0
    For i = 0 To (Len(CGC) - 1)
      Letra = Mid(CGC, i + 1, 1)
      If Letra <> "." And Letra <> "-" And Letra <> "/" Then
        dg(J) = Val(Letra)
        J = J + 1
      End If
    Next i
   
    df1 = 5 * dg(0) + 4 * dg(1) + 3 * dg(2) + 2 * dg(3) + 9 * dg(4) + 8 * dg(5) + 7 * dg(6) + 6 * dg(7) + 5 * dg(8) + 4 * dg(9) + 3 * dg(10) + 2 * dg(11)
    df2 = df1 / 11
    df3 = Int(df2) * 11
    Resto = df1 - df3
    If Resto = 0 Or Resto = 1 Then
        pridig = 0
    Else
        pridig = 11 - Resto
    End If

    df1 = 6 * dg(0) + 5 * dg(1) + 4 * dg(2) + 3 * dg(3) + 2 * dg(4) + 9 * dg(5) + 8 * dg(6) + 7 * dg(7) + 6 * dg(8) + 5 * dg(9) + 4 * dg(10) + 3 * dg(11) + 2 * pridig
         
    df2 = df1 / 11
    df3 = Int(df2) * 11
    Resto = df1 - df3
    
    If Resto = 0 Or Resto = 1 Then
        segdig = 0
    Else
        segdig = 11 - Resto
    End If
                   
    If pridig <> dg(12) Or segdig <> dg(13) Then
        Valid_CGC = False
        Exit Function
    End If
    
    Valid_CGC = True
    
End Function

Function Valid_CPF(CPF As String) As Boolean
   Dim i, J, df1, df2, df3, resto1, pridig, segdig, dg(18), Resto
   Dim Letra As String

   J = 0
   For i = 0 To (Len(CPF) - 1)
      Letra = Mid(CPF, i + 1, 1)
      If Letra <> "." And Letra <> "/" And Letra <> "-" Then
        dg(J) = Val(Letra)
        J = J + 1
      End If
   Next

   df1 = 1 * dg(0) + 2 * dg(1) + 3 * dg(2) + 4 * dg(3) + 5 * dg(4) + 6 * dg(5) + 7 * dg(6) + 8 * dg(7) + 9 * dg(8)
   df2 = df1 / 11
   df3 = Int(df2) * 11
   Resto = df1 - df3
   If Resto >= 10 Then
      pridig = 0
   Else
      pridig = Resto
   End If

   df1 = 0 * dg(0) + 1 * dg(1) + 2 * dg(2) + 3 * dg(3) + 4 * dg(4) + 5 * dg(5) + 6 * dg(6) + 7 * dg(7) + 8 * dg(8) + 9 * pridig
         
   df2 = df1 / 11
   df3 = Int(df2) * 11
   Resto = df1 - df3
   
   If Resto >= 10 Then
      segdig = 0
   Else
      segdig = Resto
   End If
                   
   If pridig <> dg(9) Or segdig <> dg(10) Then
      Valid_CPF = False
      Exit Function

   End If

    Valid_CPF = True

End Function

