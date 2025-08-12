Attribute VB_Name = "modEMail"
Option Explicit

'30/01/2009 - mpdea
'Funções de email

Public Type ConfigEnvioEmail
  ServidorSmtp As String
  ServidorPop3 As String
  Autenticacao As Boolean
  AutenticacaoPop3 As Boolean
  Usuario As String
  Senha As String
  NomeExibicaoRemetente As String
  EmailRemetente As String
End Type

Public Function LoadConfigEnvioEmail(ByVal intFilial As Integer) As ConfigEnvioEmail
  Dim rstEmail As Recordset
  Dim strSQL As String
  Dim objConfigEnvioEmail As ConfigEnvioEmail
  
  
  On Error GoTo ErrHandler
  
  
  strSQL = "SELECT * FROM Email WHERE Filial = " & intFilial
  Set rstEmail = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstEmail
    If Not (.BOF And .EOF) Then
      objConfigEnvioEmail.ServidorSmtp = .Fields("ServidorSmtp").Value & ""
      objConfigEnvioEmail.ServidorPop3 = .Fields("ServidorPop3").Value & ""
      objConfigEnvioEmail.Autenticacao = .Fields("Autenticacao").Value
      objConfigEnvioEmail.AutenticacaoPop3 = .Fields("AutenticacaoPop3").Value
      objConfigEnvioEmail.Usuario = .Fields("Usuario").Value & ""
      objConfigEnvioEmail.Senha = .Fields("Senha").Value & ""
      objConfigEnvioEmail.NomeExibicaoRemetente = .Fields("NomeExibicaoRemetente").Value & ""
      objConfigEnvioEmail.EmailRemetente = .Fields("EmailRemetente").Value & ""
    End If
    .Close
  End With
  Set rstEmail = Nothing
  
  'Retorno da função
  LoadConfigEnvioEmail = objConfigEnvioEmail
  
  Exit Function
  
ErrHandler:
  'Fecha tabela
  If Not rstEmail Is Nothing Then
    rstEmail.Close
    Set rstEmail = Nothing
  End If
  'Exibe mensagem de erro
  Err.Raise Err.Number
  
End Function

Public Sub EnviarEmailModeloTicket(ByVal strTicket As String, ByVal intFilial As Integer, _
  ByVal lngSequencia As Long, ByVal lngCodigoCliente As Long)
  
  Dim str_message As String
  Dim str_nome As String
  Dim str_email As String
  
  GetEmailDetailsCliFor lngCodigoCliente, str_nome, str_email
  Imprime_Ticket strTicket, intFilial, lngSequencia, True, str_message

  Dim frm_email As New frmEMailEnviar
  
  frm_email.LoadEmail str_nome, str_email, "", str_message
  frm_email.Show

End Sub
