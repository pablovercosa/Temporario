Attribute VB_Name = "modLicencas"
Option Explicit

'------------------------------------------------------------------------------
'19/09/2005 - mpdea
'Cases CheckSerialCaseMod
'Personaliza��es para clientes espec�ficos
'Descri��o na fun��o LoadCases_CheckSerialCaseMod (abaixo)
Public g_blnIndicePrecoEntrada As Boolean
Public g_blnGravaCustoPrecoListaSemIPI As Boolean
Public g_bln5CasasDecimais As Boolean
Public g_blnDiferimento As Boolean
Public g_bln3CasasDecimais As Boolean '30/04/2007 - Anderson - Implementa��o de 3 casas decimais de acordo com o n�mero de s�rie do cliente
Public g_blnInformarNossoNumero As Boolean '05/06/2007 - Anderson - Indica se o � para informar o Nosso Numero, muito utilizado quando se trabalha com boletos pr�-impressos
Public g_blnSadigWeb As Boolean '18/07/2007 - Anderson - Utilizado para exportar dados para SadigWeb
'------------------------------------------------------------------------------

'10/09/2007 - Anderson
'Vari�vel utilizada para determinar se o sistema deve gerar Log
Public g_bolSystemLog As Boolean

'25/09/2007 - Anderson
'Vari�vel utilizada para informar se o cliente utiliza o c�digo de barras na impress�o do Carne
Public g_bolCarneCodigoBarras As Boolean

'19/10/2007 - Anderson
'Customiza��o para verificar o lucro m�nimo permitido no momento de dar o desconto do produto
Public g_bolLucroMinimoClasse As Boolean

'31/10/2007 - Anderson
'Customiza��o de Relat�rio de produtos a comprar para Kings Cross
Public g_bolRelatorioCompra As Boolean

'19/09/2005 - mpdea
'Carrega os cases CheckSerialCaseMod globais
Public Sub LoadCases_CheckSerialCaseMod()
  
  'Data.............: 19/09/2005
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descri��o........: �ndice para c�lculo do Pre�o de Entrada
  'N�meros de s�ries: Centro Tintas (QS38230-471)
  '                   Henrique, Bertamoni & Cia. Ltda. (QS71259-458, QS71260-794)
  g_blnIndicePrecoEntrada = CheckSerialCaseMod("QS38230-471", "QS71259-458", "QS71260-794")
  
  'Data.............: 22/09/2005
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descri��o........: Gravar o pre�o de Custo no campo Pre�o de Lista sem IPI
  '                   utilizado na pasta C�lculos no Cadastro de Produtos
  'N�meros de s�ries: Centro Tintas (QS38230-471)
  '                   Henrique, Bertamoni & Cia. Ltda. (QS71259-458, QS71260-794)
  'Altera��es.......: 17/05/2006 - mpdea
  '                   Liberado personaliza��o para todos os clientes
  g_blnGravaCustoPrecoListaSemIPI = True 'CheckSerialCaseMod("QS38230-471", "QS71259-458", "QS71260-794")
  
  'Data.............: 23/09/2005
  'Desenvolvedor....: mpdea
  'Solicitante......: Embalavi
  'Descri��o........: Formatar valores com 5 casas decimais
  'Observa��es......: Centralizado verifica��o do serial em uma �nica vari�vel
  'N�meros de s�ries: Embalavi ("QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162")
  '                   Celopel ("QS31757-201")
  '                   S�o Miguel Livraria e Legalzices ("QS71315-761")
  '                   17/07/2006 - mpdea
  '                   Inclu�do case Ferbock ("QS31231-478")
  '                   05/09/2006 - Anderson
  '                   Inclus�o de cases:
  '                           QS61086-437 - Auto Servi�o Alves Ltda.
  '                           QS61083-185 - E. R. da Silva Mercadinho
  '                           QS61034-809 - Expedito S. Menezes
  '                           QS61082-101 - Posto Olinda
  '                           QS61085-353 - Wanderley e Claudenier
  '                           QS61094-361 - Petrojal
  '                           QS61089-689 - Revendedora de Combust�veis Santa Maria
  '                           QS61091-109 - Santa Maria Revendedora de Combust�veis
  '                           QS61092-193 - Posto Quatro de Outrubro
  '                           QS61084-269 - Campelo e Pimentel Ltda
  '                           QS61017-877 - Mercadinho POP
  '                           QS71366-305 - 14 BIS
  '                           QS39647-190 - Monteiro e Moraes Parafusos
  '                           QS31753-865 - Monteiro e Moraes Parafusos
  '                           QS73038-198 - Armazem dos Fios Ltda.
  '                           QS34428-021 - Nucamp Nutri��o Animal Ltda
  '                           QS36688-609 - Actel
  '                           QS37243-804 - Agrofarm Importadora e Exportadora de Produtos Veterin�rio LTDA
  '                           QS71124-755 - Almenir A. Agliardi ME
'''  g_bln5CasasDecimais = CheckSerialCaseMod("QS31306-629", "QS31571-867", "QS31572-951", _
'''                                           "QS31581-959", "QS33016-722", "QS33458-286", _
'''                                           "QS37456-162", "QS31757-201", "QS71315-761", _
'''                                           "QS31231-478", "QS61086-437", "QS61083-185", _
'''                                           "QS61034-809", "QS61082-101", "QS61085-353", _
'''                                           "QS61094-361", "QS61089-689", "QS61091-109", _
'''                                           "QS61092-193", "QS61084-269", "QS61017-877", _
'''                                           "QS71366-305", "QS39647-190", "QS31753-865", _
'''                                           "QS73038-198", "QS34428-021", "QS36688-609", _
'''                                           "QS37243-804", "QS72385-416", "QS73520-469")
  
  'Data.............: 30/04/2007
  'Desenvolvedor....: Anderson
  'Solicitante......: Candy Clean
  'Descri��o........: Formatar valores com 3 casas decimais
  'Observa��es......: Centralizado verifica��o do serial em uma �nica vari�vel
  'N�meros de s�ries: QS37957-281 - Candy Clean-Prod. de Equipamentos
  '                   QS71124-755 - Almenir A. Agliardi ME
  '                   QS38649-347 - Big Compra 12/12/2007
  '                   11/04/2008 - mpdea
  '                   QS33398-647 - Joraci Moras Burim - EPP
'''  g_bln3CasasDecimais = CheckSerialCaseMod("QS37957-281", "QS71124-755", "QS38649-347", _
'''                                           "QS33398-647")
  
  'Data.............: 23/05/2006
  'Desenvolvedor....: mpdea
  'Solicitante......: Embalavi
  'Descri��o........: Tratamento do Diferimento
  'Observa��es......: Centralizado verifica��o do serial em uma �nica vari�vel
  'N�meros de s�ries: Embalavi ("QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162")
  '                   JCS Assessoria e Com�rcio Exterior Ltda ("QS73005-670")
  g_blnDiferimento = CheckSerialCaseMod("QS31306-629", "QS31571-867", "QS31572-951", _
                                        "QS31581-959", "QS33016-722", "QS33458-286", _
                                        "QS37456-162", "QS73005-670")

  'Data.............: 05/06/2007
  'Desenvolvedor....: Anderson
  'Solicitante......: Agrotama - Technomax
  'Descri��o........: Informar Nosso Numero na impress�o dos boletos
  'N�meros de s�ries: QS73070-894 - Agrotama
  g_blnInformarNossoNumero = CheckSerialCaseMod("QS73070-894")
  
  'Data.............: 18/07/2007
  'Desenvolvedor....: Anderson
  'Solicitante......: Gurgel & Leite
  'Descri��o........: Exportar dados para o sistema da SadigWeb
  'N�meros de s�ries: QS31734-765 - Gurgel & Leite
  g_blnSadigWeb = CheckSerialCaseMod("QS31734-765")
  
  'Data.............: 10/09/2007
  'Desenvolvedor....: Anderson
  'Solicitante......: Agrotama
  'Descri��o........: Gera arquivo log em arquivo texto.
  'N�meros de s�ries: QS73070-894 - Agrotama
  '                   QS34903-452 - Thomazelli Filhas e Cia Ltda.
  g_bolSystemLog = CheckSerialCaseMod("QS73070-894", "QS34903-452")
  
  'Data.............: 25/09/2007
  'Solicitante......: Naativa
  'Desenvolvedor....: Anderson
  'Descri��o........: Otimizar o pagamento das mensalidades do carn� atrav�s do c�digo de barras
  'N�meros de s�ries: QS73159-473 Naativa
  '                   QS39820-432 Centro Visual Comandulli Ltda - 22/10/2007 - Anderson
  '                   QS73303-523 Centro Visual Comandulli Ltda - 14/10/2008 - mpdea
  '                   QS71388-657 NOEDEL CALCADOS E CONFECCOES LTDA - 09/11/2007 - Anderson
  '                   QS73097-666 TERRA K ARTIGOS ESPORTIVOS LTDA - 09/11/2007 - Anderson
  '                   QS71370-893 DARTORA COMERCIO E VESTUARIO LTDA - 09/11/2007 - Anderson
  '                   QS73200-264 IVONE DE OLIVEIRA SILVA CONFEC��ES LTDA - 09/11/2007 - Anderson
  '                   QS73145-045 TIFFANY CONFEC��ES LTDA - 09/11/2007 - Anderson
  '                   QS73147-213 MODA LAR LTDA - 09/11/2007 - Anderson
'''  g_bolCarneCodigoBarras = CheckSerialCaseMod("QS73159-473", "QS39820-432", "QS73303-523", _
'''                                              "QS71388-657", "QS73097-666", "QS71370-893", _
'''                                              "QS73200-264", "QS73145-045", "QS73147-213")
  
  'Data.............: 19/10/2007
  'Solicitante......: Agrotama
  'Desenvolvedor....: Anderson
  'Descri��o........: Verificar o lucro m�nimo permitido no momento de dar desconto na tela de venda
  'N�meros de s�ries: QS73070-894 - Agrotama
  g_bolLucroMinimoClasse = CheckSerialCaseMod("QS73070-894")
  
  'Data.............: 30/10/2007
  'Solicitante......: Kings Cross
  'Desenvolvedor....: Anderson
  'Descri��o........: Habilita relat�rio de produtos a comprar
  'N�meros de s�ries: QS38393-282 Kings Cross - Matriz
  '                   QS38714-658 Kings Cross - Filial
  g_bolRelatorioCompra = CheckSerialCaseMod("QS38393-282", "QS38714-658")

End Sub

'29/07/2003 - mpdea
'Verifica se o n�mero de s�rie informado est� registrado em QuickStore.lic
Public Function CheckSerialCaseMod(ParamArray CheckSerial() As Variant) As Boolean
  
  #If BETA = 1 Then
    CheckSerialCaseMod = True
    Exit Function
  #End If
  
  
  Dim intFreeFile As Integer
  Dim intX As Integer
  Dim intSerial As Integer
  Dim strSerialRegistrado() As String
  Dim strSerialLinha() As String
  Dim strLinha As String
'  Dim objQuickinfo As QuickInfo.IQuickInfo
  
  
  On Error GoTo ErrHandler
  
  ReDim strSerialRegistrado(0)
  'Set objQuickinfo = New QuickInfo.QuickInfoCls
  
  'PILATTI/MAURO 2018-SETEMBRO-23 COMENTAMOS (DESCOMENTADO DIA 24 POIS DA ERRO NO TICKET N�O FISCAL NA D EMBALAGENS DE PG)
  'Carrega os n�meros de s�ries registrados
  intFreeFile = FreeFile
  Open gsDefaultPath & "QuickStore.lic" For Input As #intFreeFile
  Do Until EOF(intFreeFile)
    Line Input #intFreeFile, strLinha
    'QX00000-000
    If Len(strLinha) > 11 Then
      strSerialLinha() = Split(strLinha, " ")
      'Valida o nr. de s�rie
'      If objQuickinfo.IsValidLiberacao( _
'        gsNomeEmpresa, gsCGCCPF, strSerialLinha(0), strSerialLinha(1)) Then

        ReDim Preserve strSerialRegistrado(intX)
        strSerialRegistrado(intX) = strSerialLinha(0)
        intX = intX + 1
'      End If
    End If
  Loop
  Close #intFreeFile
  'Set objQuickinfo = Nothing

  'Compara o serial para verifica��o
  For intSerial = LBound(CheckSerial) To UBound(CheckSerial)
    For intX = LBound(strSerialRegistrado) To UBound(strSerialRegistrado)

      If CheckSerial(intSerial) = strSerialRegistrado(intX) Then
        'N�mero v�lido
        CheckSerialCaseMod = True
        Exit Function
      End If

    Next intX
  Next intSerial
  
  ' pilatti/mauro incluimos esta linha
  'CheckSerialCaseMod = True
  
  Exit Function
  
ErrHandler:
  Close
  CheckSerialCaseMod = False
  
End Function

Public Function gbConsoleLicencas(ByVal sPrefixes As String) As Boolean
  Dim oQuickInfo As IQuickInfo
  Dim rs As Recordset
  Dim sTexto As String
  
  If Dir(gsConsLicFileName) = "" Then
    gsTitle = LoadResString(101)
    gsMsg = LoadResString(230) & gsConsLicFileName
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    gbProdutoRegistrado = False
    gbConsoleLicencas = False
    Exit Function
  End If
  
  Call ExecCmd(gsConsLicFileName)

  Set rs = db.OpenRecordset("ZZZ", dbOpenDynaset)
  On Error Resume Next
  gsNomeEmpresa = rs("Nome")
  gsCGCCPF = rs("CGCCPF")
  rs.Close
  Set rs = Nothing
  
  'Carrega Licen�as do Produto e o �ltimo N�mero de S�rie do Produto
  gnMaxUsers = gnReadQuickLic(sPrefixes)
  'Verifica se � uma vers�o de demonstra��o
  Set oQuickInfo = New QuickInfoCls
  gbDemoVersion = oQuickInfo.IsDemoVersion(gsNumSerie)
  Set oQuickInfo = Nothing
  
  gbProdutoRegistrado = IsProdutoRegistrado()
  Call GetMDIMainCaption
  frmMain.Caption = LoadResString(5) & " " & gsMainCaption
  
  If IsProdutoRegistrado() Then
    If Not gbDemoVersion Then
      If InStr(1, gsCGCCPF, "/") > 0 Then
        sTexto = " (CNPJ: "
      Else
        sTexto = " (CPF: "
      End If
      gsNomeEmpresa = Trim(rs("Nome").Value) & sTexto & gsCGCCPF & ")"
    Else
      gsNomeEmpresa = Trim(rs("Nome").Value) & " (VERS�O DE DEMONSTRA��O)"
    End If
  Else
    If gbDemoVersion Then
      gsNomeEmpresa = Trim(rs("Nome").Value) & " (VERS�O DE DEMONSTRA��O)"
    Else
      gsNomeEmpresa = Trim(rs("Nome").Value) & " (C�PIA N�O REGISTRADA)"
    End If
  End If
  
  Call GetGlobals
  
  gbConsoleLicencas = True

End Function

Public Function gnReadQuickLic(ByVal sPrefixes As String) As Integer
  Dim nFreeFile As Integer
  Dim sRecord As String
  Dim nCtLic As Integer
  Dim sTomo() As String
  Dim sPrefix As String
  Dim sNumAux As String
  Dim oQuickInfo As IQuickInfo
  Dim sNumSerie() As String
  Dim sNumTest As String
  Dim nI As Integer
  Dim nJ As Integer
  
  On Error GoTo ErrRead
  
  'sPrefixes = "QS", "QF", etc...
  
  Set oQuickInfo = New QuickInfoCls
  
  gsNumSerie = ""
  
  If Dir(gsQuickLicFileName) = "" Then
    gsTitle = LoadResString(224)
    gsMsg = LoadResString(225) & "'" & gsQuickLicFileName & "'"
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    gnReadQuickLic = 0
    Exit Function
  End If
  
  nFreeFile = FreeFile
  Open gsQuickLicFileName For Input As #nFreeFile
  
  nI = -1
  
  Do While Not EOF(nFreeFile)
    DoEvents
    Line Input #nFreeFile, sRecord
    sRecord = Trim(sRecord)
    If Len(Trim(sRecord)) > 0 Then
      If Mid(sRecord, 1, 1) <> ";" Then
        sTomo = Split(sRecord, " ", -1, vbTextCompare)
        nI = nI + 1
        ReDim Preserve sNumSerie(nI) As String
        sNumSerie(nI) = sTomo(0)
      End If
    End If
  Loop
  
  
  '10/02/2003 - mpdea
  'Adicionado verifica��o de conv�nios diferentes
  If nI <> -1 Then
    For nI = 0 To UBound(sNumSerie) - 1
      sNumTest = sNumSerie(nI)
      For nJ = nI + 1 To UBound(sNumSerie)
        If sNumTest = sNumSerie(nJ) Or _
          (gintGetConvenio(sNumTest) <> gintGetConvenio(sNumSerie(nJ))) Then
        
          gsTitle = LoadResString(104)
          gsMsg = "Tabela de N�meros de S�ries inconsistente. Verifique valores e valide as licen�as atuais."
          gnStyle = vbOKOnly + vbExclamation
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
        End If
      Next nJ
    Next nI
  End If
  
  Close #nFreeFile
  Open gsQuickLicFileName For Input As #nFreeFile
  
  nCtLic = 0
  Do While Not EOF(nFreeFile)
    DoEvents
    Line Input #nFreeFile, sRecord
    sRecord = Trim(sRecord)
    If Mid(sRecord, 1, 1) <> ";" Then
      sTomo = Split(sRecord, " ", -1, vbTextCompare)
      If UBound(sTomo) = 1 Then
        sPrefix = Left(sTomo(0), 2)
        If InStr(sPrefixes, sPrefix) > 0 Then
          If oQuickInfo.IsValidNumSerie(sPrefix, Mid(sTomo(0), 3, 9)) = True Then
            gbDemoVersion = oQuickInfo.IsDemoVersion(sTomo(0))
            If Not gbDemoVersion Then
              If oQuickInfo.IsValidLiberacao(gsNomeEmpresa, gsCGCCPF, sTomo(0), sTomo(1)) = True Then
                sNumAux = sTomo(0)
                nCtLic = nCtLic + 1
              End If
            Else
              If nCtLic = 0 Then
                sNumAux = sTomo(0)
                nCtLic = nCtLic + 1
              End If
            End If
          End If
        End If
      End If
    End If
  Loop
  Close #nFreeFile
  
  If Len(sNumAux) > 0 Then
    gsNumSerie = sNumAux
  Else
    nCtLic = 0
  End If
  
  
  '27/01/2003 - mpdea
  'Quick Store n�o registrado - verifica serial inicial para o conv�nio
  If nCtLic = 0 And UBound(sTomo) = 0 Then
    If sTomo(0) <> "" Then
      sNumAux = Trim(sTomo(0))
      sPrefix = Left(sNumAux, 2)
      If oQuickInfo.IsValidNumSerie(sPrefix, Mid(sNumAux, 3, 9)) Then
        gsNumSerie = sNumAux
      End If
    End If
  End If
  
  
  gnReadQuickLic = nCtLic
  'gnMaxUsers = nCtLic
  
  Exit Function
  
ErrRead:
  gnReadQuickLic = -1

End Function

Public Function IsProdutoRegistrado() As Boolean
  IsProdutoRegistrado = (gnMaxUsers > 0) Or gbDemoVersion
End Function

Public Sub GetMDIMainCaption()
'''  If IsProdutoRegistrado() Then
'''    If Not gbDemoVersion Then
'''      gsMainCaption = " [" & LoadResString(10) & " " & CStr(gnMaxUsers) & " Usu�rios]"
'''    Else
'''      gsMainCaption = " - " & LoadResString(12)
'''    End If
'''  Else
'''    If gbDemoVersion Then
'''      gsMainCaption = " - " & LoadResString(12)
'''    Else
'''      gsMainCaption = " - " & LoadResString(11)
'''    End If
'''  End If
  gsMainCaption = ""
End Sub
