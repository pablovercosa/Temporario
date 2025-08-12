VERSION 5.00
Begin VB.Form frmXML 
   Appearance      =   0  'Flat
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Quick Manager NFe - Power"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmXML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   14070
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_atalho 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Transmitir XML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   11880
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3780
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.CommandButton cmd_formatarVisualXML 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Melhorar Visualização"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   9360
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2250
      Width           =   2130
   End
   Begin VB.CommandButton cmd_localizaErro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Localizar Erro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1620
      Width           =   11490
   End
   Begin VB.TextBox txt_xmlErro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   470
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   11490
   End
   Begin VB.CommandButton cmd_regras 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Regras e Críticas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   11880
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1665
      Width           =   2130
   End
   Begin VB.CommandButton cmd_modelos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Modelos em XML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   11880
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   2130
   End
   Begin VB.CommandButton cmd_transmitirXML 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Transmitir XML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   11880
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2835
      Width           =   2130
   End
   Begin VB.CommandButton cmd_atualizarXML 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Atualizar XML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   11880
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2250
      Width           =   2130
   End
   Begin VB.TextBox tb_xml 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6540
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "frmXML.frx":4E95A
      Top             =   2160
      Width           =   11715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFA324&
      Caption         =   "Quick Store Manager NFe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F7F7F7&
      Height          =   780
      Left            =   0
      TabIndex        =   8
      Top             =   45
      Width           =   14100
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione o texto que você quer localizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   855
      Width           =   5055
   End
End
Attribute VB_Name = "frmXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sXML As String
Public sXML_Erro As String
Public xNomeArquivoXML As String
Public sSequencia As String
Public iOrigemChamador As Integer  ' 1-Tela frmNFe Aba Erros/Críticas     2-Tela frmNFe Aba Notas Fiscais

Private Sub cmd_atalho_Click()
On Error GoTo Erro:

  Dim objNFe As clsNFe
  Dim iIndice As Integer
  Dim iIndice2 As Integer
  Dim arquivoLote As String
  Dim strSQL As String
  Dim sNumNFe As String
  Dim sSerieNFe As String
  Dim sCNPJ_EmitenteXML As String
  
  Set objNFe = New clsNFe
  
'  Dim ff As Integer
'  ff = FreeFile
'  Open xNomeArquivoXML For Input As #ff
'
'  Dim Linha As String
'  While EOF(ff) = False
'      Linha = ""
'      Line Input #ff, Linha
'      arquivoLote = arquivoLote + Linha
'  Wend
'  Close #ff

  arquivoLote = tb_xml.Text
  arquivoLote = Replace(arquivoLote, vbCrLf, "")
  arquivoLote = Replace(arquivoLote, "          ", " ")
  arquivoLote = Replace(arquivoLote, "          ", " ")
  arquivoLote = Replace(arquivoLote, "          ", " ")
  arquivoLote = Replace(arquivoLote, "         ", " ")
  arquivoLote = Replace(arquivoLote, "        ", " ")
  arquivoLote = Replace(arquivoLote, "       ", " ")
  arquivoLote = Replace(arquivoLote, "      ", " ")
  arquivoLote = Replace(arquivoLote, "     ", " ")
  arquivoLote = Replace(arquivoLote, "     ", " ")
  arquivoLote = Replace(arquivoLote, "     ", " ")
  arquivoLote = Replace(arquivoLote, "    ", " ")
  
  'Obter os parametros necessários para chamar o método de envio
  objNFe.sXML_40 = arquivoLote
  arquivoLote = Replace(arquivoLote, "UTF-16", "UTF-8")
  
  '<CNPJ>04152403000107</CNPJ>
  iIndice = InStr(1, arquivoLote, "<CNPJ>")
  iIndice2 = InStr(iIndice + 6, arquivoLote, "</CNPJ>")
  sCNPJ_EmitenteXML = Mid(arquivoLote, iIndice + 6, iIndice2 - (iIndice + 6))
  
  '<nNF>5596</nNF>
  iIndice = InStr(1, arquivoLote, "<nNF>")
  iIndice2 = InStr(iIndice + 5, arquivoLote, "</nNF>")
  sNumNFe = Mid(arquivoLote, iIndice + 5, iIndice2 - (iIndice + 5))
  
  '<serie>1</serie>
  iIndice = InStr(1, arquivoLote, "<serie>")
  iIndice2 = InStr(iIndice + 7, arquivoLote, "</serie>")
  sSerieNFe = Mid(arquivoLote, iIndice + 7, iIndice2 - (iIndice + 7))
  
  sSequencia = 1
  
  '************* tratamento atalho provisorio
  ' ROSIBRAS ATACADISTA = 80262645000131
  ' ROSIBRAS MERCADO = 11613307000184
  If sCNPJ_EmitenteXML <> "80262645000131" And sCNPJ_EmitenteXML <> "11613307000184" Then
  
      MsgBox "CNPJ Emitente sem permissão!", vbInformation, "Atenção"
      Exit Sub
  End If
  '*************
  
  objNFe.EnviarXML_SEFAZ sSequencia, sNumNFe, sSerieNFe, "55", gnCodFilial, sCNPJ_EmitenteXML

  Set objNFe = Nothing
  
  MsgBox "XML NFe transmitido com sucesso!", vbInformation, "Sucesso na Transmissão"
  
  Exit Sub
Erro:
  MsgBox "Erro na transmissão do XML NFe. Descrição Erro: " & Err.Description, vbCritical, "Erro na Transmissão"
End Sub

Private Sub cmd_atualizarXML_Click()
On Error GoTo Erro:
    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    
    'Abrir arquivo .xml (para escrita , se não existir cria)
    Set ts = fso.OpenTextFile(xNomeArquivoXML, ForWriting, True)
    
    tb_xml.Text = Replace(tb_xml.Text, "UTF-16", "UTF-8")
    tb_xml.Text = Replace(tb_xml.Text, "standalone=""no""", "")
    
    ts.Write tb_xml.Text
    ts.Close
    Set ts = Nothing
    
    MsgBox "Arquivo XML atualizado com sucesso.", vbExclamation
    
    Exit Sub
Erro:
  MsgBox "Erro ao tentar atualizar o arquivo XML " & xNomeArquivoXML & ". Descrição do Erro: " & Err.Description, vbCritical

End Sub

Private Sub cmd_formatarVisualXML_Click()
  Dim lngContavbCrLf As Long
  Dim X As Long
  Dim iIndice As Long
  Dim iIndice2 As Long
  Dim sXML_formatar As String
  Dim sXML_formatarAux As String
  Dim i As Integer
  
  sXML_formatarAux = ""
  sXML_formatar = tb_xml.Text
  
  X = 1
  While InStr(X, sXML_formatar, "</") > 0
      iIndice = InStr(X, sXML_formatar, "</")
      iIndice2 = InStr(iIndice, sXML_formatar, ">")
      
      sXML_formatarAux = Mid(sXML_formatar, 1, iIndice2 + 1) & vbCrLf & Mid(sXML_formatar, iIndice2 + 1, Len(sXML_formatar) - iIndice2)
      
      sXML_formatar = sXML_formatarAux
      X = iIndice2
  Wend
  
  If InStr(1, sXML_formatar, "<prod>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<prod>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 6) & vbCrLf & Mid(sXML_formatar, iIndice + 6, Len(sXML_formatar) - (iIndice + 5))
      
      For i = 0 To 30
        'busca mais um
        If InStr(iIndice + 6, sXML_formatar, "<prod>") > 0 Then
            iIndice = InStr(iIndice + 6, sXML_formatar, "<prod>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 6) & vbCrLf & Mid(sXML_formatar, iIndice + 6, Len(sXML_formatar) - (iIndice + 5))
        End If
      Next i
  End If

  If InStr(1, sXML_formatar, "<emit>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<emit>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 6) & vbCrLf & Mid(sXML_formatar, iIndice + 6, Len(sXML_formatar) - (iIndice + 5))
  End If

  If InStr(1, sXML_formatar, "<ICMS>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<ICMS>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 6) & vbCrLf & Mid(sXML_formatar, iIndice + 6, Len(sXML_formatar) - (iIndice + 5))
  
      For i = 0 To 30
        'busca mais um
        If InStr(iIndice + 6, sXML_formatar, "<ICMS>") > 0 Then
            iIndice = InStr(iIndice + 6, sXML_formatar, "<ICMS>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 6) & vbCrLf & Mid(sXML_formatar, iIndice + 6, Len(sXML_formatar) - (iIndice + 5))
        End If
      Next i
  End If

  If InStr(1, sXML_formatar, "<dest>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<dest>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 6) & vbCrLf & Mid(sXML_formatar, iIndice + 6, Len(sXML_formatar) - (iIndice + 5))
  End If

  If InStr(1, sXML_formatar, "<enderEmit>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<enderEmit>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 11) & vbCrLf & Mid(sXML_formatar, iIndice + 11, Len(sXML_formatar) - (iIndice + 10))
  End If

  If InStr(1, sXML_formatar, "<enderDest>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<enderDest>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 11) & vbCrLf & Mid(sXML_formatar, iIndice + 11, Len(sXML_formatar) - (iIndice + 10))
  End If
  
  If InStr(1, sXML_formatar, "<PIS>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<PIS>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 5) & vbCrLf & Mid(sXML_formatar, iIndice + 5, Len(sXML_formatar) - (iIndice + 4))
  
      For i = 0 To 30
        'busca mais um
        If InStr(iIndice + 5, sXML_formatar, "<PIS>") > 0 Then
            iIndice = InStr(iIndice + 5, sXML_formatar, "<PIS>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 5) & vbCrLf & Mid(sXML_formatar, iIndice + 5, Len(sXML_formatar) - (iIndice + 4))
        End If
      Next i
  End If
  
  If InStr(1, sXML_formatar, "<PISNT>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<PISNT>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 7) & vbCrLf & Mid(sXML_formatar, iIndice + 7, Len(sXML_formatar) - (iIndice + 6))
  
      For i = 0 To 30
        'busca mais um
        If InStr(iIndice + 7, sXML_formatar, "<PISNT>") > 0 Then
            iIndice = InStr(iIndice + 7, sXML_formatar, "<PISNT>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 7) & vbCrLf & Mid(sXML_formatar, iIndice + 7, Len(sXML_formatar) - (iIndice + 6))
        End If
      Next i
  End If
  
  
  If InStr(1, sXML_formatar, "<imposto>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<imposto>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 9) & vbCrLf & Mid(sXML_formatar, iIndice + 9, Len(sXML_formatar) - (iIndice + 8))
  
      For i = 0 To 30
        'busca mais um
        If InStr(iIndice + 9, sXML_formatar, "<imposto>") > 0 Then
            iIndice = InStr(iIndice + 9, sXML_formatar, "<imposto>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 9) & vbCrLf & Mid(sXML_formatar, iIndice + 9, Len(sXML_formatar) - (iIndice + 8))
        End If
      Next i
  End If
  
  If InStr(1, sXML_formatar, "<ICMSSN102>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<ICMSSN102>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 11) & vbCrLf & Mid(sXML_formatar, iIndice + 11, Len(sXML_formatar) - (iIndice + 10))
  
      For i = 0 To 30
        'busca mais um
        If InStr(iIndice + 11, sXML_formatar, "<ICMSSN102>") > 0 Then
            iIndice = InStr(iIndice + 11, sXML_formatar, "<ICMSSN102>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 11) & vbCrLf & Mid(sXML_formatar, iIndice + 11, Len(sXML_formatar) - (iIndice + 10))
        End If
      Next i
  End If
  
  
  If InStr(1, sXML_formatar, "<dup>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<dup>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 5) & vbCrLf & Mid(sXML_formatar, iIndice + 5, Len(sXML_formatar) - (iIndice + 4))
  
      For i = 0 To 12 '12 dup chega
        'busca mais um
        If InStr(iIndice + 5, sXML_formatar, "<dup>") > 0 Then
            iIndice = InStr(iIndice + 5, sXML_formatar, "<dup>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 5) & vbCrLf & Mid(sXML_formatar, iIndice + 5, Len(sXML_formatar) - (iIndice + 4))
        End If
      Next i
  End If
  
  If InStr(1, sXML_formatar, "<COFINS>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<COFINS>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 8) & vbCrLf & Mid(sXML_formatar, iIndice + 8, Len(sXML_formatar) - (iIndice + 7))
  
      For i = 0 To 30
        'busca mais um
        If InStr(iIndice + 8, sXML_formatar, "<COFINS>") > 0 Then
            iIndice = InStr(iIndice + 8, sXML_formatar, "<COFINS>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 8) & vbCrLf & Mid(sXML_formatar, iIndice + 8, Len(sXML_formatar) - (iIndice + 7))
        End If
      Next i
  End If
  
  If InStr(1, sXML_formatar, "<COFINSNT>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<COFINSNT>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 10) & vbCrLf & Mid(sXML_formatar, iIndice + 10, Len(sXML_formatar) - (iIndice + 9))
  
      For i = 0 To 30
        'busca mais um
        If InStr(iIndice + 10, sXML_formatar, "<COFINSNT>") > 0 Then
            iIndice = InStr(iIndice + 10, sXML_formatar, "<COFINSNT>")
            sXML_formatar = Mid(sXML_formatar, 1, iIndice + 10) & vbCrLf & Mid(sXML_formatar, iIndice + 10, Len(sXML_formatar) - (iIndice + 9))
        End If
      Next i
  End If
  
  If InStr(1, sXML_formatar, "<ICMSTot>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<ICMSTot>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 9) & vbCrLf & Mid(sXML_formatar, iIndice + 9, Len(sXML_formatar) - (iIndice + 8))
  End If
  
  If InStr(1, sXML_formatar, "<total>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<total>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 7) & vbCrLf & Mid(sXML_formatar, iIndice + 7, Len(sXML_formatar) - (iIndice + 6))
  End If
  
  If InStr(1, sXML_formatar, "<transp>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<transp>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 8) & vbCrLf & Mid(sXML_formatar, iIndice + 8, Len(sXML_formatar) - (iIndice + 7))
  End If
  
  If InStr(1, sXML_formatar, "<transporta>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<transporta>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 12) & vbCrLf & Mid(sXML_formatar, iIndice + 12, Len(sXML_formatar) - (iIndice + 11))
  End If

  If InStr(1, sXML_formatar, "<detPag>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<detPag>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 8) & vbCrLf & Mid(sXML_formatar, iIndice + 8, Len(sXML_formatar) - (iIndice + 7))
  
      'busca mais um
      If InStr(iIndice + 8, sXML_formatar, "<detPag>") > 0 Then
          iIndice = InStr(iIndice + 8, sXML_formatar, "<detPag>")
          sXML_formatar = Mid(sXML_formatar, 1, iIndice + 8) & vbCrLf & Mid(sXML_formatar, iIndice + 8, Len(sXML_formatar) - (iIndice + 7))
      End If
  End If
  
  If InStr(1, sXML_formatar, "<infAdic>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<infAdic>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 9) & vbCrLf & Mid(sXML_formatar, iIndice + 9, Len(sXML_formatar) - (iIndice + 8))
  End If
  
  'www.portalfiscal.inf.br/nfe">
  If InStr(1, sXML_formatar, "cal.inf.br/nfe") > 0 Then
      iIndice = InStr(1, sXML_formatar, "cal.inf.br/nfe")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  
      'www.portalfiscal.inf.br/nfe">  busca mais um
      If InStr(iIndice + 16, sXML_formatar, "cal.inf.br/nfe") > 0 Then
          iIndice = InStr(iIndice + 16, sXML_formatar, "cal.inf.br/nfe")
          sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
      End If
  End If

  If InStr(1, sXML_formatar, "<vol>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<vol>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 5) & vbCrLf & Mid(sXML_formatar, iIndice + 5, Len(sXML_formatar) - (iIndice + 4))
  End If
  If InStr(1, sXML_formatar, "<pag>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<pag>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 5) & vbCrLf & Mid(sXML_formatar, iIndice + 5, Len(sXML_formatar) - (iIndice + 4))
  End If
  If InStr(1, sXML_formatar, "<ide>") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<ide>")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 5) & vbCrLf & Mid(sXML_formatar, iIndice + 5, Len(sXML_formatar) - (iIndice + 4))
  End If

  '<det nItem="1">
  If InStr(1, sXML_formatar, "<det nItem=""1"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""1"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""2"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""2"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""3"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""3"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""4"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""4"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""5"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""5"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""6"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""6"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""7"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""7"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""8"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""8"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""9"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""9"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 15) & vbCrLf & Mid(sXML_formatar, iIndice + 15, Len(sXML_formatar) - (iIndice + 14))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""10"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""10"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""11"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""11"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""12"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""12"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""13"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""13"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""14"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""14"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""15"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""15"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""16"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""16"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""17"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""17"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""18"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""18"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""19"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""19"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If

  If InStr(1, sXML_formatar, "<det nItem=""20"">") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<det nItem=""20"">")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 16) & vbCrLf & Mid(sXML_formatar, iIndice + 16, Len(sXML_formatar) - (iIndice + 15))
  End If



  '<infNFe versao="4.00" Id="NFe41180904152403000107550010000056361151043126">
  If InStr(1, sXML_formatar, "<infNFe versao=""4.00"" Id=") > 0 Then
      iIndice = InStr(1, sXML_formatar, "<infNFe versao=""4.00"" Id=")
      sXML_formatar = Mid(sXML_formatar, 1, iIndice + 75) & vbCrLf & Mid(sXML_formatar, iIndice + 75, Len(sXML_formatar) - (iIndice + 74))
  End If
 
  tb_xml.Text = sXML_formatar
  
End Sub

Private Sub cmd_localizaErro_Click()
  Dim sTextoSelecionado As String
  Dim lngIndice As Long
  sTextoSelecionado = LTrim(RTrim(txt_xmlErro.SelText))

  If sTextoSelecionado <> "" Then
  
    lngIndice = InStr(1, sXML, sTextoSelecionado)
    If lngIndice > 0 Then
      tb_xml.SetFocus
      tb_xml.SelStart = lngIndice - 1
      tb_xml.SelLength = Len(sTextoSelecionado)
    End If
  End If

End Sub

Private Sub cmd_sair_Click()
  Unload Me
End Sub

Private Sub cmd_transmitirXML_Click()
On Error GoTo Erro:

  Dim objNFe As clsNFe
  Dim iIndice As Integer
  Dim iIndice2 As Integer
  Dim arquivoLote As String
  Dim strSQL As String
  Dim sNumNFe As String
  Dim sSerieNFe As String
  Dim sCNPJ_EmitenteXML As String
  
  Set objNFe = New clsNFe
  
  Dim ff As Integer
  ff = FreeFile
  Open xNomeArquivoXML For Input As #ff

  Dim Linha As String
  While EOF(ff) = False
      Linha = ""
      Line Input #ff, Linha
      arquivoLote = arquivoLote + Linha
  Wend
  Close #ff
  
  'Obter os parametros necessários para chamar o método de envio
  objNFe.sXML_40 = arquivoLote
  arquivoLote = Replace(arquivoLote, "UTF-16", "UTF-8")
  
  '<CNPJ>04152403000107</CNPJ>
  iIndice = InStr(1, arquivoLote, "<CNPJ>")
  iIndice2 = InStr(iIndice + 6, arquivoLote, "</CNPJ>")
  sCNPJ_EmitenteXML = Mid(arquivoLote, iIndice + 6, iIndice2 - (iIndice + 6))
  
  '<nNF>5596</nNF>
  iIndice = InStr(1, arquivoLote, "<nNF>")
  iIndice2 = InStr(iIndice + 5, arquivoLote, "</nNF>")
  sNumNFe = Mid(arquivoLote, iIndice + 5, iIndice2 - (iIndice + 5))
  
  '<serie>1</serie>
  iIndice = InStr(1, arquivoLote, "<serie>")
  iIndice2 = InStr(iIndice + 7, arquivoLote, "</serie>")
  sSerieNFe = Mid(arquivoLote, iIndice + 7, iIndice2 - (iIndice + 7))
  
  objNFe.EnviarXML_SEFAZ sSequencia, sNumNFe, sSerieNFe, "55", gnCodFilial, sCNPJ_EmitenteXML

  Set objNFe = Nothing
  
  MsgBox "XML NFe transmitido com sucesso!", vbInformation, "Sucesso na Transmissão"
  
  Exit Sub
Erro:
  MsgBox "Erro na transmissão do XML NFe. Descrição Erro: " & Err.Description, vbCritical, "Erro na Transmissão"
  
End Sub

Private Sub Form_Load()

  If iOrigemChamador = 2 Then   ' 1-Tela frmNFe Aba Erros/Críticas     2-Tela frmNFe Aba Notas Fiscais
      cmd_atualizarXML.Visible = False
      cmd_transmitirXML.Visible = False
  
        
  End If

  tb_xml.Text = sXML
  txt_xmlErro.Text = sXML_Erro
  
  'Se igual a '0'...sem permissao de visualizar esta tela
  If gAbreModuloXML = 0 Then
      tb_xml.Text = ""
      txt_xmlErro.Text = ""
      cmd_localizaErro.Enabled = False
      tb_xml.Enabled = False
      txt_xmlErro.Enabled = False
      cmd_transmitirXML.Visible = False
      cmd_atualizarXML.Visible = False
      cmd_regras.Visible = False
      cmd_modelos.Visible = False
  End If

End Sub

