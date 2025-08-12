VERSION 5.00
Begin VB.Form frmConsultaAcp 
   Caption         =   "Consulta Associação Comercial do Paraná"
   ClientHeight    =   3570
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6090
   Icon            =   "frmConsultaAcp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCheque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3960
      MaxLength       =   12
      TabIndex        =   17
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pessoa"
      Height          =   975
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optJuridica 
         Caption         =   "Jurídica"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFisica 
         Caption         =   "Física"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtValor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   14
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&lar"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Confirmar"
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtDocumento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtAg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   15
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtBanco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   960
      MaxLength       =   3
      TabIndex        =   13
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   975
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optDetalhada 
         Caption         =   "Detalhada"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optSimples 
         Caption         =   "Simplificada"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consulta"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optSicCheque 
         Caption         =   "SIC (CGC) + Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optSic 
         Caption         =   "SIC (CGC)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optScpcCheque 
         Caption         =   "SCPC + Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optScpc 
         Caption         =   "SCPC"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optCheque 
         Caption         =   "Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lblCep 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1320
      TabIndex        =   25
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Cheque"
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Valor"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblMensagem 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label lblDoc 
      Caption         =   "Doc"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Agência"
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   1830
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Banco"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2190
      Width           =   855
   End
End
Attribute VB_Name = "frmConsultaAcp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCliFor As Recordset
Dim rsSaidas As Recordset


Private Sub cmdCancel_Click()
 Inet.Cancel
 Screen.MousePointer = vbDefault
 Unload Me
End Sub

Private Sub cmdConfirma_Click()
 
Dim sString As String
Dim sResposta As String
Dim sUrl As String
Dim sReq As String
Dim nCodErroCon As Long
Dim sMensErro As String
Dim sRet As String
Dim sProxyAdress As String

If txtDocumento.Text = "" Then
     MsgBox "Não se esqueça de preencher o campo CGC/RG.", vbInformation, "Atenção"
     Exit Sub
  End If
  
If txtAg.Text = "" Then
   MsgBox "Não se esqueça de preencher o campo Agência.", vbInformation, "Atenção"
   Exit Sub
End If

If txtBanco.Text = "" Or txtBanco.Text = "0" Then
   MsgBox "Não se esqueça de preencher o campo Banco.", vbInformation, "Atenção"
   Exit Sub
End If

DoEvents
lblMensagem.Caption = "Consultando..."
Screen.MousePointer = vbHourglass


 sRet = GetSetting("QuickStore", "ACP", "Proxy", "")
 sProxyAdress = GetSetting("QuickStore", "ACP", "Adress", "")



 If sRet = "True" Then
    Inet.AccessType = icNamedProxy
    Inet.Proxy = sProxyAdress
    Inet.UserName = "Leandro"
    Inet.Password = "8918"
    
 Else
    Inet.AccessType = icUseDefault
 End If


sUrl = "http://www.acpr.com.br/consultat.asp?strcons="

sString = Monta_String

sReq = sUrl & sString

sResposta = Inet.OpenURL(sReq, icString)
nCodErroCon = Inet.ResponseCode
sMensErro = Inet.ResponseInfo


If nCodErroCon <> 0 Then
   DisplayMsg "Problemas de conexão."
   Exit Sub
End If

MsgBox sResposta, vbInformation, "Resposta Consulta ACP"




Call Mostra_Resp(sResposta)

Screen.MousePointer = vbDefault



End Sub

Private Sub Form_Load()

Dim sCGC As String



Call CenterForm(Me)

optSimples.Value = True
optFisica.Value = True
optJuridica.Value = False



Call DadosdoCliente

txtBanco.Text = frmRecebimento.Grade_Cheque.Columns(0).Text
txtCheque.Text = frmRecebimento.Grade_Cheque.Columns(1).Text
txtValor.Text = Format(frmRecebimento.Grade_Cheque.Columns(3).Text, "###,###,##0.00")

End Sub


Private Sub mnuVerAlt_Click()

frmConfiguraAcp.Show vbModal


End Sub

Private Sub txtAg_GotFocus()
 
 txtAg.SelStart = 0
 txtAg.SelLength = Len(txtAg.Text)
End Sub

Private Sub txtAg_KeyPress(KeyAscii As Integer)
KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub txtBanco_GotFocus()
 
 txtBanco.SelStart = 0
 txtBanco.SelLength = Len(txtBanco.Text)

End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
 KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub txtCheque_GotFocus()

 txtCheque.SelStart = 0
 txtCheque.SelLength = Len(txtCheque.Text)
End Sub

Private Sub txtDocumento_GotFocus()

txtDocumento.SelStart = 0
txtDocumento.SelLength = Len(txtDocumento.Text)
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 45 Or KeyAscii > 57 Then KeyAscii = 0
 
 End Sub

Private Sub txtValor_GotFocus()

 txtValor.SelStart = 0
 txtValor.SelLength = Len(txtValor.Text)
 
 End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub DadosdoCliente()

 
  Dim sDoc As String
  Dim lSeq As Long
  Dim sCep As String
  
  
  Set rsSaidas = db.OpenRecordset("Saídas", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
 
  
  lSeq = frmRecebimento.L_Sequência.Caption
  rsSaidas.Index = "Sequência"
  rsSaidas.Seek "=", gnCodFilial, lSeq
  If rsSaidas.NoMatch Then Exit Sub
  
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", rsSaidas("Cliente")
  If rsCliFor.NoMatch Then Exit Sub
  
  If Not IsNull(rsCliFor("CGC")) Then
      sDoc = rsCliFor("CGC")
      sDoc = Replace(sDoc, ".", "")
      sDoc = Replace(sDoc, "/", "")
      sDoc = Replace(sDoc, "-", "")
      sDoc = Trim(sDoc)
      txtDocumento.Text = sDoc
  End If
    
  If rsCliFor("Física_Jurídica") = "J" Then
     lblDoc.Caption = "CGC."
     optJuridica.Value = True
  Else
     lblDoc.Caption = "RG."
     optFisica.Value = True
  End If
  
  If Not IsNull(rsCliFor("Cep")) Then
    sCep = rsCliFor("Cep")
    sCep = Replace(sCep, ".", "")
    sCep = Replace(sCep, "/", "")
    sCep = Replace(sCep, "-", "")
    sCep = Trim(sCep)
    
    lblCep.Caption = sCep
  End If
  
  rsSaidas.Close
  rsCliFor.Close
  Set rsSaidas = Nothing
  Set rsCliFor = Nothing

End Sub

Public Function Monta_String() As String

  Dim sConn As String
  Dim sSS As String
  Dim sT As String
  Dim sM As String
  Dim sVV As String
  Dim sValor As String
  Dim nI As Integer
  Dim sAux As String
  Dim sAux1 As String
  Dim sData As String
  Dim sTime As String
  Dim sNsu As String
  Dim sBanco As String
  Dim sAg As String
  Dim sCheque As String
  Dim sOrigem As String
  Dim sRetorno As String
  Dim sDoc As String
  Dim nDia As Integer
  Dim sCep As String
  
      
  sConn = "9020" 'código da transação
  sConn = sConn + "3220000108C1800C"
  
  
  If optCheque.Value = True Then
     sSS = "01"
  ElseIf optScpcCheque.Value = True Then
     sSS = "02"
  ElseIf optScpc.Value = True Then
     sSS = "03"
  ElseIf optSic.Value = True Then
     sSS = "04"
  ElseIf optSicCheque.Value = True Then
     sSS = "05"
  End If
  
  sConn = sConn & sSS 'tipo da consulta
  sT = "0"
  sConn = sConn & sT 'tipo de terminal 0= pdv
  
  If optSimples.Value = True Then
     sM = "S"
  Else
     sM = "D"
  End If
  
  sConn = sConn & sM 'tipo de consulta
  
  sVV = "01" 'versão
  
  sConn = sConn & sVV
  
  sValor = txtValor.Text * 100
  
  
  nI = 12 - Len(sValor)
  sAux = String(nI, "0")
  
  sValor = sAux & sValor
  sConn = sConn & sValor
  
  
  sData = Format(Date, "DD/MM")
  sData = Trim(Replace(sData, "/", ""))
  
  sTime = Time
  sTime = Trim(Replace(sTime, ":", ""))
  
  sConn = sConn & sData & sTime
  
  sNsu = "000000"
  
  sConn = sConn & sNsu 'código da transação opcional
  
  
  If optScpc.Value = True Or optSic.Value = True Then
    sBanco = "000"
    sAg = "0000"
    sCheque = "000000000000"
  Else
    sBanco = "00000000000000000000"
    sBanco = sBanco & txtBanco.Text
    sBanco = Right(sBanco, 3)
     
    sAg = "000000000000000000000"
    sAg = sAg & txtAg.Text
    sAg = Right(sAg, 4)
    
    sCheque = "00000000000000000000000000000000000000000"
    sCheque = sCheque & txtCheque.Text
    sCheque = Right(sCheque, 12)
  End If
  
  
  sConn = sConn & "084" & sBanco & sAg & sCheque & "TERMINAL" 'BANCO, AG, CÓDIGO DE IDENTIFICAÇÃO DO TERMINAL
  
  
  sRetorno = GetSetting("QuickStore", "ACP", "Senha", "")
  sOrigem = sRetorno
  
  sConn = sConn & sOrigem
  
  sOrigem = ""
  sRetorno = ""
  sRetorno = GetSetting("QuickStore", "ACP", "User", "")

  sOrigem = sOrigem & Trim(sRetorno) & "  "


  sConn = sConn & sOrigem
  
  If optFisica.Value = True Then
     sDoc = "0122" & txtDocumento.Text
  Else
     sDoc = "0151" & txtDocumento.Text
  End If
  
  sConn = sConn & sDoc 'pega cgc ou rg
  
  sConn = sConn & "076" 'pega código da moeda => 076
  
  sConn = sConn & "008TTTTTTTT" 'pega o tipo de terminal, qq valor começando com 008
  
  
  sAux = ""
  
  
  
  If IsNull(lblCep.Caption) Or lblCep.Caption = "" Then
      sCep = "00000000"
  Else
      sCep = Left(lblCep.Caption, 8)
  End If
  
  sAux = "016" & sCep
  
  sData = ""
  sData = frmRecebimento.Grade_Cheque.Columns(2).Text
  sData = Format(sData, "DDMMYY") 'data do cheque
  
  sAux = sAux & sData
  
  
  sData = ""
  sData = frmRecebimento.Grade_Cheque.Columns(2).Text
  If sData = "" Then
     sData = Date
  End If
  nDia = CDate(sData) - Date 'qte de dias para o cheque
  
  sAux1 = "0000000000000000000"
  sAux1 = sAux1 & nDia
  sAux1 = Right(sAux1, 3)
  
  sAux = sAux & sAux1
  
  
  sAux = sAux & "01" 'qtde de cheques
  
  sConn = sConn & sAux
  
  Monta_String = sConn

End Function

Public Function Mostra_Resp(ByVal sResp As String)


Dim sTrans As String
Dim sMapa As String
Dim sProc As String
Dim sValor As String
Dim sDataEnvio As String
Dim sNsu As String
Dim sHoraResp As String
Dim sDataResp As String
Dim sIf As String
Dim sCodResp As String
Dim sCodOrg As String
Dim sMoeda As String
Dim sTextConsulta As String
Dim sIdConsulta As String


sTrans = Left(sResp, 4)
sMapa = Mid(sResp, 5, 32)
sProc = Mid(sResp, 37, 6)
sValor = Mid(sResp, 43, 12)
sDataEnvio = Mid(sResp, 55, 10)
sNsu = Mid(sResp, 65, 6)
sHoraResp = Mid(sResp, 71, 6)
sDataResp = Mid(sResp, 77, 4)
sIf = Mid(sResp, 81, 10)
sCodResp = Mid(sResp, 91, 2)
sCodOrg = Mid(sResp, 93, 15)
sMoeda = Mid(sResp, 108, 3)


sIdConsulta = Left(sResp, 12)


'mostrar código e mensagem de resposta

If sCodResp = "00" Then
   DisplayMsg "Nada Consta - Não existe restrição."
ElseIf sResp = "03" Then
   DisplayMsg "Estabelecimento não autorizado para acesso."
ElseIf sResp = "43/76" Then
   DisplayMsg "Restrição no sistema consultado."
ElseIf sResp = "89" Then
   DisplayMsg "Mensagem Genérica."
End If

End Function
