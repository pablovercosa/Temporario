VERSION 5.00
Begin VB.Form frmXML_NFCe 
   Caption         =   "Quick Manager NFCe - Power"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmXML_NFCe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   14025
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_xmlErro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   1065
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   11490
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
      Height          =   6690
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "frmXML_NFCe.frx":4E95A
      Top             =   1950
      Width           =   11715
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
      Left            =   11850
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
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
      Left            =   11850
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1380
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quick Store Manager NFCe - CUPOM FISCAL"
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
      TabIndex        =   3
      Top             =   0
      Width           =   14100
   End
End
Attribute VB_Name = "frmXML_NFCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sXML As String
Public sXML_Erro As String
Public sCNPJ As String
Public sSequencia As String
Public sStatusDoCupomFiscalContingencia As String
Public bChamadorNFCeNormal As Boolean

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

Private Sub cmd_transmitirXML_Click()
On Error GoTo Erro
    Dim sXML_nfce As String
    Dim sRetorno As String
    Dim rsSaidaSEFAZ As Recordset
    Dim iIndice1 As Long
    Dim iIndice2 As Long
    Dim sStatus As String
    Dim sDetalheAutorizacao As String
    Dim sExMessage As String

    sXML_nfce = tb_xml.Text
    sXML_nfce = Replace(sXML_nfce, vbCrLf, "")
    sXML_nfce = Replace(sXML_nfce, "UTF-16", "UTF-8")
    
    ' Excluir do XML a tag:
    '<infNFeSupl>
    '  <qrCode>http://www.fazenda.pr.gov.br/nfce/qrcode?p=41190304152403000107650010000059449398505819|2|2|15|60.00|43437350755343497236494a795077434b792b422f7a55527143453d|3|a50f4fcd1ea8f0a87d8d725bd7d5ba9dd37c4eb9</qrCode>
    '  <urlChave>http://www.fazenda.pr.gov.br/nfce/consulta</urlChave>
    '</infNFeSupl>
    iIndice1 = InStr(1, sXML_nfce, "<infNFeSupl>")
    If iIndice1 > 0 Then
        iIndice1 = InStr(1, sXML_nfce, "</infNFe>")
        If iIndice1 > 0 Then
          sXML_nfce = Mid(sXML_nfce, iIndice1, iIndice1 + 9) & "</NFe>"
        End If
    End If
    
    ' Via SOAP
    If bSoapClient_MSSoapInit_NFCe = False Then
      Set soapclient_NFCe = New SoapClient30
      soapclient_NFCe.MSSoapInit sSoapClient_MSSoapInit_NFCe
      soapclient_NFCe.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL_NFCe
      bSoapClient_MSSoapInit_NFCe = True
    End If


    sRetorno = soapclient_NFCe.Autoriza_Xml(sCNPJ, sXML_nfce)
    
    sRetorno = Replace(sRetorno, vbCrLf, "")
    
    iIndice1 = InStr(1, sRetorno, "<statusAutorizacao>")
    If iIndice1 > 0 Then
      iIndice2 = InStr(1, sRetorno, "</statusAutorizacao>")
      sStatus = Mid(sRetorno, iIndice1 + 19, iIndice2 - (iIndice1 + 19))
    Else
      sStatus = ""
    End If
              
    iIndice1 = InStr(1, sRetorno, "<detalheAutorizacao>")
    If iIndice1 > 0 Then
      iIndice2 = InStr(1, sRetorno, "</detalheAutorizacao>")
      sDetalheAutorizacao = Mid(sRetorno, iIndice1 + 20, iIndice2 - (iIndice1 + 20))
    Else
      sDetalheAutorizacao = ""
    End If
              
    iIndice1 = InStr(1, sRetorno, "<exMessage>")
    If iIndice1 > 0 Then
      iIndice2 = InStr(1, sRetorno, "</exMessage>")
      sExMessage = Mid(sRetorno, iIndice1 + 11, iIndice2 - (iIndice1 + 11))
    Else
      sExMessage = ""
    End If

    If sStatus = "Erro" Then
        MsgBox "NFCe com ERRO:" & vbCrLf & sDetalheAutorizacao & vbCrLf & sExMessage, vbInformation, "NFCe Posição de retorno"
    End If
    
    MsgBox sRetorno, vbInformation, "NFCe Posição de retorno"
    
    'Atualizar tabela Saídas e a Grid
    Set rsSaidaSEFAZ = db.OpenRecordset("Select * from [Saídas] where Filial = " & gnCodFilial & " And Sequência = " & sSequencia & "")
      
    rsSaidaSEFAZ.Edit
    
    If bChamadorNFCeNormal = False Then
        ' Chamou da Aba de CONTINGENCIA DE NFCe
        rsSaidaSEFAZ!retNFCe_contingencia = sRetorno
        rsSaidaSEFAZ!NFCe_contingencia_status = sStatus
    Else
        ' Chamou da Aba de NFCe Normal
        
        '***** Quando for implementar...tem que tratar igual o metodo EnviarXML_SEFAZ da classe clsNFCe
        ' pois tem que verificar o retorno da chamada do Autoriza_Xml para saber se neste momento a sefaz autorizou
        ' este cupom de forma normal ou se a fazenda agora esta em contingencia....
        
        ' ....
        'If InStr(1, sRetorno, "<emContingencia>false</emContingencia>") > 0 Then
        '    If InStr(sRetorno, "<statusAutorizacao>OK</statusAutorizacao>") Then
        '.....
        'else
        '.....
    End If
    rsSaidaSEFAZ.Update
    rsSaidaSEFAZ.Close
    Set rsSaidaSEFAZ = Nothing
    
    Exit Sub
Erro:
  MsgBox "Erro na função de envio do CUPOM FISCAL para a Fazenda. " & Err.Number & " " & Err.Description, vbInformation, "Erro de Envio do Cupom Fiscal"
End Sub

Private Sub Form_Load()

  If bChamadorNFCeNormal = True Then
      Dim iIndice1 As Long
      Dim iIndice2 As Long
      Dim sDetalheAutorizacao As String
      Dim sDetalheCancelamento As String
      Dim sExMessage As String
  
      cmd_transmitirXML.Visible = False
      Label1.BackColor = &H999999
      Label1.ForeColor = &H80000008
      cmd_formatarVisualXML.BackColor = &H999999
      
      tb_xml.Text = sXML
      
      iIndice1 = InStr(1, sXML, "<detalheAutorizacao>")
      If iIndice1 > 0 Then
        iIndice2 = InStr(1, sXML, "</detalheAutorizacao>")
        sDetalheAutorizacao = Mid(sXML, iIndice1 + 20, iIndice2 - (iIndice1 + 20))
      
        iIndice1 = InStr(1, Mid(sDetalheAutorizacao, 1, 4), "100")
        If iIndice1 > 0 Then
          sDetalheAutorizacao = "SUCESSO " + sDetalheAutorizacao
        Else
          sDetalheAutorizacao = "REJEIÇÃO " + sDetalheAutorizacao
        End If
      
      Else
        sDetalheAutorizacao = ""
      End If
      
      iIndice1 = InStr(1, sXML, "<detalheCancelamento>")
      If iIndice1 > 0 Then
        iIndice2 = InStr(1, sXML, "</detalheCancelamento>")
        sDetalheCancelamento = Mid(sXML, iIndice1 + 21, iIndice2 - (iIndice1 + 21))

        iIndice1 = InStr(1, Mid(sDetalheCancelamento, 1, 4), "135")
        If iIndice1 > 0 Then
          sDetalheCancelamento = "SUCESSO " + sDetalheCancelamento
        Else
          sDetalheCancelamento = "REJEIÇÃO " + sDetalheCancelamento
        End If
      Else
        sDetalheCancelamento = ""
      End If
      
      If sDetalheAutorizacao = "" And sDetalheCancelamento = "" Then
          iIndice1 = InStr(1, sXML, "<exMessage>")
          If iIndice1 > 0 Then
            iIndice2 = InStr(1, sXML, "</exMessage>")
            sExMessage = Mid(sXML, iIndice1 + 11, iIndice2 - (iIndice1 + 11))
            txt_xmlErro.Text = "Erro/Rejeição: " + vbCrLf + "               " + sExMessage
            Exit Sub
          Else
            sDetalheCancelamento = ""
          End If
      End If
      
      txt_xmlErro.Text = "Autorização: " + vbCrLf + "               " + sDetalheAutorizacao
      If sDetalheCancelamento <> "" Then
        txt_xmlErro.Text = txt_xmlErro.Text + vbCrLf + "Cancelamento: " + vbCrLf + "               " + sDetalheCancelamento
      End If
  Else
      If sStatusDoCupomFiscalContingencia = "OK" Then
          cmd_transmitirXML.Enabled = False
      End If
    
      tb_xml.Text = sXML
      txt_xmlErro.Text = "ERRO/INCONSISTÊNCIA: " & sXML_Erro
  
  End If

End Sub
