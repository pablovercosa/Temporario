VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelMalaAutorizacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mala Direta de Contratos (Autorizações de Publicidades)"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmRelMalaAutorizacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   7440
   Begin VB.Frame Frame6 
      Caption         =   "Ajuste de margens da impressora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   3615
      Begin ComctlLib.Slider sldSuperior 
         Height          =   1695
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   2990
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -7
         Max             =   7
      End
      Begin ComctlLib.Slider sldEsquerda 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   327682
         LargeChange     =   1
         Min             =   -7
         Max             =   7
      End
      Begin VB.Label lblEsquerda 
         Alignment       =   2  'Center
         Caption         =   "Esquerda = padrão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblSuperior 
         Alignment       =   2  'Center
         Caption         =   "Superior = padrão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Frame fraDica 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   -120
      Width           =   8175
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Etiquetas Mala Direta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblDica 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelMalaAutorizacoes.frx":058A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   600
         TabIndex        =   13
         Top             =   480
         Width           =   6735
      End
   End
   Begin VB.Frame fraSaida 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3840
      TabIndex        =   11
      Top             =   1920
      Width           =   3495
      Begin Crystal.CrystalReport crtView 
         Left            =   2880
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Período da Data de Assinatura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3840
      TabIndex        =   8
      Top             =   960
      Width           =   3495
      Begin MSMask.MaskEdBox mskDataAssinaturaFin 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para obter calendário"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDataAssinaturaIni 
         Height          =   315
         Left            =   600
         TabIndex        =   0
         ToolTipText     =   "Pressione F2 para obter calendário"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   420
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   420
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmRelMalaAutorizacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nAdjustLeft As Integer
Dim nAdjustTop As Integer

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim strSQL                         As String
  Dim strSelection                   As String
  Dim strNome                        As String
  Dim strEndereco                    As String
  Dim strBairro                      As String
  Dim strComplemento                 As String
  Dim strCidade                      As String
  Dim strEstado                      As String
  Dim strCep                         As String
  Dim rstMala                        As Recordset
  Dim rsttblRelMalaAutorizacoes      As Recordset
  Dim rsttblRelMalaAutorizacoesFatur As Recordset
  
  Dim nMarginTop As Integer
  Dim nMarginBottom As Integer
  Dim nMarginRight As Integer
  Dim nMarginLeft As Integer
 
  'Margens
  nMarginTop = 720  '1.27 cm * 567 twips por cm
  nMarginBottom = nMarginTop
  nMarginLeft = 272 '0.48 cm * 567 twips por cm
  nMarginRight = nMarginLeft
  
  nMarginTop = nMarginTop + nAdjustTop
  nMarginBottom = nMarginBottom - nAdjustTop
    
  nMarginLeft = nMarginLeft + nAdjustLeft
  nMarginRight = nMarginRight - nAdjustLeft
    
  If nMarginLeft < 0 Then
    nMarginLeft = 0
  End If
  If nMarginRight < 0 Then
    nMarginRight = 0
  End If


  'Limpando as tabelas temporárias tblRelMalaAutorizacoes e tblRelMalaAutorizacoesFatur
  dbTemp.Execute "DELETE * FROM tblRelMalaAutorizacoes"
  dbTemp.Execute "DELETE * FROM tblRelMalaAutorizacoesFatur"
  
  'Tratamento para os campos Datas
  If Not IsDate(mskDataAssinaturaIni.Text) Then
    MsgBox "Data Inicial inválida, verifique.", vbExclamation, "Quick Store"
    mskDataAssinaturaIni.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(mskDataAssinaturaFin.Text) Then
    MsgBox "Data Final inválida, verifique.", vbExclamation, "Quick Store"
    mskDataAssinaturaFin.SetFocus
    Exit Sub
  End If

  If CDate(mskDataAssinaturaFin.Text) < CDate(mskDataAssinaturaIni.Text) Then
    MsgBox "Data Final menor que a Inicial, verifique.", vbExclamation, "Quick Store"
    mskDataAssinaturaFin.SetFocus
    Exit Sub
  End If

  strSQL = "SELECT * "
  strSQL = strSQL & " FROM Contrato, Programacao, Cli_For "
  strSQL = strSQL & " WHERE Contrato.[Data Assinatura] >=#" & CDate(mskDataAssinaturaIni.Text) & "#"
  strSQL = strSQL & " AND Contrato.[Data Assinatura] <=#" & CDate(mskDataAssinaturaFin.Text) & "#"
  strSQL = strSQL & " AND Programacao.[Num Autorizacao] = Contrato.[Num Autorizacao] "
  strSQL = strSQL & " AND Cli_For.Código = Contrato.[Cod Cliente] "
  strSQL = strSQL & " AND Programacao.[Gerar Etiqueta] = TRUE "
  strSQL = strSQL & " AND Programacao.Faturado = TRUE "
  strSQL = strSQL & " ORDER BY Programacao.[Num Autorizacao], Programacao.MesX "

  Set rstMala = db.OpenRecordset(strSQL, dbOpenDynaset)
  Set rsttblRelMalaAutorizacoes = dbTemp.OpenRecordset("SELECT * FROM tblRelMalaAutorizacoes", dbOpenDynaset)
  Set rsttblRelMalaAutorizacoesFatur = dbTemp.OpenRecordset("SELECT * FROM tblRelMalaAutorizacoesFatur", dbOpenDynaset)

  With rstMala
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
          rsttblRelMalaAutorizacoes.AddNew
          rsttblRelMalaAutorizacoes.Fields("Codigo").Value = .Fields("Código").Value
          
          If Len(.Fields("Nome").Value) > 70 Then
            rsttblRelMalaAutorizacoes.Fields("Nome").Value = Left((.Fields("Nome").Value & ""), 70)
          Else
            rsttblRelMalaAutorizacoes.Fields("Nome").Value = (.Fields("Nome").Value & "")
          End If
          
          If Len(.Fields("Endereço").Value) > 100 Then
            rsttblRelMalaAutorizacoes.Fields("Endereco").Value = Left((.Fields("Endereço").Value), 100)
          Else
            rsttblRelMalaAutorizacoes.Fields("Endereco").Value = .Fields("Endereço").Value & ""
          End If
          
          If Len(.Fields("Bairro").Value) > 30 Then
            rsttblRelMalaAutorizacoes.Fields("Bairro").Value = Left((.Fields("Bairro").Value), 30)
          Else
            rsttblRelMalaAutorizacoes.Fields("Bairro").Value = .Fields("Bairro").Value & ""
          End If
          
          If Len(.Fields("Complemento")) > 20 Then
            rsttblRelMalaAutorizacoes.Fields("Complemento").Value = Left((.Fields("Complemento").Value), 20)
          Else
            rsttblRelMalaAutorizacoes.Fields("Complemento").Value = .Fields("Complemento").Value & ""
          End If
          
          If Len(.Fields("Cidade").Value) > 40 Then
            rsttblRelMalaAutorizacoes.Fields("Cidade").Value = Left((.Fields("Cidade").Value), 40)
          Else
            rsttblRelMalaAutorizacoes.Fields("Cidade").Value = .Fields("Cidade").Value & ""
          End If
          
          rsttblRelMalaAutorizacoes.Fields("Estado").Value = .Fields("Estado").Value & ""
          
          If Len(.Fields("CEP").Value) > 9 Then
            rsttblRelMalaAutorizacoes.Fields("Cep").Value = Left((.Fields("CEP").Value), 9)
          Else
            rsttblRelMalaAutorizacoes.Fields("Cep").Value = .Fields("CEP").Value & ""
          End If
          
          rsttblRelMalaAutorizacoes.Update
      
        .MoveNext
      Loop
      
    Else
      MsgBox "Nenhuma informação foi encontrada dentro da seleção escolhida por você.", vbExclamation, "Quick Store"
      Exit Sub
    End If
    .Close
  End With

  Set rstMala = Nothing
  
  rsttblRelMalaAutorizacoes.Close
  Set rsttblRelMalaAutorizacoes = Nothing
  
  'Abrir novamente e agrupar para ocorrer o filtro
  Set rsttblRelMalaAutorizacoes = dbTemp.OpenRecordset("SELECT Codigo, Nome, Endereco, Bairro, Complemento, Cidade, Estado, Cep FROM tblRelMalaAutorizacoes GROUP BY Codigo, Nome, Endereco, Bairro, Complemento, Cidade, Estado, Cep ", dbOpenDynaset, dbReadOnly)
  
  With rsttblRelMalaAutorizacoes
    .MoveFirst
    .MoveLast
    .MoveFirst
    
    Do Until .EOF
      rsttblRelMalaAutorizacoesFatur.AddNew
      rsttblRelMalaAutorizacoesFatur.Fields("Codigo").Value = .Fields("Codigo").Value
      rsttblRelMalaAutorizacoesFatur.Fields("Nome").Value = .Fields("Nome").Value & ""
      rsttblRelMalaAutorizacoesFatur.Fields("Endereco").Value = .Fields("Endereco").Value & ""
      rsttblRelMalaAutorizacoesFatur.Fields("Bairro").Value = .Fields("Bairro").Value & ""
      rsttblRelMalaAutorizacoesFatur.Fields("Complemento").Value = .Fields("Complemento").Value & ""
      rsttblRelMalaAutorizacoesFatur.Fields("Cidade").Value = .Fields("Cidade").Value & ""
      rsttblRelMalaAutorizacoesFatur.Fields("Estado").Value = .Fields("Estado").Value & ""
      rsttblRelMalaAutorizacoesFatur.Fields("Cep").Value = .Fields("Cep").Value & ""
      rsttblRelMalaAutorizacoesFatur.Update
      
    .MoveNext
    Loop
  
  End With
  
  rsttblRelMalaAutorizacoes.Close
  rsttblRelMalaAutorizacoesFatur.Close
  Set rsttblRelMalaAutorizacoes = Nothing
  Set rsttblRelMalaAutorizacoesFatur = Nothing
  
  Call TransferirDados
  
  With crtView
    .Reset
    .ReportFileName = gsReportPath & "MALA1.rpt"
    '.DataFiles(0) = gsTempDBFileName 'Fará a busca para o relatório em apenas uma tabela temporária
    .DataFiles(0) = gsQuickDBFileName
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crtView
    
    'strSelection = "{tblRelMalaAutorizacoesFatur.Nome} <> '' "
    
    .MarginTop = nMarginTop
    .MarginBottom = nMarginBottom
    .MarginLeft = nMarginLeft
    .MarginRight = nMarginRight
    
    '.SelectionFormula = strSelection
    
    .Destination = IIf(optSaidaVideo.Value, crptToWindow, crptToPrinter)
    
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crtView)
    
    .WindowState = crptMaximized
    .Action = 1
  End With

  Call StatusMsg("")

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

Private Sub mskDataAssinaturaFin_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataAssinaturaFin.Text = frmCalendario.gsDateCalender(mskDataAssinaturaFin.Text)
  End If
End Sub

Private Sub mskDataAssinaturaFin_LostFocus()
  mskDataAssinaturaFin.Text = Ajusta_Data(mskDataAssinaturaFin.Text)
End Sub

Private Sub mskDataAssinaturaIni_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataAssinaturaIni.Text = frmCalendario.gsDateCalender(mskDataAssinaturaIni.Text)
  End If
End Sub

Private Sub mskDataAssinaturaIni_LostFocus()
  mskDataAssinaturaIni.Text = Ajusta_Data(mskDataAssinaturaIni.Text)
End Sub
Private Function nCheckValues(ByVal nType As TipoMargem) As Integer
  Dim nPosition As Integer
  Dim lblText As Label
  
  If nType = tmEsquerda Then
    Set lblText = lblEsquerda
    nPosition = sldEsquerda.Value
  Else
    Set lblText = lblSuperior
    nPosition = sldSuperior.Value
  End If
  
  If nPosition = 0 Then
    nCheckValues = 0
  Else
    nCheckValues = CInt(nPosition / 10 * 567)
  End If
  lblText.Caption = IIf(nType = tmSuperior, "Superior", "Esquerda") & _
    IIf(nPosition = 0, " = padrão", " = " & IIf(nPosition > 0, "+", "") & nPosition & " mm")
End Function

Private Sub sldEsquerda_Change()
  Call sldEsquerda_Click
End Sub

Private Sub sldEsquerda_Click()
  nAdjustLeft = nCheckValues(tmEsquerda)
End Sub

Private Sub sldEsquerda_Scroll()
  Call sldEsquerda_Click
End Sub

Private Sub sldSuperior_Change()
  Call sldSuperior_Click
End Sub

Private Sub sldSuperior_Click()
  nAdjustTop = nCheckValues(tmSuperior)
End Sub

Private Sub sldSuperior_Scroll()
  Call sldSuperior_Click
End Sub

Private Sub TransferirDados()
  Dim rstTemp            As Recordset
  Dim rstMalaDiretaTempo As Recordset

  'Esvaziar a table [Mala Direta - Tempo]
  db.Execute "DELETE * FROM [Mala Direta - Tempo]"

  Set rstTemp = dbTemp.OpenRecordset("tblRelMalaAutorizacoesFatur", dbOpenDynaset)
  Set rstMalaDiretaTempo = db.OpenRecordset("Mala Direta - Tempo", dbOpenDynaset)
  
  With rstTemp
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        
        rstMalaDiretaTempo.AddNew
        rstMalaDiretaTempo.Fields("Cliente").Value = .Fields("Codigo").Value
        'rstMalaDiretaTempo.Fields("Ordem").Value AUTONUMÉRICO
        'rstMalaDiretaTempo.Fields("Nome").Value = Left(("A/C: " & .Fields("Nome").Value), 30)
        rstMalaDiretaTempo.Fields("Nome").Value = "A/C: Setor Financeiro"
        rstMalaDiretaTempo.Update
      
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstTemp = Nothing
  
  rstMalaDiretaTempo.Close
  Set rstMalaDiretaTempo = Nothing

End Sub
