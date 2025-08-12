VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmImprimeMala 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Impressão de Etiquetas de Mala Direta"
   ClientHeight    =   3945
   ClientLeft      =   3750
   ClientTop       =   2280
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ImprimeMala.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3945
   ScaleWidth      =   9825
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      TabIndex        =   15
      Top             =   1860
      Width           =   5325
      Begin VB.OptionButton optCodigo 
         Appearance      =   0  'Flat
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   750
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton O_Cep 
         Appearance      =   0  'Flat
         Caption         =   "CEP"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2820
         TabIndex        =   9
         Top             =   1080
         Width           =   870
      End
      Begin VB.OptionButton O_Bairro 
         Appearance      =   0  'Flat
         Caption         =   "Bairro"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2820
         TabIndex        =   8
         Top             =   690
         Width           =   960
      End
      Begin VB.OptionButton O_Cidade 
         Appearance      =   0  'Flat
         Caption         =   "Cidade"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2820
         TabIndex        =   7
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   750
         TabIndex        =   6
         Top             =   690
         Width           =   885
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ajuste de margens da impressora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   60
      TabIndex        =   12
      Top             =   960
      Width           =   4305
      Begin ComctlLib.Slider sldSuperior 
         Height          =   1695
         Left            =   2820
         TabIndex        =   2
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
         Left            =   330
         TabIndex        =   1
         Top             =   540
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   327682
         LargeChange     =   1
         Min             =   -7
         Max             =   7
      End
      Begin VB.Label lblSuperior 
         Alignment       =   2  'Center
         Caption         =   "Superior = padrão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2100
         TabIndex        =   14
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblEsquerda 
         Alignment       =   2  'Center
         Caption         =   "Esquerda = padrão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   13
         Top             =   1140
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   11
      Top             =   960
      Width           =   5325
      Begin VB.OptionButton O_Impressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2730
         TabIndex        =   4
         Top             =   330
         Width           =   1215
      End
      Begin VB.OptionButton O_Vídeo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   780
         TabIndex        =   3
         Top             =   330
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton B_Emite 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Gerar Etiquetas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3390
      Width           =   9705
   End
   Begin Crystal.CrystalReport rptRel 
      Left            =   9210
      Top             =   3690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   $"ImprimeMala.frx":4E95A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   9705
   End
End
Attribute VB_Name = "frmImprimeMala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nAdjustLeft As Integer
Dim nAdjustTop As Integer

Private Sub B_Emite_Click()
  Dim nMarginTop As Integer
  Dim nMarginBottom As Integer
  Dim nMarginRight As Integer
  Dim nMarginLeft As Integer
 
  Rem Margens
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
  
  With rptRel
    .DataFiles(0) = gsQuickDBFileName
    If O_Vídeo.Value Then
      .Destination = crptToWindow
    Else
      .Destination = crptToPrinter
    End If
    .ReportFileName = gsReportPath & "Mala1.RPT"
    .MarginTop = nMarginTop
    .MarginBottom = nMarginBottom
    .MarginLeft = nMarginLeft
    .MarginRight = nMarginRight
    If optCodigo.Value Then
      .SortFields(0) = "+{Cli_for.Código}"
    ElseIf O_Nome.Value Then
      .SortFields(0) = "+{Cli_for.Nome}"
    ElseIf O_Cidade.Value Then
      .SortFields(0) = "+{Cli_for.Cidade}"
    ElseIf O_Bairro.Value Then
      .SortFields(0) = "+{Cli_for.Bairro}"
    ElseIf O_Cep.Value Then
      .SortFields(0) = "+{Cli_for.CEP}"
    End If
    .WindowState = crptMaximized
    MousePointer = vbHourglass
    Call StatusMsg("Aguarde, imprimindo...")
  
  
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", rptRel)
  
    
    .Action = 1
  End With
  Call StatusMsg("")
  MousePointer = vbDefault

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

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

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
