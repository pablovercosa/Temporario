VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmImprimeRemetente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Etiquetas de Remetente"
   ClientHeight    =   5670
   ClientLeft      =   1275
   ClientTop       =   690
   ClientWidth     =   7005
   HelpContextID   =   1690
   Icon            =   "ImprimeRemetente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5670
   ScaleWidth      =   7005
   Begin VB.TextBox CEP 
      Height          =   315
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   6
      Top             =   2835
      Width           =   1335
   End
   Begin VB.TextBox Estado 
      Height          =   315
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2460
      Width           =   615
   End
   Begin VB.TextBox Cidade 
      Height          =   315
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   4
      Top             =   2463
      Width           =   2775
   End
   Begin VB.TextBox Ende3 
      Height          =   315
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   3
      Top             =   2091
      Width           =   4455
   End
   Begin VB.TextBox Ende2 
      Height          =   315
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   2
      Top             =   1725
      Width           =   4455
   End
   Begin VB.TextBox Ende1 
      Height          =   315
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1347
      Width           =   4455
   End
   Begin VB.TextBox Nome 
      Height          =   315
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   0
      Top             =   975
      Width           =   4455
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ajuste de margens da impressora"
      Height          =   2295
      Left            =   60
      TabIndex        =   14
      Top             =   3300
      Width           =   3255
      Begin ComctlLib.Slider sldSuperior 
         Height          =   1695
         Left            =   2160
         TabIndex        =   8
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
         TabIndex        =   7
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
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblSuperior 
         Alignment       =   2  'Center
         Caption         =   "Superior = padrão"
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   3450
      TabIndex        =   13
      Top             =   3285
      Width           =   1455
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton B_Emite 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   5580
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   6330
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label8 
      Caption         =   "CEP :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2835
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Estado :"
      Height          =   315
      Left            =   4305
      TabIndex        =   20
      Top             =   2535
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Cidade :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2490
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Endereço 2 :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1785
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Endereço 3 :"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2145
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Endereço 1 :"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Nome :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"ImprimeRemetente.frx":058A
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmImprimeRemetente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ajusta_Ver As Integer
Dim Ajusta_Hor As Integer
Dim Margem_Sup As Long
Dim Margem_Inf As Long
Dim Margem_Dir As Long
Dim Margem_Esq As Long

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
  Ajusta_Ver = nCheckValues(tmEsquerda)
End Sub

Private Sub sldEsquerda_Scroll()
  Call sldEsquerda_Click
End Sub

Private Sub sldSuperior_Change()
  Call sldSuperior_Click
End Sub

Private Sub sldSuperior_Click()
  Ajusta_Hor = nCheckValues(tmSuperior)
End Sub

Private Sub sldSuperior_Scroll()
  Call sldSuperior_Click
End Sub

Private Sub B_Emite_Click()
 Dim Str_Rel As String, Str1 As String
 
 Margem_Sup = 720  '1.27 cm * 567 twips por cm
 Margem_Inf = Margem_Sup
 Margem_Esq = 272 '0.48 cm * 567 twips por com
 Margem_Dir = Margem_Esq

 Rem  Seta Valores e Manda Relatório

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If O_Vídeo = True Then Rel.Destination = 0
 If O_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 Str1 = gsReportPath & "Mala2.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel


 Rem Margens
 Margem_Sup = Margem_Sup + Ajusta_Hor
 Margem_Inf = Margem_Inf - Ajusta_Hor
   
 Margem_Esq = Margem_Esq + Ajusta_Ver
 Margem_Dir = Margem_Dir - Ajusta_Ver
   
 If Margem_Esq < 0 Then Margem_Esq = 0
 If Margem_Dir < 0 Then Margem_Dir = 0
   
 Rel.MarginTop = Margem_Sup
 Rel.MarginBottom = Margem_Inf
 Rel.MarginLeft = Margem_Esq
 Rel.MarginRight = Margem_Dir

 If IsNull(Nome.Text) Then Nome.Text = " "
 If IsNull(Ende1.Text) Then Ende1.Text = " "
 If IsNull(Ende2.Text) Then Ende2.Text = " "
 If IsNull(Ende3.Text) Then Ende3.Text = " "
 If IsNull(Cidade.Text) Then Cidade.Text = " "
 If IsNull(Estado.Text) Then Estado.Text = " "
 If IsNull(CEP.Text) Then CEP.Text = " "


 Str_Rel = "nome = '" + Nome.Text + "'"
 Rel.Formulas(0) = Str_Rel
 
 Str_Rel = "ende1 = '" + Ende1.Text + "'"
 Rel.Formulas(1) = Str_Rel
 
 Str_Rel = "ende2 = '" + Ende2.Text + "'"
 Rel.Formulas(2) = Str_Rel
 
 Str_Rel = "ende3 = '" + Ende3.Text + "'"
 Rel.Formulas(3) = Str_Rel
 
 Str_Rel = "cidade = '" + Cidade.Text + "'"
 Rel.Formulas(4) = Str_Rel
 
 Str_Rel = "estado = '" + Estado.Text + "'"
 Rel.Formulas(5) = Str_Rel
 
 Str_Rel = "Cep = '" + CEP.Text + "'"
 Rel.Formulas(6) = Str_Rel
 
 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
 
 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub


Private Sub CEP_KeyPress(KeyAscii As Integer)
 If KeyAscii = 34 Then KeyAscii = 0
 If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Cidade_KeyPress(KeyAscii As Integer)
 If KeyAscii = 34 Then KeyAscii = 0
 If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Ende1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 34 Then KeyAscii = 0
 If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Ende2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 34 Then KeyAscii = 0
 If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Ende3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 34 Then KeyAscii = 0
 If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Estado_KeyPress(KeyAscii As Integer)
 If KeyAscii = 34 Then KeyAscii = 0
 If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

Private Sub Nome_KeyPress(KeyAscii As Integer)
 If KeyAscii = 34 Then KeyAscii = 0
 If KeyAscii = 39 Then KeyAscii = 0
End Sub
