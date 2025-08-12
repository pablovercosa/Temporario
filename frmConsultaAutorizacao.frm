VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmConsultaAutorizacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar as Autorizações de Publicidade"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsultaAutorizacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11760
   Begin VB.CommandButton cmdConsultar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Consultar"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6070
      Width           =   1455
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6070
      Width           =   1455
   End
   Begin VB.Frame fraZ 
      Caption         =   "Intervalo de Nº de Autorizações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   17
      Top             =   1080
      Width           =   3700
      Begin VB.TextBox txtNumFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNumIni 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   480
         MaxLength       =   8
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   1800
         TabIndex        =   19
         Top             =   420
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   420
         Width           =   195
      End
   End
   Begin VB.Frame fraY 
      Caption         =   "Situação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   14
      Top             =   1080
      Width           =   3700
      Begin VB.OptionButton optTodas 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optNaoFaturadas 
         Appearance      =   0  'Flat
         Caption         =   "Não Confirmado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1380
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optFaturadas 
         Appearance      =   0  'Flat
         Caption         =   "Confirmado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraX 
      Caption         =   "Data de Assinatura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   3700
      Begin MSMask.MaskEdBox mskDataAssinaturaIni 
         Height          =   315
         Left            =   480
         TabIndex        =   0
         ToolTipText     =   "Pressione F2 para obter calendário."
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDataAssinaturaFin 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para obter calendário."
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   1800
         TabIndex        =   15
         Top             =   420
         Width           =   90
      End
   End
   Begin VB.Frame fraResultados 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   -120
      Width           =   14415
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Visualização dos Contratos Cadastradas no Sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblDica 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmConsultaAutorizacao.frx":058A
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   11220
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdResultado 
      Height          =   3900
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   11415
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   8
      BevelColorFrame =   -2147483632
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      AllowRowSizing  =   0   'False
      RowHeight       =   423
      ExtraHeight     =   26
      Columns.Count   =   8
      Columns(0).Width=   1349
      Columns(0).Caption=   "Data"
      Columns(0).Name =   "DataAssinatura"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   7
      Columns(0).FieldLen=   256
      Columns(1).Width=   1799
      Columns(1).Caption=   "Autorização"
      Columns(1).Name =   "NumeroAutorizacao"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   741
      Columns(2).Caption=   "Mês"
      Columns(2).Name =   "MesX"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   256
      Columns(3).Width=   1984
      Columns(3).Caption=   "Programação"
      Columns(3).Name =   "Programacao"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1826
      Columns(4).Caption=   "Valor Program."
      Columns(4).Name =   "VlProgramacao"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      Columns(5).Width=   3228
      Columns(5).Caption=   "Cliente"
      Columns(5).Name =   "Cliente"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3228
      Columns(6).Caption=   "Rádio"
      Columns(6).Name =   "Radio"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3228
      Columns(7).Caption=   "Agência"
      Columns(7).Name =   "Agencia"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   20135
      _ExtentY        =   6879
      _StockProps     =   79
      Caption         =   "Autorizações Existentes"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaAutorizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_lngNumAutorizacao    As Long
Private m_intMesX              As Integer
Private m_dblSomaValorRecebido As Double

Private Sub cmdConsultar_Click()
  Dim rstContrato    As Recordset
  Dim strSQL         As String
  
  'Validação das Datas...
  If Not IsDate(mskDataAssinaturaIni.Text) Then
    MsgBox "Entrar com Data Inicial válida.", vbExclamation, "Quick Store"
    mskDataAssinaturaIni.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(mskDataAssinaturaFin.Text) Then
    MsgBox "Entrar com Data Final válida.", vbExclamation, "Quick Store"
    mskDataAssinaturaFin.SetFocus
    Exit Sub
  End If
  
  If CDate(mskDataAssinaturaFin.Text) < CDate(mskDataAssinaturaIni.Text) Then
    MsgBox "Data Final menor que a Data Inicial.", vbExclamation, "Quick Store"
    mskDataAssinaturaFin.SetFocus
    Exit Sub
  End If
  'Fim da Validação das Datas
  
  'Preparando a grid...
  With grdResultado
    'Não permite atualizar o layout da grid
    .Redraw = False
    'Limpa o grid
    .RemoveAll
    'Permite atualizar o layout da grid
    .Redraw = True
  End With
  'Fim da preparação da grid
  
  'Início do SELECT...
  strSQL = " SELECT Contrato.[Num Autorizacao] AS AutorizacaoPai, Programacao.[Num Autorizacao] AS AutorizacaoFilho, * "
  strSQL = strSQL & " FROM Contrato, Programacao "
  strSQL = strSQL & " WHERE Contrato.[Data Assinatura] >= #" & CDate(mskDataAssinaturaIni.Text) & "#"
  strSQL = strSQL & " AND Contrato.[Data Assinatura] <= #" & CDate(mskDataAssinaturaFin.Text) & "#"
  strSQL = strSQL & " AND Programacao.[Num Autorizacao] = Contrato.[Num Autorizacao] "
  
  If IsNumeric(txtNumIni.Text) And IsNumeric(txtNumFin.Text) Then
    strSQL = strSQL & " AND Contrato.[Num Autorizacao] >= " & CLng(txtNumIni.Text)
    strSQL = strSQL & " AND Contrato.[Num Autorizacao] <= " & CLng(txtNumFin.Text)
  End If
  
  If optFaturadas.Value Then
    strSQL = strSQL & " AND Programacao.Faturado = TRUE "
  End If
  
  If optNaoFaturadas.Value Then
    strSQL = strSQL & " AND Programacao.Faturado = FALSE "
  End If
  
  strSQL = strSQL & " ORDER BY Contrato.[Data Assinatura], Programacao.[Num Autorizacao], Programacao.MesX "
  'Fim do SELECT
  
  Set rstContrato = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstContrato
    If Not (.BOF And .EOF) Then
      .MoveFirst
    
      Do Until .EOF
        grdResultado.AddNew
        
        grdResultado.Columns("DataAssinatura").Text = .Fields("Data Assinatura").Value
        grdResultado.Columns("NumeroAutorizacao").Text = .Fields("AutorizacaoFilho").Value
        grdResultado.Columns("MesX").Text = .Fields("MesX").Value
        
        m_intMesX = .Fields("MesX").Value
        m_lngNumAutorizacao = .Fields("AutorizacaoFilho").Value
        
        grdResultado.Columns("Programacao").Text = .Fields("Programacao").Value
        grdResultado.Columns("VlProgramacao").Text = .Fields("Valor Total").Value
        
        grdResultado.Columns("Cliente").Text = .Fields("Cod Cliente").Value & " - " & GetNomeCliFor(.Fields("Cod Cliente").Value)
        grdResultado.Columns("Radio").Text = .Fields("Cod Radio").Value & " - " & GetNomeRadio(.Fields("Cod Radio").Value)
        grdResultado.Columns("Agencia").Text = .Fields("Cod Fornecedor").Value & " - " & GetNomeCliFor(.Fields("Cod Fornecedor").Value)
      
        grdResultado.Update
        
        .MoveNext
      Loop
    
    Else
      MsgBox "Nenhuma informação foi encontrada dentro dos critérios escolhidos.", vbInformation, "Quick Store"
      mskDataAssinaturaIni.SetFocus
    End If
    .Close
  End With
  
  Set rstContrato = Nothing
  
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

Private Function GetNomeCliFor(lngCodigo As Long) As String
  Dim rstCliFor As Recordset
  
  Set rstCliFor = db.OpenRecordset("SELECT Nome FROM Cli_For WHERE Código = " & lngCodigo, dbOpenDynaset)
  
  With rstCliFor
    GetNomeCliFor = IIf((.BOF And .EOF), "<_não_cadastrado>", .Fields("Nome").Value & "")
    .Close
  End With
  
  Set rstCliFor = Nothing
End Function

Private Function GetNomeRadio(lngCodigo As Long) As String
  Dim rstRadio As Recordset
  
  Set rstRadio = db.OpenRecordset("SELECT Nome FROM Radio WHERE Código = " & lngCodigo, dbOpenDynaset)

  With rstRadio
    GetNomeRadio = IIf((.BOF And .EOF), "<_não_cadastrado>", .Fields("Nome").Value & "")
    .Close
  End With

  Set rstRadio = Nothing
  
End Function

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


