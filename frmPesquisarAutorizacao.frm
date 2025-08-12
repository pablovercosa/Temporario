VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPesquisarAutorizacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisar"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesquisarAutorizacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   8895
   Begin VB.Frame Frame2 
      Caption         =   "Período da Data de Assinatura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      TabIndex        =   13
      Top             =   1440
      Width           =   4325
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H0000C0C0&
         Caption         =   "&Pesquisar"
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
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskDataAssinaturaIni 
         Height          =   315
         Left            =   600
         TabIndex        =   3
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
      Begin MSMask.MaskEdBox mskDataAssinaturaFin 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   1920
         TabIndex        =   15
         Top             =   420
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   420
         Width           =   195
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
      Left            =   -120
      TabIndex        =   7
      Top             =   -120
      Width           =   9135
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisa de Autorização de Publicidade"
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
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblDica 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   7980
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Pesquisa "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4325
      Begin VB.OptionButton optFornecedor 
         Appearance      =   0  'Flat
         Caption         =   "Código da Agência de Publicidade"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtPesquisa 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1000
         Width           =   3015
      End
      Begin VB.OptionButton optNumero 
         Appearance      =   0  'Flat
         Caption         =   "Número de Autorização"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdResultado 
      Bindings        =   "frmPesquisarAutorizacao.frx":058A
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   8655
      ScrollBars      =   2
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
      Col.Count       =   4
      BevelColorFrame =   -2147483632
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      RowHeight       =   423
      ExtraHeight     =   26
      Columns.Count   =   4
      Columns(0).Width=   2355
      Columns(0).Caption=   "Número"
      Columns(0).Name =   "Numero"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   200
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3387
      Columns(1).Caption=   "Rádio"
      Columns(1).Name =   "Radio"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   4604
      Columns(2).Caption=   "Cliente"
      Columns(2).Name =   "Cliente"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   4445
      Columns(3).Caption=   "Agência"
      Columns(3).Name =   "Fornecedor"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      _ExtentX        =   15266
      _ExtentY        =   4895
      _StockProps     =   79
      Caption         =   "Resultados da Pesquisa"
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
   Begin VB.Label Label1 
      Caption         =   "Duplo Click na linha carregará a informação na tela de Autorização"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   6135
   End
   Begin VB.Label lblNomeCliente 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   8655
   End
End
Attribute VB_Name = "frmPesquisarAutorizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'23/01/2004 - Desenvolvedores Daniel And Maikel
'Case: STC

Private m_strSQL      As String

Private Sub cmdPesquisar_Click()
  Dim rstAutorizacoes  As Recordset
  Dim lngNumero        As Long
  Dim lngCodFornecedor As Long
  
  If Not IsNumeric(txtPesquisa.Text) Then Exit Sub
    
  If optNumero.Value Then                   'Pesquisa pelo Número
    lngNumero = CLng(txtPesquisa.Text)
  Else                                      'Pesquisa pelo CodFornecedor
    lngCodFornecedor = CLng(txtPesquisa.Text)
  End If
  
  m_strSQL = " SELECT [Num Autorizacao], [Cod Radio], [Cod Cliente], [Cod Fornecedor] "
  
  If optNumero.Value Then
    m_strSQL = m_strSQL & " FROM Contrato WHERE [Num Autorizacao] = " & lngNumero
  Else
    m_strSQL = m_strSQL & " FROM Contrato WHERE [Cod Fornecedor] = " & lngCodFornecedor
  End If
  
  'Período de Data de Assinatura Ini e Fin
  If Not IsDate(mskDataAssinaturaIni.Text) Then
    MsgBox "Data Inicial inválida.", vbExclamation, "Quick Store"
    mskDataAssinaturaIni.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(mskDataAssinaturaFin.Text) Then
    MsgBox "Data Final inválida.", vbExclamation, "Quick Store"
    mskDataAssinaturaFin.SetFocus
    Exit Sub
  End If
  
  'Populando a str com o período Ini e Fin
  m_strSQL = m_strSQL & " AND [Data Assinatura] >= #" & (mskDataAssinaturaIni.Text) & "#"
  m_strSQL = m_strSQL & " AND [Data Assinatura] <= #" & (mskDataAssinaturaFin.Text) & "#"
  
  m_strSQL = m_strSQL & " ORDER BY [Num Autorizacao]"
  
  Set rstAutorizacoes = db.OpenRecordset(m_strSQL, dbOpenDynaset)
  
  With grdResultado
    'Não permite atualizar o layout do grid
    .Redraw = False
    'Limpa o grid
    .RemoveAll
    'Permite atualizar o layout do grid
    .Redraw = True
  End With
  
  With rstAutorizacoes
    'Se o recordset estiver vazio
    If (.BOF And .EOF) Then
      MsgBox "Nenhum registro encontrado segundo os critérios informados, verifique !", vbInformation, "Quick Store"
      Exit Sub
    End If
    
    .MoveFirst
    
    Do Until .EOF
      
      grdResultado.AddNew
      
      grdResultado.Columns("Numero").Text = .Fields("Num Autorizacao").Value
      grdResultado.Columns("Radio").Text = .Fields("Cod Radio").Value & " - " & GetNomeRadio(.Fields("Cod Radio").Value)
      grdResultado.Columns("Cliente").Text = .Fields("Cod Cliente").Value & " - " & GetNomeCliFor(.Fields("Cod Cliente").Value)
      grdResultado.Columns("Fornecedor").Text = .Fields("Cod Fornecedor").Value & " - " & GetNomeCliFor(.Fields("Cod Fornecedor").Value)
      
      grdResultado.Update
      
      .MoveNext
    Loop
  End With
End Sub

Private Sub Form_Load()
  Dim rstClientes As Recordset
  
  Call CenterForm(Me)
'  frmCliFor.g_lngCodCliente
'  Set rstClientes = db.OpenRecordset("SELECT Código, Nome FROM Cli_For WHERE Código = " & 1, dbOpenDynaset)
'
'  With rstClientes
'    If Not (.BOF And .EOF) Then
'      lblNomeCliente.Caption = .Fields("Código").Value & " - " & .Fields("Nome").Value
'      lblDica.Caption = "Através dessa tela você poderá encontrar as autorizações de publicidade para o cliente " & .Fields("Nome").Value
'    End If
'    .Close
'  End With
'
'  Set rstClientes = Nothing
End Sub

Private Sub grdResultado_DblClick()
  frmAutorizacaoPublicidade.txtNumAutorizacao.Text = grdResultado.Columns(0).Text
  frmAutorizacaoPublicidade.txtNumAutorizacao_LostFocus
  grdResultado.SetFocus
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

Private Sub txtPesquisa_Change()
'  cmdPesquisar.Enabled = Len(Trim(txtPesquisa.Text)) > 0
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
