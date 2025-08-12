VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmClassificacaoClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Classificação dos Clientes"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClassificacaoClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   7815
   Begin VB.Frame fraG 
      Caption         =   "Grupos de Classificação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   7815
      Begin MSMask.MaskEdBox mskLimiteIni1 
         Height          =   315
         Left            =   4440
         TabIndex        =   2
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         MaxLength       =   17
         Format          =   "##,###,###,###,##0.00"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtNome4 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   480
         MaxLength       =   40
         TabIndex        =   10
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtNome3 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   480
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtNome2 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   480
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtNome1 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   480
         MaxLength       =   40
         TabIndex        =   1
         Top             =   840
         Width           =   3735
      End
      Begin MSMask.MaskEdBox mskLimiteFin1 
         Height          =   315
         Left            =   6120
         TabIndex        =   3
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         MaxLength       =   17
         Format          =   "##,###,###,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskLimiteIni2 
         Height          =   315
         Left            =   4440
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648384
         MaxLength       =   17
         Format          =   "##,###,###,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskLimiteFin2 
         Height          =   315
         Left            =   6120
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648384
         MaxLength       =   17
         Format          =   "##,###,###,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskLimiteIni3 
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         MaxLength       =   17
         Format          =   "##,###,###,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskLimiteFin3 
         Height          =   315
         Left            =   6120
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         MaxLength       =   17
         Format          =   "##,###,###,###,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskLimiteIni4 
         Height          =   315
         Left            =   4440
         TabIndex        =   11
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648384
         MaxLength       =   17
         Format          =   "##,###,###,###,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "...Ao Infinito"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6120
         TabIndex        =   23
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Limite Final:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6120
         TabIndex        =   22
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Limite Inicial:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4440
         TabIndex        =   21
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "4)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2340
         Width           =   150
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "3)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1860
         Width           =   150
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "2)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1380
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "1)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   900
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Grupo:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   600
         Width           =   1170
      End
   End
   Begin VB.Frame fraX 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7815
      Begin VB.Data datParametros 
         Caption         =   "datParametros"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   4935
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmClassificacaoClientes.frx":058A
         Height          =   315
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   855
         DataFieldList   =   "Filial"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Columns(0).Width=   3200
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   8454143
         DataFieldToDisplay=   "Filial"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   300
      End
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmClassificacaoClientes.frx":05A6
   End
End
Attribute VB_Name = "frmClassificacaoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rstGrupos    As Recordset 'Este recordset estará fazendo as devidas buscas na tabela GruposDeClientes

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstParametros As Recordset
  
  txtNomeFilial.Text = ""
  
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstParametros = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & CByte(cboFilial.Text), dbOpenSnapshot)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      txtNomeFilial.Text = .Fields("Nome").Value & ""
    End If
    .Close
  End With

  Set rstParametros = Nothing
  
  '22/07/2004 - Implementado rotina
  'Esta private irá verificar se existe o registro na base
  'caso exista mostrará
  If Len(txtNomeFilial.Text) > 0 Then Call VerificarSeExisteRegistro

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
  If KeyAscii = 13 Then
     SendKeys "{Tab}"
     KeyAscii = 0
 End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Set rstGrupos = db.OpenRecordset("GruposDeClientes", dbOpenDynaset)
  
  datParametros.DatabaseName = gsQuickDBFileName
    
  Me.Show
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstGrupos
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
  
End Sub

Private Sub MoveLast()
  On Error Resume Next
  
  With rstGrupos
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
  
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  
  With rstGrupos
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With

End Sub

Private Sub MoveNext()
  On Error Resume Next
  
  With rstGrupos
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With

End Sub

Private Sub DeleteRecord()
  Dim intResposta  As Integer
  Dim dblAuxCodigo As Double
  Dim strAuxStr    As String
  
  If IsNull(Num_Registro) Then
    Beep
    MsgBox "Não existe nenhuma Classificação de Clientes para apagar.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  strAuxStr = "Deseja realmente apagar esta Classificação de Clientes ? "
  intResposta = MsgBox(strAuxStr, 20, "ATENÇÃO!!")
  If intResposta = 6 Then
    rstGrupos.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim intErro As Integer
  
  On Error GoTo Processa_Erro
  
  '--------------[Validações]------------------------
  If Len(txtNomeFilial.Text) <= 0 Then intErro = True
  If intErro = True Then
    DisplayMsg "Filial incorreta, verifique."
    cboFilial.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Len(txtNome1.Text) <= 0 Then intErro = True
  If intErro = True Then
    DisplayMsg "Nome do Grupo 1 está vazio, verifique."
    txtNome1.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Len(txtNome2.Text) <= 0 Then intErro = True
  If intErro = True Then
    DisplayMsg "Nome do Grupo 2 está vazio, verifique."
    txtNome2.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Len(txtNome3.Text) <= 0 Then intErro = True
  If intErro = True Then
    DisplayMsg "Nome do Grupo 3 está vazio, verifique."
    txtNome3.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Len(txtNome4.Text) <= 0 Then intErro = True
  If intErro = True Then
    DisplayMsg "Nome do Grupo 4 está vazio, verifique."
    txtNome4.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Not IsNumeric(mskLimiteIni1.Text) Then intErro = True
  If intErro = True Then
    DisplayMsg "Limite Inicial do Grupo 1 não é numérico, verifique."
    mskLimiteIni1.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Not IsNumeric(mskLimiteFin1.Text) Then intErro = True
  If intErro = True Then
    DisplayMsg "Limite Final do Grupo 1 não é numérico, verifique."
    mskLimiteFin1.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Not IsNumeric(mskLimiteIni2.Text) Then intErro = True
  If intErro = True Then
    DisplayMsg "Limite Inicial do Grupo 2 não é numérico, verifique."
    mskLimiteIni2.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Not IsNumeric(mskLimiteFin2.Text) Then intErro = True
  If intErro = True Then
    DisplayMsg "Limite Final do Grupo 2 não é numérico, verifique."
    mskLimiteFin2.SetFocus
    Exit Sub
  End If

  intErro = False
  If Not IsNumeric(mskLimiteIni3.Text) Then intErro = True
  If intErro = True Then
    DisplayMsg "Limite Inicial do Grupo 3 não é numérico, verifique."
    mskLimiteIni3.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Not IsNumeric(mskLimiteFin3.Text) Then intErro = True
  If intErro = True Then
    DisplayMsg "Limite Final do Grupo 3 não é numérico, verifique."
    mskLimiteFin3.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If Not IsNumeric(mskLimiteIni4.Text) Then intErro = True
  If intErro = True Then
    DisplayMsg "Limite Inicial do Grupo 4 não é numérico, verifique."
    mskLimiteIni4.SetFocus
    Exit Sub
  End If
  
  If Not VerificaValores Then Exit Sub
  
  '--------------[Fim das Validações]----------------
  
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
   With rstGrupos
     If IsNull(Num_Registro) Then
        .AddNew
        .Fields("Filial") = CByte(cboFilial.Text)
     Else
       .Edit
     End If
     
     .Fields("NomeG1").Value = txtNome1.Text & ""
     .Fields("NomeG2").Value = txtNome2.Text & ""
     .Fields("NomeG3").Value = txtNome3.Text & ""
     .Fields("NomeG4").Value = txtNome4.Text & ""
     .Fields("LimiteIniG1").Value = CDbl(Format(mskLimiteIni1.Text, "##,###,###,###.00"))
     .Fields("LimiteIniG2").Value = CDbl(Format(mskLimiteIni2.Text, "##,###,###,###.00"))
     .Fields("LimiteIniG3").Value = CDbl(Format(mskLimiteIni3.Text, "##,###,###,###.00"))
     .Fields("LimiteIniG4").Value = CDbl(Format(mskLimiteIni4.Text, "##,###,###,###.00"))
     .Fields("LimiteFinG1").Value = CDbl(Format(mskLimiteFin1.Text, "##,###,###,###.00"))
     .Fields("LimiteFinG2").Value = CDbl(Format(mskLimiteFin2.Text, "##,###,###,###.00"))
     .Fields("LimiteFinG3").Value = CDbl(Format(mskLimiteFin3.Text, "##,###,###,###.00"))
     .Fields("CodigoG1").Value = 1 'Além de nome cada grupo terá um código implícito para
     .Fields("CodigoG2").Value = 2 'buscas e pesquisas
     .Fields("CodigoG3").Value = 3
     .Fields("CodigoG4").Value = 4

     .Update
     Num_Registro = .LastModified
     .Bookmark = Num_Registro
   End With
  
  Call StatusMsg("")
  
  Exit Sub
  
Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar gravar registro."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & " - " & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  
  cboFilial.Text = ""
  txtNomeFilial.Text = ""
  txtNome1.Text = ""
  txtNome2.Text = ""
  txtNome3.Text = ""
  txtNome4.Text = ""
  
  mskLimiteIni1.Mask = ""
  mskLimiteIni1.Text = ""
  mskLimiteIni2.Mask = ""
  mskLimiteIni2.Text = ""
  mskLimiteIni3.Mask = ""
  mskLimiteIni3.Text = ""
  mskLimiteIni4.Mask = ""
  mskLimiteIni4.Text = ""

  mskLimiteFin1.Mask = ""
  mskLimiteFin1.Text = ""
  mskLimiteFin2.Mask = ""
  mskLimiteFin2.Text = ""
  mskLimiteFin3.Mask = ""
  mskLimiteFin3.Text = ""
  
  cboFilial.SetFocus
  Num_Registro = Null
  
  If Not rstGrupos.EOF Then
    On Error Resume Next
    rstGrupos.MoveFirst
    rstGrupos.MovePrevious
    On Error GoTo 0
  End If

End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpFirst"
      Call MoveFirst
    Case "miOpPrevious"
      Call MovePrevious
    Case "miOpNext"
      Call MoveNext
    Case "miOpLast"
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstGrupos.Close
  Set rstGrupos = Nothing
End Sub

Sub ShowRecord()
  With rstGrupos
    cboFilial.Text = .Fields("Filial").Value
    cboFilial_LostFocus
    txtNome1.Text = .Fields("NomeG1") & ""
    txtNome2.Text = .Fields("NomeG2") & ""
    txtNome3.Text = .Fields("NomeG3") & ""
    txtNome4.Text = .Fields("NomeG4") & ""
    mskLimiteIni1.Text = .Fields("LimiteIniG1").Value
    mskLimiteIni2.Text = .Fields("LimiteIniG2").Value
    mskLimiteIni3.Text = .Fields("LimiteIniG3").Value
    mskLimiteIni4.Text = .Fields("LimiteIniG4").Value
    mskLimiteFin1.Text = .Fields("LimiteFinG1").Value
    mskLimiteFin2.Text = .Fields("LimiteFinG2").Value
    mskLimiteFin3.Text = .Fields("LimiteFinG3").Value
    
    Num_Registro = .Bookmark
  End With
End Sub

Private Function VerificaValores() As Boolean
  'Esta função tem a finalidade de verificar se os valores iniciais são superiores aos finais
  VerificaValores = True

  If CDbl(mskLimiteIni1.Text) >= CDbl(mskLimiteFin1.Text) Then
    VerificaValores = False
    MsgBox "Limite Final do Grupo 1 é inferior ao Limite Inicial, verifique.", vbExclamation, "Quick Store"
    mskLimiteFin1.SetFocus
    Exit Function
  End If
  
  If CDbl(mskLimiteIni2.Text) >= CDbl(mskLimiteFin2.Text) Then
    VerificaValores = False
    MsgBox "Limite Final do Grupo 2 é inferior ao Limite Inicial, verifique.", vbExclamation, "Quick Store"
    mskLimiteFin2.SetFocus
    Exit Function
  End If
    
  If CDbl(mskLimiteIni3.Text) >= CDbl(mskLimiteFin3.Text) Then
    VerificaValores = False
    MsgBox "Limite Final do Grupo 3 é inferior ao Limite Inicial, verifique.", vbExclamation, "Quick Store"
    mskLimiteFin3.SetFocus
    Exit Function
  End If
  
  'O ideal é que o Limite Inicial 1 seja igual a zero para que não haja
  'clientes sem faixa de classificação, conforme testes realizados e verificados
  'na base de dados
  If CDbl(mskLimiteIni1.Text) > 0 Then
    VerificaValores = False
    MsgBox "Para que não haja clientes sem faixa de classificação, é necessário que o Limite Ini de G1 seja 0.", vbExclamation, "Quick Store"
    mskLimiteIni1.SetFocus
    Exit Function
  End If
  
  'O 4 não pode ser menor que o 3, 2, 1
  If CDbl(mskLimiteIni4.Text) <= CDbl(mskLimiteFin3.Text) Then
    VerificaValores = False
    MsgBox "Limite Inicial do Grupo 4 é inferior ou igual ao Limite Final do Grupo 3, verifique.", vbExclamation, "Quick Store"
    mskLimiteIni4.SetFocus
    Exit Function
  End If
  
  If CDbl(mskLimiteIni4.Text) <= CDbl(mskLimiteFin2.Text) Then
    VerificaValores = False
    MsgBox "Limite Inicial do Grupo 4 é inferior ou igual ao Limite Final do Grupo 2, verifique.", vbExclamation, "Quick Store"
    mskLimiteIni4.SetFocus
    Exit Function
  End If
                               
  If CDbl(mskLimiteIni4.Text) <= CDbl(mskLimiteFin1.Text) Then
    VerificaValores = False
    MsgBox "Limite Inicial do Grupo 4 é inferior ou igual ao Limite Final do Grupo 1, verifique.", vbExclamation, "Quick Store"
    mskLimiteIni4.SetFocus
    Exit Function
  End If
  
  'O 3 não poderá ser menor igual ao 2 e 1
  If CDbl(mskLimiteFin3.Text) <= CDbl(mskLimiteFin2.Text) Then
    VerificaValores = False
    MsgBox "Limite Final do Grupo 3 é inferior ou igual ao Limite Final do Grupo 2, verifique.", vbExclamation, "Quick Store"
    mskLimiteFin3.SetFocus
    Exit Function
  End If
  
  If CDbl(mskLimiteFin3.Text) <= CDbl(mskLimiteFin1.Text) Then
    VerificaValores = False
    MsgBox "Limite Final do Grupo 3 é inferior ou igual ao Limite Final do Grupo 1, verifique.", vbExclamation, "Quick Store"
    mskLimiteFin3.SetFocus
    Exit Function
  End If
  
  'O 2 não poderá ser menor igual ao 1
  If CDbl(mskLimiteFin2.Text) <= CDbl(mskLimiteFin1.Text) Then
    VerificaValores = False
    MsgBox "Limite Final do Grupo 2 é inferior ou igual ao Limite Final do Grupo 1, verifique.", vbExclamation, "Quick Store"
    mskLimiteFin2.SetFocus
    Exit Function
  End If
  
End Function

Private Sub VerificarSeExisteRegistro()
  Dim rstGrupos As Recordset
  Dim strQuery  As String
  
  strQuery = "SELECT * FROM GruposDeClientes "
  strQuery = strQuery & " WHERE Filial = " & CByte(cboFilial.Text)
  
  Set rstGrupos = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstGrupos.RecordCount = 0 Then Exit Sub
  
  With rstGrupos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      txtNome1.Text = .Fields("NomeG1").Value & ""
      txtNome2.Text = .Fields("NomeG2").Value & ""
      txtNome3.Text = .Fields("NomeG3").Value & ""
      txtNome4.Text = .Fields("NomeG4").Value & ""
      mskLimiteIni1.Text = .Fields("LimiteIniG1").Value
      mskLimiteIni2.Text = .Fields("LimiteIniG2").Value
      mskLimiteIni3.Text = .Fields("LimiteIniG3").Value
      mskLimiteIni4.Text = .Fields("LimiteIniG4").Value
      mskLimiteFin1.Text = .Fields("LimiteFinG1").Value
      mskLimiteFin2.Text = .Fields("LimiteFinG2").Value
      mskLimiteFin3.Text = .Fields("LimiteFinG3").Value
      
    End If
    .Close
  End With

  Set rstGrupos = Nothing
  
End Sub


