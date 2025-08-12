VERSION 5.00
Begin VB.Form frmVendaRap1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFA324&
   BorderStyle     =   0  'None
   Caption         =   "Venda Rápida - Operador de Caixa"
   ClientHeight    =   6585
   ClientLeft      =   2685
   ClientTop       =   1230
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00666666&
   Icon            =   "VendaRap1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleMode       =   0  'User
   ScaleWidth      =   15808.26
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Prosseguir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   270
      MaskColor       =   &H00E5E5E5&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4770
      Width           =   11385
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   270
      MaskColor       =   &H00E5E5E5&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5535
      Width           =   11385
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   315
      TabIndex        =   4
      Top             =   2970
      Width           =   11355
      Begin VB.TextBox Senha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00999999&
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   8
         PasswordChar    =   "•"
         TabIndex        =   2
         Top             =   90
         Width           =   2775
      End
      Begin VB.TextBox Caixa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   15
         MaxLength       =   2
         TabIndex        =   1
         Top             =   810
         Width           =   1470
      End
      Begin VB.Label Vendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1725
         TabIndex        =   9
         Top             =   1380
         Visible         =   0   'False
         Width           =   9660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   75
         TabIndex        =   8
         Top             =   150
         Width           =   600
      End
      Begin VB.Label Nome_Caixa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " <- Selecione um caixa..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1590
         TabIndex        =   7
         Top             =   810
         Width           =   9660
      End
      Begin VB.Label Cod_Operador 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   150
         TabIndex        =   6
         Top             =   1380
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Destino 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.ListBox Lista1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2100
      Left            =   315
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   450
      Width           =   11265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escolha o Operador e o Caixa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F7F7F7&
      Height          =   435
      Left            =   3390
      TabIndex        =   3
      Top             =   -45
      Width           =   4755
   End
End
Attribute VB_Name = "frmVendaRap1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsFuncionarios As Recordset
Private rsCaixas As Recordset
Private rsParametros As Recordset
Private itContador As Integer
Private stNomeLogado As String

Private Sub Caixa_LostFocus()

  On Error GoTo ErrHandler


  Call StatusMsg("")
  Nome_Caixa.Caption = ""
  If IsNull(Caixa.Text) Then Exit Sub
  If Caixa.Text = "" Then Exit Sub
  If Not IsNumeric(Caixa.Text) Then Exit Sub
  If Val(Caixa.Text) < 1 Or Val(Caixa.Text) > 99 Then Exit Sub
  
  
  rsCaixas.Index = "Caixa"
  rsCaixas.Seek "=", Val(Caixa.Text)
  If rsCaixas.NoMatch Then
    DisplayMsg "Caixa não cadastrado."
    Exit Sub
  End If
  
  Nome_Caixa.Caption = rsCaixas("Descrição") & ""
 
  Exit Sub
 
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
 
End Sub


Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
  
  On Error GoTo ErrHandler


  Senha_LostFocus
  
  DoEvents
  
  Call StatusMsg("")
  
  If Vendedor.Caption = "" Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "Escolha um operador antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Cod_Operador.Caption)
  If rsFuncionarios.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Funcionário não existe."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If CriptografaSenha(Senha.Text) <> rsFuncionarios("ValorP") Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "Senha incorreta."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  
  If Nome_Caixa.Caption = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Caixa não digitado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If Caixa.Enabled Then Caixa.SetFocus
    Exit Sub
  End If
  
'  If Destino.Caption = "VENDA RÁPIDA" Then
    Call StatusMsg("Aguarde...")
    DoEvents
    
    Senha.Text = ""
    
    
    '16/01/2006 - mpdea
    'Escolha da tela de Venda Rápida a ser utilizada
    If g_frmVendaRapida Is Nothing Then
      If rsParametros.Fields("VR_Tela_CheckOut").Value Then
'''        'Minimiza a tela principal para evitar exibições incorretas por foco
'''        frmMain.WindowState = vbMinimized
'''        'Utilização da tela de Venda Rápida em tela cheia
'''        Set g_frmVendaRapida = frmVendaRap2_CheckOut
          
          'Tela Padrão
          Set g_frmVendaRapida = frmVendaRap2
          frmPesquisaProduto2.Show

      Else
          'Tela Padrão
          Set g_frmVendaRapida = frmVendaRap2
      End If
    End If
    
    
    Funcionario = Cod_Operador.Caption
    g_frmVendaRapida.Nome_Operador.Caption = " " & Vendedor.Caption
    g_frmVendaRapida.Cod_Operador.Caption = Cod_Operador.Caption
    'frmVendaRap2.Cod_Caixa.Caption = Caixa.Text
    '22/10/2004 - Daniel
    'Flexibilidade de troca de caixa
    'Case: Solicitado por Casagrande
    g_frmVendaRapida.cboCaixa.Text = Caixa.Text
    g_frmVendaRapida.Nome_Caixa.Caption = " " & Nome_Caixa.Caption
    
    g_frmVendaRapida.Combo_Vendedor = gnUserCode
    g_frmVendaRapida.Nome_Vendedor = stNomeLogado
    
    g_frmVendaRapida.Show
    g_frmVendaRapida.CheckMovimentacao
    
    Call StatusMsg("")
    
    Unload Me
    
'  End If
  
  Exit Sub
 
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Form_Load()
  Dim Fim As Integer
  Dim Cód As Long
  Dim Aux As String, Aux2 As String
  Dim iContador As Integer
  
  
  On Error GoTo ErrHandler
  
  
  iContador = 0
  
  Call CenterForm(Me)
  
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsCaixas = db.OpenRecordset("Caixas em Uso")
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  rsFuncionarios.Index = "Código"
  Fim = False
  Cód = 0
  Do
    rsFuncionarios.Seek ">", Cód
    If rsFuncionarios.NoMatch Then Fim = True
    If Fim = False Then
      Cód = rsFuncionarios("Código")
      '06/06/2005 - Daniel
      'Adicionado cláusula And rsFuncionarios("Ativo") = True
      'na linha abaixo
      If rsFuncionarios("Liberado") = True And rsFuncionarios("Ativo") = True And rsFuncionarios("isPrestServ") = False Then
        Aux = rsFuncionarios("Nome")
        Aux2 = "0000" + LTrim(str(rsFuncionarios("Código")))
        Aux2 = Right$(Aux2, 4)
        'Aux = Aux + " - " + Aux2
        Aux = Aux2 & " - " & Aux
        Lista1.AddItem Aux
        If Cód = gnUserCode Then
            itContador = iContador
            stNomeLogado = rsFuncionarios("Nome")
        End If
        iContador = iContador + 1
     End If
    End If
  Loop Until Fim = True
  
  Vendedor.Caption = ""
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then
    MsgBox ("Erro. Filial não encontrada.")
    Exit Sub
  End If
  
  If gbCaixas = False Then
    Caixa.Text = 1
    Caixa_LostFocus
    Caixa.Enabled = False
  End If
  
  Lista1.ListIndex = itContador
  Senha.Text = gSenhaUsuarioLogado
 
  Exit Sub
 
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsFuncionarios.Close
  rsCaixas.Close
  rsParametros.Close
  Set rsFuncionarios = Nothing
  Set rsCaixas = Nothing
  Set rsParametros = Nothing
End Sub

Private Sub Lista1_Click()
  Dim Str_Aux As String
  Dim Tamanho As Integer
  
  Str_Aux = Lista1.List(Lista1.ListIndex)
  Tamanho = Len(Str_Aux)
  
  'Cod_Operador.Caption = Val(Right(Str_Aux, 4))
  Cod_Operador.Caption = Val(Left(Str_Aux, 4))
 
  Str_Aux = Trim(Right(Str_Aux, (Tamanho - 6)))
  Vendedor.Caption = Str_Aux
  gsVendedorVR = Str_Aux
  
'  If Caixa.Enabled Then
'    Caixa.SetFocus
'  Else
'''    Senha.SetFocus
'  End If
End Sub

Private Sub Senha_LostFocus()
  If IsNull(Senha.Text) Then Exit Sub
  If Senha.Text = "" Then Exit Sub
End Sub
