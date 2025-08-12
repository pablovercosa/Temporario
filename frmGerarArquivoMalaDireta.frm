VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmGerarArquivoMalaDireta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Arquivo para Mala Direta"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGerarArquivoMalaDireta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   6480
   Begin VB.Frame fraA 
      Caption         =   "Ajuste de margens da impressora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   37
      Top             =   4080
      Width           =   6400
      Begin VB.Frame fraSaida 
         Caption         =   "Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4680
         TabIndex        =   41
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton optVideo 
            Caption         =   "Vídeo"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optImpressora 
            Caption         =   "Impressora"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.OptionButton optCEP 
         Caption         =   "CEP"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton optBairro 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton optCidade 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optCodigo 
         Caption         =   "Código"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H0000C0C0&
         Caption         =   "Im&primir"
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Click neste botão para gerar as informações em relatório ou imprimir direto."
         Top             =   1800
         Width           =   1575
      End
      Begin ComctlLib.Slider sldEsquerda 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   327682
         LargeChange     =   1
         Min             =   -7
         Max             =   7
      End
      Begin ComctlLib.Slider sldSuperior 
         Height          =   1695
         Left            =   2280
         TabIndex        =   16
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
      Begin Crystal.CrystalReport rptRel 
         Left            =   360
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ordem:"
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
         Left            =   3480
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblSuperior 
         Alignment       =   2  'Center
         Caption         =   "Superior = padrão"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblEsquerda 
         Alignment       =   2  'Center
         Caption         =   "Esquerda = padrão"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gerar a Informação como:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      TabIndex        =   36
      Top             =   2280
      Width           =   3270
      Begin VB.OptionButton optArquivo 
         Caption         =   "Arquivo TXT"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optRelatorio 
         Caption         =   "Relatório"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Clientes Ativos / Inativos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   35
      Top             =   2280
      Width           =   3120
      Begin VB.OptionButton optTodos 
         Caption         =   "Imprimir todos"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optAtivos 
         Caption         =   "Imprimir somente ativos"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optInativos 
         Caption         =   "Imprimir somente inativos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdGerar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Gerar Arquivo"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Click neste botão para gerar as informações em arquivo .TXT"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Frame fraX 
      Caption         =   "Intervalo de Datas de Aniversário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   27
      Top             =   960
      Width           =   6400
      Begin VB.ComboBox cboMesFin 
         Height          =   315
         ItemData        =   "frmGerarArquivoMalaDireta.frx":058A
         Left            =   4920
         List            =   "frmGerarArquivoMalaDireta.frx":05B2
         TabIndex        =   8
         Top             =   660
         Width           =   855
      End
      Begin VB.ComboBox cboDiaFin 
         Height          =   315
         ItemData        =   "frmGerarArquivoMalaDireta.frx":05E6
         Left            =   3690
         List            =   "frmGerarArquivoMalaDireta.frx":0647
         TabIndex        =   7
         Top             =   660
         Width           =   855
      End
      Begin VB.ComboBox cboMesIni 
         Height          =   315
         ItemData        =   "frmGerarArquivoMalaDireta.frx":06C7
         Left            =   1800
         List            =   "frmGerarArquivoMalaDireta.frx":06EF
         TabIndex        =   6
         Top             =   660
         Width           =   855
      End
      Begin VB.ComboBox cboDiaIni 
         Height          =   315
         ItemData        =   "frmGerarArquivoMalaDireta.frx":0723
         Left            =   565
         List            =   "frmGerarArquivoMalaDireta.frx":0784
         TabIndex        =   5
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Mês Final"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4920
         TabIndex        =   34
         Top             =   450
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mês Inicial"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1800
         TabIndex        =   33
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "|-- ao --|"
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
         Left            =   2767
         TabIndex        =   32
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "de"
         Height          =   195
         Left            =   4635
         TabIndex        =   31
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dia Final"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3720
         TabIndex        =   30
         Top             =   450
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "de"
         Height          =   195
         Left            =   1515
         TabIndex        =   29
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dia Inicial"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   600
         TabIndex        =   28
         Top             =   450
         Width           =   675
      End
   End
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
      Height          =   855
      Left            =   0
      TabIndex        =   26
      ToolTipText     =   "Selecione o Grupo conforme critério estipulado na tela de Classificação de Clientes."
      Top             =   0
      Width           =   6400
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "&Consultar Grupo"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optG4 
         Caption         =   "4"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optG3 
         Caption         =   "3"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optG2 
         Caption         =   "2"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optG1 
         Caption         =   "1"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   2640
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmGerarArquivoMalaDireta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_intAdjustLeft As Integer
Dim m_intAdjustTop  As Integer
Dim m_datInicial    As Date
Dim m_datFinal      As Date

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdConsultar_Click()
  'Primeiro é validado se o usuário que está clicando é responsável
  'ou não pelo Martketing da empresa
  Dim rstFuncionarios As Recordset
  Dim strQuery        As String
  
  strQuery = "SELECT Código, Marketing "
  strQuery = strQuery & " FROM Funcionários "
  strQuery = strQuery & " WHERE Código = " & gnUserCode
  
  Set rstFuncionarios = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      If .Fields("Marketing").Value Then
        frmClassificacaoClientes.Show
      Else
        MsgBox "Usuário não responsável pelo Marketing.", vbExclamation, "Atenção"
      End If
    End If
    .Close
  End With

  Set rstFuncionarios = Nothing

End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdGerar_Click()
  'Validamos os dados informados em primeiro momento
  If Not ValidarDados Then Exit Sub
  
  Screen.MousePointer = vbHourglass
  
  Call CarregarInformacoes
  
  Screen.MousePointer = vbDefault
  
  Call StatusMsg("Criação do arquivo de mala direta...")
  
  Call CriarArquivo
  
  Call StatusMsg("")
  
  Call CriarArquivoComErros
  
End Sub

Private Sub cmdImprimir_Click()
  'Validamos os dados informados em primeiro momento
  If Not ValidarDados Then Exit Sub

  Call CarregarInformacoes

  'Mostrar primeiro ao usuário a lista de clientes com
  'data de aniversário incorreta e em seguida carregar
  'o relatório
  Call CriarArquivoComErros

  Screen.MousePointer = vbHourglass

  Call CarregarMalaDiretaTempo
  
  Screen.MousePointer = vbDefault
  
  Call Impressao

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  cmdGerar.Visible = False
  fraA.Visible = True
  Me.Height = 6900
  
End Sub

Private Function ValidarDados() As Boolean
  Dim strAuxi       As String
  Dim intAnoInicial As Integer
  Dim intAnoFinal   As Integer
  Dim strMensagem   As String
  
  ValidarDados = True
  
  intAnoInicial = CInt(Right(Data_Atual, 4))
  intAnoFinal = CInt(Right(Data_Atual, 4))
  
  '-------[Dia]-------
  If IsNumeric(cboDiaIni.Text) Then
    If (cboDiaIni.Text) <= "0" Or cboDiaIni.Text > "31" Then
      ValidarDados = False
      MsgBox "Dia inicial inválido, verifique.", vbExclamation, "Quick Store"
      cboDiaIni.SetFocus
      Exit Function
    End If
  Else
    ValidarDados = False
    MsgBox "Dia inicial inválido, verifique.", vbExclamation, "Quick Store"
    cboDiaIni.SetFocus
    Exit Function
  End If
  
  If IsNumeric(cboDiaFin.Text) Then
    If (cboDiaFin.Text) <= "0" Or cboDiaFin.Text > "31" Then
      ValidarDados = False
      MsgBox "Dia Final inválido, verifique.", vbExclamation, "Quick Store"
      cboDiaFin.SetFocus
      Exit Function
    End If
  Else
    ValidarDados = False
    MsgBox "Dia Final inválido, verifique.", vbExclamation, "Quick Store"
    cboDiaFin.SetFocus
    Exit Function
  End If
  '-------[Dia]-------
  
  '-------[Mês]-------
  If IsNumeric(cboMesIni.Text) Then
    If (cboMesIni.Text) <= "0" Or (cboMesIni.Text) >= "13" Then
      ValidarDados = False
      MsgBox "Mês inicial inválido, verifique.", vbExclamation, "Quick Store"
      cboMesIni.SetFocus
      Exit Function
    End If
  Else
    ValidarDados = False
    MsgBox "Mês inicial inválido, verifique.", vbExclamation, "Quick Store"
    cboMesIni.SetFocus
    Exit Function
  End If
  
  If IsNumeric(cboMesFin.Text) Then
    If (cboMesFin.Text) <= "0" Or (cboMesFin.Text) >= "13" Then
      ValidarDados = False
      MsgBox "Mês Final inválido, verifique.", vbExclamation, "Quick Store"
      cboMesFin.SetFocus
      Exit Function
    End If
  Else
    ValidarDados = False
    MsgBox "Mês Final inválido, verifique.", vbExclamation, "Quick Store"
    cboMesFin.SetFocus
    Exit Function
  End If
  '-------[Mês]-------
  
  'Caso a Function continue como True verificaremos se esta data é válida
  'Concatenação dos valores para formar dd/mm/yyyy
  
  '-------[Fevereiro]-------
  If Month(Data_Atual) = 2 Then
    If (cboDiaIni.Text) >= "30" Then  'Pode ocorrer 28 ou 29 superior nunca
      ValidarDados = False
      MsgBox "Dia inicial inválido, verifique.", vbExclamation, "Quick Store"
      cboDiaIni.SetFocus
      Exit Function
    End If
    
    If (cboDiaFin.Text) >= "30" Then  'Pode ocorrer 28 ou 29 superior nunca
      ValidarDados = False
      MsgBox "Dia Final inválido, verifique.", vbExclamation, "Quick Store"
      cboDiaFin.SetFocus
      Exit Function
    End If
  End If
  
  '-------[ABR/JUN/SET/NOV]-------
  If Month(Data_Atual) = 4 Or Month(Data_Atual) = 6 Or Month(Data_Atual) = 9 Or Month(Data_Atual) = 11 Then
    If (cboDiaIni.Text) >= "31" Then  'Pode ocorrer 30 superior nunca
      ValidarDados = False
      MsgBox "Dia inicial inválido, mês corrente possui 30 dias.", vbExclamation, "Quick Store"
      cboDiaIni.SetFocus
      Exit Function
    End If
    
    If (cboDiaFin.Text) >= "31" Then  'Pode ocorrer 30 superior nunca
      ValidarDados = False
      MsgBox "Dia Final inválido, mês corrente possui 30 dias.", vbExclamation, "Quick Store"
      cboDiaFin.SetFocus
      Exit Function
    End If
  End If
  
  'Passou pelas duas condições
  'então:
  strAuxi = (cboDiaIni.Text) & "/" & (cboMesIni.Text) & "/" & CStr(intAnoInicial)
  
  If IsDate(strAuxi) Then
    m_datInicial = Format(strAuxi, "dd/mm/yyyy")
  Else
    ValidarDados = False
    MsgBox "Data inicial está inválida, verifique.", vbExclamation, "Quick Store"
    cboDiaIni.SetFocus
    Exit Function
  End If
  
  strAuxi = (cboDiaFin.Text) & "/" & (cboMesFin.Text) & "/" & CStr(intAnoFinal)
  
  If IsDate(strAuxi) Then
    m_datFinal = Format(strAuxi, "dd/mm/yyyy")
  Else
    ValidarDados = False
    MsgBox "Data final está inválida, verifique.", vbExclamation, "Quick Store"
    cboDiaFin.SetFocus
    Exit Function
  End If
  
  'Validação do ano
  If m_datInicial > m_datFinal Then
    ValidarDados = False
    cboDiaIni.SetFocus
    
    strMensagem = "Data Inicial maior que a Data Final, ou você está "
    strMensagem = strMensagem & vbCrLf & "tentando emitir de um ano para o outro, exemplo: "
    strMensagem = strMensagem & vbCrLf & "de Dez/2004 a Jan/2005, caso seja isto, você deverá gerar "
    strMensagem = strMensagem & vbCrLf & "arquivos distintos um para cada ano." & Chr(13)
    
    MsgBox strMensagem, vbExclamation, "Atenção"
    
    Exit Function
  End If
  
End Function

Private Sub CriarArquivo()
  Dim rstMalaExportacao As Recordset
  Dim strQuery          As String
  Dim strAuxi1          As String
  Dim strAuxi2          As String
  Dim strNomeArquivo    As String
  Dim intResp           As Integer
  
  strAuxi1 = gsReportPath & "MALA"
  strAuxi2 = Format(Date, "dd/mm/yy")
  strAuxi1 = strAuxi1 & Left(strAuxi2, 2)
  strAuxi1 = strAuxi1 & Mid(strAuxi2, 4, 2)
  strAuxi1 = strAuxi1 & Mid(strAuxi2, 7, 4)
  
  strAuxi1 = strAuxi1 & ".txt"
  
  Dialog1.FileName = strAuxi1
  
  On Error GoTo Erro_Gravar
  
    With Dialog1
      .CancelError = True
      .DialogTitle = "Salvar arquivo para a Mala Direta como"
      .DefaultExt = "txt"
      .InitDir = gsDefaultPath
      .Filter = "Arquivo para Mala Direta | *.txt"
      .Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
      .ShowSave
    End With
    
  On Error GoTo 0
  
    strNomeArquivo = Dialog1.FileName
    If Dir(strNomeArquivo) <> "" Then
      intResp = MsgBox("Já existe este arquivo, deseja sobrescrever ?", vbQuestion + vbOKCancel, "Atenção")
      If intResp = vbCancel Then
        DisplayMsg "Geração de arquivo cancelada."
        Exit Sub
      End If
    End If
  
    strQuery = "SELECT * FROM MalaExportacao "
    strQuery = strQuery & " WHERE DataIncorreta = FALSE " 'Para não trazer informações incompletas
    strQuery = strQuery & " AND Nascimento >= #" & Format(m_datInicial, "mm/dd/yyyy") & "#"
    strQuery = strQuery & " AND Nascimento <= #" & Format(m_datFinal, "mm/dd/yyyy") & "#"
    strQuery = strQuery & " ORDER BY Codigo "
    
    Set rstMalaExportacao = dbTemp.OpenRecordset(strQuery, dbOpenDynaset)
  
    With rstMalaExportacao
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Open strNomeArquivo For Output As #1
            
          Do Until .EOF
            'Lay Out do arquivo de mala direta
            'Campo            Tamanho       Posição
            '
            'Nome               70             01 a  70
            'Endereco          100             71 a 170
            'Complemento        20            171 a 190
            'Bairro             30            191 a 220
            'CEP                09            221 a 229
            'Cidade             40            230 a 269
            'Estado             02            270 a 271
            '
            Print #1, Left((.Fields("Nome").Value & "" & String(70, " ")), 70) & _
                      Left((.Fields("Endereco").Value & "" & String(100, " ")), 100) & _
                      Left((.Fields("Complemento").Value & "" & String(20, " ")), 20) & _
                      Left((.Fields("Bairro").Value & "" & String(30, " ")), 30) & _
                      Left((.Fields("CEP").Value & "" & String(9, " ")), 9) & _
                      Left((.Fields("Cidade").Value & "" & String(40, " ")), 40) & _
                      Left((.Fields("Estado").Value & "" & String(2, " ")), 2)
          
          .MoveNext
          Loop
      
        Close #1
        
      End If
      .Close
    End With
  
    Set rstMalaExportacao = Nothing
    
    DisplayMsg "Geração efetuada com Sucesso."
    Exit Sub

Erro_Gravar:
  MsgBox "Impossível gerar o arquivo. Processo não executado.", vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Sub CriarArquivoComErros()
  'Esta rotina criará um arquivo onde mostrará nele os clientes
  'que estejam com a data de nascimento inválida
  Dim rstMalaExportacao As Recordset
  Dim strQuery          As String
  Dim strAuxi1          As String
  Dim strAuxi2          As String
  Dim strNomeArquivo    As String
  Dim intResp           As Integer
  
  strQuery = "SELECT * FROM MalaExportacao "
  strQuery = strQuery & " WHERE DataIncorreta = TRUE " 'Onde a data está zerada ou o mês vazio
  strQuery = strQuery & " ORDER BY Codigo "
  
  Set rstMalaExportacao = dbTemp.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstMalaExportacao.RecordCount = 0 Then
    rstMalaExportacao.Close
    Set rstMalaExportacao = Nothing
    Exit Sub 'Sai fora e não gera arquivo com erros
  Else
    Sleep (1000) 'Aguardará 01 segundo
    MsgBox "Ocorrerá geração de arquivo com Clientes que possuem datas inválidas.", vbExclamation, "Atenção"
  End If
  
  strAuxi1 = gsReportPath & "ERRO-MALA"
  strAuxi2 = Format(Date, "dd/mm/yy")
  strAuxi1 = strAuxi1 & Left(strAuxi2, 2)
  strAuxi1 = strAuxi1 & Mid(strAuxi2, 4, 2)
  strAuxi1 = strAuxi1 & Mid(strAuxi2, 7, 4)
  
  strAuxi1 = strAuxi1 & ".txt"
  
  Dialog1.FileName = strAuxi1
  
  On Error GoTo Erro_Gravar
  
    With Dialog1
      .CancelError = True
      .DialogTitle = "Salvar arquivo de Clientes com datas inválidas como"
      .DefaultExt = "txt"
      .InitDir = gsDefaultPath
      .Filter = "Arquivo com erros de Clientes da Mala Direta | *.txt"
      .Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
      .ShowSave
    End With
    
  On Error GoTo 0
  
    strNomeArquivo = Dialog1.FileName
    If Dir(strNomeArquivo) <> "" Then
      intResp = MsgBox("Já existe este arquivo, deseja sobrescrever ?", vbQuestion + vbOKCancel, "Atenção")
      If intResp = vbCancel Then
        DisplayMsg "Geração de arquivo cancelada."
        Exit Sub
      End If
    End If
    
    '-----[Recordset]-----
    With rstMalaExportacao
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Open strNomeArquivo For Output As #1

          Print #1, "----------------------------------------------------------"
          Print #1, "Os seguintes clientes abaixo precisam ser atualizados pois"
          Print #1, "apresentaram 'Data de Nascimento' inválida em Contatos.   "
          Print #1, "Eles pertencem ao grupo que você selecionou, verifique.   "
          Print #1, "----------------------------------------------------------"
        
          Do Until .EOF
            Print #1, Right(((String(5, " ")) & .Fields("Codigo").Value), 6) & " - " & .Fields("Nome").Value & ""
          
          .MoveNext
          Loop
        
        Close #1
    
      End If
      .Close
    End With
    
    Set rstMalaExportacao = Nothing

    DisplayMsg "Geração efetuada com Sucesso."
    Exit Sub

Erro_Gravar:
  MsgBox "Impossível gerar o arquivo. Processo não executado.", vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Sub CarregarInformacoes()
  Dim rstClientes       As Recordset
  Dim rstMalaExportacao As Recordset
  Dim strSQL            As String
  Dim bytCodGrupo       As Byte
  Dim strAuxi           As String
  '13/01/2005 - Daniel
  'Variável que verificará se a data é válida para não ocorrer
  'absurdos como '31/09/...' Setembro tem 30 dias
  Dim strData           As String
  Dim intAnoAtual       As Integer
  Dim bytAuxi           As Byte
  
  intAnoAtual = Year(Data_Atual)
  
  'A tabela MalaExportacao foi desenvolvida especialmente para
  'este case da TV Shopping
  dbTemp.Execute "DELETE * FROM MalaExportacao"
  
  If optG1.Value Then bytCodGrupo = 1
  If optG2.Value Then bytCodGrupo = 2
  If optG3.Value Then bytCodGrupo = 3
  If optG4.Value Then bytCodGrupo = 4
  
  strSQL = "SELECT Cli_For.Código, Cli_For.Nome, Cli_For.Endereço, Cli_For.Complemento, Cli_For.Bairro, Cli_For.CEP, Cli_For.Cidade, Cli_For.Estado, Cli_For.CodGrupo, Cli_For.Inativo, "
  strSQL = strSQL & " Contatos.Contato AS NomeContato, Contatos.[Dia Aniversário] AS Dia, Contatos.[Mês Aniversário] AS Mes "
  strSQL = strSQL & " FROM Cli_For, Contatos "
  strSQL = strSQL & " WHERE Cli_For.Código = Contatos.Cliente "
  strSQL = strSQL & " AND Cli_For.Tipo = 'C' " 'Clientes
  strSQL = strSQL & " AND Cli_For.CodGrupo = " & bytCodGrupo 'Grupo
  'Adicionado linha para filtrar Ativos / Inativos / Todos
  If optAtivos.Value Then strSQL = strSQL & " AND NOT Cli_For.Inativo "
  If optInativos.Value Then strSQL = strSQL & " AND Cli_For.Inativo "
  '-------------------------------------------------------------------
  strSQL = strSQL & " ORDER BY Cli_For.Código "
  
  Set rstClientes = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  'Abrindo a tabela temporária
  Set rstMalaExportacao = dbTemp.OpenRecordset("MalaExportacao", dbOpenDynaset)
  
  
  With rstClientes
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        rstMalaExportacao.AddNew
          '04/08/2004 - Daniel
          'Adicionado Tratamento dos campos
          rstMalaExportacao.Fields("Codigo").Value = .Fields("Código").Value
          rstMalaExportacao.Fields("Nome").Value = .Fields("NomeContato").Value & ""  '.Fields("Nome").Value & ""
          
          If Len(.Fields("Endereço").Value) > 100 Then
            rstMalaExportacao.Fields("Endereco").Value = Left((.Fields("Endereço").Value & ""), 100)
          Else
            rstMalaExportacao.Fields("Endereco").Value = .Fields("Endereço").Value & ""
          End If
          
          If Len(.Fields("Complemento").Value) > 20 Then
            rstMalaExportacao.Fields("Complemento").Value = Left((.Fields("Complemento").Value & ""), 20)
          Else
            rstMalaExportacao.Fields("Complemento").Value = .Fields("Complemento").Value & ""
          End If
          
          If Len(.Fields("Bairro").Value) > 30 Then
            rstMalaExportacao.Fields("Bairro").Value = Left((.Fields("Bairro").Value & ""), 30)
          Else
            rstMalaExportacao.Fields("Bairro").Value = .Fields("Bairro").Value & ""
          End If
          
          If Len(.Fields("CEP").Value) > 9 Then
            rstMalaExportacao.Fields("CEP").Value = Left((.Fields("CEP").Value), 9)
          Else
            rstMalaExportacao.Fields("CEP").Value = .Fields("CEP").Value
          End If
          
          If Len(.Fields("Cidade").Value) > 40 Then
            rstMalaExportacao.Fields("Cidade").Value = Left((.Fields("Cidade").Value & ""), 40)
          Else
            rstMalaExportacao.Fields("Cidade").Value = .Fields("Cidade").Value & ""
          End If
          
          rstMalaExportacao.Fields("Estado").Value = .Fields("Estado").Value & ""
          
          If .Fields("Dia").Value <> 0 And Len(.Fields("Mes").Value) > 0 Then
            Select Case .Fields("Mes").Value
              Case "JAN"
                strAuxi = "01"
              Case "FEV"
                strAuxi = "02"
              Case "MAR"
                strAuxi = "03"
              Case "ABR"
                strAuxi = "04"
              Case "MAI"
                strAuxi = "05"
              Case "JUN"
                strAuxi = "06"
              Case "JUL"
                strAuxi = "07"
              Case "AGO"
                strAuxi = "08"
              Case "SET"
                strAuxi = "09"
              Case "OUT"
                strAuxi = "10"
              Case "NOV"
                strAuxi = "11"
              Case Else
                strAuxi = "12"
            End Select
          
            '13/01/2005 - Daniel
            'Verificação se a data é válida, encontramos na base de
            'dados da TV contatos cadastrados com data de aniversário
            '31/SET isto é dado incorreto pois SET tem 30 dias
            strData = .Fields("Dia").Value & "/" & strAuxi & "/" & intAnoAtual 'O ano não vai importar para a query
          
            If IsDate(strData) Then
              rstMalaExportacao.Fields("Nascimento").Value = Format((CDate(.Fields("Dia").Value & "/" & strAuxi & "/" & intAnoAtual)), "dd/mm/yyyy") 'O ano não vai importar para a query
            Else
              'Colocamos a data para o último dia útil do mês
              'Exemplo: Se estiver 31/SET ficará 30/SET
              If strAuxi = "02" Then 'FEV
                bytAuxi = (.Fields("Dia").Value) - 3
                rstMalaExportacao.Fields("Nascimento").Value = Format((CDate(bytAuxi & "/" & strAuxi & "/" & intAnoAtual)), "dd/mm/yyyy")
              Else
                bytAuxi = (.Fields("Dia").Value) - 1
                rstMalaExportacao.Fields("Nascimento").Value = Format((CDate(bytAuxi & "/" & strAuxi & "/" & intAnoAtual)), "dd/mm/yyyy")
              End If
            End If
            'Antiga linha comentada em 13/01/2005
            'rstMalaExportacao.Fields("Nascimento").Value = Format((CDate(.Fields("Dia").Value & "/" & strAuxi & "/2004")), "dd/mm/yyyy") 'O ano não vai importar para a query
          End If
          
          'Tratamento para clientes com Data de Nascimento Incorreta
          If .Fields("Dia").Value = 0 Or Len(.Fields("Mes").Value) <= 0 Then
            rstMalaExportacao.Fields("DataIncorreta").Value = True
          Else
            rstMalaExportacao.Fields("DataIncorreta").Value = False
          End If
          
          rstMalaExportacao.Fields("Grupo").Value = .Fields("CodGrupo").Value
        rstMalaExportacao.Update
      
      .MoveNext
      Loop
      
    End If
    .Close
    rstMalaExportacao.Close
  End With
  
  Set rstClientes = Nothing
  Set rstMalaExportacao = Nothing

End Sub

Private Sub optArquivo_Click()
  cmdGerar.Visible = True
  fraA.Visible = False
  Me.Height = 4455
End Sub

Private Sub optRelatorio_Click()
  cmdGerar.Visible = False
  fraA.Visible = True
  Me.Height = 6900
End Sub

Private Sub CarregarMalaDiretaTempo()
  Dim rstMalaDiretaTempo As Recordset   'QS.mdb
  Dim rstMalaExportacao  As Recordset   'Temp.mdb
  Dim strQuery           As String

  'Limpamos o contéudo da table [Mala Direta - Tempo]
  db.Execute "DELETE * FROM [Mala Direta - Tempo]"

  'Nesta procedure carregaremos apenas as informações
  'onde o cliente está com a data de aniversário válida
  strQuery = "SELECT * FROM MalaExportacao"
  strQuery = strQuery & " WHERE DataIncorreta = FALSE "
  strQuery = strQuery & " AND Nascimento >= #" & Format(m_datInicial, "mm/dd/yyyy") & "#"
  strQuery = strQuery & " AND Nascimento <= #" & Format(m_datFinal, "mm/dd/yyyy") & "#"
  strQuery = strQuery & " ORDER BY Codigo "

  Set rstMalaExportacao = dbTemp.OpenRecordset(strQuery, dbOpenDynaset)
  Set rstMalaDiretaTempo = db.OpenRecordset("Mala Direta - Tempo", dbOpenDynaset)
  
  If rstMalaExportacao.RecordCount = 0 Then
    rstMalaExportacao.Close
    rstMalaDiretaTempo.Close
    Set rstMalaExportacao = Nothing
    Set rstMalaDiretaTempo = Nothing
    
    MsgBox "Não foram encontradas informações dentro da seleção efetuada, verifique.", vbExclamation, "Quick Store"
    Exit Sub
  End If
  
  
  With rstMalaExportacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        rstMalaDiretaTempo.AddNew
        rstMalaDiretaTempo.Fields("Cliente").Value = .Fields("Codigo").Value
        rstMalaDiretaTempo.Fields("Nome").Value = Left(("A/C: " & .Fields("Nome").Value & ""), 30)
        rstMalaDiretaTempo.Update
      
      .MoveNext
      Loop
    
    End If
    .Close
  End With
  
  Set rstMalaExportacao = Nothing
  
  rstMalaDiretaTempo.Close
  Set rstMalaDiretaTempo = Nothing

End Sub

Private Sub Impressao()
  Dim nMarginTop As Integer
  Dim nMarginBottom As Integer
  Dim nMarginRight As Integer
  Dim nMarginLeft As Integer
 
  'Margens
  nMarginTop = 720  '1.27 cm * 567 twips por cm
  nMarginBottom = nMarginTop
  nMarginLeft = 272 '0.48 cm * 567 twips por cm
  nMarginRight = nMarginLeft
  
  nMarginTop = nMarginTop + m_intAdjustTop
  nMarginBottom = nMarginBottom - m_intAdjustTop
    
  nMarginLeft = nMarginLeft + m_intAdjustLeft
  nMarginRight = nMarginRight - m_intAdjustLeft
    
  If nMarginLeft < 0 Then
    nMarginLeft = 0
  End If
  If nMarginRight < 0 Then
    nMarginRight = 0
  End If
  
  With rptRel
    .DataFiles(0) = gsQuickDBFileName
    If optVideo.Value Then
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
    ElseIf optNome.Value Then
      .SortFields(0) = "+{Cli_for.Nome}"
    ElseIf optCidade.Value Then
      .SortFields(0) = "+{Cli_for.Cidade}"
    ElseIf optBairro.Value Then
      .SortFields(0) = "+{Cli_for.Bairro}"
    ElseIf optCEP.Value Then
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

Private Sub sldEsquerda_Change()
  Call sldEsquerda_Click
End Sub

Private Sub sldEsquerda_Click()
  m_intAdjustLeft = nCheckValues(tmEsquerda)
End Sub

Private Sub sldEsquerda_Scroll()
  Call sldEsquerda_Click
End Sub

Private Sub sldSuperior_Change()
  Call sldSuperior_Click
End Sub

Private Sub sldSuperior_Click()
  m_intAdjustTop = nCheckValues(tmSuperior)
End Sub

Private Sub sldSuperior_Scroll()
  Call sldSuperior_Click
End Sub

