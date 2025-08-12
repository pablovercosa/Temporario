VERSION 5.00
Begin VB.Form frmPlanodeContas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plano de Contas Financeiro"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlanodeContas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   10185
   Begin VB.CheckBox chkCentro 
      Caption         =   "OUTRAS"
      Height          =   255
      Index           =   63
      Left            =   6720
      TabIndex        =   65
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "DEPRECIAÇÃO"
      Height          =   255
      Index           =   62
      Left            =   6720
      TabIndex        =   64
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "JUROS E MULTAS"
      Height          =   255
      Index           =   61
      Left            =   6720
      TabIndex        =   63
      Top             =   5040
      Width           =   3255
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "APLICAÇÕES"
      Height          =   255
      Index           =   60
      Left            =   6720
      TabIndex        =   62
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton cmdDesmarcar 
      Caption         =   "D&esmarcar todos"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "CMV - CUSTO DAS MERCADORIAS"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "ADIANTAMENTOS A FORNECEDORES"
      Height          =   255
      Index           =   59
      Left            =   6720
      TabIndex        =   61
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "EMPRÉSTIMOS"
      Height          =   255
      Index           =   58
      Left            =   6720
      TabIndex        =   60
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "PERMUTAS"
      Height          =   255
      Index           =   57
      Left            =   6720
      TabIndex        =   59
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "IPI"
      Height          =   255
      Index           =   56
      Left            =   6720
      TabIndex        =   58
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "CONTRIBUIÇÃO SINDICAL"
      Height          =   255
      Index           =   55
      Left            =   6720
      TabIndex        =   57
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "IRRF"
      Height          =   255
      Index           =   54
      Left            =   6720
      TabIndex        =   56
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "INSS"
      Height          =   255
      Index           =   53
      Left            =   6720
      TabIndex        =   55
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "RETENÇÕES"
      Height          =   255
      Index           =   52
      Left            =   6720
      TabIndex        =   54
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "AQUISIÇÃO DE IMOBILIZADO"
      Height          =   255
      Index           =   51
      Left            =   6720
      TabIndex        =   53
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "INVESTIMENTOS"
      Height          =   255
      Index           =   50
      Left            =   6720
      TabIndex        =   52
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "IMPOSTOS ATRASADOS"
      Height          =   255
      Index           =   49
      Left            =   6720
      TabIndex        =   51
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "CONTRIBUIÇÕES ATRASADAS"
      Height          =   255
      Index           =   48
      Left            =   6720
      TabIndex        =   50
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "DONATIVOS"
      Height          =   255
      Index           =   47
      Left            =   6720
      TabIndex        =   49
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "INFORMÁTICA"
      Height          =   255
      Index           =   46
      Left            =   6720
      TabIndex        =   48
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "SEGUROS"
      Height          =   255
      Index           =   45
      Left            =   6720
      TabIndex        =   47
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "VIAGENS E EXPOSIÇÕES"
      Height          =   255
      Index           =   44
      Left            =   6720
      TabIndex        =   46
      Top             =   960
      Width           =   2175
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "UNIFORMES"
      Height          =   255
      Index           =   43
      Left            =   6720
      TabIndex        =   45
      Top             =   720
      Width           =   2295
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "COMISSARIA"
      Height          =   255
      Index           =   42
      Left            =   3600
      TabIndex        =   44
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "FRETES"
      Height          =   255
      Index           =   41
      Left            =   3600
      TabIndex        =   43
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "DESCONTOS CONCEDIDOS"
      Height          =   255
      Index           =   40
      Left            =   3600
      TabIndex        =   42
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "ASSOCIATIVAS"
      Height          =   255
      Index           =   39
      Left            =   3600
      TabIndex        =   41
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "INDENIZAÇÕES"
      Height          =   255
      Index           =   38
      Left            =   3600
      TabIndex        =   40
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "CARTORÁRIAS E JUDICIAIS"
      Height          =   255
      Index           =   37
      Left            =   3600
      TabIndex        =   39
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "CONSULTIVAS E JURÍDICAS"
      Height          =   255
      Index           =   36
      Left            =   3600
      TabIndex        =   38
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "BANCÁRIAS E FINANCEIRAS"
      Height          =   255
      Index           =   35
      Left            =   3600
      TabIndex        =   37
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "TRANSPORTE"
      Height          =   255
      Index           =   34
      Left            =   3600
      TabIndex        =   36
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "TELEFONES / CORREIOS"
      Height          =   255
      Index           =   33
      Left            =   3600
      TabIndex        =   35
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "EXPEDIENTE E CONSERVAÇÃO"
      Height          =   255
      Index           =   32
      Left            =   3600
      TabIndex        =   34
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "ÁGUA / LUZ / GÁS"
      Height          =   255
      Index           =   31
      Left            =   3600
      TabIndex        =   33
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "ALUGUÉIS / CONDOMÍNIO / IPTU"
      Height          =   255
      Index           =   30
      Left            =   3600
      TabIndex        =   32
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "GASTOS FUNCIONAIS"
      Height          =   255
      Index           =   29
      Left            =   3600
      TabIndex        =   31
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "ALIMENTAÇÃO FUNCIONAL"
      Height          =   255
      Index           =   28
      Left            =   3600
      TabIndex        =   30
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "INSS - CONTRIBUIÇÃO EMPRESA"
      Height          =   255
      Index           =   27
      Left            =   3600
      TabIndex        =   29
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "PRÓ-LABORES"
      Height          =   255
      Index           =   26
      Left            =   3600
      TabIndex        =   28
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "AUTÔNOMOS"
      Height          =   255
      Index           =   25
      Left            =   3600
      TabIndex        =   27
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "COOPERATIVADOS"
      Height          =   255
      Index           =   24
      Left            =   3600
      TabIndex        =   26
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "CURSOS PROFISSIONAIS"
      Height          =   255
      Index           =   23
      Left            =   3600
      TabIndex        =   25
      Top             =   960
      Width           =   2415
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "BOLSA ESTÁGIO"
      Height          =   255
      Index           =   22
      Left            =   3600
      TabIndex        =   24
      Top             =   720
      Width           =   2295
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "FGTS"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   23
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "VERBAS RESCISÓRIAS"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "13º SALÁRIO"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "FÉRIAS"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "SALÁRIOS"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "EMBALAGEM"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "MARKETING INSTITUCIONAL"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "MARKETING PROMOCIONAL"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   3375
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "COMISSÕES S/ SERVIÇOS"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "COMISSÕES S/ ROYALTIES"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   3135
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "COMISSÕES S/ MERCADORIAS"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "SIMPLES"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "IRPJ"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "CSLL"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "ISS"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "ICMS"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "FINSOCIAL"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "PIS"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "COFINS"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkCentro 
      Caption         =   "CSV - CUSTO DOS SERVIÇOS"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdAdicionar 
      BackColor       =   &H0000C0C0&
      Caption         =   "A&dicionar Seleção"
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "&Selecionar todos"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selecione os novos Centros de Custo que você deseja incluir no cadastro de Centro de Custo:"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6765
   End
End
Attribute VB_Name = "frmPlanodeContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdicionar_Click()
  If Not ValidarCampos Then Exit Sub
  Call CriarCentrosCusto
End Sub

Private Sub cmdDesmarcar_Click()
  Dim bytC As Byte
  
  For bytC = 1 To 63
    chkCentro(bytC).Value = vbUnchecked
  Next bytC
End Sub

Private Sub cmdSelecionar_Click()
  Dim bytC As Byte
  
  For bytC = 1 To 63
    chkCentro(bytC).Value = vbChecked
  Next bytC
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

Private Sub CriarCentrosCusto()
  '12/05/2005 - Daniel
  '
  'Esta rotina adicionará novos centros de custo com
  'origem na estrutura: Plano de Contas Financeiro
  '
  'Solicitante..: Carlos (Consultor OSM)
  Dim rstCentro   As Recordset
  Dim rstMaxCodCC As Recordset
  Dim strSQL      As String
  Dim intLastCode As Integer
  Dim bytC        As Byte
  Dim strArray(1 To 63) As String
  Dim bytCont     As Byte
  Dim strTexto    As String
  Dim bytT        As Byte
  
  On Error GoTo TratarErro

  Call StatusMsg("Aguarde inserindo dados em Centro de Custos...")
  Screen.MousePointer = vbHourglass
  
  intLastCode = 0
  bytCont = 0
  
  'Buscar o maior Código cadastrado
  strSQL = "SELECT MAX(Código)AS Maior FROM [Centros de Custo]"

  Set rstMaxCodCC = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstMaxCodCC
    If Not (.BOF And .EOF) Then
      .MoveFirst
      intLastCode = CInt(0 & .Fields("Maior").Value)
    End If
    .Close
  End With

  Set rstMaxCodCC = Nothing
  'Fim da busca
  
  For bytC = 1 To 63
    If chkCentro(bytC).Value = vbChecked Then
      If Not blnExiste(Trim(chkCentro(bytC).Caption)) Then
        intLastCode = intLastCode + 1
        
        Set rstCentro = db.OpenRecordset("Centros de Custo", dbOpenDynaset)
        
        With rstCentro
          .AddNew
            .Fields("Código").Value = intLastCode
            .Fields("Nome").Value = Trim(chkCentro(bytC).Caption) & ""
            .Fields("Data Alteração").Value = Format(Date, "DD/MM/YYYY")
            .Fields("Ativo").Value = True
          .Update
          .Close
        End With
        
        Set rstCentro = Nothing
      Else
        bytCont = bytCont + 1 'Se bytCont igual a 1 ou + já teve alguém que ficou de fora...
        
        strArray(bytCont) = "(" & chkCentro(bytC).Caption & ")" & " "
      End If
    End If
    
  Next bytC
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  If bytCont >= 1 Then
    strTexto = "E não foram adicionados: " & vbCrLf
  
    For bytT = 1 To 63
      If Len(strArray(bytT)) > 0 Then strTexto = strTexto & strArray(bytT)
    Next bytT
    
    MsgBox strTexto, vbExclamation, "Os seguintes Centros já existem:"
  End If
  
  Exit Sub
  
TratarErro:
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Function blnExiste(ByVal strNome As String) As Boolean
  Dim rstCentro As Recordset
  Dim strSQL    As String
  
  blnExiste = False
  
  strSQL = "SELECT Código, Nome FROM [Centros de Custo] WHERE Nome = '" & strNome & "'"
  
  Set rstCentro = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  If rstCentro.RecordCount <> 0 Then blnExiste = True
  
  rstCentro.Close
  Set rstCentro = Nothing
  
End Function

Private Function ValidarCampos() As Boolean
  Dim bytC As Byte
  
  For bytC = 1 To 63
    If chkCentro(bytC).Value = vbChecked Then ValidarCampos = True
    If ValidarCampos Then Exit For
  Next bytC
  
  If Not ValidarCampos Then MsgBox "Nenhum campo marcado, verifique.", vbExclamation, "Quick Store"

End Function
