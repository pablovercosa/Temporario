VERSION 5.00
Begin VB.Form frmReativacaoCentroIndividualmente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reativação de Centro de Custo Individualmente"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReativacaoCentroIndividualmente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3660
   ScaleWidth      =   5475
   Begin VB.CommandButton cmdAtivar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Ativar Selecionados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ListBox lstCCInativo 
      Height          =   2085
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label lblMsg 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Centro(s) de Custo(s) Inativo(s):"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmReativacaoCentroIndividualmente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAtivar_Click()
  Dim intX      As Integer
  Dim blnExiste As Boolean
  Dim strMsg    As String
  
  On Error GoTo TratarErro
  
  'Validar se selecionou alguém
  With lstCCInativo
    For intX = 0 To .ListCount - 1
      If .Selected(intX) Then blnExiste = True
      If blnExiste Then Exit For 'Se existir um já poderemos prosseguir...
    Next intX
  End With
  
  If blnExiste Then
    With lstCCInativo
      For intX = 0 To .ListCount - 1
        If .Selected(intX) Then Call AtivarCentro(lstCCInativo.List(intX) & "")
      Next intX
    End With
  
    Call StatusMsg("")
    
    strMsg = "Ativação concluída com Sucesso !" & vbCrLf & vbCrLf
    strMsg = strMsg & "Observação: Para visualizar corretamente os centros ativados, click no" & vbCrLf
    strMsg = strMsg & "botão 'Atualizar às linhas' do menu principal da tela de cadastro de Centro de Custo."
  
    MsgBox strMsg, vbInformation, "Quick Store"
  Else
    MsgBox "Nenhum ítem foi selecionado, verifique.", vbExclamation, "Quick Store"
  End If
  
  Exit Sub

TratarErro:
  MsgBox "Erro ao <Ativar>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Sub

Private Sub AtivarCentro(ByVal strNome As String)
  Dim rstCentro As Recordset
  Dim strSQL    As String
  Dim strMsg    As String
  
  On Error GoTo TratarErro
  
  strMsg = ""
  strMsg = "Aguarde reativando Centro: " & strNome
  
  Call StatusMsg(strMsg)
  
  strSQL = "UPDATE [Centros de Custo] SET Ativo = TRUE, [Data Alteração] = #" & Format(Data_Atual, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " WHERE Nome = '" & strNome & "'"
  
  db.Execute strSQL
  
  strMsg = "Centro " & strNome & " reativado com sucesso."
  
  Call StatusMsg(strMsg)
  
  Exit Sub

TratarErro:
  Call StatusMsg("")
  MsgBox "Erro em Private <AtivarCentro>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Call BuscarCentroInativo
End Sub

Private Sub BuscarCentroInativo()
  Dim rstCC  As Recordset
  Dim strSQL As String

  On Error GoTo TratarErro
  
  strSQL = "SELECT Nome FROM [Centros de Custo] WHERE NOT Ativo ORDER BY Nome"
  
  Set rstCC = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  If rstCC.RecordCount = 0 Then
    lblMsg.Caption = "<< Não existem Centros desativados !!! >>"
    cmdAtivar.Enabled = False
    rstCC.Close
    Set rstCC = Nothing
    
    Exit Sub
  Else
    lblMsg.Caption = ""
    cmdAtivar.Enabled = True
  End If

  With rstCC
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        lstCCInativo.AddItem (.Fields("Nome").Value & "")
      
       .MoveNext
      Loop
    
    End If
    .Close
  End With
  
  Set rstCC = Nothing

  Exit Sub

TratarErro:
  MsgBox "Erro em Private <BuscarCentroInativo>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Sub
