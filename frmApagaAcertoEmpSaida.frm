VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmApagaAcertoEmpSaida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apaga Acerto de Empréstimos de Saída"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "frmApagaAcertoEmpSaida.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   7020
   Begin VB.CheckBox chkConcluido 
      Caption         =   "&Somente acertos concluídos"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Frame Frame6 
      Caption         =   "Período"
      Height          =   795
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   5145
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3675
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   285
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   6
         Top             =   375
         Width           =   885
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton B_Apaga 
      Caption         =   "Apagar"
      Height          =   400
      Left            =   5565
      TabIndex        =   3
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label Nome_Filial 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Filial :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   270
      Width           =   495
   End
End
Attribute VB_Name = "frmApagaAcertoEmpSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Apaga_Click()
  Dim strSQL As String
  
  On Error GoTo ErrHandler
  
  'Verifica Datas
  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Data inicial inválida, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Data final inválida, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual a data final."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  'Solicita senha do Gerente
  If Not frmGerente.gbSenhaGerente Then Exit Sub
  
  DoEvents
  
  strSQL = "DELETE * FROM [Consignação Saída] WHERE Filial = " & gnCodFilial & _
           " AND [Data Operação] BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & _
           "# AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "# AND Concluído = " & _
           IIf(chkConcluido.Value = vbChecked, True, False)
  
  'Executa ação
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde...")
  db.Execute strSQL, dbFailOnError
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  MsgBox "Registros apagados: " & db.RecordsAffected, vbInformation, "Informação"
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Data_Ini_LostFocus()
  Data_Ini.Text = Ajusta_Data(Data_Ini.Text)
End Sub

Private Sub Data_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
  End Select
End Sub

Private Sub Data_Fim_LostFocus()
  Data_Fim.Text = Ajusta_Data(Data_Fim.Text)
End Sub

Private Sub Data_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
  End Select
End Sub

Private Sub Form_Load()
  
  Call StatusMsg("")

  Call CenterForm(Me)
    
  Nome_Filial.Caption = str(gnCodFilial) & " - " & gsGetNameFilial(gnCodFilial)
  
End Sub
