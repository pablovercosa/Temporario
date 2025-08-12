VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmApagaRecebidas 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Apagar Contas Recebidas"
   ClientHeight    =   2775
   ClientLeft      =   1830
   ClientTop       =   1560
   ClientWidth     =   7035
   ForeColor       =   &H80000008&
   Icon            =   "ApagaReceber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   7035
   Begin VB.CommandButton B_Apaga 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Apagar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   6810
   End
   Begin MSMask.MaskEdBox Dia 
      Height          =   345
      Left            =   5220
      TabIndex        =   0
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   690
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   $"ApagaReceber.frx":4E95A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   75
      TabIndex        =   5
      Top             =   1260
      Width           =   6825
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Apagar contas recebidas com data de vencimento até (inclusive) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Nome_Filial 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   660
      TabIndex        =   3
      Top             =   165
      Width           =   6240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   210
      Width           =   495
   End
End
Attribute VB_Name = "frmApagaRecebidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset

Private Sub B_Apaga_Click()
 Dim Resposta As Integer
 Dim sSql As String

 Call StatusMsg("")

 '11/07/2007 - Anderson
 'Criação de Log e Mensagem de confirmação para o usuário.
 If MsgBox("Tem certeza que deseja executar esta operação?", vbYesNo + vbQuestion, "Atenção") = vbYes Then

   If Not IsDate(Dia.Text) Then
     DisplayMsg "Data inválida, verifique."
     Exit Sub
   End If
  
  '11/07/2007 - Anderson
  'Criação de log para registro de exclusão de registro
  'Efetua registro do Log
  db.Execute "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & _
    Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '" & Left("Usu:" & gnUserCode & " Exc cts recebidas com DtVc até " & Dia.Text, 80) & "', 'CNT_REC: exc-DT man')", dbFailOnError
  
   sSql = "Delete * From [Contas a Receber] Where Tipo = 'R' and [Valor Recebido] <> 0 And Vencimento <= DateValue('" + Dia.Text + "')"
   sSql = sSql + " And Filial = " + str(gnCodFilial)
   db.Execute sSql
    
    '10/09/2007 - Anderson
    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, sSql, "frmApagaRecebidas_B_Apaga_Click", "Contas a Receber", g_strArquivoSystemLog
    End If

   DisplayMsg "Foram apagados " + str(db.RecordsAffected) + " lançamentos."

  End If
  
End Sub

Private Sub Dia_LostFocus()
  Dia.Text = Ajusta_Data(Dia.Text)
End Sub

Private Sub Dia_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Dia.Text = frmCalendario.gsDateCalender(Dia.Text)
  End Select
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  Nome_Filial.Caption = gnCodFilial & "-" & rsParametros("Nome")
End Sub

