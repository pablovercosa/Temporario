VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmApagaCheques 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Apaga Cheques Recebidos"
   ClientHeight    =   2400
   ClientLeft      =   1830
   ClientTop       =   1560
   ClientWidth     =   7230
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
   Icon            =   "ApagaCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2400
   ScaleWidth      =   7230
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
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1830
      Width           =   6945
   End
   Begin MSMask.MaskEdBox Dia 
      Height          =   345
      Left            =   5580
      TabIndex        =   0
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   660
      Width           =   1500
      _ExtentX        =   2646
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
      Caption         =   $"ApagaCheques.frx":4E95A
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   150
      TabIndex        =   5
      Top             =   1170
      Width           =   6975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Apagar cheques recebidos com data de vencimento até (inclusive) :"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   720
      Width           =   5205
   End
   Begin VB.Label Nome_Filial 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   630
      TabIndex        =   3
      Top             =   180
      Width           =   6450
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmApagaCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset

Private Sub B_Apaga_Click()
  Dim sSql As String
  
  
  Call StatusMsg("")
  
  
  If Not IsDate(Dia.Text) Then
    DisplayMsg "Data inválida, verifique."
    Exit Sub
  End If
  
  sSql = "Delete * From [Contas a Receber] Where Tipo = 'C' and (([Valor Recebido] <> 0) Or (Valor = 0 And [Valor Recebido] = 0)) And Vencimento <= DateValue('" + Dia.Text + "')"
  sSql = sSql + " And Filial = " + str(gnCodFilial)
  db.Execute sSql
  
  'Efetua registro do Log
  db.Execute "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & _
    Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '" & Left("Usu:" & gnUserCode & " Exc chqs recebidos com DtVc até " & Dia.Text, 80) & "', 'CNT_REC: exc-DT cheq')", dbFailOnError
  
  
  '10/09/2007 - Anderson
  'Gera arquivo log do sistema
  If g_bolSystemLog Then
    SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, sSql, "frmApagaCheques_B_Apaga_Click", "Contas a Receber", g_strArquivoSystemLog
  End If
  
  If db.RecordsAffected > 0 Then
    DisplayMsg "Foram apagados " + str(db.RecordsAffected) + " lançamentos."
  Else
    DisplayMsg "Nenhum lançamento foi apagado."
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
