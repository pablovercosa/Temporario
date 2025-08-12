VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmApagaPagas 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Apagar Contas Pagas"
   ClientHeight    =   2355
   ClientLeft      =   1965
   ClientTop       =   1980
   ClientWidth     =   6525
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
   Icon            =   "ApagaPagar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   6525
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
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1755
      Width           =   6270
   End
   Begin Threed.SSPanel Mensagem 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2370
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   450
      _StockProps     =   15
      ForeColor       =   16711680
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin MSMask.MaskEdBox Dia 
      Height          =   345
      Left            =   4920
      TabIndex        =   0
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   630
      Width           =   1440
      _ExtentX        =   2540
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
      Format          =   "dd//mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"ApagaPagar.frx":4E95A
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1110
      Width           =   6315
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apagar contas pagas com data de vencimento até (inclusive) :"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   690
      Width           =   4785
   End
   Begin VB.Label Nome_Filial 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   570
      TabIndex        =   2
      Top             =   120
      Width           =   5805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   435
   End
End
Attribute VB_Name = "frmApagaPagas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsContas_Pagar As Recordset

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
    Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '" & Left("Usu:" & gnUserCode & " - Exclusão de contas pagas", 80) & "', 'CNT_PAG: excluir')", dbFailOnError

   sSql = "Delete * From [Contas a Pagar] Where  Vencimento <= DateValue('" & Dia.Text + "')" & "And [Valor Pago] <> 0"
   sSql = sSql & " And Filial = " & str(gnCodFilial)
   db.Execute sSql
   
   DisplayMsg "Foram apagados " + str(db.RecordsAffected) + " lançamentos."

  End If
  
End Sub

Private Sub Dia_LostFocus()
  Dia.Text = Ajusta_Data(Dia.Text)
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  Nome_Filial.Caption = gnCodFilial & "-" & rsParametros("Nome")
End Sub
