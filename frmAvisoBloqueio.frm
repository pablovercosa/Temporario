VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmAvisoBloqueio 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3735
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5310
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAvisoBloqueio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDispose 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4800
      Top             =   3240
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5055
      Begin Threed.SSPanel sspStatus 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Favor aguardar, sincronizando..."
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.PictureBox picProgress 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   4785
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdRepetir 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Repetir"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   -120
      TabIndex        =   2
      Top             =   -120
      Width           =   5535
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAvisoBloqueio.frx":000C
         ForeColor       =   &H00808080&
         Height          =   855
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Aguarde..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   1350
         Left            =   240
         Picture         =   "frmAvisoBloqueio.frx":0094
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5160
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "frmAvisoBloqueio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_enuResult As VbMsgBoxResult

Public Sub ShowTentativas(ByVal intTentativas As Integer)
  
  If Not g_objAvisoBloqueioDisabledObject Is Nothing Then
    g_objAvisoBloqueioDisabledObject.Enabled = False
  End If
  
  tmrDispose.Enabled = False
  tmrDispose.Enabled = True
  
  Call ShowProgress(intTentativas / 30 * 100)
  
  If Me.Height <> 3000 Then
    Me.Height = 3000
    sspStatus.Caption = "Favor aguardar, sincronizando..."
    cmdRepetir.Visible = False
    cmdCancelar.Visible = False
  End If
  Me.Show
  Me.Refresh
  
End Sub

Public Function ShowRetryCancel() As VbMsgBoxResult
  
  If Not g_objAvisoBloqueioDisabledObject Is Nothing Then
    g_objAvisoBloqueioDisabledObject.Enabled = True
  End If
  
  tmrDispose.Enabled = False
  
  sspStatus.Caption = "Tentativas esgotadas, deseja repetir?"
  picProgress.Cls
  cmdRepetir.Visible = True
  cmdCancelar.Visible = True
  Me.Height = 3765
  
  Me.Hide
  Me.Show vbModal
  
  DoEvents
  
  ShowRetryCancel = m_enuResult

End Function

Private Sub cmdCancelar_Click()
  m_enuResult = vbCancel
  Unload Me
End Sub

Private Sub cmdRepetir_Click()
  m_enuResult = vbRetry
  Me.Hide
End Sub

Private Sub Form_Load()
  'Coloca a tela em modo modal e acima de todas
  Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not g_objAvisoBloqueioDisabledObject Is Nothing Then
    g_objAvisoBloqueioDisabledObject.Enabled = True
  End If
  Set frmAvisoBloqueio = Nothing
End Sub

Private Sub tmrDispose_Timer()
  Unload Me
End Sub

Private Sub ShowProgress(ByVal sngPercent As Single, Optional ByVal strText As String)
  
  picProgress.Line (0, 0)-((sngPercent / 100) * picProgress.Width, picProgress.Height), &H80FF&, BF
  picProgress.Line ((sngPercent / 100) * picProgress.Width, 0)-(picProgress.Width, picProgress.Height), picProgress.BackColor, BF
  
  With picProgress
    .CurrentX = (.ScaleWidth - .TextWidth(strText)) / 2
    .CurrentY = (.ScaleHeight - .TextHeight(strText)) / 2
  End With
  picProgress.Print strText
  picProgress.Refresh

End Sub
