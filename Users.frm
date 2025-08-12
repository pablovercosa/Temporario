VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estações Atualmente Conectadas"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "Users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   6345
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5760
      Top             =   3360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   -120
      TabIndex        =   8
      Top             =   -120
      Width           =   6495
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"Users.frx":058A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   2040
         TabIndex        =   10
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Estações Conectadas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   360
         Picture         =   "Users.frx":061B
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refazer Lista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3960
      TabIndex        =   0
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3960
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtCtUsers 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3975
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2565
      Width           =   2265
   End
   Begin VB.TextBox txtMaxUsers 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2265
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Estações:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Número de Conexões Ativas:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3945
      TabIndex        =   5
      Top             =   2295
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Licenças Atual:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3930
      TabIndex        =   3
      Top             =   1410
      Width           =   1905
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdRefresh_Click()
  Dim nI As Integer
  Dim sUser As String
  Screen.MousePointer = vbHourglass
  gsCurrentUsers = gsGetMDBUsers(gsQuickTMPFileName)
  gnCtCurrentUsers = 0
  lstUsers.Clear
  For nI = LBound(gsCurrentUsers) To UBound(gsCurrentUsers)
    DoEvents
    sUser = gsCurrentUsers(nI)
    If Len(Trim(sUser)) = 0 Then
      Exit For
    End If
    lstUsers.AddItem Format(nI, "#00") & " - " & sUser
    gnCtCurrentUsers = gnCtCurrentUsers + 1
  Next nI
  lstUsers.Refresh
  txtMaxUsers.Text = CStr(gnMaxUsers)
  txtCtUsers.Text = CStr(gnCtCurrentUsers)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

Private Sub lstUsers_GotFocus()
  cmdOK.SetFocus
End Sub

Private Sub Timer1_Timer()
  cmdRefresh_Click
End Sub

Private Sub txtCtUsers_GotFocus()
  cmdOK.SetFocus
End Sub

Private Sub txtMaxUsers_GotFocus()
  cmdOK.SetFocus
End Sub
