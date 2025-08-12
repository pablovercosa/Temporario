VERSION 5.00
Begin VB.Form frmAgenda 
   Appearance      =   0  'Flat
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   0  'None
   Caption         =   "Sua Agenda Quick Store"
   ClientHeight    =   6645
   ClientLeft      =   1755
   ClientTop       =   1755
   ClientWidth     =   11175
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "WeblySleek UI Semibold"
      Size            =   8.25
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00E5E5E5&
   HelpContextID   =   1800
   Icon            =   "Agenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6645
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      FillColor       =   &H00FFFFFF&
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
      Height          =   6540
      Left            =   0
      ScaleHeight     =   6510
      ScaleWidth      =   11145
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmd_tips 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Habilita Tips"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9030
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   780
         Width           =   1950
      End
      Begin VB.ListBox lstPend 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F7F7&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   4440
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   1305
         Width           =   10890
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5895
         Width           =   10890
      End
      Begin VB.Label Semana 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   615
         Left            =   4590
         TabIndex        =   5
         Top             =   90
         Width           =   3750
      End
      Begin VB.Label Dia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   615
         Left            =   225
         TabIndex        =   4
         Top             =   90
         Width           =   3750
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Painel de Informações"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   495
         Left            =   4245
         TabIndex        =   3
         Top             =   900
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_tips_Click()
On Error GoTo Erro

    SaveSetting "QuickStore", "Options", "Show Tips", 1
    Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Dim i As Integer
 
  Call CenterForm(Me)
  
 Dia.Caption = Format(Data_Atual, "dd/mm/yyyy")
 Semana.Caption = WeekdayName(Weekday(Dia.Caption, vbUseSystemDayOfWeek), False, vbUseSystemDayOfWeek)
 
' i = Weekday(Data_Atual)
' If i = 1 Then Semana.Caption = "Domingo"
' If i = 2 Then Semana.Caption = "Segunda"
' If i = 3 Then Semana.Caption = "Terça"
' If i = 4 Then Semana.Caption = "Quarta"
' If i = 5 Then Semana.Caption = "Quinta"
' If i = 6 Then Semana.Caption = "Sexta"
' If i = 7 Then Semana.Caption = "Sábado"
 
End Sub

