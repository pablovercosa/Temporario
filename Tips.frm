VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Dicas Úteis"
   ClientHeight    =   4335
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tips.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Apresentar o Dicas Úteis no início"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   2655
      Value           =   1  'Checked
      Width           =   3540
   End
   Begin VB.CommandButton cmdNextTip 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Próxima"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3735
      Width           =   6795
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   75
      Picture         =   "Tips.frx":4E95A
      ScaleHeight     =   2505
      ScaleWidth      =   6765
      TabIndex        =   3
      Top             =   75
      Width           =   6795
      Begin VB.Label Label1 
         Caption         =   "Que tal isto..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   555
         TabIndex        =   5
         Top             =   165
         Width           =   1485
      End
      Begin VB.Label lblTipText 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   6435
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3150
      Width           =   6795
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

'30/01/2009 - mpdea
'Adaptado para o novo menu
'Key: Q7MENU
Private Sub chkLoadTipsAtStartup_Click()
  ' save whether or not this form should be displayed at startup
  SaveSetting "QuickStore", "Options", "Show Tips", chkLoadTipsAtStartup.Value
'  If chkLoadTipsAtStartup.Value = 1 Then
'    frmMain.ActiveBar1.Tools("miTips").Checked = True
'  Else
'    frmMain.ActiveBar1.Tools("miTips").Checked = False
'  End If
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  ' Set the checkbox, this will force the value to be written back out to the registry
  Me.chkLoadTipsAtStartup.Value = vbChecked

  ' Seed Rnd
  Randomize

  If gnNumConvenio = 31 Then
'    Me.chkLoadTipsAtStartup.Value = 0
    Me.chkLoadTipsAtStartup.Visible = True
  Else
    Me.chkLoadTipsAtStartup.Value = 1
    Me.chkLoadTipsAtStartup.Visible = False
  End If

  ' Read in the tips file and display a tip at random.
  Call LoadTips(gsTipFile)
    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
