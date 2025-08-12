VERSION 5.00
Begin VB.Form frm_manterAtivo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   495
   ClientLeft      =   195
   ClientTop       =   105
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Solução QuickStore...stand by"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3405
   End
End
Attribute VB_Name = "frm_manterAtivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Activate()

  DoEvents
  Sleep 1500
  DoEvents
  
  Unload Me
End Sub

