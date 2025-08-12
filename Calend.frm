VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolha uma Data"
   ClientHeight    =   3225
   ClientLeft      =   6210
   ClientTop       =   3990
   ClientWidth     =   3165
   Icon            =   "Calend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3225
   ScaleWidth      =   3165
   Begin MSComCtl2.MonthView mtvCalender 
      Height          =   2460
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   4339
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777088
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   12640511
      ScrollRate      =   1
      StartOfWeek     =   69599233
      TitleBackColor  =   16753444
      TitleForeColor  =   -2147483639
      TrailingForeColor=   8421504
      CurrentDate     =   36172
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsData As String
Dim gbCancel As Boolean

Public Function gsDateCalender(ByVal sDataOrigem As String) As String
  On Error Resume Next
  mtvCalender.Value = Format(Now, "dd/mm/yyyy")
  mtvCalender.Value = sDataOrigem
  On Error GoTo 0
  
  Call AdjustCalender(1, 1)
  gsData = sDataOrigem
  Me.Show vbModal
  gsDateCalender = gsFormatDate(gsData)
  Set frmCalendario = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    gbCancel = True
    Unload Me
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc(vbCrLf) Then
    KeyAscii = 0
    gsData = mtvCalender.Value
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  '28/01/2009 - mpdea
  'Coloca a tela em modo modal e acima de todas
  Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub mtvCalender_DateClick(ByVal DateClicked As Date)
  gsData = DateClicked
  Unload Me
End Sub

Private Sub mtvCalender_DblClick()
  If mtvCalender.MonthColumns = 1 Then
    Call AdjustCalender(2, 2)
  Else
    Call AdjustCalender(1, 1)
  End If
End Sub

Private Sub AdjustCalender(ByVal nCol As Integer, ByVal nRow As Integer)
  With mtvCalender
    .MonthColumns = nCol
    .MonthRows = nRow
    Me.Width = .Width + 90
    Me.Height = .Height + 475
  End With
  Me.Refresh
End Sub
