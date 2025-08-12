VERSION 5.00
Begin VB.Form frmImprimeEtiquetaImagem 
   Caption         =   " Imagem"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImprimeEtiquetaImagem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9165
      Left            =   -30
      ScaleHeight     =   9135
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   0
      Width           =   6765
   End
End
Attribute VB_Name = "frmImprimeEtiquetaImagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sImagem As String
Public lLargura As Long
Public lAltura As Long
Public sTituloPagina As String

Private Sub Form_Load()
On Error GoTo Erro

  Me.ScaleMode = 3
  Me.Width = lLargura + 2
  Me.Height = lAltura + 2
  
  If Not IsNull(sTituloPagina) And sTituloPagina <> "" Then
      Me.Caption = " " & sTituloPagina
  End If

  Picture1.Width = lLargura
  Picture1.Height = lAltura

  On Error Resume Next
  Picture1.Picture = LoadPicture(App.Path & sImagem)
  
  Exit Sub
Erro:
  MsgBox "Erro na carga da tela de imagem " & Err.Number & " " & Err.Description, vbInformation, "Atençao"
End Sub
