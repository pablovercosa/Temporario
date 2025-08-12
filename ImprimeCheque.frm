VERSION 5.00
Begin VB.Form frmImprimeCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Cheques"
   ClientHeight    =   2640
   ClientLeft      =   720
   ClientTop       =   1515
   ClientWidth     =   7710
   Icon            =   "ImprimeCheque.frx":0000
   LinkTopic       =   "Impressão de Cheques"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      Caption         =   "Usar a data :"
      Height          =   675
      Left            =   135
      TabIndex        =   13
      Top             =   1845
      Width           =   3450
      Begin VB.OptionButton Data_Atual 
         Caption         =   "atual"
         Height          =   255
         Left            =   1575
         TabIndex        =   15
         Top             =   285
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton Data_Cheque 
         Caption         =   "do cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.CommandButton B_Próximo 
      Caption         =   "Próximo >>"
      Height          =   480
      Left            =   6315
      TabIndex        =   12
      Top             =   1470
      Width           =   1340
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   480
      Left            =   6315
      TabIndex        =   11
      Top             =   2070
      Width           =   1340
   End
   Begin VB.Label Cheque 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Data 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5040
      TabIndex        =   9
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Valor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3720
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Num_Cheque 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Banco 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   1335
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Data"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Valor"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Número"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Banco"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cheque"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"ImprimeCheque.frx":058A
      Height          =   675
      Left            =   165
      TabIndex        =   0
      Top             =   90
      Width           =   7440
   End
End
Attribute VB_Name = "frmImprimeCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Posição As Integer

Private Sub B_Imprime_Click()
  Dim Valor_Num As Double
  Dim Dia_Str As String
  
  If Data_Atual.Value = True Then
    Dia_Str = Date
  Else
    Dia_Str = Data.Caption
  End If
  
  Valor_Num = Retorna_Valor(Valor.Caption)
  If Valor_Num <> 0 Then
    Call Imprime_Cheque("()", Banco.Caption, Dia_Str, Valor_Num)
    DisplayMsg "Cheque impresso."
  End If
  
End Sub


Private Sub B_Próximo_Click()
 Posição = Posição + 1
 If Posição >= 50 Then
   DisplayMsg "Final dos cheques."
   Exit Sub
 End If

 Valor.Caption = Format(frmRecebimento.Exporta_Valor(Posição), "###,###,##0.00")
 Banco.Caption = frmRecebimento.Exporta_Banco(Posição)
 Num_Cheque.Caption = frmRecebimento.Exporta_Cheque(Posição)
 Data.Caption = frmRecebimento.Exporta_Data(Posição)
 
 Cheque.Caption = (Posição + 1)
 
 Call StatusMsg("")
 If Retorna_Valor(Valor.Caption) = 0 Then
   DisplayMsg "Este cheque não pode ser impresso, não existe valor."
 End If
 
End Sub

Private Sub Form_Load()
 
  Call CenterForm(Me)
  
  Posição = 0
  
  Valor.Caption = Format(frmRecebimento.Exporta_Valor(0), "###,###,##0.00")
  Banco.Caption = frmRecebimento.Exporta_Banco(0)
  Num_Cheque.Caption = frmRecebimento.Exporta_Cheque(0)
  Data.Caption = frmRecebimento.Exporta_Data(0)
  
  Cheque.Caption = (Posição + 1)
  
  If Retorna_Valor(Valor.Caption) = 0 Then
    DisplayMsg "Este cheque não pode ser impresso, não existe valor."
  End If
 
End Sub


