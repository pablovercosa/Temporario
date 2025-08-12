VERSION 5.00
Begin VB.Form frmFundoCombateAPobreza 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   4
      Left            =   6360
      TabIndex        =   15
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   6720
      TabIndex        =   14
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   13
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtChaveRef 
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtChaveRef 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Valor do ICMS Interestadual para a UF do remetente*"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Valor do ICMS Interestadual para a UF de destino"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Valor do ICMS relativo ao Fundo de Combate à Pobreza (FCP) da UF de destino"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Percentual provisório de partilha do ICMS Interestadua*"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Alíquota interestadual das UF envolvidas"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Alíquota interna da UF de destino"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Percentual do ICMS relativo ao Fundo de Combate à Pobreza (FCP) na UF de destino*"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblRef 
      Caption         =   "Valor da BC do ICMS na UF de destino*"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmFundoCombateAPobreza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

