VERSION 5.00
Begin VB.Form frmAquisicaoEstrategicoRel 
   Caption         =   "Solução QuickStore"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAquisicaoEstrategicoRel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAquisicaoEstrategicoRel.frx":4E95A
   ScaleHeight     =   8370
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Aquisições de módulos extras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   2430
      Width           =   4020
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAquisicaoEstrategicoRel.frx":6FE52
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1935
      TabIndex        =   7
      Top             =   3105
      Width           =   9060
   End
   Begin VB.Line Line1 
      X1              =   4950
      X2              =   7335
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Line Line2 
      X1              =   3375
      X2              =   6255
      Y1              =   5625
      Y2              =   5625
   End
   Begin VB.Line Line3 
      X1              =   3375
      X2              =   5670
      Y1              =   6075
      Y2              =   6075
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Infopar A3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5265
      TabIndex        =   6
      Top             =   7425
      Width           =   1635
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAquisicaoEstrategicoRel.frx":6FF3E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1125
      TabIndex        =   5
      Top             =   6435
      Width           =   9870
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "* Posição dos seus principais clientes por grandeza de itens com valor gasto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1935
      TabIndex        =   4
      Top             =   5805
      Width           =   9060
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "* Posição dos seus principais fornecedores por grandeza de itens com valor pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1935
      TabIndex        =   3
      Top             =   5355
      Width           =   9060
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "* Posição em tempo real dos produtos mais vendidos por grandeza de itens vendidos com valor faturado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1935
      TabIndex        =   2
      Top             =   4680
      Width           =   9060
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Conteúdo do módulo estratégico:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1125
      TabIndex        =   1
      Top             =   4320
      Width           =   10185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Conteúdo do módulo NFe XML:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1125
      TabIndex        =   0
      Top             =   2745
      Width           =   3570
   End
End
Attribute VB_Name = "frmAquisicaoEstrategicoRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
