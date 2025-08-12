VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCheckStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Produto sem estoque suficiente para movimentação"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheckStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvwStock 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6588
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4140
      Width           =   7365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESTOQUE INSUFICIENTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   7455
   End
End
Attribute VB_Name = "frmCheckStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'09/10/2002 - mpdea
'Adicionado form para exibição dos produtos com estoque insuficiente para
'movimentação

Private Sub cmdClose_Click()
  Unload Me
End Sub

Public Sub ShowStockInsufficient()
  Dim intX As Integer
  Dim itemX As ListItem
  
  For intX = LBound(typCheckStock) To UBound(typCheckStock)
    With typCheckStock(intX)
      If .blnStockInsufficient Then
        Set itemX = lvwStock.ListItems.Add(, , .strCode)
        itemX.SubItems(1) = .dblStock
        'itemX.SubItems(2) = .dblQuantity
        
        ' PILATTI INICIO 2017/07/03
        Dim vAuxI As Integer
        Dim vAuxI2 As Integer
        Dim vAuxPreco As String
        
        vAuxI = InStr(.dblQuantity, ",")
        vAuxI2 = Len(.dblQuantity)
        If vAuxI2 > (vAuxI + 3) Then
          vAuxPreco = Mid(.dblQuantity, 1, vAuxI + 3)
        End If
        ' PILATTI FIM
        itemX.SubItems(2) = vAuxPreco
      End If
    End With
  Next intX
  
  Me.Show vbModal
  
End Sub

Private Sub Form_Load()
  
  With lvwStock
    .View = lvwReport
    .LabelEdit = lvwManual
    With .ColumnHeaders
      .Add , , "Código do Produto", 2700
      .Add , , "Estoque atual", 1750, lvwColumnRight
      .Add , , "Quantidade informada", 1750, lvwColumnRight
    End With
  End With
  
End Sub
