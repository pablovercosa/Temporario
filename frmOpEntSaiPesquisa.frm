VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOpEntSaiPesquisa 
   Caption         =   " Pesquisa de Operações"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   Icon            =   "frmOpEntSaiPesquisa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10170
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid gridOperacoes 
      Height          =   6390
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   11271
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483648
      BackColorFixed  =   12632256
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483648
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
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
   End
End
Attribute VB_Name = "frmOpEntSaiPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iOrigemOperacao As Integer   ' 1 = Origem Tela Operações Saída
                                    ' 2 = Origem Tela Operações Entrada

Private Sub Form_Load()
On Error GoTo Erro
  Dim rsOper As Recordset

  gridOperacoes.ColWidth(0) = 0
  gridOperacoes.ColWidth(1) = 1500
  gridOperacoes.ColWidth(2) = 5000

  gridOperacoes.Row = 0
  gridOperacoes.TextMatrix(0, 1) = "Código"
  gridOperacoes.TextMatrix(0, 2) = "Nome da operação"
  
  If iOrigemOperacao = 1 Then
      ' Carregar todas as Operações de Saída
  Else
      ' Carregar todas as Operações de Entrada
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro na abertura da Tela de pesquisa de Operações " & Err.Description & " - " & Err.Number, vbInformation, "Atenção"
  
End Sub

Private Sub gridOperacoes_Click()
On Error GoTo Erro

  Dim sCodigoOperacao As String
  Dim sNomeOperacao As String

  If gridOperacoes.RowSel > 0 Then
      sCodigoOperacao = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 1)
      sNomeOperacao = gridOperacoes.TextMatrix(gridOperacoes.RowSel, 2)
      
      If sCodigoOperacao = "" Then
        MsgBox "Selecione um registro na grade!", vbInformation, "Atenção"
        Exit Sub
      End If
  End If

  Exit Sub
Erro:
  MsgBox "Erro na seleção da Operação " & Err.Description & " - " & Err.Number, vbInformation, "Atenção"
End Sub
