VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTabelasDePrecosProdutos 
   Caption         =   " Tabelas de preços do Produto"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTabelasDePrecosProdutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_acatarValor 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Acatar Valor"
      Height          =   465
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2610
      Width           =   5955
   End
   Begin MSFlexGridLib.MSFlexGrid gridTabelaPrecos 
      Height          =   1770
      Left            =   1080
      TabIndex        =   2
      Top             =   780
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3122
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
   Begin VB.Label lbl_nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   5955
   End
   Begin VB.Label lbl_codigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   2085
   End
End
Attribute VB_Name = "frmTabelasDePrecosProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public codigoProduto As String
Public nomeProduto As String
Public valorProdutoAcatado As String

Private Sub cmd_acatarValor_Click()
  If gridTabelaPrecos.RowSel > 0 Then
      valorProdutoAcatado = gridTabelaPrecos.TextMatrix(gridTabelaPrecos.RowSel, 2)
  End If
  Unload Me
End Sub

Private Sub Form_Load()

  lbl_codigo.Caption = codigoProduto
  lbl_nome.Caption = nomeProduto
  
  gridTabelaPrecos.ColWidth(0) = 1
  gridTabelaPrecos.ColWidth(1) = 1700
  gridTabelaPrecos.ColWidth(2) = 1700

  gridTabelaPrecos.Row = 0
  gridTabelaPrecos.TextMatrix(0, 1) = "Tabela de Preço"
  gridTabelaPrecos.TextMatrix(0, 2) = "Valor R$"
  
  If codigoProduto <> "0" And codigoProduto <> "" Then
  
    Dim sSQL As String
    Dim rsTabelas As Recordset
    If Len(codigoProduto) > 0 Then
        sSQL = "Select P.Tabela, P.Preço from Preços P, AcessoTabelasDePrecosProdutos A "
        If Funcionario <> "" Then
            sSQL = sSQL & " where A.Usuario = " & Funcionario
        Else
            sSQL = sSQL & " where A.Usuario = " & gnUserCode
        End If
        
        sSQL = sSQL & " And A.Tabela = P.Tabela "
        sSQL = sSQL & " And P.Produto ='" & codigoProduto & "'"
    
        Set rsTabelas = db.OpenRecordset(sSQL, dbOpenDynaset)
        If rsTabelas.EOF And rsTabelas.BOF Then
            Exit Sub
        End If
        
        rsTabelas.MoveFirst
        While Not rsTabelas.EOF
            gridTabelaPrecos.AddItem vbTab & rsTabelas.Fields("Tabela").Value & vbTab & _
                            FormatNumber(rsTabelas.Fields("Preço").Value, 2) & vbTab
            rsTabelas.MoveNext
        Wend
        rsTabelas.Close
        Set rsTabelas = Nothing
    End If
  End If
  
End Sub

Private Sub gridTabelaPrecos_DblClick()
    cmd_acatarValor_Click
End Sub
