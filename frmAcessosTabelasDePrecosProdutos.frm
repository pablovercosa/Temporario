VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAcessosTabelasDePrecosProdutos 
   Caption         =   " Acessos Tabelas De Precos Produtos"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcessosTabelasDePrecosProdutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_listarOpVinculadas 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listar 'tabelas vinculadas' ao usuário"
      Height          =   430
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   4425
   End
   Begin VB.CommandButton cmd_novo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vincular"
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3420
      Width           =   2175
   End
   Begin VB.CommandButton cmd_excluir 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Excluir vínculo"
      Height          =   430
      Left            =   2310
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3420
      Width           =   2175
   End
   Begin VB.ComboBox cmb_tabelas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4230
      Width           =   4425
   End
   Begin VB.CommandButton cmd_salvar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4710
      Width           =   4425
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5220
      Width           =   4425
   End
   Begin MSFlexGridLib.MSFlexGrid grade_tabelasPrecos 
      Height          =   2715
      Left            =   60
      TabIndex        =   5
      Top             =   630
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   4789
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColor       =   12648447
      BackColorSel    =   12582912
      ForeColorSel    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Tabelas de Preços"
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   3960
      Width           =   1965
   End
End
Attribute VB_Name = "frmAcessosTabelasDePrecosProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sUsuario As String
Private s_Funcao As String

Private Sub cmd_cancelar_Click()
  grade_tabelasPrecos.Enabled = True
  cmd_novo.Enabled = True
  cmd_excluir.Enabled = True
  cmb_tabelas.Enabled = False
  cmd_salvar.Enabled = False

  cmb_tabelas.ListIndex = -1
  s_Funcao = ""
End Sub

Private Sub cmd_excluir_Click()

  Dim lnResponse As Long

  If grade_tabelasPrecos.RowSel < 1 Then
    MsgBox "Selecione um registro na grade.", vbInformation
    Exit Sub
  End If

  lnResponse = MsgBox("Deseja realmente excluir o registro?", vbYesNo, "Atenção")
  If lnResponse = vbNo Then
    Exit Sub
  End If

  db.Execute "Delete from AcessoTabelasDePrecosProdutos where Usuario = " & sUsuario & " and Tabela = '" & grade_tabelasPrecos.TextMatrix(grade_tabelasPrecos.RowSel, 1) & "' "

  MsgBox "Registro desvinculado com sucesso.", vbInformation, "Sucesso"
  
  cmd_listarOpVinculadas_Click
End Sub

Private Sub cmd_listarOpVinculadas_Click()
On Error GoTo Erro
 
  Dim rsTab As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
 
  grade_tabelasPrecos.Rows = 1
  grade_tabelasPrecos.Row = 0
  
  strSQL = "SELECT Tabela from AcessoTabelasDePrecosProdutos "
  strSQL = strSQL & " where usuario = " & sUsuario
  
  Set rsTab = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rsTab.EOF And rsTab.BOF) Then
    rsTab.MoveFirst
  End If
  While Not rsTab.EOF
  
      grade_tabelasPrecos.AddItem 0 & vbTab & rsTab.Fields(0).Value
                      
      rsTab.MoveNext
  Wend
  rsTab.Close
  Set rsTab = Nothing
  
  Exit Sub
Erro:
    MsgBox Err.Description, vbInformation, "Erro na rotina botão listar vinculados"
End Sub

Private Sub cmd_novo_Click()

  grade_tabelasPrecos.Enabled = False
  cmd_excluir.Enabled = False
  cmb_tabelas.Enabled = True
  cmd_salvar.Enabled = True
  
  cmb_tabelas.ListIndex = -1

  s_Funcao = "INCLUIR"
  cmb_tabelas.SetFocus
End Sub

Private Sub cmd_salvar_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim sDado As String

  If s_Funcao = "INCLUIR" Then
  
    If cmb_tabelas.ListIndex < 0 Then
      MsgBox "Selecione uma Tabela de Preços na combo.", vbInformation, "Atenção"
      Exit Sub
    End If
  
    sSql = "Insert into AcessoTabelasDePrecosProdutos (usuario, tabela) VALUES (" & sUsuario & ",'"
    sSql = sSql & cmb_tabelas.Text & "') "
    
    db.Execute sSql
    
    If db.RecordsAffected > 0 Then
      'MsgBox "Registro salvo com sucesso", vbInformation, "Sucesso"
    Else
      MsgBox "Registro já existe na base ou algum parâmetro que vc digitou esta inconsistente", vbInformation, "Atenção"
    End If
    cmd_cancelar_Click
    cmd_listarOpVinculadas_Click
  End If

  Exit Sub
Erro:
  MsgBox "Erro ao tentar salvar o registro...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Load()
  grade_tabelasPrecos.ColWidth(0) = 1
  grade_tabelasPrecos.ColWidth(1) = 2700

  grade_tabelasPrecos.Row = 0
  grade_tabelasPrecos.TextMatrix(0, 1) = "Tabela de Preço"
  
  Dim rsTab As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Tabela from [Tabela de Preços] "
  Set rsTab = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rsTab.EOF And rsTab.BOF) Then
    rsTab.MoveFirst
  End If
  While Not rsTab.EOF
      cmb_tabelas.AddItem rsTab.Fields(0).Value
      rsTab.MoveNext
  Wend
  rsTab.Close
  Set rsTab = Nothing
  
  cmd_listarOpVinculadas_Click
End Sub

