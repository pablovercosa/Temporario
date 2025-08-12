VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProdutosCFOP 
   Caption         =   " Produto x CFOPs (Operações)"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdutosCFOP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5850
      Width           =   11565
   End
   Begin VB.CommandButton cmd_salvar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5340
      Width           =   11565
   End
   Begin VB.TextBox txt_cso 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   11
      Top             =   4845
      Width           =   1305
   End
   Begin VB.TextBox txt_cfop 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   8010
      MaxLength       =   4
      TabIndex        =   10
      Top             =   4845
      Width           =   1305
   End
   Begin VB.ComboBox cmb_operacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4860
      Width           =   7725
   End
   Begin VB.CommandButton cmd_excluir 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Excluir vínculo de Operação/CFOP do produto"
      Height          =   430
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4020
      Width           =   3825
   End
   Begin VB.CommandButton cmd_alterar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Alterar Operação/CFOP vinculada ao produto"
      Height          =   430
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4020
      Width           =   3825
   End
   Begin VB.CommandButton cmd_novo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vincular nova Operação/CFOP ao produto"
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4020
      Width           =   3825
   End
   Begin MSFlexGridLib.MSFlexGrid grade_produtosCFOPs 
      Height          =   2805
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   1
      Cols            =   5
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
   Begin VB.CommandButton cmd_listarOpVinculadas 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listar Operações/CFOPs vinculados"
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   11565
   End
   Begin VB.CommandButton cmd_acharProduto 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   10770
      Picture         =   "frmProdutosCFOP.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox txt_nomeProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   3000
      TabIndex        =   2
      Top             =   60
      Width           =   7725
   End
   Begin VB.TextBox txt_codigoProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   390
      Left            =   720
      TabIndex        =   1
      Top             =   60
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "Operação Saída/Entrada"
      Height          =   285
      Left            =   60
      TabIndex        =   15
      Top             =   4590
      Width           =   1965
   End
   Begin VB.Label lbl_cso 
      Caption         =   "CSO"
      Height          =   195
      Left            =   9720
      TabIndex        =   13
      Top             =   4560
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "CFOP"
      Height          =   225
      Left            =   8010
      TabIndex        =   12
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Código"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "frmProdutosCFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodigoProduto As String
Public nomeProduto As String
Dim i_CodOperacao As Integer
Dim s_Funcao As String

Private Sub cmd_acharProduto_Click()
  cmd_cancelar_Click
  grade_produtosCFOPs.Rows = 1
  nChamaConsulta = 4
  
  Dim obj As Form
  Set obj = New frmPesquisaProduto
  obj.Show
End Sub

Private Sub cmd_alterar_Click()

  If grade_produtosCFOPs.RowSel < 1 Then
    MsgBox "Selecione um registro na grade.", vbInformation
    Exit Sub
  End If
  
  cmd_acharProduto.Enabled = False
  txt_codigoProduto.Enabled = False
  grade_produtosCFOPs.Enabled = False
  cmd_novo.Enabled = False
  cmd_excluir.Enabled = False
  cmb_operacao.Enabled = False
  txt_cfop.Enabled = True
  txt_cso.Enabled = True
  cmd_salvar.Enabled = True

  i_CodOperacao = grade_produtosCFOPs.TextMatrix(grade_produtosCFOPs.RowSel, 1)
  txt_cfop.Text = grade_produtosCFOPs.TextMatrix(grade_produtosCFOPs.RowSel, 3)
  txt_cso.Text = grade_produtosCFOPs.TextMatrix(grade_produtosCFOPs.RowSel, 4)
  
  cmb_operacao.Text = grade_produtosCFOPs.TextMatrix(grade_produtosCFOPs.RowSel, 1) & " - " & grade_produtosCFOPs.TextMatrix(grade_produtosCFOPs.RowSel, 2)
  
  s_Funcao = "ALTERAR"
  txt_cfop.SetFocus
  
End Sub

Private Sub cmd_cancelar_Click()
  cmd_acharProduto.Enabled = True
  txt_codigoProduto.Enabled = True
  grade_produtosCFOPs.Enabled = True
  cmd_novo.Enabled = True
  cmd_alterar.Enabled = True
  cmd_excluir.Enabled = True
  cmb_operacao.Enabled = False
  txt_cfop.Enabled = False
  txt_cso.Enabled = False
  cmd_salvar.Enabled = False

  cmb_operacao.ListIndex = -1
  txt_cfop.Text = ""
  txt_cso.Text = ""
  s_Funcao = ""

End Sub

Private Sub cmd_excluir_Click()

  Dim iCodOperacao As Integer
  Dim lnResponse As Long

  If grade_produtosCFOPs.RowSel < 1 Then
    MsgBox "Selecione um registro na grade.", vbInformation
    Exit Sub
  End If

  lnResponse = MsgBox("Deseja realmente excluir o registro?", vbYesNo, "Atenção")
  If lnResponse = vbNo Then
    Exit Sub
  End If

  iCodOperacao = grade_produtosCFOPs.TextMatrix(grade_produtosCFOPs.RowSel, 1)
  
  db.Execute "Delete from ProdutoCFOP where CodOperacao = " & iCodOperacao & " and CodProduto = '" & LTrim(RTrim(txt_codigoProduto.Text)) & "' "

  MsgBox "Registro desvinculado com sucesso.", vbInformation, "Sucesso"
  
  cmd_listarOpVinculadas_Click
End Sub

Private Sub cmd_listarOpVinculadas_Click()
On Error GoTo Erro
 
  Dim rsCFOP_OpSaida As Recordset
  Dim rsCFOP_OpEntrada As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
 
  If LTrim(RTrim(txt_codigoProduto.Text)) = "" Then
    DisplayMsg "Escolha um produto."
    'txt_codigoProduto.SetFocus
    Exit Sub
  End If

  grade_produtosCFOPs.Rows = 1
  grade_produtosCFOPs.Row = 0
  
  strSQL = "SELECT P.CodProduto, P.CodOperacao, P.CFOP, P.CSO, O.Nome from ProdutoCFOP P, [Operações Saída] O "
  strSQL = strSQL & " where P.CodProduto = '" & LTrim(RTrim(txt_codigoProduto.Text)) & "' "
  strSQL = strSQL & " and P.CodOperacao = O.Código"
  
  Set rsCFOP_OpSaida = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rsCFOP_OpSaida.EOF And rsCFOP_OpSaida.BOF) Then
    rsCFOP_OpSaida.MoveFirst
  End If
  While Not rsCFOP_OpSaida.EOF
  
      grade_produtosCFOPs.AddItem 0 & vbTab & rsCFOP_OpSaida.Fields(1).Value & vbTab & _
                      rsCFOP_OpSaida.Fields(4).Value & vbTab & _
                      rsCFOP_OpSaida.Fields(2).Value & vbTab & _
                      rsCFOP_OpSaida.Fields(3).Value
                      
      rsCFOP_OpSaida.MoveNext
  Wend
  rsCFOP_OpSaida.Close
  Set rsCFOP_OpSaida = Nothing


  strSQL = "SELECT P.CodProduto, P.CodOperacao, P.CFOP, P.CSO, O.Nome from ProdutoCFOP P, [Operações Entrada] O "
  strSQL = strSQL & " where P.CodProduto = '" & LTrim(RTrim(txt_codigoProduto.Text)) & "' "
  strSQL = strSQL & " and P.CodOperacao = O.Código"
  
  Set rsCFOP_OpEntrada = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rsCFOP_OpEntrada.EOF And rsCFOP_OpEntrada.BOF) Then
    rsCFOP_OpEntrada.MoveFirst
  End If
  While Not rsCFOP_OpEntrada.EOF
  
      grade_produtosCFOPs.AddItem 0 & vbTab & rsCFOP_OpEntrada.Fields(1).Value & vbTab & _
                      rsCFOP_OpEntrada.Fields(4).Value & vbTab & _
                      rsCFOP_OpEntrada.Fields(2).Value & vbTab & _
                      rsCFOP_OpEntrada.Fields(3).Value
                      
      rsCFOP_OpEntrada.MoveNext
  Wend
  rsCFOP_OpEntrada.Close
  Set rsCFOP_OpEntrada = Nothing

  grade_produtosCFOPs.RowSel = 0

  Exit Sub
Erro:
  If Not (rsCFOP_OpSaida Is Nothing) Then
      rsCFOP_OpSaida.Close
      Set rsCFOP_OpSaida = Nothing
  End If
  
  If Not (rsCFOP_OpEntrada Is Nothing) Then
      rsCFOP_OpEntrada.Close
      Set rsCFOP_OpEntrada = Nothing
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_novo_Click()

  cmd_acharProduto.Enabled = False
  txt_codigoProduto.Enabled = False
  grade_produtosCFOPs.Enabled = False
  cmd_alterar.Enabled = False
  cmd_excluir.Enabled = False
  cmb_operacao.Enabled = True
  txt_cfop.Enabled = True
  txt_cso.Enabled = True
  cmd_salvar.Enabled = True

  'cmb_operacao.Text = ""
  cmb_operacao.ListIndex = -1
  txt_cfop.Text = ""
  txt_cso.Text = ""

  s_Funcao = "INCLUIR"
  cmb_operacao.SetFocus
End Sub

Private Sub cmd_salvar_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim sDado As String
  Dim sDadoArray() As String

  If s_Funcao = "ALTERAR" Then
    sSql = "Update ProdutoCFOP set CFOP='" & LTrim(RTrim(txt_cfop.Text)) & "', CSO='" & LTrim(RTrim(txt_cso.Text)) & "' "
    sSql = sSql & " WHERE CodProduto='" & LTrim(RTrim(txt_codigoProduto.Text)) & "' and CodOperacao=" & i_CodOperacao
    
    db.Execute sSql
    
    MsgBox "Registro salvo com sucesso", vbInformation, "Sucesso"
    
    cmd_cancelar_Click
    cmd_listarOpVinculadas_Click
  ElseIf s_Funcao = "INCLUIR" Then
  
    If cmb_operacao.ListIndex < 0 Then
      MsgBox "Selecione uma Operação de Saída/Entrada na combo.", vbInformation, "Atenção"
      Exit Sub
    End If
  
    sDado = cmb_operacao.Text
    sDadoArray = Split(sDado, " - ")
  
    sSql = "Insert into ProdutoCFOP (CodProduto, CodOperacao, CFOP, CSO) VALUES ('" & LTrim(RTrim(txt_codigoProduto.Text)) & "',"
    sSql = sSql & sDadoArray(0) & ",'" & LTrim(RTrim(txt_cfop.Text)) & "','" & LTrim(RTrim(txt_cso.Text)) & "') "
    
    db.Execute sSql
    
    If db.RecordsAffected > 0 Then
      MsgBox "Registro salvo com sucesso", vbInformation, "Sucesso"
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

Public Sub Form_Load()
On Error GoTo Erro:

  grade_produtosCFOPs.ColWidth(0) = 10
  grade_produtosCFOPs.ColWidth(1) = 2000
  grade_produtosCFOPs.ColWidth(2) = 5000
  grade_produtosCFOPs.ColWidth(3) = 2000
  grade_produtosCFOPs.ColWidth(4) = 2000
  
  grade_produtosCFOPs.Row = 0
  grade_produtosCFOPs.TextMatrix(0, 1) = "Operação"
  grade_produtosCFOPs.TextMatrix(0, 2) = "Nome Operação"
  grade_produtosCFOPs.TextMatrix(0, 3) = "CFOP"
  grade_produtosCFOPs.TextMatrix(0, 4) = "CSO"
  
  Dim rsCFOP_OpSaida As Recordset
  Dim rsCFOP_OpEntrada As Recordset
  Dim rsParam As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT CodigoRegimeTributario from [Parâmetros Filial] where filial = " & gnCodFilial
  Set rsParam = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If (rsParam.Fields("CodigoRegimeTributario").Value) = 3 Then
    'Empresa regime LUCRO REAL
    lbl_cso.Visible = False
    txt_cso.Visible = False
  Else
    'Empresa regime SIMPLES NACIONAL
    lbl_cso.Visible = True
    txt_cso.Visible = True
  End If
  rsParam.Close
  Set rsParam = Nothing
      
  
  strSQL = "SELECT Código, Nome from [Operações Saída] "
  Set rsCFOP_OpSaida = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rsCFOP_OpSaida.EOF And rsCFOP_OpSaida.BOF) Then
    rsCFOP_OpSaida.MoveFirst
  End If
  While Not rsCFOP_OpSaida.EOF
      cmb_operacao.AddItem rsCFOP_OpSaida.Fields(0).Value & " - " & rsCFOP_OpSaida.Fields(1).Value
      rsCFOP_OpSaida.MoveNext
  Wend
  rsCFOP_OpSaida.Close
  Set rsCFOP_OpSaida = Nothing

  strSQL = "SELECT Código, Nome from [Operações Entrada] "
  Set rsCFOP_OpEntrada = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not (rsCFOP_OpEntrada.EOF And rsCFOP_OpEntrada.BOF) Then
    rsCFOP_OpEntrada.MoveFirst
  End If
  While Not rsCFOP_OpEntrada.EOF
      cmb_operacao.AddItem rsCFOP_OpEntrada.Fields(0).Value & " - " & rsCFOP_OpEntrada.Fields(1).Value
      rsCFOP_OpEntrada.MoveNext
  Wend
  rsCFOP_OpEntrada.Close
  Set rsCFOP_OpEntrada = Nothing
  
  If Not IsNull(CodigoProduto) And CodigoProduto <> "" Then
    txt_codigoProduto.Text = CodigoProduto
    txt_nomeProduto.Text = nomeProduto
  End If
  
  cmd_listarOpVinculadas_Click

  Exit Sub
Erro:
  If Not (rsCFOP_OpSaida Is Nothing) Then
      rsCFOP_OpSaida.Close
      Set rsCFOP_OpSaida = Nothing
  End If
  
  If Not (rsCFOP_OpEntrada Is Nothing) Then
      rsCFOP_OpEntrada.Close
      Set rsCFOP_OpEntrada = Nothing
  End If

  MsgBox "Erro ao realizar carga da tela...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub grade_produtosCFOPs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  grade_produtosCFOPs.Redraw = False
End Sub

Private Sub grade_produtosCFOPs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  grade_produtosCFOPs.RowSel = grade_produtosCFOPs.Row
  grade_produtosCFOPs.Redraw = True
End Sub

Private Sub txt_codigoProduto_LostFocus()
  If Len(LTrim(RTrim(txt_codigoProduto.Text))) > 0 Then
      Dim sSql As String
      Dim rsProduto As Recordset
      sSql = "Select Nome from Produtos where código='" & LTrim(RTrim(txt_codigoProduto.Text)) & "' "
      Set rsProduto = db.OpenRecordset(sSql, dbOpenDynaset)
      If rsProduto.RecordCount > 0 Then
        txt_nomeProduto.Text = rsProduto.Fields(0).Value
      Else
        txt_nomeProduto.Text = ""
        MsgBox "Produto inexistente", vbInformation, "Atenção"
      End If
      rsProduto.Close
      Set rsProduto = Nothing
  Else
      txt_nomeProduto.Text = ""
      grade_produtosCFOPs.Rows = 1
  End If
End Sub
