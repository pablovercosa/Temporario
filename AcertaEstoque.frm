VERSION 5.00
Begin VB.Form frmAcertaEstoque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerta Quantidades de Estoque"
   ClientHeight    =   3390
   ClientLeft      =   1920
   ClientTop       =   2280
   ClientWidth     =   8205
   Icon            =   "AcertaEstoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "AcertaEstoque.frx":058A
   ScaleHeight     =   3390
   ScaleWidth      =   8205
   Begin VB.CommandButton B_Acerta 
      Caption         =   "&Acertar"
      Height          =   400
      Left            =   6705
      TabIndex        =   4
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Acertar Estoque "
      Height          =   1455
      Left            =   105
      TabIndex        =   1
      Top             =   825
      Width           =   7950
      Begin VB.OptionButton O_Todos 
         Caption         =   "&Todos os produtos que tem diferença."
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   6795
      End
      Begin VB.OptionButton O_Consertar 
         Caption         =   "&Somente dos produtos que tiveram o campo ""Consertar"" marcado na tela ""Informação da Contagem""."
         Height          =   405
         Left            =   105
         TabIndex        =   2
         Top             =   375
         Value           =   -1  'True
         Width           =   7575
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"AcertaEstoque.frx":0825
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7905
   End
End
Attribute VB_Name = "frmAcertaEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsContagem As Recordset
Dim rsProdutos As Recordset
Dim rsEstoque  As Recordset

Private Sub B_Acerta_Click()
  Dim Resposta       As Integer
  Dim Código         As String
  Dim Tamanho        As Integer
  Dim Cor            As Integer
  Dim Conta          As Long
  Dim Criar_Registro As Integer
  Dim Estoque_Final  As Single
  Dim Mes_Atual      As Integer
  Dim Ano_Atual      As Integer
  
  Call StatusMsg("")
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Este processo não poderá ser desfeito, deseja prosseguir?"
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    DisplayMsg "Estoque não foi atualizado."
    Exit Sub
  End If
  
  On Error GoTo ErrTrans
  
  Screen.MousePointer = vbHourglass
  
  Call ws.BeginTrans
  
  Código = ""
  Tamanho = 0
  Cor = 0
  Conta = 0
  rsProdutos.Index = "Código"
  rsContagem.Index = "Código"

Lp1:
  If gbAcertaGrade = True Then
    rsContagem.Seek ">", Código, Tamanho, Cor
  Else
    rsContagem.Seek ">", Código
  End If
  
  If rsContagem.NoMatch Then GoTo Fim_Lp
  Código = rsContagem("Código")
  
  If gbAcertaGrade = True Then
    Tamanho = rsContagem("Tamanho")
    Cor = rsContagem("Cor")
  End If
  
  'Verifica se a filial de origem é a mesma que está logado
  If rsContagem("Empresa") <> gnCodFilial Then GoTo Lp1
  
  If rsContagem("Diferença") = 0 Then GoTo Lp1
  
  If O_Consertar.Value = True Then
    If rsContagem("Consertar") = False Then GoTo Lp1
  End If
  
  rsProdutos.Seek "=", rsContagem("Código")
  If rsProdutos.NoMatch Then GoTo Lp1
  
  Conta = Conta + 1
  
  Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
  
  Rem Acha Último Estoque deste produto
  Criar_Registro = False
  Estoque_Final = 0
  rsEstoque.Index = "Produto"
  rsEstoque.Seek "=", rsContagem("Empresa"), Data_Atual, rsContagem("Código"), Tamanho, Cor, 0
  
  If Not rsEstoque.NoMatch Then
    Estoque_Final = rsEstoque("Estoque Final")
  End If
  
  If rsEstoque.NoMatch Then
    
    rsEstoque.Index = "Data"
    rsEstoque.Seek "<", rsContagem("Empresa"), rsContagem("Código"), Tamanho, Cor, 0, Data_Atual
    If rsEstoque.NoMatch Then Criar_Registro = True
    If Not rsEstoque.NoMatch Then
      If rsEstoque("Filial") = rsContagem("Empresa") And rsEstoque("Produto") = rsContagem("Código") And rsEstoque("Tamanho") = 0 And rsEstoque("Cor") = 0 And rsEstoque("Edição") = 0 Then
        Criar_Registro = True
        Estoque_Final = rsEstoque("Estoque Final")
      End If
    End If
  
    rsEstoque.AddNew
    rsEstoque("Filial") = rsContagem("Empresa")
    rsEstoque("Data") = Data_Atual
    rsEstoque("Produto") = rsContagem("Código")
    rsEstoque("Tamanho") = Tamanho
    rsEstoque("Cor") = Cor
    rsEstoque("Edição") = 0
    rsEstoque("Classe") = rsProdutos("Classe")
    rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
    rsEstoque("Estoque Anterior") = Estoque_Final
    rsEstoque.Update
    
    rsEstoque.Index = "Produto"
    rsEstoque.Seek "=", rsContagem("Empresa"), Data_Atual, rsContagem("Código"), Tamanho, Cor, 0
  
  End If
  
  'Verifica se a real diferença está correta
  If rsContagem("Qtde Estoque") <> Estoque_Final Then
    With rsContagem
      .Edit
      .Fields("Qtde Estoque") = Estoque_Final
      .Fields("Diferença") = .Fields("Digitado") - Estoque_Final
      .Update
    End With
    If gbAcertaGrade = True Then
      rsContagem.Seek "=", Código, Tamanho, Cor
    Else
      rsContagem.Seek "=", Código
    End If
  End If
  
  Rem neste ponto esta com o registro de estoque
  Rem no buffer, agora soma com os valores da movimentação
  rsEstoque.Edit
  If rsContagem("Diferença") < 0 Then
    rsEstoque("Ajuste Saída") = rsEstoque("Ajuste Saída") + Abs(rsContagem("Diferença"))
  End If
  
  If rsContagem("Diferença") > 0 Then
    rsEstoque("Ajuste Entra") = rsEstoque("Ajuste Entra") + Abs(rsContagem("Diferença"))
  End If
  
  Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
  Estoque_Final = Estoque_Final - rsEstoque("Transf Saída") + rsEstoque("Transf Entra")
  Estoque_Final = Estoque_Final - rsEstoque("Ajuste Saída") + rsEstoque("Ajuste Entra")
  Estoque_Final = Estoque_Final - rsEstoque("Grátis Saída") + rsEstoque("Grátis Entra")
  Estoque_Final = Estoque_Final - rsEstoque("Empre Saída") + rsEstoque("Empre Entra")
  Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolução")
  
  If rsProdutos("Estoque") = False Then
    Estoque_Final = 0
  End If
  
  rsEstoque("Estoque Final") = Estoque_Final
  rsEstoque.Update
  
  If gbAcertaGrade Then
    Call Grava_Estoque_Final(rsContagem("Empresa"), rsProdutos("Código"), Tamanho, Cor, 0, Estoque_Final, CDate(Data_Atual))
  Else
    Call Grava_Estoque_Final(rsContagem("Empresa"), rsProdutos("Código"), 0, 0, 0, Estoque_Final, CDate(Data_Atual))
  End If
  
  rsContagem.Edit
  rsContagem("Diferença") = 0
  rsContagem("Qtde Estoque") = rsContagem("Digitado")
  rsContagem("Consertar") = False
  rsContagem.Update
  
  GoTo Lp1
  
Fim_Lp:

  '---[ Gera Log do usuário ]---'
      g_GravaLog Data_Atual, "Acerto de Estoque, DQ(" & Data_Atual & "), DW(" & Date & "),Funcionário: " & _
                            gnUserCode & " - " & gsUserName, "ACERTO ESTOQUE"
  '---[ Gera Log do usuário ]---'
  
  Call ws.CommitTrans
  Screen.MousePointer = vbDefault
  DisplayMsg "Fim de processo. Registros atualizados : " + str(Conta)
  Exit Sub
  
ErrTrans:
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Acertar Estoque."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly & vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  On Error Resume Next
  Call ws.Rollback
  Exit Sub

End Sub

Private Sub Form_Load()
  Dim sSql As String
  Dim sCaption As String
  
  Call CenterForm(Me)
  If gbAcertaGrade Then
    sSql = "Contagem Grade"
    sCaption = "(Produtos com Grade)"
  Else
    sSql = "Contagem"
    sCaption = ""
  End If
  Me.Caption = "Acerta Estoque " & sCaption
  Set rsContagem = dbTemp.OpenRecordset(sSql)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsContagem.Close
  rsProdutos.Close
  rsEstoque.Close
  Set rsContagem = Nothing
  Set rsProdutos = Nothing
  Set rsEstoque = Nothing
End Sub


