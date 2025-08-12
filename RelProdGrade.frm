VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelProdGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Produtos com Grade"
   ClientHeight    =   5295
   ClientLeft      =   1875
   ClientTop       =   1680
   ClientWidth     =   7620
   Icon            =   "RelProdGrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   7620
   Begin VB.Frame Frame2 
      Caption         =   "Período"
      Height          =   855
      Left            =   2400
      TabIndex        =   19
      Top             =   3480
      Width           =   5055
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "Data Final :"
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Data Inicial :"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   2190
      Visible         =   0   'False
      Width           =   2115
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelProdGrade.frx":058A
      DataSource      =   "Data2"
      Height          =   285
      Left            =   1005
      TabIndex        =   1
      Top             =   2295
      Width           =   855
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8705
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1746
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.CheckBox O_Zero 
      Caption         =   "Não mostrar produtos com estoque/vendas igual a 0 (zero)"
      Height          =   225
      Left            =   225
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4740
      Width           =   4635
   End
   Begin VB.CheckBox O_Classe 
      Caption         =   "Relatório separado por classe / subclasse"
      Height          =   225
      Left            =   225
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4425
      Visible         =   0   'False
      Width           =   3585
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   7080
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   6120
      TabIndex        =   7
      Top             =   4455
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2685
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   2715
      Visible         =   0   'False
      Width           =   1695
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "RelProdGrade.frx":059E
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1005
      TabIndex        =   2
      Top             =   2820
      Width           =   855
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8837
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1984
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
      Begin VB.OptionButton O_Estoque 
         Caption         =   "Estoque"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   480
         Width           =   1065
      End
      Begin VB.OptionButton O_Vendas 
         Caption         =   "Vendas"
         Height          =   225
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.ListBox Lista 
      Height          =   1815
      Left            =   210
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   315
      Width           =   3165
   End
   Begin VB.Label Nome_Filial 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2055
      TabIndex        =   18
      Top             =   2295
      Width           =   4200
   End
   Begin VB.Label Label8 
      Caption         =   "Filial :"
      Height          =   225
      Left            =   270
      TabIndex        =   17
      Top             =   2295
      Width           =   645
   End
   Begin VB.Label Label6 
      Caption         =   "Use 0 para todas as classes."
      Height          =   225
      Left            =   1005
      TabIndex        =   16
      Top             =   3135
      Width           =   3795
   End
   Begin VB.Label Nome_Classe 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2055
      TabIndex        =   14
      Top             =   2820
      Width           =   4200
   End
   Begin VB.Label Label3 
      Caption         =   "Classe :"
      Height          =   225
      Left            =   270
      TabIndex        =   13
      Top             =   2820
      Width           =   750
   End
   Begin VB.Label Selecionados 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5490
      TabIndex        =   11
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "Tamanhos Selecionados :"
      Height          =   225
      Left            =   3465
      TabIndex        =   10
      Top             =   1890
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione até 15 tamanhos a serem usados no relatório. Tamanhos não selecionados serão colocados na coluna ""Outros""."
      Height          =   645
      Left            =   3570
      TabIndex        =   9
      Top             =   315
      Width           =   3795
   End
End
Attribute VB_Name = "frmRelProdGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tab_Tamanhos(15) As Integer

Dim rsTamanhos As Recordset
Dim rsProdutos As Recordset
Dim rsEstoque As Recordset
Dim rsTempo As Recordset
Dim rsEstoque_Final As Recordset
Dim rsGrade As Recordset
Dim rsParametros As Recordset
Dim rsClasses As Recordset
Dim rsCores As Recordset




Sub Rel_Estoque()
  Dim Tamanhos As Integer
  Dim sSql As String
  Dim i As Integer
  Dim J As Integer
  Dim Str_Aux As String
  Dim Aux_Produto As String
  Dim Aux_Data As Date
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Estoque As Double
  Dim Estoque_Lng As Long
  Dim Cod_Completo As String
  Dim Erro As Integer
  
  Tamanhos = 0
  Erase Tab_Tamanhos


  Call StatusMsg("Aguarde, preparando arquivo ...")
  sSql = "Delete * From [Rel Grade]"
  dbTemp.Execute sSql
  Call StatusMsg("")
  
 
 
  For i = 0 To (Lista.ListCount - 1)
    If Lista.Selected(i) = True Then
      Tamanhos = Tamanhos + 1
      Str_Aux = Lista.List(i)
      Str_Aux = Left(Str_Aux, 3)
      J = Val(Str_Aux)
      Tab_Tamanhos(Tamanhos) = J
    End If
  Next i
 
   
'  Aux_Data = CDate(Data_Ini.Text)
  Aux_Produto = 0
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edição = 0
  rsTempo.Index = "Código"
  rsProdutos.Index = "Código"
  rsGrade.Index = "Original"
  
  Aux_Produto = "0"
Lp1:
  rsProdutos.Seek ">", Aux_Produto
  If rsProdutos.NoMatch Then GoTo Fim_Lp1

  Aux_Produto = rsProdutos("Código")
  If rsProdutos("Tipo") <> "G" Then GoTo Lp1
  If Nome_Classe.Caption <> "" Then
    If rsProdutos("Classe") <> Val(Combo_Classe.Text) Then GoTo Lp1
  End If
  
  Call StatusMsg("Verificando produto " + rsProdutos("Nome"))
  DoEvents
  
  
  
  Cod_Completo = ""
LP2:
  rsGrade.Seek ">", Aux_Produto, Cod_Completo
  If rsGrade.NoMatch Then GoTo Lp1
  If rsGrade("Código Original") <> Aux_Produto Then GoTo Lp1
  
  Cod_Completo = rsGrade("Código")
  
  Str_Aux = Right(rsGrade("Código"), 6)
  Aux_Tamanho = Left(Str_Aux, 3)
  Aux_Cor = Right(Str_Aux, 3)
  
  
  Estoque = Acha_Estoque(Combo_Filial.Text, Aux_Produto, Aux_Tamanho, Aux_Cor, 0, Erro)


  'Verificar se é da classe
  rsTempo.Seek "=", Aux_Produto, Aux_Cor
  If rsTempo.NoMatch Then
    rsTempo.AddNew
      rsTempo("Código") = Aux_Produto
      rsTempo("Cód Cor") = Aux_Cor
      rsTempo("Nome") = rsProdutos("Nome")
      rsCores.Seek "=", Aux_Cor
      If Not rsCores.NoMatch Then rsTempo("Nome Cor") = rsCores("Nome")
      
  Else
    rsTempo.Edit
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(1) Then
    rsTempo("Tamanho1") = rsTempo("Tamanho1") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(2) Then
    rsTempo("Tamanho2") = rsTempo("Tamanho2") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(3) Then
    rsTempo("Tamanho3") = rsTempo("Tamanho3") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(4) Then
    rsTempo("Tamanho4") = rsTempo("Tamanho4") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(5) Then
    rsTempo("Tamanho5") = rsTempo("Tamanho5") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(6) Then
    rsTempo("Tamanho6") = rsTempo("Tamanho6") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(7) Then
    rsTempo("Tamanho7") = rsTempo("Tamanho7") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(8) Then
    rsTempo("Tamanho8") = rsTempo("Tamanho8") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(9) Then
    rsTempo("Tamanho9") = rsTempo("Tamanho9") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(10) Then
    rsTempo("Tamanho10") = rsTempo("Tamanho10") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(11) Then
    rsTempo("Tamanho11") = rsTempo("Tamanho11") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(12) Then
    rsTempo("Tamanho12") = rsTempo("Tamanho12") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(13) Then
    rsTempo("Tamanho13") = rsTempo("Tamanho13") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(14) Then
    rsTempo("Tamanho14") = rsTempo("Tamanho14") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(15) Then
    rsTempo("Tamanho15") = rsTempo("Tamanho15") + Estoque
    GoTo Fim_Tamanhos
  End If
  
  rsTempo("Outros") = rsTempo("Outros") + Estoque
  
  
Fim_Tamanhos:
  rsTempo.Update
  GoTo LP2
  
Fim_Lp1:
  

End Sub

Sub Rel_Vendas()
  Dim Tamanhos As Integer
  Dim sSql As String
  Dim i As Integer
  Dim J As Integer
  Dim Str_Aux As String
  Dim Aux_Produto As String
  Dim Aux_Data As Date
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long

  Tamanhos = 0
  Erase Tab_Tamanhos


  Call StatusMsg("Aguarde, preparando arquivo ...")
  sSql = "Delete * From [Rel Grade]"
  dbTemp.Execute sSql
  Call StatusMsg("")
  
 
 
  For i = 0 To (Lista.ListCount - 1)
    If Lista.Selected(i) = True Then
      Tamanhos = Tamanhos + 1
      Str_Aux = Lista.List(i)
      Str_Aux = Left(Str_Aux, 3)
      J = Val(Str_Aux)
      Tab_Tamanhos(Tamanhos) = J
    End If
  Next i
 
   
  rsEstoque.Index = "Produto"
  Aux_Data = CDate(Data_Ini.Text)
  Aux_Produto = 0
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edição = 0
  rsTempo.Index = "Código"
  rsProdutos.Index = "Código"
  
  
    
  Call StatusMsg("Verificando produto " + rsProdutos("Nome"))
  DoEvents

  
Lp1:
  rsEstoque.Seek ">", gnCodFilial, Aux_Data, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição
  If rsEstoque.NoMatch Then GoTo Fim_Lp1
  If rsEstoque("Filial") <> gnCodFilial Then GoTo Fim_Lp1
  If rsEstoque("Data") > CDate(Data_Fim.Text) Then GoTo Fim_Lp1
  
  Aux_Data = rsEstoque("Data")
  Aux_Produto = rsEstoque("Produto")
  Aux_Tamanho = rsEstoque("Tamanho")
  Aux_Cor = rsEstoque("Cor")
  Aux_Edição = rsEstoque("Edição")
  
  rsProdutos.Seek "=", Aux_Produto
  If rsProdutos.NoMatch Then GoTo Lp1
  
  If rsProdutos("Tipo") <> "G" Then GoTo Lp1

  If Nome_Classe.Caption <> "" Then
    If rsProdutos("Classe") <> Val(Combo_Classe.Text) Then GoTo Lp1
  End If
  

  'Verificar se é da classe
  rsTempo.Seek "=", Aux_Produto, Aux_Cor
  If rsTempo.NoMatch Then
    rsTempo.AddNew
      rsTempo("Código") = Aux_Produto
      rsTempo("Cód Cor") = Aux_Cor
      rsTempo("Nome") = rsProdutos("Nome")
      rsCores.Seek "=", Aux_Cor
      If Not rsCores.NoMatch Then rsTempo("Nome Cor") = rsCores("Nome")
      
  Else
    rsTempo.Edit
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(1) Then
    rsTempo("Tamanho1") = rsTempo("Tamanho1") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(2) Then
    rsTempo("Tamanho2") = rsTempo("Tamanho2") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(3) Then
    rsTempo("Tamanho3") = rsTempo("Tamanho3") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(4) Then
    rsTempo("Tamanho4") = rsTempo("Tamanho4") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(5) Then
    rsTempo("Tamanho5") = rsTempo("Tamanho5") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(6) Then
    rsTempo("Tamanho6") = rsTempo("Tamanho6") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(7) Then
    rsTempo("Tamanho7") = rsTempo("Tamanho7") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(8) Then
    rsTempo("Tamanho8") = rsTempo("Tamanho8") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(9) Then
    rsTempo("Tamanho9") = rsTempo("Tamanho9") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(10) Then
    rsTempo("Tamanho10") = rsTempo("Tamanho10") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(11) Then
    rsTempo("Tamanho11") = rsTempo("Tamanho11") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(12) Then
    rsTempo("Tamanho12") = rsTempo("Tamanho12") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(13) Then
    rsTempo("Tamanho13") = rsTempo("Tamanho13") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(14) Then
    rsTempo("Tamanho14") = rsTempo("Tamanho14") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  If Aux_Tamanho = Tab_Tamanhos(15) Then
    rsTempo("Tamanho15") = rsTempo("Tamanho15") + rsEstoque("Vendas")
    GoTo Fim_Tamanhos
  End If
  
  rsTempo("Outros") = rsTempo("Outros") + rsEstoque("Vendas")
  
  
Fim_Tamanhos:
  rsTempo.Update
  GoTo Lp1
  
Fim_Lp1:
  
End Sub

Private Sub B_Imprime_Click()
  Dim Str1 As String
  Dim Str_Rel As String
  Dim Valor_Gravar As Long
  Dim sSql As String
  Dim i As Integer
  
  If Nome_Filial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    Combo_Filial.SetFocus
    Exit Sub
  End If
  
  If Val(Selecionados.Caption) = 0 Then
    DisplayMsg "Escolha os tamanhos a serem usados antes."
    Exit Sub
  End If
 
  If Val(Selecionados.Caption) > 15 Then
    DisplayMsg "Foram escolhidos mais de 15 tamanhos. Retire alguns e tente novamente."
    Exit Sub
  End If
 
  If O_Vendas.Value = True Then
    If Not IsDate(Data_Ini.Text) Then
      DisplayMsg "Data inicial inválida."
      Data_Ini.SetFocus
      Exit Sub
    End If
    If Not IsDate(Data_Fim.Text) Then
      DisplayMsg "Data final inválida."
      Data_Fim.SetFocus
      Exit Sub
    End If
  
    If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
      DisplayMsg "Data inicial deve ser inferior à data final."
      Data_Ini.SetFocus
      Exit Sub
    End If
  End If
  
  
  Call StatusMsg("Aguarde ...")
  DoEvents
  
  If O_Vendas.Value = True Then Rel_Vendas
  If O_Estoque.Value = True Then Rel_Estoque
 

  If O_Zero.Value = 1 Then
    Call StatusMsg("Aguarde, apagando produtos com estoque / vendas igual a 0.")
    sSql = "Delete * From [Rel Grade] Where Tamanho1 = 0 AND Tamanho2 = 0 AND Tamanho3 = 0"
    sSql = sSql + " AND Tamanho4 = 0 AND Tamanho5 = 0 AND Tamanho6 = 0 AND Tamanho7 = 0"
    sSql = sSql + " AND Tamanho8 = 0 AND Tamanho9 = 0 AND Tamanho10 = 0 AND Tamanho11 = 0"
    sSql = sSql + " AND Tamanho12 = 0 AND Tamanho13 = 0 AND Tamanho14 = 0 AND Tamanho15 = 0 AND Outros = 0"
    dbTemp.Execute sSql
    Call StatusMsg("")
  End If
  
  
  
  
  
 Rem  Nome do BD
  With Rel1
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
  End With

 Rem Saída
 'If B_Vídeo = True Then Rel1.Destination = 0
 'If B_Impressora = True Then Rel1.Destination = 1
 Rel1.Destination = 0
 
 Rem Nome do arquivo .rpt
 Str1 = gsReportPath & "ESTOQ_GG.RPT"
 
 Rel1.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 For i = 1 To 15
  Str_Rel = "Tam" + Trim(str(i))
  Str_Rel = Str_Rel + " = '" + Trim(str(Tab_Tamanhos(i))) + "'"
  Rel1.Formulas((i - 1)) = Str_Rel
 Next i

 If O_Vendas.Value = True Then Rel1.Formulas(15) = "Titulo = 'Relatório de Vendas - Produtos com Grade'"
 If O_Estoque.Value = True Then Rel1.Formulas(15) = "Titulo = 'Relatório de Estoque - Produtos com Grade'"

 Rem Seleção
' Str_Rel = ""
' If O_Inativos.Value = 0 Then
'  Str_Rel = "{Produtos.Desativado} = False"
' End If
' Rel1.SelectionFormula = Str_Rel
 
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 
' Rel1.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
' Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
' Rel1.Formulas(1) = Str_Rel
 
 
 
 
 
 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  

 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault
  
  



End Sub

Private Sub Combo_Classe_CloseUp()
 
 Combo_Classe.Text = Combo_Classe.Columns(1).Text
 Combo_Classe_LostFocus

End Sub

Private Sub Combo_Classe_InitColumnProps()

' Combo_Classe.Text = Combo_Classe.Columns(1).Text
' Combo_Classe_LostFocus

End Sub

Private Sub Combo_Classe_LostFocus()

  Nome_Classe.Caption = ""
  If IsNull(Combo_Classe.Text) Then Exit Sub
  If Combo_Classe.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Classe.Text) Then Exit Sub
  If Val(Combo_Classe.Text) < 1 Then Exit Sub
  If Val(Combo_Classe.Text) > 9999 Then Exit Sub
  
  rsClasses.Index = "Código"
  rsClasses.Seek "=", Val(Combo_Classe.Text)
  If rsClasses.NoMatch Then Exit Sub
  
  Nome_Classe.Caption = rsClasses("Nome") & ""
  
  
  
  
End Sub


Private Sub Combo_Filial_CloseUp()

 Combo_Filial.Text = Combo_Filial.Columns(1).Text
 Combo_Filial_LostFocus
 
End Sub

Private Sub Combo_Filial_LostFocus()

 Nome_Filial.Caption = ""
 
 If IsNull(Combo_Filial.Text) Then Exit Sub
 If Combo_Filial.Text = "" Then Exit Sub
 If Not IsNumeric(Combo_Filial.Text) Then Exit Sub
 If Val(Combo_Filial.Text) < 1 Or Val(Combo_Filial.Text) > 99 Then Exit Sub
 
 rsParametros.Index = "Filial"
 rsParametros.Seek "=", Val(Combo_Filial.Text)
 If rsParametros.NoMatch Then Exit Sub
 
 Nome_Filial.Caption = rsParametros("Nome") & ""
 
End Sub

Private Sub Data_Ini_LostFocus()
  Data_Ini.Text = Ajusta_Data(Data_Ini.Text)
End Sub

Private Sub Data_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
  End Select
End Sub

Private Sub Data_Fim_LostFocus()
  Data_Fim.Text = Ajusta_Data(Data_Fim.Text)
End Sub

Private Sub Data_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
  End Select
End Sub


Private Sub Data_Ini1_Change()

End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim Aux_Str As String
  
  Call CenterForm(Me)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName

  Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)
  
  Set rsTempo = dbTemp.OpenRecordset("Rel Grade")
  
  Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  
  Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)

  rsCores.Index = "Código"

  i = 0
  rsTamanhos.Index = "Código"
Lp1:
  rsTamanhos.Seek ">", i
  If rsTamanhos.NoMatch Then GoTo Fim_Tamanhos
  i = rsTamanhos("Código")
  Aux_Str = Format(rsTamanhos("Código"), "000")
  Aux_Str = Aux_Str + " - " + rsTamanhos("Nome")
  Lista.AddItem Aux_Str
  GoTo Lp1
  
Fim_Tamanhos:

End Sub

Private Sub Lista_Click()
  Dim i As Integer
  Dim J As Integer
  
  
  For i = 0 To (Lista.ListCount - 1)
   If Lista.Selected(i) = True Then
     J = J + 1
   End If
  Next i
  
  Selecionados.Caption = J

End Sub

Private Sub O_Estoque_Click()
 Data_Ini.Enabled = False
 Data_Fim.Enabled = False
 
End Sub

Private Sub O_Vendas_Click()
 Data_Ini.Enabled = True
 Data_Fim.Enabled = True
End Sub

