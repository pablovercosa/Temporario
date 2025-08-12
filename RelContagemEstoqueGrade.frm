VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelContagemGrade 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Contagem de Estoque - Grade"
   ClientHeight    =   3435
   ClientLeft      =   3135
   ClientTop       =   3000
   ClientWidth     =   5955
   ForeColor       =   &H80000008&
   Icon            =   "RelContagemEstoqueGrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3435
   ScaleWidth      =   5955
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   400
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opções"
      Height          =   1320
      Left            =   105
      TabIndex        =   17
      Top             =   1080
      Width           =   5745
      Begin VB.CheckBox O_Classe 
         Caption         =   "Separar por classe"
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   255
         Width           =   1935
      End
      Begin VB.CheckBox O_Inativos 
         Caption         =   "Considerar produtos inativos"
         Height          =   255
         Left            =   165
         TabIndex        =   3
         Top             =   585
         Width           =   3015
      End
      Begin VB.CheckBox O_Zero 
         Caption         =   "Não considerar produtos com estoque zero"
         Height          =   225
         Left            =   165
         TabIndex        =   4
         Top             =   945
         Width           =   3690
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2070
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   4125
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   855
      Left            =   1560
      TabIndex        =   14
      Top             =   2490
      Width           =   1215
      Begin VB.OptionButton O_nome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H0000C0C0&
      Caption         =   "Im&primir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   150
      TabIndex        =   13
      Top             =   2490
      Width           =   1335
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   75
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4140
      Visible         =   0   'False
      Width           =   1740
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "RelContagemEstoqueGrade.frx":058A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   630
      Width           =   735
      DataFieldList   =   "Nome"
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
      Columns(0).Width=   7858
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1455
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   4200
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelContagemEstoqueGrade.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   165
      Width           =   735
      DataFieldList   =   "Nome"
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
      Columns(0).Width=   8467
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1561
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Classe 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1920
      TabIndex        =   16
      Top             =   660
      Width           =   3900
   End
   Begin VB.Label Label2 
      Caption         =   "Classe :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   675
      Width           =   660
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   11
      Top             =   225
      Width           =   585
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Top             =   165
      Width           =   3900
   End
End
Attribute VB_Name = "frmRelContagemGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsProdutos As Recordset
Dim rsEstoque_Final As Recordset
Dim TB2_Contagem As Recordset
Dim rsClasses As Recordset
Dim rsSub_Classes As Recordset
Dim rsCores As Recordset
Dim rsTamanhos As Recordset
Dim rsGrade As Recordset

Private Sub B_Imprime_Click()
 Dim Termina As Integer
 Dim Val2 As Integer
 Dim Erro As Integer
 Dim Str1 As String
 Dim Str2 As String
 Dim Str3 As String
 Dim Str_Data1 As String
 Dim Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 Dim Produto As String
 Dim Aux_Produto As String
 Dim Completo As String
 Dim Tamanho As Integer
 Dim Cor As Integer
 Dim Edição As Long
 Dim Tipo As Integer
 Dim sSql As String
 Dim Estoque As Double
 Dim Aux_Data As Variant
 Dim Aux_Classe As Integer
 Dim Aux_Sub As Integer
 Dim Nome_Cla As String
 Dim Nome_Sub As String
 Dim Nome_Cor As String
 Dim Nome_Tam As String
 
 
 Call StatusMsg("")

 Rem Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a filial."
   Combo.SetFocus
   Exit Sub
 End If


 If Filial_Liberada <> 0 Then
   If Val(Combo.Text) <> Filial_Liberada Then
     DisplayMsg "Funcionário não tem acesso a esta filial."
     Exit Sub
   End If
 End If


 Rem apaga os valores existente no arquivo


 
 Rem apaga pesquisa anterior desta filial do arquivo temporario
 Call StatusMsg("Aguarde, preparando arquivo temporário ...")

 sSql = "Delete * From [Contagem Grade]"
 dbTemp.Execute sSql

 Call StatusMsg("")
 

 Rem Le estoque e joga no temporário
 rsProdutos.Index = "Código"
 rsEstoque_Final.Index = "Produto"
 Termina = False
 Produto = ""
 Call StatusMsg("Aguarde, contando estoque...")

 rsClasses.Index = "Código"
 rsSub_Classes.Index = "Código"
 rsGrade.Index = "Original"
 rsCores.Index = "Código"
 rsTamanhos.Index = "Código"
 
LP1S:
  rsProdutos.Seek ">", Produto
  If rsProdutos.NoMatch Then GoTo Imprime
  Produto = rsProdutos("Código")
  If Produto = "0" Then GoTo LP1S
  
  If Nome_Classe.Caption <> "" Then
    If rsProdutos("Classe") <> Val(Combo_Classe.Text) Then GoTo LP1S
  End If
  
  If rsProdutos("Desativado") = True And O_Inativos.Value = 0 Then GoTo LP1S

  If rsProdutos("Tipo") <> "G" Then GoTo LP1S

  If rsProdutos("Fracionado") = True Then GoTo LP1S


  Rem    Tem um produto com grade disponível
  Rem    Agora deve achar todas as cores e tamanhos possíveis
  Rem    e seus respectivos estoques
  Completo = ""
LP2:
  rsGrade.Seek ">", Produto, Completo
  If rsGrade.NoMatch Then GoTo LP1S
  If rsGrade("Código Original") <> Produto Then GoTo LP1S
  
  Completo = rsGrade("Código")
  
  Acha_Produto Completo, Aux_Produto, Tamanho, Cor, Edição, Tipo, Erro
  
  If Erro <> 0 Then GoTo LP2
  
  Estoque = Acha_Estoque(Val(Combo.Text), Produto, Tamanho, Cor, 0, Erro)
  If Erro > 1 Then GoTo LP2
  
  If Estoque = 0 Then
   If O_Zero.Value = 1 Then GoTo LP2
  End If
  
  
  Call StatusMsg("Aguarde, gravando arquivo temporário, produto " & (Produto))
     
     
  rsClasses.Seek "=", rsProdutos("Classe")
  If rsClasses.NoMatch Then
     Nome_Cla = "Classe não cadastrada"
  Else
     Nome_Cla = rsClasses("Nome")
  End If
  
  rsSub_Classes.Seek "=", rsProdutos("Sub Classe")
  If rsSub_Classes.NoMatch Then
    Nome_Sub = "Subclasse não cadastrada"
  Else
    Nome_Sub = rsSub_Classes("Nome")
  End If
  
  rsCores.Seek "=", Cor
  If rsCores.NoMatch Then
    Nome_Cor = "Cor não cadastrada"
  Else
    Nome_Cor = rsCores("Nome")
  End If
  
  rsTamanhos.Seek "=", Tamanho
  If rsTamanhos.NoMatch Then
    Nome_Tam = "Tamanho não cadastrado"
  Else
    Nome_Tam = rsTamanhos("Nome")
  End If
  
  
  
  TB2_Contagem.AddNew
     TB2_Contagem("Código") = Produto
     TB2_Contagem("Código Ordenação") = rsProdutos("Código Ordenação")
     TB2_Contagem("Nome") = rsProdutos("Nome")
     TB2_Contagem("Classe") = rsProdutos("Classe")
     TB2_Contagem("Nome Classe") = Nome_Cla
     TB2_Contagem("Sub Classe") = rsProdutos("Sub Classe")
     TB2_Contagem("Nome Sub") = Nome_Sub
     TB2_Contagem("Unidade") = rsProdutos("Unidade Venda")
     TB2_Contagem("Qtde Estoque") = Estoque
     TB2_Contagem("Empresa") = Val(Combo.Text)
     TB2_Contagem("Cor") = Cor
     TB2_Contagem("Nome Cor") = Nome_Cor
     TB2_Contagem("Tamanho") = Tamanho
     TB2_Contagem("Nome Tamanho") = Nome_Tam
  TB2_Contagem.Update
  
  GoTo LP2

Imprime:
 Call StatusMsg("")

 Rem  Nome do BD
  With Rel1
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
  End With

 Rem Saída
 If B_Vídeo = True Then Rel1.Destination = 0
 If B_Impressora = True Then Rel1.Destination = 1
 Rem If B_Arquivo = True Then
 Rem    frmMenu.Relatório.Destination = 2
 Rem    frmMenu.Relatório.PrintFileName = T_Arquivo.Text
 Rem End If

 Rem Nome do arquivo .rpt
 If O_Classe.Value = 0 Then
    Str1 = gsReportPath & "CONTAG1G.RPT"
 End If
 If O_Classe.Value = 1 Then
    Str1 = gsReportPath & "CONTAG2G.RPT"
 End If
 Rel1.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 Rem Seleção
 Str_Rel = ""
' If O_Inativos.Value = 0 Then
'  Str_Rel = "{Produtos.Desativado} = False"
' End If
 Rel1.SelectionFormula = Str_Rel
 
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 
 Rel1.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 Rel1.Formulas(1) = Str_Rel
 
 
 If O_Classe.Value = 1 Then
   If O_Código.Value = True Then
     Rel1.SortFields(0) = "+{Contagem Grade.Classe}"
     Rel1.SortFields(1) = "+{Contagem Grade.Sub Classe}"
     Rel1.SortFields(2) = "+{Contagem Grade.Código Ordenação}"
     Rel1.SortFields(3) = "+{Contagem Grade.Tamanho}"
     Rel1.SortFields(4) = "+{Contagem Grade.Cor}"
   End If
   If O_Nome.Value = True Then
     Rel1.SortFields(0) = "+{Contagem Grade.Classe}"
     Rel1.SortFields(1) = "+{Contagem Grade.Sub Classe}"
     Rel1.SortFields(2) = "+{Contagem Grade.Nome}"
     Rel1.SortFields(3) = "+{Contagem Grade.Tamanho}"
     Rel1.SortFields(4) = "+{Contagem Grade.Cor}"
   End If
 End If

 If O_Classe.Value = 0 Then
   If O_Código.Value = True Then
     Rel1.SortFields(0) = "+{Contagem Grade.Código Ordenação}"
     Rel1.SortFields(1) = "+{Contagem Grade.Tamanho}"
     Rel1.SortFields(2) = "+{Contagem Grade.Cor}"
     Rel1.SortFields(3) = ""
     Rel1.SortFields(4) = ""
   End If
   If O_Nome.Value = True Then
     Rel1.SortFields(0) = "+{Contagem Grade.Nome}"
     Rel1.SortFields(1) = "+{Contagem Grade.Tamanho}"
     Rel1.SortFields(2) = "+{Contagem Grade.Cor}"
     Rel1.SortFields(3) = ""
     Rel1.SortFields(4) = ""
   End If
 End If
 
 

 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  

 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

 Exit Sub


End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub Combo_Classe_CloseUp()
Combo_Classe.Text = Combo_Classe.Columns(1).Text
Combo_Classe_LostFocus

End Sub

Private Sub Combo_Classe_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Classe_LostFocus()
  Call StatusMsg("")
  Nome_Classe.Caption = ""
  If IsNull(Combo_Classe.Text) Then Exit Sub
  If Combo_Classe.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Classe.Text) Then Exit Sub
  If Val(Combo_Classe.Text) < 0 Then Exit Sub
  If Val(Combo_Classe.Text) > 9999 Then Exit Sub

  rsClasses.Index = "Código"
  rsClasses.Seek "=", Val(Combo_Classe.Text)
  If rsClasses.NoMatch Then Exit Sub
  Nome_Classe.Caption = rsClasses("Nome")

End Sub

Private Sub Combo_CloseUp()
Combo.Text = Combo.Columns(1).Text
Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
  Nome_Empresa.Caption = ""
  If IsNull(Combo.Text) Then Exit Sub
  If Combo.Text = "" Then Exit Sub
  If Not IsNumeric(Combo.Text) Then Exit Sub
  If Val(Combo.Text) < 0 Then Exit Sub
  If Val(Combo.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
 Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
 Set rsEstoque_Final = db.OpenRecordset("Estoque Final", , dbReadOnly)
 Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
 Set rsSub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)
 Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
 Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
 
 Set TB2_Contagem = dbTemp.OpenRecordset("Contagem Grade")

 Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)

 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName

 Combo.Text = gnCodFilial

End Sub
