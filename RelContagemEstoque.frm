VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelContagem 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Contagem de Estoque"
   ClientHeight    =   3930
   ClientLeft      =   3645
   ClientTop       =   3330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "RelContagemEstoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3930
   ScaleWidth      =   7215
   Begin VB.CommandButton cmdFechar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Fechar"
      Height          =   435
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3420
      Width           =   7035
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opções"
      Height          =   1275
      Left            =   105
      TabIndex        =   17
      Top             =   960
      Width           =   7035
      Begin VB.CheckBox O_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Separar por classe"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   2
         Top             =   270
         Width           =   2235
      End
      Begin VB.CheckBox O_Inativos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Considerar produtos inativos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   3
         Top             =   585
         Width           =   3120
      End
      Begin VB.CheckBox O_Zero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Não considerar produtos com estoque zero"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   4
         Top             =   915
         Width           =   3765
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   4245
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   585
      Left            =   3660
      TabIndex        =   14
      Top             =   2280
      Width           =   3465
      Begin VB.OptionButton O_nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Im&primir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2940
      Width           =   7035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   585
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   3465
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1500
         TabIndex        =   6
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   210
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4245
      Visible         =   0   'False
      Width           =   1740
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "RelContagemEstoque.frx":4E95A
      DataSource      =   "Data2"
      Height          =   345
      Left            =   660
      TabIndex        =   1
      Top             =   555
      Width           =   885
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
      BackColorOdd    =   16777152
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
      _ExtentX        =   1561
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   4080
      Top             =   4200
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
      Bindings        =   "RelContagemEstoque.frx":4E96E
      DataSource      =   "Data1"
      Height          =   345
      Left            =   660
      TabIndex        =   0
      Top             =   120
      Width           =   885
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
      BackColorOdd    =   16777152
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
      _ExtentX        =   1561
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Nome_Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1590
      TabIndex        =   16
      Top             =   555
      Width           =   5565
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Classe"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   615
      Width           =   465
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   11
      Top             =   180
      Width           =   300
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1590
      TabIndex        =   12
      Top             =   120
      Width           =   5565
   End
End
Attribute VB_Name = "frmRelContagem"
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
 Dim Tamanho As Integer
 Dim Cor As Integer
 Dim sSql As String
 Dim Estoque As Double
 Dim Aux_Data As Variant
 Dim Aux_Classe As Integer
 Dim Aux_Sub As Integer
 Dim Aux_Produto As Double
 Dim Nome_Cla As String
 Dim Nome_Sub As String
 
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

 sSql = "Delete * From Contagem"
 dbTemp.Execute sSql

 Call StatusMsg("")
 

 Rem Le estoque e joga no temporário
 rsProdutos.Index = "Código"
 rsEstoque_Final.Index = "Produto"
 Termina = False
 Produto = ""
 Call StatusMsg("Aguarde, contando estoque.")

 rsClasses.Index = "Código"
 rsSub_Classes.Index = "Código"

LP1S:
  rsProdutos.Seek ">", Produto
  If rsProdutos.NoMatch Then GoTo Imprime
  Produto = rsProdutos("Código")
  
  '14/01/2005 - Daniel
  'Em algumas bases de dados o campo Produtos.Código está
  'aparecendo com caracteres incorretos tais como ...
  '
  'Case: São Francisco Móveis e Eletro. de Olinda - PE
  If Len(Produto) > 20 Then Produto = "0"
  '-------------------------------------------------------
  
  If Produto = "0" Then GoTo LP1S
  
  If Nome_Classe.Caption <> "" Then
    If rsProdutos("Classe") <> Val(Combo_Classe.Text) Then GoTo LP1S
  End If
  
  If rsProdutos("Desativado") = True And O_Inativos.Value = 0 Then GoTo LP1S

  If rsProdutos("Tipo") <> "N" Then GoTo LP1S

  'If rsProdutos("Fracionado") = True Then GoTo LP1S
     
  Estoque = 0
  rsEstoque_Final.Seek "=", Val(Combo.Text), Produto, 0, 0, 0
  If Not rsEstoque_Final.NoMatch Then Estoque = rsEstoque_Final("Estoque Atual")
  
  If O_Zero.Value = 1 Then
    If Estoque = 0 Then GoTo LP1S
  End If
  
  
     
  Call StatusMsg("Aguarde, gravando arquivo temporário, produto " + (Produto))
 
    
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
  
  
  TB2_Contagem.AddNew
     TB2_Contagem("Código") = Produto
     TB2_Contagem("Código Ordenação") = rsProdutos("Código Ordenação")
     TB2_Contagem("Nome") = rsProdutos("Nome")
     TB2_Contagem("Classe") = rsProdutos("Classe")
     TB2_Contagem("Nome Classe") = Nome_Cla
     TB2_Contagem("Sub Classe") = rsProdutos("Sub Classe")
     TB2_Contagem("Nome Sub") = Nome_Sub
     TB2_Contagem("Unidade") = rsProdutos("Unidade Venda")
     TB2_Contagem("Fracionado") = rsProdutos("Fracionado")
     TB2_Contagem("Qtde Estoque") = Estoque
     TB2_Contagem("Empresa") = Val(Combo.Text)
  TB2_Contagem.Update
       
  
  GoTo LP1S

Imprime:

 Call StatusMsg("")

' Rem  Nome do BD
'  With Rel1
'    .DataFiles(0) = gsTempDBFileName
'    .DataFiles(1) = gsQuickDBFileName
'  End With
  
  
  '31/10/2002 - mpdea
  'Corrigido associação com a localização das bases de dados
  With Rel1
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsTempDBFileName
    .DataFiles(2) = gsQuickDBFileName
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
    Str1 = gsReportPath & "CONTAGE1.RPT"
 End If
 If O_Classe.Value = 1 Then
    Str1 = gsReportPath & "CONTAGE2.RPT"
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
 '    Rel1.SortFields(0) = "+{Estoque - Tempo.Classe}"
 '    Rel1.SortFields(1) = "+{Estoque - Tempo.Produto}"
     Rel1.SortFields(0) = "+{Contagem.Classe}"
     Rel1.SortFields(1) = "+{Contagem.Sub Classe}"
     Rel1.SortFields(2) = "+{Contagem.Código Ordenação}"
   End If
   If O_Nome.Value = True Then
     Rel1.SortFields(0) = "+{Contagem.Classe}"
     Rel1.SortFields(1) = "+{Contagem.Sub Classe}"
     Rel1.SortFields(2) = "+{Contagem.Nome}"
   End If
 End If

 If O_Classe.Value = 0 Then
   If O_Código.Value = True Then
     Rel1.SortFields(0) = "+{Contagem.Código Ordenação}"
     Rel1.SortFields(1) = ""
     Rel1.SortFields(2) = ""
   End If
   If O_Nome.Value = True Then
     Rel1.SortFields(0) = "+{Contagem.Nome}"
     Rel1.SortFields(1) = ""
     Rel1.SortFields(2) = ""
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
  
  Set TB2_Contagem = dbTemp.OpenRecordset("Contagem")
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  Combo.Text = gnCodFilial
  Combo_LostFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsProdutos.Close
  rsEstoque_Final.Close
  rsClasses.Close
  rsSub_Classes.Close
  Set rsParametros = Nothing
  Set rsProdutos = Nothing
  Set rsEstoque_Final = Nothing
  Set rsClasses = Nothing
  Set rsSub_Classes = Nothing

End Sub
