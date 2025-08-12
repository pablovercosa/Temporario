VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEstoque2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Estoque Geral"
   ClientHeight    =   3540
   ClientLeft      =   1485
   ClientTop       =   1530
   ClientWidth     =   8775
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
   Icon            =   "RelEstoqueTodos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   8775
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      Height          =   720
      Left            =   4755
      TabIndex        =   16
      Top             =   2070
      Width           =   3915
      Begin VB.OptionButton O_Edição 
         Appearance      =   0  'Flat
         Caption         =   "Com Edição"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2550
         TabIndex        =   19
         Top             =   300
         Width           =   1245
      End
      Begin VB.OptionButton O_Grade 
         Appearance      =   0  'Flat
         Caption         =   "Com Grade"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1200
         TabIndex        =   18
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton O_Normal 
         Appearance      =   0  'Flat
         Caption         =   "Normal"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Value           =   -1  'True
         Width           =   900
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
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   630
      Visible         =   0   'False
      Width           =   2115
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "RelEstoqueTodos.frx":4E95A
      DataSource      =   "Data2"
      Height          =   345
      Left            =   1125
      TabIndex        =   1
      ToolTipText     =   "Use 0 para todos"
      Top             =   540
      Width           =   1200
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
      Columns(1).Width=   1879
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2117
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   720
      Left            =   2520
      TabIndex        =   11
      Top             =   2070
      Width           =   2205
      Begin VB.OptionButton O_nome 
         Appearance      =   0  'Flat
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1260
         TabIndex        =   13
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CheckBox O_Inativos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Imprimir também os produtos inativos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   4
      Top             =   1725
      Width           =   3915
   End
   Begin VB.CheckBox O_Estoque 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Não imprimir produtos com estoque igual a 0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   1395
      Width           =   4515
   End
   Begin VB.CheckBox O_Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Separado por Classe/Subclasse"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   1065
      Width           =   3090
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
      Height          =   465
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2970
      Width           =   8565
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   720
      Left            =   105
      TabIndex        =   8
      Top             =   2070
      Width           =   2385
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1140
         TabIndex        =   10
         Top             =   300
         Width           =   1185
      End
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   300
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
      Left            =   4590
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   240
      Visible         =   0   'False
      Width           =   1740
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   6570
      Top             =   1050
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
      Bindings        =   "RelEstoqueTodos.frx":4E96E
      DataSource      =   "Data1"
      Height          =   345
      Left            =   1125
      TabIndex        =   0
      Top             =   120
      Width           =   1200
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
      _ExtentX        =   2117
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Nome_Fornecedor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2400
      TabIndex        =   15
      Top             =   540
      Width           =   6270
   End
   Begin VB.Label Label2 
      Caption         =   "Fornecedor"
      Height          =   225
      Left            =   105
      TabIndex        =   14
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   6
      Top             =   165
      Width           =   855
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   6270
   End
End
Attribute VB_Name = "frmRelEstoque2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsTempo As Recordset
Dim rsEstoque As Recordset
Dim rsProdutos As Recordset
Dim rsForn_Prod As Recordset
Dim rsEstoque2 As Recordset
Dim rsCliFor As Recordset
Dim rsGrade As Recordset
Dim rsEdicoes As Recordset



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
 Dim Tamanho2 As Integer
 Dim Cor As Integer
 Dim Cor2 As Integer
 Dim sSql As String
 Dim Estoque As Double
 Dim Aux_Data As Variant
 Dim Aux_Classe As Integer
 Dim Aux_Sub_Classe As Integer
 Dim Aux_Produto As String
 Dim Edição As Long
 Dim Edição2 As Long
 Dim Aux_Prod As String
 Dim Aux_Completo As String
 
 
 
 Call StatusMsg("")

 Rem Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a empresa."
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

 sSql = "Delete * From [ZZZVendas]"
 db.Execute sSql

 Call StatusMsg("")
 

 Rem Le estoque e joga no temporário
 rsEstoque2.Index = "Produto"
 rsProdutos.Index = "Código"
 rsForn_Prod.Index = "Produto"
 rsTempo.Index = "Produto"
 rsGrade.Index = "Original"
 rsEdicoes.Index = "Produto"
 Termina = False
 Produto = 0
 Tamanho = 0
 Cor = 0
 Edição = 0
 Call StatusMsg("Aguarde, contando estoque.")

 
 Aux_Prod = ""
 
LP1S:
  rsProdutos.Seek ">", Aux_Prod
  If rsProdutos.NoMatch Then GoTo Imprime
    
  '14/01/2005 - Daniel
  'Em algumas bases de dados o campo Produtos.Código está
  'com caracteres incorretos tais como 
  'Solicitação de manutenção do código pelo cliente
  'São Francisco Móveis e Eletro. de Olinda - PE
  Dim bytAsc                  As Byte
  Dim intAuxiliar             As Integer
  Dim blnPossuiCaracIncorreto As Boolean
  
  For intAuxiliar = 1 To Len(rsProdutos("Código").Value)
    bytAsc = Asc(Mid(rsProdutos("Código").Value, intAuxiliar, 1))
    
    If bytAsc = 127 Then '127 = 
      blnPossuiCaracIncorreto = True
      Exit For
    End If
  Next intAuxiliar
  
  If Not blnPossuiCaracIncorreto Then
    Aux_Prod = rsProdutos("Código")
    If Aux_Prod = "0" Then GoTo LP1S
  Else
    blnPossuiCaracIncorreto = False
    
    Aux_Prod = "0"
    GoTo LP1S
  End If
  
  'Antigo código comentado em 14/01/2005 - Daniel
  'Aux_Prod = rsProdutos("Código")
  'If Aux_Prod = "0" Then GoTo LP1S
  
  '------------------------------------------------------
  
  Rem Verifica Fornecedor
  If Nome_Fornecedor.Caption <> "" Then
    rsForn_Prod.Seek "=", rsProdutos("Código"), Val(Combo_Fornecedor.Text)
    If rsForn_Prod.NoMatch Then GoTo LP1S
  End If

  Call StatusMsg("Verificando produto " + Aux_Prod)
  

  If rsProdutos("Tipo") = "N" And O_Normal.Value = True Then
     Estoque = 0
     Aux_Data = Null
     rsEstoque2.Seek "=", Val(Combo.Text), Aux_Prod, 0, 0, 0
     If Not rsEstoque2.NoMatch Then
       Estoque = rsEstoque2("Estoque Atual")
       Aux_Data = rsEstoque2("Última Data")
     End If

     rsTempo.Seek "=", Aux_Prod, 0, 0, 0
     If Not rsTempo.NoMatch Then rsTempo.Delete
     rsTempo.AddNew
       rsTempo("Produto") = Aux_Prod
       rsTempo("Tamanho") = 0
       rsTempo("Cor") = 0
       rsTempo("Edição") = 0
     
       rsTempo("Classe") = rsProdutos("Classe")
       rsTempo("Sub Classe") = rsProdutos("Sub Classe")
       rsTempo("Vendas") = Estoque
       rsTempo("Última Data") = Format(Aux_Data, "dd/mm/yyyy")
     rsTempo.Update
  End If
  
  
  
  '
  '  PRODUTOS COM GRADE
  ' --------------------
  If rsProdutos("Tipo") = "G" And O_Grade.Value = True Then
     Aux_Completo = ""
     
Lp_Grade:
     rsGrade.Seek ">", Aux_Prod, Aux_Completo
     If rsGrade.NoMatch Then GoTo Cont_Edição
     If rsGrade("Código Original") <> Aux_Prod Then GoTo Cont_Edição
     Aux_Completo = rsGrade("Código")
     Erro = Len(Aux_Prod)
     Erro = Len(Aux_Completo) - Erro
     Str1 = Right(Aux_Completo, Erro)
     Tamanho = Left(Str1, 3)
     Cor = Right(Str1, 3)
     
     Estoque = 0
     Aux_Data = Null
     
     rsEstoque2.Seek "=", Val(Combo.Text), Aux_Prod, Tamanho, Cor, 0
     If Not rsEstoque2.NoMatch Then
       Estoque = rsEstoque2("Estoque Atual")
       Aux_Data = rsEstoque2("Última Data")
     End If

     If O_Grade.Value = False Then
       Tamanho2 = 0
       Cor2 = 0
     Else
       Tamanho2 = Tamanho
       Cor2 = Cor
     End If
     
     rsTempo.Seek "=", Aux_Prod, Tamanho2, Cor2, 0
     If Not rsTempo.NoMatch Then
        rsTempo.Edit
        
        rsTempo("Vendas") = rsTempo("Vendas") + Estoque
        rsTempo("Última Data") = Format(Aux_Data, "dd/mm/yyyy")
     Else
       rsTempo.AddNew
       
       rsTempo("Produto") = Aux_Prod
       rsTempo("Tamanho") = Tamanho2
       rsTempo("Cor") = Cor2
       rsTempo("Edição") = 0
     
       rsTempo("Classe") = rsProdutos("Classe")
       rsTempo("Sub Classe") = rsProdutos("Sub Classe")
       rsTempo("Vendas") = Estoque
       rsTempo("Última Data") = Format(Aux_Data, "dd/mm/yyyy")
     End If
     
     rsTempo.Update
     
     GoTo Lp_Grade
  End If
     
     


  '
  'Produtos com Edição
Cont_Edição:
  If rsProdutos("Tipo") = "E" And O_Edição.Value = True Then
     Edição = 0
     
Lp_Edição:
     rsEdicoes.Seek ">", Aux_Prod, Edição
     If rsEdicoes.NoMatch Then GoTo LP1S
     If rsEdicoes("Produto") <> Aux_Prod Then GoTo LP1S
     Edição = rsEdicoes("Código")
     
        
     Estoque = 0
     Aux_Data = Null
     
     rsEstoque2.Seek "=", Val(Combo.Text), Aux_Prod, 0, 0, Edição
     If Not rsEstoque2.NoMatch Then
       Estoque = rsEstoque2("Estoque Atual")
       Aux_Data = rsEstoque2("Última Data")
     End If

     If O_Edição.Value = False Then
       Edição2 = 0
     Else
       Edição2 = Edição
     End If
     
     rsTempo.Seek "=", Aux_Prod, 0, 0, Edição2
     If Not rsTempo.NoMatch Then
        rsTempo.Edit
        
        rsTempo("Vendas") = rsTempo("Vendas") + Estoque
        rsTempo("Última Data") = Format(Aux_Data, "dd/mm/yyyy")
     Else
       rsTempo.AddNew
       
       rsTempo("Produto") = Aux_Prod
       rsTempo("Tamanho") = 0
       rsTempo("Cor") = 0
       rsTempo("Edição") = Edição2
     
       rsTempo("Classe") = rsProdutos("Classe")
       rsTempo("Sub Classe") = rsProdutos("Sub Classe")
       rsTempo("Vendas") = Estoque
       rsTempo("Última Data") = Format(Aux_Data, "dd/mm/yyyy")
     End If
     
     rsTempo.Update
     
     GoTo Lp_Edição
  End If

  GoTo LP1S



Imprime:
 

 If O_Estoque.Value = 1 Then
   Call StatusMsg("Aguarde, apagando produtos com estoque 0...")

   sSql = "Delete * From [ZZZVendas] Where Vendas = 0"
   db.Execute sSql
 End If
 
 Call StatusMsg("")

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel1.DataFiles(0) = Str1

 Rem Saída
 If B_Vídeo = True Then Rel1.Destination = 0
 If B_Impressora = True Then Rel1.Destination = 1
 Rem If B_Arquivo = True Then
 Rem    frmMenu.Relatório.Destination = 2
 Rem    frmMenu.Relatório.PrintFileName = T_Arquivo.Text
 Rem End If

 Rem Nome do arquivo .rpt
 If O_Classe.Value = 0 Then
   If O_Normal.Value = True Then Str1 = gsReportPath & "ESTOQ_N.RPT"
   If O_Grade.Value = True Then Str1 = gsReportPath & "ESTOQ_G.RPT"
   If O_Edição.Value = True Then Str1 = gsReportPath & "ESTOQ_E.RPT"
 End If
 If O_Classe.Value = 1 Then
   If O_Normal.Value = True Then Str1 = gsReportPath & "ESTOQ_NC.RPT"
   If O_Grade.Value = True Then Str1 = gsReportPath & "ESTOQ_GC.RPT"
   If O_Edição.Value = True Then Str1 = gsReportPath & "ESTOQ_EC.RPT"
 End If
 Rel1.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 Rem Seleção
 Str_Rel = ""
 If O_Inativos.Value = 0 Then
  Str_Rel = "{Produtos.Desativado} = False"
 End If
 Rel1.SelectionFormula = Str_Rel
 
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 
 Rel1.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 Rel1.Formulas(1) = Str_Rel
 
 
 If O_Classe.Value = 1 Then
   If O_Código.Value = True Then
     Rel1.SortFields(0) = "+{ZZZVendas.Classe}"
     Rel1.SortFields(1) = "+{Produtos.Código Ordenação}"
   End If
   If O_Nome.Value = True Then
     Rel1.SortFields(0) = "+{ZZZVendas.Classe}"
     Rel1.SortFields(1) = "+{Produtos.Nome}"
   End If
 End If

 If O_Classe.Value = 0 Then
   If O_Código.Value = True Then
     Rel1.SortFields(0) = "+{Produtos.Código Ordenação}"
     Rel1.SortFields(1) = ""
   End If
   If O_Nome.Value = True Then
     Rel1.SortFields(0) = "+{Produtos.Nome}"
     Rel1.SortFields(1) = ""
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

Private Sub Combo_CloseUp()
Combo.Text = Combo.Columns(1).Text
Combo_LostFocus
End Sub

Private Sub Combo_Fornecedor_CloseUp()

Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
Combo_Fornecedor_LostFocus

End Sub

Private Sub Combo_Fornecedor_LostFocus()
  Nome_Fornecedor.Caption = ""
  If IsNull(Combo_Fornecedor.Text) Then Exit Sub
  If Combo_Fornecedor.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
  If Val(Combo_Fornecedor.Text) < 0 Then Exit Sub
  If Val(Combo_Fornecedor.Text) > 99999999 Then Exit Sub

  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Val(Combo_Fornecedor.Text)
  If rsCliFor.NoMatch Then Exit Sub
  Nome_Fornecedor.Caption = rsCliFor("Nome") & ""

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
  Set rsTempo = db.OpenRecordset("ZZZVendas")
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsForn_Prod = db.OpenRecordset("Forn_Prod", , dbReadOnly)
  Set rsEstoque2 = db.OpenRecordset("Estoque Final", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsEdicoes = db.OpenRecordset("Edições", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  Combo.Text = gnCodFilial
  
  If gbGrade = False Then O_Grade.Enabled = False
  If gbEdicao = False Then O_Edição.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

  rsParametros.Close
  rsTempo.Close
  rsProdutos.Close
  rsForn_Prod.Close
  rsEstoque2.Close
  rsCliFor.Close
  rsGrade.Close
  rsEdicoes.Close
  
  
  Set rsParametros = Nothing
  Set rsTempo = Nothing
  Set rsProdutos = Nothing
  Set rsForn_Prod = Nothing
  Set rsEstoque2 = Nothing
  Set rsCliFor = Nothing
  Set rsGrade = Nothing
  Set rsEdicoes = Nothing
  
End Sub
