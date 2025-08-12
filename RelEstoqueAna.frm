VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEstoqueAna 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Estoque Analitico"
   ClientHeight    =   2640
   ClientLeft      =   1440
   ClientTop       =   1875
   ClientWidth     =   6975
   ForeColor       =   &H80000008&
   Icon            =   "RelEstoqueAna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   6975
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   1575
      TabIndex        =   16
      Top             =   1680
      Width           =   2010
      Begin VB.OptionButton O_Valores 
         Caption         =   "Estoque e valores"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   555
         Width           =   1800
      End
      Begin VB.OptionButton O_Estoque 
         Caption         =   "Somente estoque"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   255
         Value           =   -1  'True
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   525
         Width           =   1095
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   255
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   5505
      TabIndex        =   8
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   210
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3570
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   2775
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.ComboBox Nome_Mês 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1125
      Width           =   2295
   End
   Begin VB.TextBox Ano 
      Height          =   315
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1140
      Width           =   615
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   6420
      Top             =   930
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Prod 
      Bindings        =   "RelEstoqueAna.frx":058A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1095
      TabIndex        =   1
      Top             =   540
      Width           =   1815
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
      Columns(0).Width=   8229
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3493
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelEstoqueAna.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
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
      Columns(0).Width=   9340
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1614
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Filial:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   14
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3015
      TabIndex        =   13
      Top             =   120
      Width           =   3840
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Ano :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3735
      TabIndex        =   12
      Top             =   1170
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Mês :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Nome_Prod 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3000
      TabIndex        =   10
      Top             =   525
      Width           =   3855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Produto :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   615
      Width           =   735
   End
End
Attribute VB_Name = "frmRelEstoqueAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsProdutos As Recordset
Dim rsParametros As Recordset
Dim rsEstoque As Recordset
Dim rsZZZ4 As Recordset

Private Sub Ano_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub


Private Sub B_Imprime_Click()
  Dim Aux As Integer
  Dim Mês_Aux As Integer
  Dim Mês As Integer
  Dim Ano_Aux As Integer
  Dim Venda_Aux As Long
  Dim Nome_Aux As String
  Dim Ult_Mês As Integer
  Dim Ult_Ano As Integer
  Dim Pen_Mês As Integer
  Dim Pen_Ano As Integer
  Dim Ult_Mês_AA As Integer
  Dim Ult_Ano_AA As Integer
  Dim Mês_AA As Integer
  Dim Ano_AA As Integer
  Dim Prox_Mês_AA As Integer
  Dim Prox_Ano_AA As Integer
  Dim Termina As Integer
  Dim Produto As String
  Dim Estoque As Double
  Dim Vendas_Aux As Double
  Dim Dias_Corridos As Integer
  Dim Dias_Faltantes As Integer
  Dim Data_Aux As String
  Dim Str1 As String
  Dim Str_Rel As String
  Dim Str_Aux As String
  Dim sSql As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Contador As Long
  Dim Data_Ini As Date
  Dim Data_Fim As Date
  Dim Venda_Valor_Aux As Double
  Dim Aux1 As String
  


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



  If Nome_Mês.Text = "" Then
    DisplayMsg "Escolha um mês da lista."
    Nome_Mês.SetFocus
    Exit Sub
  End If

  If IsNull(Combo_Prod.Text) Then
    Combo_Prod.Text = ""
  End If
  

  If IsNull(Nome_Prod.Caption) Or Nome_Prod.Caption = "" Then
   
  '  Call StatusMsg("Escolha um produto."
  '  Combo_Prod.SetFocus
  '  Exit Sub
   
  End If

  If IsNull(Ano.Text) Then
    DisplayMsg "Escolha um ano."
    Ano.SetFocus
    Exit Sub
  End If

  If Not IsNumeric(Ano.Text) Or Ano.Text = "" Then
    DisplayMsg "Escolha um ano."
    Ano.SetFocus
    Exit Sub
  End If

  If Val(Ano.Text) < 1995 Or Val(Ano.Text) > 2200 Then
    DisplayMsg "Digite o ano, com 4 dígitos (ex. 1997)"
    Ano.SetFocus
    Exit Sub
  End If
  
  


  If Nome_Mês.Text = "Janeiro" Then Mês = 1
  If Nome_Mês.Text = "Fevereiro" Then Mês = 2
  If Nome_Mês.Text = "Março" Then Mês = 3
  If Nome_Mês.Text = "Abril" Then Mês = 4
  If Nome_Mês.Text = "Maio" Then Mês = 5
  If Nome_Mês.Text = "Junho" Then Mês = 6
  If Nome_Mês.Text = "Julho" Then Mês = 7
  If Nome_Mês.Text = "Agosto" Then Mês = 8
  If Nome_Mês.Text = "Setembro" Then Mês = 9
  If Nome_Mês.Text = "Outubro" Then Mês = 10
  If Nome_Mês.Text = "Novembro" Then Mês = 11
  If Nome_Mês.Text = "Dezembro" Then Mês = 12

  GoSub Calcula_Mes





  Rem Agora zera o arquivo zzz4

  Call StatusMsg("Aguarde, preparando arquivo temporário ...")
  sSql = "Delete * From Analítico"
  dbTemp.Execute sSql
  Call StatusMsg("")

  Rem Começa o loop de produtos

  Produto = Combo_Prod.Text
  If Produto = "" Then Produto = "0"
  
  rsProdutos.Index = "Código"
  rsEstoque.Index = "Data2"

  If Produto <> "" Then
    rsProdutos.Seek "=", Produto
  End If
    
Lp1:
  If Combo_Prod.Text = "" Then
    rsProdutos.Seek ">", Produto
  End If

  If rsProdutos.NoMatch Then GoTo Fim
  Produto = rsProdutos("Código")

  Call StatusMsg("Lendo produto " + Produto)


  Mês_Aux = Mês
  Ano_Aux = Val(Ano.Text)
  Venda_Aux = 0
  Venda_Valor_Aux = 0
  GoSub Acha_Venda
  
  rsZZZ4.AddNew
    rsZZZ4("Produto") = Produto
    rsZZZ4("Nome") = rsProdutos("Nome")
    rsZZZ4("Unidade Venda") = rsProdutos("Unidade Venda")
    
    rsZZZ4("Vendas Atual") = Venda_Aux
    rsZZZ4("Valor Vendas Atual") = Venda_Valor_Aux
  

  Mês_Aux = Ult_Mês
  Ano_Aux = Ult_Ano
  Venda_Aux = 0
  Venda_Valor_Aux = 0
  GoSub Acha_Venda

    rsZZZ4("Vendas Último") = Venda_Aux
    rsZZZ4("Valor Vendas Último") = Venda_Valor_Aux

  Mês_Aux = Pen_Mês
  Ano_Aux = Pen_Ano
  Venda_Aux = 0
  Venda_Valor_Aux = 0
  GoSub Acha_Venda
  
    rsZZZ4("Vendas Penúltimo") = Venda_Aux
    rsZZZ4("Valor Vendas Penúltimo") = Venda_Valor_Aux

  Mês_Aux = Ult_Mês_AA
  Ano_Aux = Ult_Ano_AA
  Venda_Aux = 0
  Venda_Valor_Aux = 0
  GoSub Acha_Venda

    rsZZZ4("Vendas Último AA") = Venda_Aux
    rsZZZ4("Valor Vendas Último AA") = Venda_Valor_Aux

  Mês_Aux = Mês_AA
  Ano_Aux = Ano_AA
  Venda_Aux = 0
  Venda_Valor_Aux = 0
  GoSub Acha_Venda
  
    rsZZZ4("Vendas AA") = Venda_Aux
    rsZZZ4("Valor Vendas AA") = Venda_Valor_Aux
  
  Mês_Aux = Prox_Mês_AA
  Ano_Aux = Prox_Ano_AA
  Venda_Aux = 0
  Venda_Valor_Aux = 0
  GoSub Acha_Venda

    rsZZZ4("Vendas Próximo AA") = Venda_Aux
    rsZZZ4("Valor Vendas Próximo AA") = Venda_Valor_Aux


  Rem Acha estoque
  Estoque = 0
  rsEstoque.Index = "Data2"
  rsEstoque.Seek ">", Val(Combo.Text), Produto, 0, 0, 0, CDate("12/12/2100")
  If Not rsEstoque.NoMatch Then
    If Val(Combo.Text) = rsEstoque("Filial") And Produto = rsEstoque("Produto") Then Estoque = rsEstoque("Estoque Final")
  End If

    rsZZZ4("Estoque") = Estoque



  rsZZZ4.Update
  If Nome_Prod.Caption = "" Then GoTo Lp1


Fim:

  Call StatusMsg("Aguarde, imprimindo ...")

' Rem  Nome do BD
'  With Rel
'    .DataFiles(0) = gsTempDBFileName
'    .DataFiles(1) = gsQuickDBFileName
'  End With

  
  '31/10/2002 - mpdea
  'Corrigido associação com a localização das bases de dados
  With Rel
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsTempDBFileName
    .DataFiles(2) = gsQuickDBFileName
  End With


 Rem Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 If O_Estoque.Value = True Then Str1 = gsReportPath & "ANALI.RPT"
 If O_Valores.Value = True Then Str1 = gsReportPath & "ANALI1.RPT"
 
 Rel.ReportFileName = Str1

 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel
 
 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 
 Rel.Formulas(0) = Str_Rel

 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 Rel.Formulas(1) = Str_Rel
 

 Str_Rel = "penul = '"
 Str_Aux = ""
 Mês_Aux = Pen_Mês
 GoSub Acha_Nome
 Str_Aux = Nome_Aux + str$(Pen_Ano)
 Str_Rel = Str_Rel + Str_Aux + "'"
 Rel.Formulas(2) = Str_Rel

 Str_Rel = "ult = '"
 Str_Aux = ""
 Mês_Aux = Ult_Mês
 GoSub Acha_Nome
 Str_Aux = Nome_Aux + str$(Ult_Ano)
 Str_Rel = Str_Rel + Str_Aux + "'"
 Rel.Formulas(3) = Str_Rel


 Str_Rel = "atual = '"
 Str_Aux = ""
 Mês_Aux = Mês
 GoSub Acha_Nome
 Str_Aux = Nome_Aux + Ano.Text
 Str_Rel = Str_Rel + Str_Aux + "'"
 Rel.Formulas(4) = Str_Rel


 Str_Rel = "ult_aa = '"
 Str_Aux = ""
 Mês_Aux = Ult_Mês_AA
 GoSub Acha_Nome
 Str_Aux = Nome_Aux + str$(Ult_Ano_AA)
 Str_Rel = Str_Rel + Str_Aux + "'"
 Rel.Formulas(5) = Str_Rel


 Str_Rel = "atual_aa = '"
 Str_Aux = ""
 Mês_Aux = Mês_AA
 GoSub Acha_Nome
 Str_Aux = Nome_Aux + str$(Ano_AA)
 Str_Rel = Str_Rel + Str_Aux + "'"
 Rel.Formulas(6) = Str_Rel


 Str_Rel = "prox_aa = '"
 Str_Aux = ""
 Mês_Aux = Prox_Mês_AA
 GoSub Acha_Nome
 Str_Aux = Nome_Aux + str$(Prox_Ano_AA)
 Str_Rel = Str_Rel + Str_Aux + "'"
 Rel.Formulas(7) = Str_Rel


 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  

 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault


Exit Sub



Acha_Venda:
  Aux1 = "01/"
  If Mês_Aux < 12 Then
    Aux1 = Aux1 + Trim(str(Mês_Aux + 1)) + "/" + Trim(str(Ano_Aux))
  Else
    Aux1 = Aux1 + "01/"
    Aux1 = Aux1 + Trim(str((Ano_Aux + 1)))
  End If
  
  Data_Fim = CDate(Aux1)
  Data_Fim = Data_Fim - 1
      
  Aux1 = "01/" + Trim(str(Mês_Aux)) + "/" + Trim(str(Ano_Aux))
  Data_Ini = CDate(Aux1) - 1
      

  Contador = 0
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edição = 0
  rsEstoque.Index = "Data"
Acha_Venda2:
  rsEstoque.Seek ">", Val(Combo.Text), Produto, Aux_Tamanho, Aux_Cor, Aux_Edição, Data_Ini
  If rsEstoque.NoMatch Then Return
  
  If rsEstoque("Data") > Data_Fim Then Return
  If rsEstoque("Filial") <> Val(Combo.Text) Then Return
  
  If rsEstoque("Produto") <> Produto Then Return
  
  Data_Ini = rsEstoque("Data")
  Aux_Tamanho = rsEstoque("Tamanho")
  Aux_Cor = rsEstoque("Cor")
  Aux_Edição = rsEstoque("Edição")
  
  Venda_Aux = Venda_Aux + rsEstoque("Vendas")
  Venda_Valor_Aux = Venda_Valor_Aux + rsEstoque("Valor Vendas")
       
  GoTo Acha_Venda2
  
  
Calcula_Mes:

  Ult_Mês = Mês - 1
  Ult_Ano = Val(Ano.Text)
  If Ult_Mês = 0 Then
    Ult_Ano = Ano - 1
    Ult_Mês = 12
  End If

  Pen_Mês = Ult_Mês - 1
  Pen_Ano = Ult_Ano
  If Pen_Mês = 0 Then
    Pen_Ano = Pen_Ano - 1
    Pen_Mês = 12
  End If

  Ult_Mês_AA = Ult_Mês
  Ult_Ano_AA = Ult_Ano - 1

  Mês_AA = Mês
  Ano_AA = Val(Ano.Text) - 1

  Prox_Mês_AA = Mês + 1
  Prox_Ano_AA = Val(Ano.Text) - 1
  If Prox_Mês_AA = 13 Then
    Prox_Ano_AA = Prox_Ano_AA + 1
    Prox_Mês_AA = 1
  End If


  Return

Acha_Nome:
  If Mês_Aux = 1 Then Nome_Aux = "Jan/"
  If Mês_Aux = 2 Then Nome_Aux = "Fev/"
  If Mês_Aux = 3 Then Nome_Aux = "Mar/"
  If Mês_Aux = 4 Then Nome_Aux = "Abr/"
  If Mês_Aux = 5 Then Nome_Aux = "Mai/"
  If Mês_Aux = 6 Then Nome_Aux = "Jun/"
  If Mês_Aux = 7 Then Nome_Aux = "Jul/"
  If Mês_Aux = 8 Then Nome_Aux = "Ago/"
  If Mês_Aux = 9 Then Nome_Aux = "Set/"
  If Mês_Aux = 10 Then Nome_Aux = "Out/"
  If Mês_Aux = 11 Then Nome_Aux = "Nov/"
  If Mês_Aux = 12 Then Nome_Aux = "Dez/"
  
  Return

End Sub




Private Sub Combo_CloseUp()
Combo.Text = Combo.Columns(1).Text
Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
  Call StatusMsg("")
 
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

Private Sub Command3D1_Click()

End Sub

Private Sub Command3D2_Click()
End Sub

Private Sub Combo_Prod_CloseUp()

 Combo_Prod.Text = Combo_Prod.Columns(1).Text
 Combo_Prod_LostFocus

End Sub

Private Sub Combo_Prod_GotFocus()
  Call StatusMsg(LoadResString(51))
End Sub

Private Sub Combo_Prod_LostFocus()
  Call StatusMsg("")
 
  Nome_Prod.Caption = ""
  If gsHandleNull(Combo_Prod.Text) = "0" Then
    Exit Sub
  End If

  rsProdutos.Index = "Código"
  rsProdutos.Seek "=", Combo_Prod.Text

  If rsProdutos.NoMatch Then Exit Sub
  Nome_Prod.Caption = rsProdutos("Nome")

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)

 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
 Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 Set rsZZZ4 = dbTemp.OpenRecordset("Analítico")
 Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)

 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName

 Combo.Text = gnCodFilial

 Nome_Mês.AddItem "Janeiro"
 Nome_Mês.AddItem "Fevereiro"
 Nome_Mês.AddItem "Março"
 Nome_Mês.AddItem "Abril"
 Nome_Mês.AddItem "Maio"
 Nome_Mês.AddItem "Junho"
 Nome_Mês.AddItem "Julho"
 Nome_Mês.AddItem "Agosto"
 Nome_Mês.AddItem "Setembro"
 Nome_Mês.AddItem "Outubro"
 Nome_Mês.AddItem "Novembro"
 Nome_Mês.AddItem "Dezembro"
End Sub


Private Sub Form_Unload(Cancel As Integer)

 rsProdutos.Close
 rsParametros.Close
 rsZZZ4.Close
 rsEstoque.Close
 
  Set rsProdutos = Nothing
 Set rsParametros = Nothing
 Set rsZZZ4 = Nothing
 Set rsEstoque = Nothing

End Sub
