VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelProdComprar 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Produtos com Estoque Abaixo do Mínimo"
   ClientHeight    =   5280
   ClientLeft      =   1725
   ClientTop       =   2175
   ClientWidth     =   6630
   ForeColor       =   &H80000008&
   HelpContextID   =   1500
   Icon            =   "RelEstoqueComprar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5280
   ScaleWidth      =   6630
   Begin VB.Frame Frame3 
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   3015
      Begin VB.CheckBox O_Classe 
         Caption         =   "Separar por classe / sub classe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   2
         Top             =   360
         Width           =   2625
      End
      Begin VB.CheckBox O_Fornecedor 
         Caption         =   "Imprimir os fornecedores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   720
         Width           =   2640
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   15
      Top             =   3600
      Width           =   1455
      Begin VB.OptionButton O_Edição 
         Caption         =   "Edição"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   960
      End
      Begin VB.OptionButton O_Grade 
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   960
      End
      Begin VB.OptionButton O_Normal 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "Imprimir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "RelEstoqueComprar.frx":058A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
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
      Columns(0).Width=   8705
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1746
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   2760
      Top             =   4800
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
      Bindings        =   "RelEstoqueComprar.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
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
      Columns(0).Width=   9234
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1720
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   -120
      TabIndex        =   17
      Top             =   -120
      Width           =   6855
      Begin VB.TextBox txtDicas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Text            =   "RelEstoqueComprar.frx":05B2
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dicas ao usuário !"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         TabIndex        =   18
         Top             =   480
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   360
         Picture         =   "RelEstoqueComprar.frx":076A
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Label Nome_Fornecedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2520
      TabIndex        =   14
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Fornecedor :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2520
      TabIndex        =   11
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2715
      Width           =   855
   End
End
Attribute VB_Name = "frmRelProdComprar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'24/01/2006 - mpdea
'Modificado relatório para que funcione corretamente

Private rsParametros As Recordset
Private rsTempo As Recordset
Private rsEstoque As Recordset
Private rsProdutos As Recordset
Private rsCliFor As Recordset
Private rsForn_Prod As Recordset
Private rsTamanhos As Recordset
Private rsCores As Recordset
Private rsClasses As Recordset
Private rsSub_Classes As Recordset
Private rsEdicoes As Recordset

'Private Sub B_Imprime_Click()
'  Dim Termina As Integer
'  Dim Val2 As Integer
'  Dim Erro As Integer
'  Dim Atual As Double
'  Dim Str1 As String
'  Dim Str2 As String
'  Dim Str3 As String
'  Dim Str_Data1 As String
'  Dim Str_Data2 As String
'  Dim Str_Rel As String
'  Dim Data1 As Variant
'  Dim Produto As String
'  Dim Produto2 As String
'  Dim Comprar As Integer
'  Dim sSql As String
'  Dim Tamanho As Integer
'  Dim Cor As Integer
'  Dim Edição As Long
'  Dim Aux_Tamanho As Integer
'  Dim Aux_Cor As Integer
'  Dim Aux_Edição As Long
'  Dim Fornecedor As Long
'
'
'
'  Call StatusMsg("")
'
'  Rem Verifica empresa
'  If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
'    DisplayMsg "Escolha a empresa."
'    Combo.SetFocus
'    Exit Sub
'  End If
'
'  If Filial_Liberada <> 0 Then
'    If Val(Combo.Text) <> Filial_Liberada Then
'      DisplayMsg "Funcionário não tem acesso a esta filial."
'      Exit Sub
'    End If
'  End If
'
'
'
'  If IsNull(Nome_Fornecedor.Caption) Then Nome_Fornecedor.Caption = ""
'  If Nome_Fornecedor.Caption = "" Then
'    Fornecedor = 0
'  Else
'    Fornecedor = Val(Combo_Fornecedor.Text)
'  End If
'
'
'
'
'  Rem apaga pesquisa anterior desta filial do arquivo temporario
'  Call StatusMsg("Aguarde, preparando arquivo temporário ...")
'  sSql = "Delete * From Comprar"
'  dbTemp.Execute sSql
'
'
'  Rem Le produtos e joga os que precisa no temporário
'  rsEstoque.Index = "Produto"
'  rsTempo.Index = "Produto"
'  rsTamanhos.Index = "Código"
'  rsCores.Index = "Código"
'  rsClasses.Index = "Código"
'  rsSub_Classes.Index = "Código"
'  rsEdicoes.Index = "Produto"
'  rsCliFor.Index = "Código"
'
'
'  Termina = False
'  Produto2 = ""
'  Call StatusMsg("Aguarde, contando estoque...")
'
'  rsProdutos.Index = "Código"
'  Do While Not Termina
'Outro_Prod:
'
'   rsProdutos.Seek ">", Produto2
'   If rsProdutos.NoMatch Then Termina = True
'  '  DoEvents
'   If Not Termina Then Produto2 = rsProdutos("Código")
'
'   '14/01/2005 - Daniel
'   'Em algumas bases de dados o campo Produtos.Código está
'   'aparecendo com caracteres incorretos tais como ...
'   'isto estava gerando o BUG com mensagem 3163
'   '
'   'Case: São Francisco Móveis e Eletro. de Olinda - PE
'   If Len(Produto2) > 20 Then Produto2 = "0"
'   '------------------------------------------------------
'
'  '  Call StatusMsg("Lendo produto " + Produto2)
'  '  If Produto2 = "1224" Then
'  '    Produto2 = Produto2
'  '  End If
'
'   If Not Termina Then
'     If Fornecedor <> 0 Then
'        rsForn_Prod.Index = "Produto"
'        rsForn_Prod.Seek "=", Produto2, Fornecedor
'        If rsForn_Prod.NoMatch Then GoTo Outro_Prod
'     End If
'     'Código de verificação incluído para o Run-time 3022 na base Temp.mdb
'     '15/01/2001 - por mpdea
'     If O_Normal.Value Then
'       If rsProdutos("Tipo") <> "N" Then GoTo Outro_Prod
'     ElseIf O_Grade.Value Then
'       If rsProdutos("Tipo") <> "G" Then GoTo Outro_Prod
'     ElseIf O_Edição.Value Then
'       If rsProdutos("Tipo") <> "E" Then GoTo Outro_Prod
'     End If
'
'     '19/08/2003 - maikel
'     '             Agora produtos inativos não saem mais no relatório de produtos a comprar
'     If rsProdutos("Desativado") = True Then GoTo Outro_Prod
'
'   End If
'   If Not Termina Then
'     Comprar = False
'     If rsProdutos("Estoque Ideal") > 0 Then
'        Rem agora procura o estoque
'        Comprar = True
'        Atual = 0
'        Tamanho = 0
'        Cor = 0
'        Edição = -1
'Le_Estoque:
'        rsEstoque.Seek ">", Val(Combo.Text), Produto2, Tamanho, Cor, Edição
'        If rsEstoque.NoMatch Then Comprar = True 'Comprar = False
'        If Not rsEstoque.NoMatch Then
'           If rsEstoque("Filial") <> Val(Combo.Text) Then Comprar = False
'           If rsEstoque("Filial") = Val(Combo.Text) Then
'             'If Produto2 = "3" Then Stop
'             'If rsEstoque("Produto") <> Produto2 Then
'               Comprar = True  'Comprar = False
'             'If rsEstoque("Produto") = Produto2 Then
'                Tamanho = rsEstoque("Tamanho")
'                Cor = rsEstoque("Cor")
'                Edição = rsEstoque("Edição")
'                If (rsEstoque("Estoque Atual") >= rsProdutos("Estoque Mínimo")) And _
'                   rsEstoque("Produto") = Produto2 Then
'                  Comprar = False
'                End If
'                If rsEstoque("Estoque Atual") < rsProdutos("Estoque Mínimo") And _
'                   rsEstoque("Produto") = Produto2 Then
'                   Atual = rsEstoque("Estoque Atual")
'                End If
'  '               If rsProdutos("Estoque Mínimo") = 0 Then Comprar = False
'               'Substitui a anterior v. 6.0.40
'               If Not rsProdutos("Estoque") Then Comprar = False
'             ' End If
'           End If
'        End If
'
'
'        If Comprar = True Then
'          Aux_Tamanho = Tamanho
'          Aux_Cor = Cor
'          Aux_Edição = Edição
'          If O_Normal.Value = True Then
'            Aux_Tamanho = 0
'            Aux_Cor = 0
'            Aux_Edição = 0
'          End If
'          If O_Grade.Value = True Then
'            Aux_Edição = 0
'          End If
'          If O_Edição.Value = True Then
'            Aux_Tamanho = 0
'            Aux_Cor = 0
'          End If
'
'          rsTempo.Seek "=", Produto2, Tamanho, Cor, Edição
'          If rsTempo.NoMatch Then
'
'
'            rsTempo.AddNew
'
'            rsTempo("Código") = Produto2
'            rsTempo("Nome") = rsProdutos("Nome")
'            rsTempo("Fracionado") = rsProdutos("Fracionado")
'            rsTempo("Unidade Venda") = rsProdutos("Unidade Venda")
'            rsTempo("Último Custo") = rsProdutos("Último Custo")
'
'            If Aux_Tamanho <> 0 Then
'              rsTamanhos.Seek "=", Aux_Tamanho
'              If Not rsTamanhos.NoMatch Then
'                rsTempo("Tamanho") = Aux_Tamanho
'                rsTempo("Nome Tamanho") = rsTamanhos("Nome")
'              End If
'            End If
'
'            If Aux_Cor <> 0 Then
'              rsCores.Seek "=", Aux_Cor
'              If Not rsCores.NoMatch Then
'                rsTempo("Cor") = Aux_Cor
'                rsTempo("Nome Cor") = rsCores("Nome")
'              End If
'            End If
'
'            If Aux_Edição <> 0 Then
'              rsEdicoes.Seek "=", Produto2, Aux_Edição
'              If Not rsEdicoes.NoMatch Then
'                rsTempo("Edição") = Aux_Edição
'                rsTempo("Nome Edição") = rsEdicoes("Nome")
'              End If
'            End If
'
'            rsClasses.Seek "=", rsProdutos("Classe")
'            If Not rsClasses.NoMatch Then
'              rsTempo("Classe") = rsProdutos("Classe")
'              rsTempo("Nome Classe") = rsClasses("Nome")
'            End If
'
'            rsSub_Classes.Seek "=", rsProdutos("Sub Classe")
'            If Not rsSub_Classes.NoMatch Then
'              rsTempo("Sub Classe") = rsProdutos("Sub Classe")
'              rsTempo("Nome Sub") = rsSub_Classes("Nome")
'            End If
'
'
'           '28/10/2005 - mpdea
'           'Modificado para que insira as informações somente uma vez, pois
'           'obtém os dados de Estoque Final
'           'Comentado a edição dos dados existentes
'           '
'           'rsTempo("Estoque") = 0
'           'rsTempo("Estoque") = rsTempo("Estoque") + Atual
'           rsTempo("Estoque") = Atual
'           rsTempo("Ideal") = rsProdutos("Estoque Ideal")
'           rsTempo.Update
'
'  '         Else
'  '
'  '           rsTempo.Edit
'          End If
'
'
'         ' Comprar = False   V 4.0.18
'        End If
'
'        '---[ Nova verificação ]---'
'         If Not (rsEstoque.NoMatch) Then
'           If rsEstoque.Fields("Produto").Value <> Produto2 Then Comprar = False
'         Else
'           Comprar = False
'         End If
'        '---[ Nova verificação ]---'
'
'        If Comprar = True Then GoTo Le_Estoque
'      End If
'   End If
'  Loop
'
'  Call StatusMsg("")
'
'  If O_Fornecedor.Value = 1 Then
'    Call StatusMsg("Aguarde, verificando fornecedores...")
'    rsTempo.Index = "Produto"
'    Produto = 0
'    Aux_Tamanho = 0
'    Aux_Cor = 0
'    Aux_Edição = 0
'    rsForn_Prod.Index = "Produto"
'
'Lp_Prod:
'    rsTempo.Seek ">", Produto, Aux_Tamanho, Aux_Cor, Aux_Edição
'    If rsTempo.NoMatch Then GoTo Imprime
'    Produto = rsTempo("Código")
'    Aux_Tamanho = rsTempo("Tamanho")
'    Aux_Cor = rsTempo("Cor")
'    Aux_Edição = rsTempo("Edição")
'
'    Fornecedor = 0
'
'    rsForn_Prod.Seek ">", Produto, Fornecedor
'    If rsForn_Prod.NoMatch Then GoTo Lp_Prod
'    If rsForn_Prod("Produto") <> Produto Then GoTo Lp_Prod
'
'    Fornecedor = rsForn_Prod("Fornecedor")
'
'    rsCliFor.Seek "=", Fornecedor
'    If rsCliFor.NoMatch Then GoTo Forn2
'
'    rsTempo.Edit
'      rsTempo("Fornece1") = Fornecedor
'      rsTempo("Nome1") = rsCliFor("Nome")
'      rsTempo("Tel1_1") = rsCliFor("Fone 1")
'      rsTempo("Tel1_2") = rsCliFor("Fone 2")
'      rsTempo("Fax1") = rsCliFor("Fax")
'    rsTempo.Update
'
'
'Forn2:
'    rsForn_Prod.Seek ">", Produto, Fornecedor
'    If rsForn_Prod.NoMatch Then GoTo Lp_Prod
'    If rsForn_Prod("Produto") <> Produto Then GoTo Lp_Prod
'
'    Fornecedor = rsForn_Prod("Fornecedor")
'
'    rsCliFor.Seek "=", Fornecedor
'    If rsCliFor.NoMatch Then GoTo Forn3
'
'    rsTempo.Edit
'      rsTempo("Fornece2") = Fornecedor
'      rsTempo("Nome2") = rsCliFor("Nome")
'      rsTempo("Tel2_1") = rsCliFor("Fone 1")
'      rsTempo("Tel2_2") = rsCliFor("Fone 2")
'      rsTempo("Fax2") = rsCliFor("Fax")
'    rsTempo.Update
'
'
'Forn3:
'    rsForn_Prod.Seek ">", Produto, Fornecedor
'    If rsForn_Prod.NoMatch Then GoTo Lp_Prod
'    If rsForn_Prod("Produto") <> Produto Then GoTo Lp_Prod
'
'    Fornecedor = rsForn_Prod("Fornecedor")
'
'    rsCliFor.Seek "=", Fornecedor
'    If rsCliFor.NoMatch Then GoTo Forn3
'
'    rsTempo.Edit
'      rsTempo("Fornece3") = Fornecedor
'      rsTempo("Nome3") = rsCliFor("Nome")
'      rsTempo("Tel3_1") = rsCliFor("Fone 1")
'      rsTempo("Tel3_2") = rsCliFor("Fone 2")
'      rsTempo("Fax3") = rsCliFor("Fax")
'    rsTempo.Update
'
'
'Forn4:
'
'    rsForn_Prod.Seek ">", Produto, Fornecedor
'    If rsForn_Prod.NoMatch Then GoTo Lp_Prod
'    If rsForn_Prod("Produto") <> Produto Then GoTo Lp_Prod
'
'    Fornecedor = rsForn_Prod("Fornecedor")
'
'    rsCliFor.Seek "=", Fornecedor
'    If rsCliFor.NoMatch Then GoTo Forn5
'
'    rsTempo.Edit
'      rsTempo("Fornece4") = Fornecedor
'      rsTempo("Nome4") = rsCliFor("Nome")
'      rsTempo("Tel4_1") = rsCliFor("Fone 1")
'      rsTempo("Tel4_2") = rsCliFor("Fone 2")
'      rsTempo("Fax4") = rsCliFor("Fax")
'    rsTempo.Update
'
'
'Forn5:
'
'  rsForn_Prod.Seek ">", Produto, Fornecedor
'  If rsForn_Prod.NoMatch Then GoTo Lp_Prod
'  If rsForn_Prod("Produto") <> Produto Then GoTo Lp_Prod
'
'  Fornecedor = rsForn_Prod("Fornecedor")
'
'  rsCliFor.Seek "=", Fornecedor
'  If rsCliFor.NoMatch Then GoTo Lp_Prod
'
'  rsTempo.Edit
'    rsTempo("Fornece5") = Fornecedor
'    rsTempo("Nome5") = rsCliFor("Nome")
'    rsTempo("Tel5_1") = rsCliFor("Fone 1")
'    rsTempo("Tel5_2") = rsCliFor("Fone 2")
'    rsTempo("Fax5") = rsCliFor("Fax")
'  rsTempo.Update
'
'  GoTo Lp_Prod
'  End If
'
'
'
'Imprime:
'
'   Rel.Reset
'
'  ' Rem  Nome do BD
'  '  With Rel
'  '    .DataFiles(0) = gsTempDBFileName
'  '    .DataFiles(1) = gsQuickDBFileName
'  '  End With
'
'
'   '31/10/2002 - mpdea
'   'Corrigido associação com a localização das bases de dados
'   With Rel
'     If O_Normal.Value Then
'       .DataFiles(0) = gsTempDBFileName
'       .DataFiles(1) = gsTempDBFileName
'       .DataFiles(2) = gsQuickDBFileName
'     End If
'     If O_Grade.Value Then
'       .DataFiles(0) = gsTempDBFileName
'     End If
'     If O_Edição.Value Then
'       .DataFiles(0) = gsTempDBFileName
'     End If
'   End With
'
'
'  Rem Saída
'  If B_Vídeo = True Then Rel.Destination = 0
'  If B_Impressora = True Then Rel.Destination = 1
'  Rem If B_Arquivo = True Then
'  Rem    frmMenu.Relatório.Destination = 2
'  Rem    frmMenu.Relatório.PrintFileName = T_Arquivo.Text
'  Rem End If
'
'  Rem Nome do arquivo .rpt
'  If O_Classe.Value = 0 Then
'     If O_Normal.Value = True Then Str1 = gsReportPath & "COMPRAR1.RPT"
'     If O_Grade.Value = True Then Str1 = gsReportPath & "COMPRAR1G.RPT"
'     If O_Edição.Value = True Then Str1 = gsReportPath & "COMPRAR1E.RPT"
'  End If
'  If O_Classe.Value = 1 Then
'     If O_Normal.Value = True Then Str1 = gsReportPath & "COMPRAR2.RPT"
'     If O_Grade.Value = True Then Str1 = gsReportPath & "COMPRAR2G.RPT"
'     If O_Edição.Value = True Then Str1 = gsReportPath & "COMPRAR2E.RPT"
'  End If
'
'  Rel.ReportFileName = Str1
'
'
'
'  Str_Rel = "nome_empresa = '"
'  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
'
'  Rel.Formulas(0) = Str_Rel
'
'  Str_Rel = "nome_filial = '"
'  Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
'
'  Rel.Formulas(1) = Str_Rel
'
'  Str_Rel = "nome_fornecedor = '"
'  If Nome_Fornecedor.Caption = "" Then
'    Str_Rel = Str_Rel + "Todos" + "'"
'  End If
'  If Nome_Fornecedor.Caption <> "" Then
'    Str_Rel = Str_Rel + Combo_Fornecedor.Text + " - " + Nome_Fornecedor.Caption + "'"
'  End If
'
'  Rel.Formulas(2) = Str_Rel
'
'  Str_Rel = "Fornecedor = '"
'  If O_Fornecedor.Value = 1 Then
'    Str_Rel = Str_Rel + "SIM'"
'  Else
'    Str_Rel = Str_Rel + "NÃO'"
'  End If
'
'  If O_Normal.Value = True Then Rel.Formulas(3) = Str_Rel
'
'
'
'  Call StatusMsg("Aguarde, imprimindo...")
'  MousePointer = vbHourglass
'
'  Rel.WindowState = crptMaximized
'
'
'  '25/07/2003 - mpdea
'  'Seta a impressora para relatório
'  Call SetPrinterName("REL", Rel)
'
'
'  '28/10/2005 - mpdea
'  'Exibe botão de configuração de impressão
'  Rel.WindowShowPrintSetupBtn = True
'
'
'  Rel.Action = 1
'
'  Call StatusMsg("")
'  MousePointer = vbDefault
'
'  Exit Sub
'
'End Sub

Private Sub FillTempData()
  Dim strSQL As String
  Dim rstGrade As Recordset
  
  Dim intFilial As Integer
  Dim lngFornecedor As Long
  Dim strTipo As String
  
  Dim blnComprar As Boolean
  Dim strCodigoProduto As String
  Dim strGrade As String
  Dim intTamanho As Integer
  Dim intCor As Integer
  Dim lngEdicao As Long
  
  Dim dblEstoqueAtual As Double
  
  
  On Error GoTo ErrHandler
  
  
  'Obtém filial
  Call IsDataType(dtInteger, Combo.Text, intFilial)
  
  'Obtém Fornecedor
  If Combo_Fornecedor.Text <> "" Then
    Call IsDataType(dtLong, Combo_Fornecedor.Text, lngFornecedor)
  End If
  
  'Tipo
  If O_Normal.Value Then
    strTipo = "N"
  ElseIf O_Grade.Value Then
    strTipo = "G"
  Else
    strTipo = "E"
  End If
  
  
  'Limpa tabela temporária
  dbTemp.Execute "DELETE FROM Comprar;", dbFailOnError
  
  
  'Índices
  rsProdutos.Index = "Código"
  rsForn_Prod.Index = "Produto"
  rsEstoque.Index = "Produto"
  rsTempo.Index = "Produto"
  rsClasses.Index = "Código"
  rsSub_Classes.Index = "Código"
  rsTamanhos.Index = "Código"
  rsCores.Index = "Código"
  rsEdicoes.Index = "Produto"
  rsCliFor.Index = "Código"
  
  'Pesquisa produtos
  strCodigoProduto = ""
  Do
    rsProdutos.Seek ">", strCodigoProduto
    If rsProdutos.NoMatch Then
      Exit Do
    Else
      strCodigoProduto = rsProdutos.Fields("Código").Value
    
      'Flag para compra do produto (padrão comprar)
      blnComprar = True
      
      'Valida Fornecedor específico
      If lngFornecedor > 0 Then
        rsForn_Prod.Seek "=", strCodigoProduto, lngFornecedor
        blnComprar = Not rsForn_Prod.NoMatch
      End If
      
      'Valida Ativo, Tipo selecionado, Controla Estoque e
      'Estoque Ideal maior que zero
      If blnComprar Then
        blnComprar = Not rsProdutos.Fields("Desativado").Value And _
          rsProdutos.Fields("Tipo").Value = strTipo And _
          rsProdutos.Fields("Estoque").Value And _
          rsProdutos.Fields("Estoque Ideal").Value > 0
      End If
      
      'Valida Estoque e insere registros na tabela temporária
      If blnComprar Then
        
        Select Case strTipo
          Case "N"
            dblEstoqueAtual = 0
            With rsEstoque
              .Seek "=", intFilial, strCodigoProduto, 0, 0, 0
              If Not rsEstoque.NoMatch Then
                dblEstoqueAtual = .Fields("Estoque Atual").Value
                blnComprar = (dblEstoqueAtual < rsProdutos.Fields("Estoque Mínimo").Value)
              End If
            End With
            
            If blnComprar Then
              Call InserirRegistroTempData(0, 0, 0, dblEstoqueAtual)
            End If
            
          Case "G"
            strSQL = "SELECT Código "
            strSQL = strSQL & "FROM [Códigos da Grade] "
            strSQL = strSQL & "WHERE [Código Original] = '" & strCodigoProduto & "' "
            strSQL = strSQL & "ORDER BY Código"
            
            Set rstGrade = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
            With rstGrade
              Do Until .EOF
                
                strGrade = rstGrade.Fields("Código").Value
                intTamanho = CInt(Mid(strGrade, Len(strGrade) - 5, 3))
                intCor = CInt(Right(strGrade, 3))
                
                dblEstoqueAtual = 0
                With rsEstoque
                  .Seek "=", intFilial, strCodigoProduto, intTamanho, intCor, 0
                  If Not rsEstoque.NoMatch Then
                    dblEstoqueAtual = .Fields("Estoque Atual").Value
                    blnComprar = (dblEstoqueAtual < rsProdutos.Fields("Estoque Mínimo").Value)
                  End If
                End With
                
                If blnComprar Then
                  Call InserirRegistroTempData(intTamanho, intCor, 0, dblEstoqueAtual)
                End If
                
                .MoveNext
              Loop
              .Close
            End With
            Set rstGrade = Nothing
            
          Case "E"
            strSQL = "SELECT Código "
            strSQL = strSQL & "FROM Edições "
            strSQL = strSQL & "WHERE Produto = '" & strCodigoProduto & "' "
            strSQL = strSQL & "ORDER BY Código"
            
            Set rstGrade = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
            With rstGrade
              Do Until .EOF
                
                lngEdicao = CLng(rstGrade.Fields("Código").Value)
                
                dblEstoqueAtual = 0
                With rsEstoque
                  .Seek "=", intFilial, strCodigoProduto, 0, 0, lngEdicao
                  If Not rsEstoque.NoMatch Then
                    dblEstoqueAtual = .Fields("Estoque Atual").Value
                    blnComprar = (dblEstoqueAtual < rsProdutos.Fields("Estoque Mínimo").Value)
                  End If
                End With
                
                If blnComprar Then
                  Call InserirRegistroTempData(0, 0, lngEdicao, dblEstoqueAtual)
                End If
                
                .MoveNext
              Loop
              .Close
            End With
            Set rstGrade = Nothing
            
        End Select
      End If
    End If
  Loop
  
  Exit Sub
  
ErrHandler:
  'Fecha tabela
  If Not rstGrade Is Nothing Then
    rstGrade.Close
    Set rstGrade = Nothing
  End If
  
  'Repassa erro
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

Private Sub InserirRegistroTempData(ByVal intTamanho As Integer, _
  ByVal intCor As Integer, ByVal lngEdicao As Long, ByVal dblEstoqueAtual As Double)
  
  Dim strSQL As String
  Dim rstFornecedores As Recordset
  Dim lngFornecedor As Long
  Dim intX As Integer

  
  On Error GoTo ErrHandler
  
  
  With rsTempo
    .AddNew
    .Fields("Código").Value = rsProdutos.Fields("Código").Value
    .Fields("Código Ordenação").Value = rsProdutos.Fields("Código Ordenação").Value
    .Fields("Nome").Value = rsProdutos.Fields("Nome").Value
    .Fields("Unidade Venda").Value = rsProdutos.Fields("Unidade Venda").Value
    .Fields("Último Custo").Value = rsProdutos.Fields("Último Custo").Value
    .Fields("Fracionado").Value = rsProdutos.Fields("Fracionado").Value
    
    With rsClasses
      .Seek "=", rsProdutos.Fields("Classe").Value
      If Not .NoMatch Then
        rsTempo.Fields("Classe").Value = rsProdutos.Fields("Classe").Value
        rsTempo.Fields("Nome Classe").Value = .Fields("Nome").Value
      End If
    End With
    
    With rsSub_Classes
      .Seek "=", rsProdutos.Fields("Sub Classe").Value
      If Not .NoMatch Then
        rsTempo.Fields("Sub Classe").Value = rsProdutos.Fields("Sub Classe").Value
        rsTempo.Fields("Nome Sub").Value = .Fields("Nome").Value
      End If
    End With
    
    Select Case rsProdutos.Fields("Tipo").Value
      Case "G"
        With rsTamanhos
          .Seek "=", intTamanho
          If Not .NoMatch Then
            rsTempo.Fields("Tamanho").Value = intTamanho
            rsTempo.Fields("Nome Tamanho").Value = .Fields("Nome").Value
          End If
        End With
        
        With rsCores
          .Seek "=", intCor
          If Not .NoMatch Then
            rsTempo.Fields("Cor").Value = intCor
            rsTempo.Fields("Nome Cor").Value = .Fields("Nome").Value
          End If
        End With
        
      Case "E"
        With rsEdicoes
          .Seek "=", rsProdutos.Fields("Código").Value, lngEdicao
          If Not .NoMatch Then
            rsTempo.Fields("Edição").Value = lngEdicao
            rsTempo.Fields("Nome Edição").Value = .Fields("Nome").Value
          End If
        End With
        
    End Select
    
    .Fields("Estoque").Value = dblEstoqueAtual
    .Fields("Ideal").Value = rsProdutos.Fields("Estoque Ideal").Value
    
    'Fornecedores
    If O_Fornecedor.Value = vbChecked Then
      
      strSQL = "SELECT Fornecedor "
      strSQL = strSQL & "FROM Forn_Prod "
      strSQL = strSQL & "WHERE Produto = '" & rsProdutos.Fields("Código").Value & "' "
      strSQL = strSQL & "ORDER BY Fornecedor"
      
      Set rstFornecedores = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      With rstFornecedores
        For intX = 1 To 5
          If Not .EOF Then
            lngFornecedor = .Fields("Fornecedor").Value
            
            With rsCliFor
              .Seek "=", lngFornecedor
              If Not .NoMatch Then
                rsTempo("Fornece" & intX) = lngFornecedor
                rsTempo("Nome" & intX) = .Fields("Nome").Value
                rsTempo("Tel" & intX & "_1") = .Fields("Fone 1").Value
                rsTempo("Tel" & intX & "_2") = .Fields("Fone 2").Value
                rsTempo("Fax" & intX) = .Fields("Fax").Value
              End If
            End With
            
            .MoveNext
          End If
        Next intX
        .Close
      End With
      Set rstFornecedores = Nothing
    End If
    
    .Update
  End With
  
  Exit Sub
  
ErrHandler:
  'Fecha tabela
  If Not rstFornecedores Is Nothing Then
    rstFornecedores.Close
    Set rstFornecedores = Nothing
  End If
  
  'Repassa erro
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

Private Sub cmdImprimir_Click()

  On Error GoTo ErrHandler
  
  
  'Verifica Filial
  If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
    DisplayMsg "Escolha a Filial."
    Combo.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(Combo.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If
  
  'Status
  Screen.MousePointer = vbHourglass
  cmdImprimir.Enabled = False
  Call StatusMsg("Obtendo informações...")

  'Preenche a tabela temporária
  Call FillTempData
  
  'Status
  Call StatusMsg("Imprimindo relatório...")
  
  'Prepara Relatório
  With Rel
    .Reset
    
    'Seta localização das bases de dados e relatório
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsTempDBFileName
    If O_Normal.Value Then
      If O_Classe.Value = vbChecked Then
        .ReportFileName = gsReportPath & "COMPRAR2.RPT"
      Else
        .ReportFileName = gsReportPath & "COMPRAR1.RPT"
      End If
      .DataFiles(2) = gsQuickDBFileName
    ElseIf O_Grade.Value Then
      If O_Classe.Value = vbChecked Then
        .ReportFileName = gsReportPath & "COMPRAR2G.RPT"
      Else
        .ReportFileName = gsReportPath & "COMPRAR1G.RPT"
      End If
    ElseIf O_Edição.Value Then
      If O_Classe.Value = vbChecked Then
        .ReportFileName = gsReportPath & "COMPRAR2E.RPT"
      Else
        .ReportFileName = gsReportPath & "COMPRAR1E.RPT"
      End If
    End If
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel
    
    'Saída
    If B_Vídeo.Value Then
      .Destination = crptToWindow
    Else
      .Destination = crptToPrinter
    End If
    
    'Fórmulas
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
    .Formulas(1) = "nome_filial = '" & Nome_Empresa.Caption & "'"
    If Nome_Fornecedor.Caption = "" Then
      .Formulas(2) = "nome_fornecedor = 'Todos'"
    Else
      .Formulas(2) = "nome_fornecedor = '" & Combo_Fornecedor.Text & " - " & Replace(Nome_Fornecedor.Caption, "'", " ") & "'"
    End If
    If O_Normal.Value Then
      If O_Fornecedor.Value = 1 Then
        .Formulas(3) = "Fornecedor = 'SIM'"
      Else
        .Formulas(3) = "Fornecedor = 'NÃO'"
      End If
    End If
    
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel)

   'Exibe relatório
    .Action = 1
  End With

  'Status
  Call StatusMsg("Pronto")
  cmdImprimir.Enabled = True
  Screen.MousePointer = vbDefault

  Exit Sub
  
ErrHandler:
  'Status
  Call StatusMsg("Erro")
  cmdImprimir.Enabled = True
  Screen.MousePointer = vbDefault
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(1).Text
  Combo_LostFocus
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
  Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_GotFocus()
  Call StatusMsg(LoadResString(50))
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
  Nome_Fornecedor.Caption = rsCliFor("Nome")
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

  On Error GoTo ErrHandler
  
  
  Call CenterForm(Me)
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsTempo = dbTemp.OpenRecordset("Comprar")
  Set rsEstoque = db.OpenRecordset("Estoque Final", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsForn_Prod = db.OpenRecordset("Forn_Prod", , dbReadOnly)
  Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsEdicoes = db.OpenRecordset("Edições", , dbReadOnly)
  
  Combo.Text = gnCodFilial
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  O_Grade.Enabled = gbGrade
  O_Edição.Enabled = gbEdicao

  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Unload(Cancel As Integer)

  On Error GoTo ErrHandler
  
  rsParametros.Close
  rsTempo.Close
  rsEstoque.Close
  rsProdutos.Close
  rsCliFor.Close
  rsForn_Prod.Close
  rsTamanhos.Close
  rsCores.Close
  rsClasses.Close
  rsSub_Classes.Close
  rsEdicoes.Close
  
  Set rsParametros = Nothing
  Set rsTempo = Nothing
  Set rsEstoque = Nothing
  Set rsProdutos = Nothing
  Set rsCliFor = Nothing
  Set rsForn_Prod = Nothing
  Set rsTamanhos = Nothing
  Set rsCores = Nothing
  Set rsClasses = Nothing
  Set rsSub_Classes = Nothing
  Set rsEdicoes = Nothing

  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub O_Edição_Click()
  O_Fornecedor.Enabled = False
  O_Fornecedor.Value = 0
End Sub

Private Sub O_Grade_Click()
  O_Fornecedor.Enabled = False
  O_Fornecedor.Value = 0
End Sub

Private Sub O_Normal_Click()
  O_Fornecedor.Enabled = True
End Sub
