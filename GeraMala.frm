VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmGeraMala 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Preparação do Arquivo de Mala Direta"
   ClientHeight    =   7245
   ClientLeft      =   2760
   ClientTop       =   1305
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1240
   Icon            =   "GeraMala.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7245
   ScaleWidth      =   10110
   Begin VB.Frame fraGrupos 
      Caption         =   "Grupos de Classificação de Clientes"
      Height          =   615
      Left            =   5100
      TabIndex        =   41
      ToolTipText     =   $"GeraMala.frx":4E95A
      Top             =   4500
      Width           =   4935
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "..."
         Height          =   300
         Left            =   3000
         TabIndex        =   26
         ToolTipText     =   "Click aqui para Consultar a Classificação de Clientes."
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optG4 
         Caption         =   "4"
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton optG3 
         Caption         =   "3"
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optG2 
         Caption         =   "2"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optG1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Preceder nome de contatos por"
      Height          =   645
      Left            =   60
      TabIndex        =   40
      Top             =   5970
      Width           =   4935
      Begin VB.TextBox Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2610
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "Ao Sr(a):"
         Top             =   210
         Width           =   1875
      End
   End
   Begin VB.Data datGrupos 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8550
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Grupos Interesse"
      Top             =   5700
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Frame Frame7 
      Caption         =   "Clientes sem Compra"
      Height          =   675
      Left            =   60
      TabIndex        =   38
      Top             =   3810
      Width           =   9975
      Begin MSMask.MaskEdBox Sem_Compra 
         Height          =   315
         Left            =   5850
         TabIndex        =   21
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Ao analisar clientes, selecionar somente os que não compram desde "
         Height          =   255
         Left            =   570
         TabIndex        =   39
         Top             =   300
         Width           =   5265
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ativos / Inativos"
      Height          =   615
      Left            =   60
      TabIndex        =   37
      Top             =   4500
      Width           =   4935
      Begin VB.OptionButton O_Ativ_Inativ 
         Appearance      =   0  'Flat
         Caption         =   "Inativos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton O_Ativ_Ativ 
         Appearance      =   0  'Flat
         Caption         =   "Ativos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1890
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton O_Ativ_Todos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Estado"
      Height          =   825
      Left            =   5100
      TabIndex        =   35
      Top             =   5130
      Width           =   4935
      Begin VB.TextBox Estado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3750
         MaxLength       =   2
         TabIndex        =   16
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Caso deseje clientes de somente um estado"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   300
         Width           =   3435
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cidade"
      Height          =   855
      Left            =   60
      TabIndex        =   33
      Top             =   5100
      Width           =   4935
      Begin VB.TextBox Cidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1110
         MaxLength       =   30
         TabIndex        =   15
         Top             =   420
         Width           =   3405
      End
      Begin VB.Label Label2 
         Caption         =   "Preencha o campo caso deseje clientes de somente uma cidade"
         Height          =   615
         Left            =   270
         TabIndex        =   34
         Top             =   180
         Width           =   4395
      End
   End
   Begin VB.Frame Quadro_Contato 
      Caption         =   "Contatos"
      Height          =   945
      Left            =   60
      TabIndex        =   31
      Top             =   1650
      Width           =   9975
      Begin VB.OptionButton Sem_Contato 
         Caption         =   "Etiquetas sem nome do contato"
         Height          =   255
         Left            =   6960
         TabIndex        =   14
         Top             =   420
         Width           =   2775
      End
      Begin VB.OptionButton Vários_Contatos 
         Caption         =   "Uma etiqueta para cada contato"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   420
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton Um_Contato 
         Caption         =   "Somente uma etiqueta por empresa/pessoa selecionada. Imprime o primeiro contato."
         Height          =   615
         Left            =   540
         TabIndex        =   13
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton B_Começa 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Preparar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6690
      Width           =   9975
   End
   Begin VB.CheckBox Limpa_Arquivo 
      Appearance      =   0  'Flat
      Caption         =   "Limpar o arquivo de mala direta"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5130
      TabIndex        =   32
      Top             =   6150
      Value           =   1  'Checked
      Width           =   3885
   End
   Begin VB.Frame Frame3 
      Caption         =   "Aniversário"
      Height          =   855
      Left            =   60
      TabIndex        =   30
      Top             =   780
      Width           =   9975
      Begin VB.ComboBox Combo_Mês 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "GeraMala.frx":4EA00
         Left            =   6990
         List            =   "GeraMala.frx":4EA28
         TabIndex        =   11
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton O_Com_Data 
         Appearance      =   0  'Flat
         Caption         =   "Com aniversário no mês de"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4500
         TabIndex        =   10
         Top             =   323
         Width           =   2385
      End
      Begin VB.OptionButton O_Sem_Data 
         Appearance      =   0  'Flat
         Caption         =   "Independe de data de aniversário"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   540
         TabIndex        =   9
         Top             =   323
         Value           =   -1  'True
         Width           =   3045
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Grupos de Interesse"
      Height          =   1185
      Left            =   60
      TabIndex        =   29
      Top             =   2610
      Width           =   9975
      Begin SSDataWidgets_B.SSDBCombo cmbGrupos 
         Bindings        =   "GeraMala.frx":4EA91
         Height          =   345
         Left            =   540
         TabIndex        =   7
         Top             =   690
         Width           =   2565
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
         Columns(0).Width=   1879
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Código"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Código"
         Columns(0).DataType=   3
         Columns(0).FieldLen=   256
         Columns(1).Width=   6668
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4524
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.OptionButton O_Sem_Grupo 
         Appearance      =   0  'Flat
         Caption         =   "Todos os grupos - mesmo que o cliente/fornecedor não tenha nenhum grupo cadastrado."
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   6960
         TabIndex        =   5
         Top             =   225
         Value           =   -1  'True
         Width           =   2925
      End
      Begin VB.OptionButton O_Todos_Grupos 
         Appearance      =   0  'Flat
         Caption         =   "Todos os grupos - somente se houver ao menos um grupo cadastrado para o cliente/fornecedor."
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   3870
         TabIndex        =   8
         Top             =   150
         Width           =   3105
      End
      Begin VB.OptionButton O_Grupo 
         Appearance      =   0  'Flat
         Caption         =   "Somente um grupo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   510
         TabIndex        =   6
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção de Tipo"
      Height          =   705
      Left            =   60
      TabIndex        =   28
      Top             =   60
      Width           =   9975
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8550
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton O_Outros 
         Appearance      =   0  'Flat
         Caption         =   "Outros"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6696
         TabIndex        =   4
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton O_Revendedor 
         Appearance      =   0  'Flat
         Caption         =   "Revendedor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4484
         TabIndex        =   3
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton O_Fornecedor 
         Appearance      =   0  'Flat
         Caption         =   "Fornecedor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2392
         TabIndex        =   2
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton O_Cliente 
         Appearance      =   0  'Flat
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   540
         TabIndex        =   1
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmGeraMala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrupos As Recordset
Dim rsMala_Direta As Recordset
Dim rsMala_Tempo As Recordset
Dim rsCliFor As Recordset
Dim rsContatos As Recordset
Private gnGrupo As Integer
'28/07/2004 - Daniel
'Personalização para a TV Shopping
Private m_blnTV As Boolean

Private Sub B_Começa_Click()
  Dim Cliente As Long
  Dim Ordem As Long
  Dim Mensa As String
  Dim Código As Long
  Dim Mês As String
  Dim Contato As String
  Dim Erro As Integer
  Dim Aux_Str As String
  Dim Str_Cidade As String
  Dim Str_Estado As String
  Dim Dia As Variant
  Dim sSql As String

  On Error GoTo Processa_Erro

  Dia = "Falso"
  If IsDate(Sem_Compra.Text) Then Dia = Sem_Compra.Text

  If O_Grupo.Value = True Then
    If gnGrupo = -1 Or Trim(cmbGrupos.Text) = "" Then
      DisplayMsg "Escolha o Grupo de Interesse."
      cmbGrupos.SetFocus
      Exit Sub
    End If
  End If
  
  If O_Com_Data.Value = True Then
    If Combo_Mês.Text = "" Then
      DisplayMsg "Escolha o mês de aniversário."
      Combo_Mês.SetFocus
      Exit Sub
    End If
  End If
   
   
  If IsNull(Cidade.Text) Then
    Str_Cidade = ""
  Else
    Str_Cidade = Cidade.Text
  End If
  
  If IsNull(Estado.Text) Then
    Str_Estado = ""
  Else
    Str_Estado = Estado.Text
  End If
   

   
  If Limpa_Arquivo.Value = 1 Then  'limpa temporário
    sSql = "Delete * From [Mala Direta - Tempo]"
    Call StatusMsg("Aguarde, apagando informações...")
    db.Execute sSql
  End If
    
  Call StatusMsg("")
  
  rsCliFor.Index = "Código"
  rsMala_Direta.Index = "Cliente"

  Código = 0
  
Lp1:
  rsCliFor.Seek ">", Código
  If rsCliFor.NoMatch Then GoTo Fim
  
  Call StatusMsg("Aguarde, verificando cliente " + str(Código))
  
     
  
  Código = rsCliFor("Código")
  If rsCliFor("Sem Mala Direta") = True Then GoTo Lp1
  
  If Str_Cidade <> "" Then
    If IsNull(rsCliFor("Cidade")) Then GoTo Lp1
    If rsCliFor("Cidade") <> Str_Cidade Then GoTo Lp1
  End If
  
  If Str_Estado <> "" Then
    If rsCliFor("Estado") <> Str_Estado Then GoTo Lp1
  End If
  
  
  If O_Cliente.Value = True Then
    If rsCliFor("Tipo") <> "C" Then GoTo Lp1
  End If
  
  If O_Fornecedor.Value = True Then
    If rsCliFor("Tipo") <> "F" Then GoTo Lp1
  End If
  
  If O_Revendedor.Value = True Then
    If rsCliFor("Tipo") <> "R" Then GoTo Lp1
  End If
  
  If O_Outros.Value = True Then
    If rsCliFor("Tipo") <> "O" Then GoTo Lp1
  End If
  
  If rsCliFor("Tipo") = "C" Then
    If Dia <> "Falso" Then
      If rsCliFor("Última Compra") >= CDate(Dia) Then GoTo Lp1
    End If
  End If
    
  
  
  If O_Grupo.Value = True Then
    If gnGrupo <> 0 Then
      rsMala_Direta.Seek "=", rsCliFor("Código"), gnGrupo
      If rsMala_Direta.NoMatch Then GoTo Lp1
    End If
  End If
  
  If O_Todos_Grupos.Value = True Then
    rsMala_Direta.Seek ">", rsCliFor("Código"), 0
    If rsMala_Direta.NoMatch Then GoTo Lp1
    If rsMala_Direta("Cliente") <> rsCliFor("Código") Then GoTo Lp1
  End If
   
  If O_Ativ_Ativ.Value = True Then
    If rsCliFor("Inativo") = True Then GoTo Lp1
  End If
  
  If O_Ativ_Inativ.Value = True Then
    If rsCliFor("Inativo") = False Then GoTo Lp1
  End If
  
  '28/07/2004 - Daniel
  'Personalização TV Shopping Brasil
  'Critério de Classificação de Clientes
  If m_blnTV And O_Cliente.Value Then
    If optG1.Value Then
      If rsCliFor("CodGrupo") <> 1 Then GoTo Lp1
    End If
    
    If optG2.Value Then
      If rsCliFor("CodGrupo") <> 2 Then GoTo Lp1
    End If
    
    If optG3.Value Then
      If rsCliFor("CodGrupo") <> 3 Then GoTo Lp1
    End If
    
    If optG4.Value Then
      If rsCliFor("CodGrupo") <> 4 Then GoTo Lp1
    End If
  End If
  '---------------------------------------------
  
  rsContatos.Index = "Cliente"
  rsContatos.Seek ">", Código, 0
  Contato = ""
  If Not rsContatos.NoMatch Then
    If Código = rsContatos("Cliente") Then Contato = rsContatos("Contato") & ""
  End If
  
  If O_Sem_Data.Value = True Then
    If Sem_Contato.Value = True Or Um_Contato.Value = True Then
      If Sem_Contato.Value = True Then Contato = ""
      rsMala_Tempo.AddNew
        rsMala_Tempo("Cliente") = Código
        Aux_Str = Nome.Text + " " + Contato
        Aux_Str = Left$(Aux_Str, 30)
        rsMala_Tempo("Nome") = Aux_Str
        If Contato = "" Then rsMala_Tempo("Nome") = ""
      rsMala_Tempo.Update
      GoTo Lp1
    End If
    
    'Faz aqui se tiver vários contatos sem data de aniversário
    Erro = False
    Ordem = 0
    rsContatos.Index = "Cliente"
    Do
      rsContatos.Seek ">", Código, Ordem
      If rsContatos.NoMatch Then Erro = True
      If Erro = False Then If IsNull(rsContatos("Contato")) Then Erro = True
      If Erro = False Then If rsContatos("Cliente") <> Código Then Erro = True
      If Erro = False Then
        Ordem = rsContatos("Seqüência")
        Contato = rsContatos("Contato")
        rsMala_Tempo.AddNew
          rsMala_Tempo("Cliente") = Código
          Aux_Str = Nome.Text + " " + Contato
          Aux_Str = Left$(Aux_Str, 30)
          rsMala_Tempo("Nome") = Aux_Str
          If Contato = "" Then rsMala_Tempo("Nome") = ""
        rsMala_Tempo.Update
      End If
    Loop Until Erro = True
    GoTo Lp1
  End If
  
  Rem Procura por data de aniversário
  If Combo_Mês.Text = "Janeiro" Then Mês = "JAN"
  If Combo_Mês.Text = "Fevereiro" Then Mês = "FEV"
  If Combo_Mês.Text = "Março" Then Mês = "MAR"
  If Combo_Mês.Text = "Abril" Then Mês = "ABR"
  If Combo_Mês.Text = "Maio" Then Mês = "MAI"
  If Combo_Mês.Text = "Junho" Then Mês = "JUN"
  If Combo_Mês.Text = "Julho" Then Mês = "JUL"
  If Combo_Mês.Text = "Agosto" Then Mês = "AGO"
  If Combo_Mês.Text = "Setembro" Then Mês = "SET"
  If Combo_Mês.Text = "Outubro" Then Mês = "OUT"
  If Combo_Mês.Text = "Novembro" Then Mês = "NOV"
  If Combo_Mês.Text = "Dezembro" Then Mês = "DEZ"

  Erro = False
  Ordem = 0
  rsContatos.Index = "Cliente"
  Do
    rsContatos.Seek ">", Código, Ordem
    If rsContatos.NoMatch Then Erro = True
    If Erro = False Then If IsNull(rsContatos("Contato")) Then Erro = True
    If Erro = False Then If rsContatos("Cliente") <> Código Then Erro = True
    If Erro = False Then
      Ordem = rsContatos("Seqüência")
      If rsContatos("Mês Aniversário") = Mês Then
      Contato = rsContatos("Contato")
        If Um_Contato.Value = True Then
          rsMala_Tempo.AddNew
            rsMala_Tempo("Cliente") = Código
            Aux_Str = Nome.Text + " " + Contato
            Aux_Str = Left$(Aux_Str, 30)
            rsMala_Tempo("Nome") = Aux_Str
            If Contato = "" Then rsMala_Tempo("Nome") = ""
          rsMala_Tempo.Update
          GoTo Lp1
        End If
          
        If Sem_Contato.Value = True Then
          rsMala_Tempo.AddNew
            rsMala_Tempo("Cliente") = Código
            rsMala_Tempo("Nome") = ""
          rsMala_Tempo.Update
          GoTo Lp1
        End If
      
       rsMala_Tempo.AddNew
         rsMala_Tempo("Cliente") = Código
         Aux_Str = Nome.Text + " " + Contato
         Aux_Str = Left$(Aux_Str, 30)
         rsMala_Tempo("Nome") = Aux_Str
         If Contato = "" Then rsMala_Tempo("Nome") = ""
       rsMala_Tempo.Update
      
      End If 'If rsContatos("Mês Aniversário") = Mês Then
    End If
  Loop Until Erro = True
  GoTo Lp1

      
Fim:

 DisplayMsg "Mala Direta preparada."
  
  
  Exit Sub
Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Preparar Etiquetas de Mala Direta")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Sub
    Case 3 'Encerrar
      End
  End Select
  
End Sub

Private Sub cmbGrupos_Click()
  gnGrupo = cmbGrupos.Columns(0).Text
End Sub

Private Sub cmbGrupos_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not cmbGrupos.DroppedDown And KeyCode = vbKeyDelete Then
    KeyCode = 0
  End If
End Sub

Private Sub cmbGrupos_KeyPress(KeyAscii As Integer)
  If Not cmbGrupos.DroppedDown Then
    KeyAscii = 0
  End If
End Sub

Private Sub cmdVerificar_Click()
  'Primeiro é validado se o usuário que está clicando é responsável
  'ou não pelo Martketing da empresa
  Dim rstFuncionarios As Recordset
  Dim strQuery        As String
  
  strQuery = "SELECT Código, Marketing "
  strQuery = strQuery & " FROM Funcionários "
  strQuery = strQuery & " WHERE Código = " & gnUserCode
  
  Set rstFuncionarios = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      If .Fields("Marketing").Value Then
        frmClassificacaoClientes.Show
      Else
        MsgBox "Usuário não responsável pelo Marketing.", vbExclamation, "Atenção"
      End If
    End If
    .Close
  End With

  Set rstFuncionarios = Nothing

End Sub

Private Sub Estado_LostFocus()
  If IsNull(Estado.Text) Then Exit Sub
  Estado.Text = UCase(Estado.Text)
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsGrupos = db.OpenRecordset("Grupos Interesse", , dbReadOnly)
  Set rsMala_Direta = db.OpenRecordset("Mala Direta", , dbReadOnly)
  Set rsMala_Tempo = db.OpenRecordset("Mala Direta - Tempo")
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsContatos = db.OpenRecordset("Contatos", , dbReadOnly)
  
  datGrupos.DatabaseName = gsQuickDBFileName
  gnGrupo = -1
  
  '28/07/2004 - Daniel
  'Adicionado personalização para a TV Shopping
  'Projeto Classificação de Clientes
  If CheckSerialCaseMod("QS39945-043", "QS39944-959", "QS40449-276") Then
    m_blnTV = True
    
    fraGrupos.Visible = True
  Else
    fraGrupos.Visible = False
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsGrupos.Close
  rsMala_Direta.Close
  rsMala_Tempo.Close
  rsCliFor.Close
  rsContatos.Close
  Set rsGrupos = Nothing
  Set rsMala_Direta = Nothing
  Set rsMala_Tempo = Nothing
  Set rsCliFor = Nothing
  Set rsContatos = Nothing
End Sub

Private Sub O_Com_Data_Click()
  Combo_Mês.Enabled = True
  Vários_Contatos.Enabled = True
End Sub

Private Sub O_Grupo_Click()
   cmbGrupos.Enabled = True
End Sub

Private Sub O_Sem_Data_Click()
 Combo_Mês.Enabled = False
 
  'Vários_Contatos.Enabled = False
  'Sem_Contato.Value = True
End Sub

Private Sub O_Sem_Grupo_Click()
  cmbGrupos.Enabled = False
End Sub

Private Sub O_Todos_Grupos_Click()
  cmbGrupos.Enabled = False
End Sub

Private Sub Sem_Compra_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Sem_Compra.Text = frmCalendario.gsDateCalender(Sem_Compra.Text)
  End Select
End Sub

Private Sub Sem_Compra_LostFocus()
  Sem_Compra.Text = Ajusta_Data(Sem_Compra.Text)
End Sub
