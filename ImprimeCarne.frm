VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmImprimeCarnes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Impressão de Carnês"
   ClientHeight    =   4215
   ClientLeft      =   3885
   ClientTop       =   2460
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1650
   Icon            =   "ImprimeCarne.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4215
   ScaleWidth      =   8790
   Begin VB.CommandButton cmd_OutraTelaCarne 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "1010101010101010101"
      BeginProperty Font 
         Name            =   "CIA Code 39 Medium Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3660
      Width           =   8610
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   8610
   End
   Begin VB.Frame Frame3 
      Caption         =   "Período"
      Height          =   795
      Left            =   75
      TabIndex        =   17
      Top             =   885
      Width           =   4845
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3330
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   195
         TabIndex        =   19
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   330
         Width           =   795
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
      Left            =   1950
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
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
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4020
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir"
      Height          =   465
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2580
      Width           =   8610
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imprimir"
      Height          =   795
      Left            =   75
      TabIndex        =   14
      Top             =   1710
      Width           =   8610
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton O_Impresso 
         Appearance      =   0  'Flat
         Caption         =   "Somente os &já impressos (reimpressão)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5070
         TabIndex        =   7
         Top             =   330
         Width           =   3135
      End
      Begin VB.OptionButton O_N_Impresso 
         Appearance      =   0  'Flat
         Caption         =   "Somente os &não impressos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2130
         TabIndex        =   6
         Top             =   330
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Emissão por"
      Height          =   795
      Left            =   4980
      TabIndex        =   13
      Top             =   870
      Width           =   3705
      Begin VB.OptionButton O_Vencimento 
         Appearance      =   0  'Flat
         Caption         =   "Data de &Vencimento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton O_Emissão 
         Appearance      =   0  'Flat
         Caption         =   "Data de &Emissão"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   300
         Width           =   1515
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "ImprimeCarne.frx":4E95A
      DataSource      =   "Data1"
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   75
      Width           =   1185
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
      Columns(0).Width=   7805
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1244
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2090
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "ImprimeCarne.frx":4E96E
      DataSource      =   "Data2"
      Height          =   345
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "Use 0 para todos os clientes"
      Top             =   480
      Width           =   1185
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
      Columns(0).Width=   8996
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1720
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2090
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Nome_Empresa 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   2250
      TabIndex        =   16
      Top             =   75
      Width           =   6435
   End
   Begin VB.Label Label6 
      Caption         =   "Filial"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Nome_Cliente 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   2250
      TabIndex        =   12
      Top             =   480
      Width           =   6435
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   525
      Width           =   615
   End
End
Attribute VB_Name = "frmImprimeCarnes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsClientes As Recordset
Dim rsReceber As Recordset

Sub Imprime_Boleto1()
  Dim Aux As Variant
  Dim Nome_Arq As String
  Dim Texto As String
  Dim Final As Integer
  Dim Str_Impre As String
  Dim Cód1, Cód2, Cód3 As Integer
  Dim Num_cod As Integer
  Dim Resposta As Long
  Dim Final_Linha As Integer
  Dim Linhas As Integer
  Dim Especial2 As Integer
  
  Call StatusMsg("")
  
  Rem Inicializa variáveis nota
  Limpa_Variáveis_Boleto
  
  
  Glob_Nota_Impressa = rsReceber("Nota")
  Glob_Nome = rsClientes("Nome") & ""
  Glob_Fantasia = rsClientes("Fantasia") & ""
  Glob_CGC = rsClientes("CGC") & ""
  Glob_Inscrição = rsClientes("Inscrição") & ""
  Glob_Data_Emissão = rsReceber("Data Emissão") & ""
  
  If IsNull(rsClientes("Endereço Cob")) Or rsClientes("Endereço Cob") = "" Then
    Glob_Endereço = rsClientes("Endereço") & ""
    Glob_NumeroEndereco = rsClientes.Fields("Endereço Número").Value & "" '23/10/2009 - mpdea
    
    Glob_Complemento = rsClientes("Complemento") & ""
    Glob_Bairro = rsClientes("Bairro") & ""
    Glob_CEP = rsClientes("Cep") & ""
    Glob_Cidade = rsClientes("Cidade") & ""
    Glob_Estado = rsClientes("Estado") & ""
  Else
    Glob_Endereço = rsClientes("Endereço Cob") & ""
    Glob_Complemento = rsClientes("Complemento Cob") & ""
    Glob_Bairro = rsClientes("Bairro Cob") & ""
    Glob_CEP = rsClientes("Cep Cob") & ""
    Glob_Cidade = rsClientes("Cidade Cob") & ""
    Glob_Estado = rsClientes("Estado Cob") & ""
  End If

  Glob_Data_Saída = str(Date)
  Glob_Fatura = rsReceber("Fatura") & ""
  Glob_Descrição = rsReceber("Descrição") & ""
  Glob_Vencimento = rsReceber("Vencimento") & ""
  Glob_Valor = rsReceber("Valor")
  Glob_Desconto = rsReceber("Desconto")
  Glob_Acréscimo = rsReceber("Acréscimo")
  Glob_Mensagem_Cli = rsClientes("Mensagem Boleto") & ""
  gsObsDoc(0) = gsObsDoc(0) & ""
  gsObsDoc(1) = gsObsDoc(1) & ""
  gsObsDoc(2) = gsObsDoc(2) & ""
  Glob_Código_Cli = rsClientes("Código")
  Glob_Sequência = rsReceber("Sequência")
  
  Rem Pegar o nome do arquivo de configuração
  Nome_Arq = gsConfigPath & gsDocFileName & ".CBB"
   
  
  Rem Comprime a impressora
  Num_cod = 0
  Str_Impre = ""
  If rsParametros("Boleto Comprimido") = True Then
     If rsParametros("Cód Comp 1") <> "" Then
       Num_cod = 1
       If rsParametros("Cód Comp 2") <> "" Then
         Num_cod = 2
         If rsParametros("Cód Comp 3") <> "" Then
           Num_cod = 3
         End If
       End If
     End If
     If Num_cod = 1 Then
       Str_Impre = Chr$(Val(rsParametros("Cód Comp 1")))
     End If
     If Num_cod = 2 Then
       Str_Impre = Chr$(Val(rsParametros("Cód Comp 1")))
       Str_Impre = Str_Impre + Chr$(Val(rsParametros("Cód Comp 2")))
     End If
     If Num_cod = 3 Then
       Str_Impre = Chr$(Val(rsParametros("Cód Comp 1")))
       Str_Impre = Str_Impre + Chr$(Val(rsParametros("Cód Comp 2")))
       Str_Impre = Str_Impre + Chr$(Val(rsParametros("Cód Comp 3")))
     End If
     If Str_Impre <> "" Then
        Str_Impre = Chr$(Len(Str_Impre) Mod 256) + Chr$(Len(Str_Impre) \ 256) + Str_Impre
        Printer.Print ""
        If Not IsWindowsNT() Then
          Resposta = Escape(Printer.hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
        Else
          Resposta = Escape32(Printer.hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
        End If
        If Resposta <= 0 Then
          DisplayMsg "Não foi possível comprimir a impressora."
          Exit Sub
        End If
     End If
  End If
  
  Rem Impressão em 1/8"
  Num_cod = 0
  Str_Impre = ""
  If rsParametros("Boleto Oitavo") = True Then
     If rsParametros("Cód Oitavo 1") <> "" Then
       Num_cod = 1
       If rsParametros("Cód Oitavo 2") <> "" Then
         Num_cod = 2
         If rsParametros("Cód Oitavo 3") <> "" Then
           Num_cod = 3
         End If
       End If
     End If
     If Num_cod = 1 Then
       Str_Impre = Chr$(Val(rsParametros("Cód Oitavo 1")))
     End If
     If Num_cod = 2 Then
       Str_Impre = Chr$(Val(rsParametros("Cód Oitavo 1")))
       Str_Impre = Str_Impre + Chr$(Val(rsParametros("Cód Oitavo 2")))
     End If
     If Num_cod = 3 Then
       Str_Impre = Chr$(Val(rsParametros("Cód Oitavo 1")))
       Str_Impre = Str_Impre + Chr$(Val(rsParametros("Cód Oitavo 2")))
       Str_Impre = Str_Impre + Chr$(Val(rsParametros("Cód Oitavo 3")))
     End If
     If Str_Impre <> "" Then
        Str_Impre = Chr$(Len(Str_Impre) Mod 256) + Chr$(Len(Str_Impre) \ 256) + Str_Impre
        Printer.Print ""
        If Not IsWindowsNT() Then
          Resposta = Escape(Printer.hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
        Else
          Resposta = Escape32(Printer.hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
        End If
        If Resposta <= 0 Then
          DisplayMsg "Não foi possível ajustar a impressora para 1/8'."
          Exit Sub
        End If
     End If
  End If
  
  
  On Error GoTo Arq_Inexiste
  Open Nome_Arq For Input As #1
  On Error GoTo 0
  
  Final = False
  Do
    Input #1, Texto
    If Texto = "*** Fim de arquivo ***" Then Final = True
    If Final = False Then
      Texto = Apaga_Aspas(Texto)
      Final_Linha = False
      If Len(Texto) < 3 Then
        DisplayMsg "Arquivo de configuração inválido."
        Exit Sub
      End If
      Especial2 = False
      If Left(Texto, 13) = "[LINHA_BRANCO" Then
        Especial2 = True
        Linhas = Val(Mid(Texto, 15))
        Do
          Printer.Print
          Linhas = Linhas - 1
        Loop Until Linhas = 0
      End If
      If Especial2 = False Then
        Str_Impre = Retorna_Texto(Texto)
        Printer.Print Str_Impre
      End If
    End If
 Loop Until Final = True
      

 Close #1
 Printer.Print
 Printer.EndDoc
 

 rsReceber.Edit
    '10/09/2007 - Anderson
    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
      "Cli:" & rsReceber("Cliente") & "- Seq:" & rsReceber("Sequência") & "- NF:" & rsReceber("Nota") & "- Venc:" & rsReceber("Vencimento") & "- Valor:" & rsReceber("Valor"), _
      "frmImprimeCarnes_ImprimeBoleto1", _
      "Contas a Receber", g_strArquivoSystemLog
    End If
   rsReceber("Impresso") = True
   rsReceber("Data Alteração") = Format(Date, "dd/mm/yyyy")
 rsReceber.Update

 
 Exit Sub
  
  
  
  
  
Arq_Inexiste:
  DisplayMsg "Arquivo de configuração não encontrado."
  Exit Sub
  
Final_Arquivo:
  'Mensagem.Caption "Nota fiscal impressa."
  Exit Sub


 

End Sub

Private Sub B_Imprime_Click()
 Dim Val1 As Integer
 Dim Val2 As Integer
 Dim Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 Dim Impressos As Integer
 Dim Aux_Vencimento As Variant
 Dim Aux_Cliente As Long
 Dim Aux_Contador As Long
 Dim Aux_Tipo As String
 Dim Resp As Integer
 Dim Nome_Arq As String
  Dim F As Form
  
 On Error GoTo ErrHandler
 
 Call StatusMsg("")

  gbToCancel = False
  
 Rem Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a empresa."
   Combo_Empresa.SetFocus
   Exit Sub
 End If

 If Filial_Liberada <> 0 Then
   If Val(Combo_Empresa.Text) <> Filial_Liberada Then
     DisplayMsg "Funcionário não tem acesso a esta filial."
     Exit Sub
   End If
 End If
 
   

 Rem Verifica Data
 Erro = False
 If IsNull(Data_Ini.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Data_Ini.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Data_Ini.SetFocus
   Exit Sub
 End If
 
 Rem Verifica Data Final
 Erro = False
 If IsNull(Data_Fim.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Data_Fim.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Data_Fim.SetFocus
   Exit Sub
 End If


 If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
   DisplayMsg "Data inicial deve ser menor ou igual a data final."
   Data_Ini.SetFocus
   Exit Sub
 End If

 If IsNull(Nome_Cliente.Caption) Then Combo_Cliente.Text = 0
 If Nome_Cliente.Caption = "" Then Combo_Cliente.Text = 0
 
 rsParametros.Index = "Filial"
 rsParametros.Seek "=", gnCodFilial
  
  Set F = New frmObsDoc
  F.Caption = "Impressão de Carnês"
  F.gsFileExt = ".CCA"
  F.Show vbModal
  Set F = Nothing
 If gsRetornoDoc <> "OK" Then
   DisplayMsg "Impressão cancelada."
   Exit Sub
 End If
 
 cmdCancel.Enabled = True
 
 Nome_Arq = gsConfigPath & gsDocFileName & ".CCA"
 If Dir(Nome_Arq) = "" Then
   DisplayMsg "Arquivo """ & Nome_Arq & """ não encontrado."
   Exit Sub
 End If
 
 rsClientes.Index = "Código"
 rsReceber.Index = "Vencimento"
 Aux_Vencimento = CDate("01/01/1980")
 Aux_Tipo = "R"
 Aux_Cliente = 0
 Aux_Contador = 0
 Impressos = 0
 
Lp1:
  If gbToCancel = True Then GoTo Fim
 rsReceber.Seek ">", Aux_Tipo, Val(Combo_Empresa.Text), Aux_Vencimento, Aux_Contador
 If rsReceber.NoMatch Then GoTo Fim
 If rsReceber("Filial") <> Val(Combo_Empresa.Text) Then GoTo Fim
 If rsReceber("Tipo") <> "R" Then GoTo Fim
 
 Aux_Vencimento = rsReceber("Vencimento")

 Aux_Contador = rsReceber("Contador")
 
 
 If rsReceber("Tipo Parcelamento") <> "T" Then GoTo Lp1
 
 
 If Val(Combo_Cliente.Text) <> 0 Then
   If Val(Combo_Cliente.Text) <> rsReceber("Cliente") Then GoTo Lp1
 End If
 
 If O_Vencimento.Value = True Then
   If rsReceber("Vencimento") < CDate(Data_Ini.Text) Then GoTo Lp1
   If rsReceber("Vencimento") > CDate(Data_Fim.Text) Then GoTo Lp1
 End If
 
 If O_Emissão.Value = True Then
   If rsReceber("Data Emissão") < CDate(Data_Ini.Text) Then GoTo Lp1
   If rsReceber("Data Emissão") > CDate(Data_Fim.Text) Then GoTo Lp1
 End If
 
 If O_N_Impresso.Value = True Then
   If rsReceber("Impresso") = True Then GoTo Lp1
 End If
 
 If O_Impresso.Value = True Then
   If rsReceber("Impresso") = False Then GoTo Lp1
 End If
 
  
 rsClientes.Seek "=", rsReceber("Cliente")
 If rsClientes.NoMatch Then GoTo Lp1
 
 
 Resp = Imprime_Carnê("R", rsReceber("Filial"), rsReceber("Vencimento"), rsReceber("Contador"), Nome_Arq)
 
 If Resp <> 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Voce deseja continuar com a impressão?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Exit Sub
    End If
 Else
   Impressos = Impressos + 1
   rsReceber.Edit
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
        "Cli:" & rsReceber("Cliente") & "- Seq:" & rsReceber("Sequência") & "- NF:" & rsReceber("Nota") & "- Venc:" & rsReceber("Vencimento") & "- Valor:" & rsReceber("Valor"), _
        "frmImprimeCarnes_B_Imprime_Click", _
        "Contas a Receber", g_strArquivoSystemLog
      End If
     rsReceber("Impresso") = True
     rsReceber("Data Alteração") = Format(Date, "dd/mm/yyyy")
   rsReceber.Update
    DisplayMsg "Carnês impressos : " + str(Impressos)
 End If
 
 GoTo Lp1


Fim:
  On Error Resume Next
  Call StatusMsg("")
  SetPrinterName "REL"
  cmdCancel.Enabled = False
  DisplayMsg "Final de impressão, foram impressos " + str(Impressos) + " carnês."
  On Error GoTo 0
  Exit Sub

ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Imprimir documento."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  SetPrinterName "REL"
  Exit Sub

End Sub

Private Sub cmd_OutraTelaCarne_Click()
    frmImprimeCarneCodigoBarras.Show
End Sub

Private Sub cmdCancel_Click()
  gbToCancel = True
End Sub

Private Sub Combo_Cliente_CloseUp()
 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 0 Or Val(Combo_Cliente.Text) > 99999999 Then Exit Sub

  rsClientes.Index = "Código"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then Exit Sub
  Nome_Cliente.Caption = rsClientes("Nome")

End Sub

Private Sub Combo_Empresa_CloseUp()
 Combo_Empresa.Text = Combo_Empresa.Columns(1).Text
 Combo_Empresa_LostFocus

End Sub

Private Sub Combo_Empresa_LostFocus()
  Nome_Empresa.Caption = ""
  If IsNull(Combo_Empresa.Text) Then Exit Sub
  If Not IsNumeric(Combo_Empresa.Text) Then Exit Sub
  If Val(Combo_Empresa.Text) < 0 Or Val(Combo_Empresa.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Empresa.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")

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

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("CLi_For", , dbReadOnly)
  Set rsReceber = db.OpenRecordset("Contas a Receber")
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName

  Combo_Empresa = gnCodFilial

End Sub


Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsClientes.Close
  rsReceber.Close
  Set rsParametros = Nothing
  Set rsClientes = Nothing
  Set rsReceber = Nothing
End Sub
