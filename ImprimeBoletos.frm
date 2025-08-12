VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmImprimeBoletos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Boletos Bancários"
   ClientHeight    =   4410
   ClientLeft      =   3885
   ClientTop       =   2460
   ClientWidth     =   6630
   HelpContextID   =   1650
   Icon            =   "ImprimeBoletos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4410
   ScaleWidth      =   6630
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   400
      Left            =   5190
      TabIndex        =   11
      Top             =   3870
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Período"
      Height          =   795
      Left            =   105
      TabIndex        =   20
      Top             =   1725
      Width           =   6435
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3690
         TabIndex        =   4
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   315
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2820
         TabIndex        =   22
         Top             =   375
         Width           =   810
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   21
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3945
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   4635
      Visible         =   0   'False
      Width           =   2325
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
      Bindings        =   "ImprimeBoletos.frx":058A
      DataSource      =   "Data3"
      Height          =   315
      Left            =   960
      TabIndex        =   2
      ToolTipText     =   "Use 0 para todas as Contas"
      Top             =   975
      Width           =   855
      DataFieldList   =   "Descrição"
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
      Columns.Count   =   3
      Columns(0).Width=   6668
      Columns(0).Caption=   "Descrição"
      Columns(0).Name =   "Descrição"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Descrição"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3096
      Columns(1).Caption=   "Conta"
      Columns(1).Name =   "Conta"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Conta"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1799
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2115
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   4500
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4485
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   3675
      TabIndex        =   10
      Top             =   3870
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imprimir :"
      Height          =   990
      Left            =   2385
      TabIndex        =   15
      Top             =   2685
      Width           =   4155
      Begin VB.OptionButton O_Todos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   255
         Width           =   960
      End
      Begin VB.OptionButton O_Impresso 
         Caption         =   "Somente os &já impressos (reimpressão)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   3015
      End
      Begin VB.OptionButton O_N_Impresso 
         Caption         =   "Somente os &não impressos"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Emissão por :"
      Height          =   990
      Left            =   105
      TabIndex        =   14
      Top             =   2670
      Width           =   2175
      Begin VB.OptionButton O_Vencimento 
         Caption         =   "Data de &Vencimento"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton O_Emissão 
         Caption         =   "Data de &Emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   570
         Width           =   1575
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "ImprimeBoletos.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   45
      Width           =   855
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
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "ImprimeBoletos.frx":05B2
      DataSource      =   "Data2"
      Height          =   315
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "Use 0 para todos os Clientes"
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
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Conta 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2250
      TabIndex        =   19
      Top             =   975
      Width           =   3405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conta :"
      Height          =   195
      Left            =   105
      TabIndex        =   18
      Top             =   1035
      Width           =   510
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2250
      TabIndex        =   17
      Top             =   45
      Width           =   3405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Filial :"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   105
      Width           =   390
   End
   Begin VB.Label Nome_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2250
      TabIndex        =   13
      Top             =   480
      Width           =   3405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente :"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   540
      Width           =   570
   End
End
Attribute VB_Name = "frmImprimeBoletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsParametros As Recordset
Dim rsClientes As Recordset
Dim rsReceber As Recordset
Dim rsContas As Recordset

Sub Imprime_Boleto1()
  Dim Aux As Variant
  Dim Nome_Arq As String
  Dim Texto As String
  Dim Final As Integer
  Dim Str_Impre As String
  Dim Cód1 As Integer
  Dim Cód2 As Integer
  Dim Cód3 As Integer
  Dim Num_cod As Integer
  Dim Resposta As Long
  Dim Final_Linha As Integer
  Dim Linhas As Integer
  Dim Especial2 As Integer
  Dim nFileNum As Integer
  
  Call StatusMsg("")
  
  On Error GoTo ErrHandler
  
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
  
  Nome_Arq = gsConfigPath & gsDocFileName & ".CBB"
  nFileNum = FreeFile
  Open Nome_Arq For Input As nFileNum
  
  Final = False
  Do
    Input #nFileNum, Texto
    If Texto = "*** Fim de arquivo ***" Then Final = True
    If Final = False Then
      Texto = Apaga_Aspas(Texto)
      Final_Linha = False
      If Len(Texto) < 3 Then
        DisplayMsg "Arquivo de configuração """ & Nome_Arq & """ inválido."
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
      

  Close #nFileNum
  Printer.Print
  Printer.EndDoc
  
  rsReceber.Edit
  '10/09/2007 - Anderson
  'Gera arquivo log do sistema
  If g_bolSystemLog Then
    SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
    "Cli:" & rsReceber("Cliente") & "- Seq:" & rsReceber("Sequência") & "- NF:" & rsReceber("Nota") & "- Venc:" & rsReceber("Vencimento") & "- Valor:" & rsReceber("Valor"), _
    "frmImprimeBoletos_ImprimeBoleto1", _
    "Contas a Receber", g_strArquivoSystemLog
  End If
  rsReceber("Impresso") = True
  rsReceber("Data Alteração") = Format(Date, "dd/mm/yyyy")
  rsReceber.Update
  
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Imprimir documento usando o arquivo de configuração """ & Nome_Arq & """."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Private Sub B_Imprime_Click()
  Dim Val1 As Integer
  Dim Val2 As Integer
  Dim Erro As Integer
  Dim Str1 As String
  Dim Str2 As String
  Dim Str3 As String
  Dim Str_Data1 As String
  Dim Str_Data2 As String
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
  Dim rstContasReceber As Recordset '05/06/2007 - Anderson
  Dim bolErroNossoNumero As Boolean '05/06/2007 - Anderson
  
  Dim strNossoNumero As String '16/05/2007 - Anderson
  
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
  F.Caption = "Impressão de Boletos"
  F.gsFileExt = ".CBB"
  F.Show vbModal
  Set F = Nothing
  If gsRetornoDoc <> "OK" Then
   DisplayMsg "Impressão cancelada."
   Exit Sub
  End If
  
  cmdCancel.Enabled = True
  
  Nome_Arq = gsConfigPath & gsDocFileName & ".CBB"
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
  
  If rsReceber("Tipo Parcelamento") <> "B" Then GoTo Lp1
  
  If Nome_Conta.Caption <> "" Then
   If rsReceber("Conta Boleto") <> Val(Combo_Conta.Text) Then GoTo Lp1
  End If
  
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
  
  '16/05/2007 - Anderson
  'Se Número de série Agrotama, informar nosso número para boletos pré-impressos
  If g_blnInformarNossoNumero And strNossoNumero = "" Then
    Do
      strNossoNumero = InputBox("Informe o Nosso Número para a impressão do boleto.", "Impressão de Boletos")
      If strNossoNumero = "" Then
        Exit Sub
      End If
      If Not IsNumeric(strNossoNumero) Then
        MsgBox "O valor digitado não é válido!", vbExclamation, "Impressão de Boletos"
      End If
    Loop Until IsNumeric(strNossoNumero)
  End If
 
  '05/06/2007 - Anderson
  'Verifica se o Nosso Número já foi emitido em outro boleto para evitar duplicidade.
  'Solicitado pelo cliente Agrotama
  If g_blnInformarNossoNumero Then
  
    'Abre registro para evitar duplicidade em nosso número
    Set rstContasReceber = db.OpenRecordset("SELECT CNAB_NossoNumero, Filial, Cliente, Vendedor, Sequência, Nota, [Data Emissão], Vencimento, Valor FROM [Contas a Receber] Where CNAB_NossoNumero='" & strNossoNumero & "'")
    
    'Informa que não existe problemas com Nosso Numero
    bolErroNossoNumero = False
    
    'Verifica se existe Nosso número no banco de dados
    If Not rstContasReceber.EOF Then
      MsgBox "Já existe um título com o Nosso Número: " & strNossoNumero & " informado em outro boleto." & Chr(13) & _
             "Favor verificar o título com os dados abaixo: " & Chr(13) & Chr(13) & _
             "Nosso Número: " & rstContasReceber("CNAB_NossoNumero") & Chr(13) & _
             "Filial: " & rstContasReceber("Filial") & Chr(13) & _
             "Cliente: " & rstContasReceber("Cliente") & Chr(13) & _
             "Vendedor: " & rstContasReceber("Vendedor") & Chr(13) & _
             "Sequência: " & rstContasReceber("Sequência") & Chr(13) & _
             "Nota: " & rstContasReceber("Nota") & Chr(13) & _
             "Data Emissão: " & rstContasReceber("Data Emissão") & Chr(13) & _
             "Vencimento: " & rstContasReceber("Vencimento") & Chr(13) & _
             "Valor: " & rstContasReceber("Valor"), vbOKOnly + vbInformation, "Impressão de Boletos"
             
      'Informa que existe um título com o mesmo Nosso Numero
      bolErroNossoNumero = True
    End If
  
    'Fecha tabela de contas a receber
    rstContasReceber.Close
    Set rstContasReceber = Nothing
    
    'Se houver duplicidade em Nosso Número, o sistema encerra.
    If bolErroNossoNumero Then
      GoTo Fim
    End If
  
  End If
  
  Resp = Imprime_Boleto("R", rsReceber("Filial"), rsReceber("Vencimento"), rsReceber("Contador"), Nome_Arq)
  
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
    rsReceber("Impresso") = True
    rsReceber("Data Alteração") = Format(Date, "dd/mm/yyyy")
    '16/05/2007 - Anderson
    'Se Número de série Agrotama, informar nosso número para boletos pré-impressos
    If CheckSerialCaseMod("QS73070-894") Then
      rsReceber("CNAB_NossoNumero") = Right(String(11, "0") & strNossoNumero, 11)
      rsReceber("CNAB_DigitoVerificador") = GetDigitoVerificador_NossoNumero(strNossoNumero, Bradesco)
      rsReceber("CNAB_Carteira") = "9"
      strNossoNumero = strNossoNumero + 1
    End If
    '10/09/2007 - Anderson
    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, _
      "Cli:" & rsReceber("Cliente") & "- Seq:" & rsReceber("Sequência") & "- NF:" & rsReceber("Nota") & "- Venc:" & rsReceber("Vencimento") & "- Valor:" & rsReceber("Valor"), _
      "frmImprimeBoletos_B_Imprime_Click", _
      "Contas a Receber", g_strArquivoSystemLog
    End If
    rsReceber.Update
    Call StatusMsg("Boletos impressos: " & str(Impressos))
  End If
  
  GoTo Lp1
  
Fim:
  SetPrinterName "REL"
  Call StatusMsg("")
  DisplayMsg "Final de impressão, foram impressos " + str(Impressos) + " boletos."
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

Private Sub Combo_Conta_CloseUp()

  Combo_Conta.Text = Combo_Conta.Columns(2).Text
  Combo_Conta_LostFocus
  

End Sub

Private Sub Combo_Conta_LostFocus()

  Nome_Conta.Caption = ""
  If IsNull(Combo_Conta.Text) Then Exit Sub
  If Combo_Conta.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Conta.Text) Then Exit Sub
  
  rsContas.Index = "Código"
  rsContas.Seek "=", Val(Combo_Conta.Text)
  If rsContas.NoMatch Then Exit Sub
  
  Nome_Conta.Caption = rsContas("Descrição") & ""
  

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
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName

  Combo_Empresa = gnCodFilial
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsClientes.Close
  rsReceber.Close
  rsContas.Close
  Set rsParametros = Nothing
  Set rsClientes = Nothing
  Set rsReceber = Nothing
  Set rsContas = Nothing
End Sub
