VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProgramaFidelidadeConsultaGerencial 
   Caption         =   " Programa Fidelidade x Consulta Gerencial"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgramaFidelidadeConsultaGerencial.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   15210
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmb_programaFidelidade 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   75
      Width           =   7185
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   90
      TabIndex        =   5
      Top             =   510
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   12991
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Posição Pontos"
      TabPicture(0)   =   "frmProgramaFidelidadeConsultaGerencial.frx":4E95A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbl_resultadoPontos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd_pesquisar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt_pontos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmb_tipoPesquisa"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Relatório de Lançamentos"
      TabPicture(1)   =   "frmProgramaFidelidadeConsultaGerencial.frx":4E976
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "grade_programas"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmd_pesqLanc"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt_cpf"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt_totalReais"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txt_totalpontos"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmb_clientes"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame2"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame3"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4470
         TabIndex        =   39
         Top             =   1590
         Width           =   10365
         Begin VB.TextBox txt_guidResgate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   5250
            TabIndex        =   26
            Top             =   180
            Width           =   3495
         End
         Begin VB.TextBox txt_sequencia 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1410
            TabIndex        =   25
            Top             =   180
            Width           =   1995
         End
         Begin VB.Label Label14 
            Caption         =   "Cod Guid Resgate"
            Height          =   255
            Left            =   3780
            TabIndex        =   41
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Nº Sequência"
            Height          =   255
            Left            =   330
            TabIndex        =   40
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   240
         TabIndex        =   37
         Top             =   780
         Width           =   4005
         Begin VB.TextBox txt_cnpjFilial 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1290
            TabIndex        =   18
            Top             =   270
            Width           =   2355
         End
         Begin VB.Label Label12 
            Caption         =   "CNPJ Filial"
            Height          =   255
            Left            =   330
            TabIndex        =   38
            Top             =   330
            Width           =   825
         End
      End
      Begin VB.ComboBox cmb_clientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   7050
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   420
         Width           =   7785
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4470
         TabIndex        =   34
         Top             =   780
         Width           =   10365
         Begin VB.CheckBox chk_compras 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Somente lançamentos de Compras?"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   300
            TabIndex        =   24
            Top             =   450
            Width           =   3135
         End
         Begin VB.CheckBox chk_resgatePelaDataRecebida 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Considerar a pesquisa pela data de recebimento do resgate"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   5250
            TabIndex        =   22
            Top             =   480
            Width           =   4935
         End
         Begin VB.OptionButton opt_resgateNaoRecebidos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Somente os NÃO recebidos"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   7410
            TabIndex        =   23
            Top             =   195
            Width           =   2505
         End
         Begin VB.OptionButton opt_resgateRecebidos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Somente os recebidos"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   5250
            TabIndex        =   21
            Top             =   195
            Width           =   2085
         End
         Begin VB.OptionButton opt_resgateTodos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Todos"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   4230
            TabIndex        =   20
            Top             =   195
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.CheckBox chk_resgate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Somente lançamentos de RESGATE?"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   300
            TabIndex        =   19
            Top             =   180
            Width           =   3225
         End
      End
      Begin VB.TextBox txt_totalpontos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2490
         TabIndex        =   32
         Top             =   6900
         Width           =   1455
      End
      Begin VB.TextBox txt_totalReais 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   29
         Top             =   6900
         Width           =   1455
      End
      Begin VB.TextBox txt_cpf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1890
         TabIndex        =   16
         Top             =   390
         Width           =   2355
      End
      Begin VB.CommandButton cmd_pesqLanc 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2250
         Width           =   14595
      End
      Begin VB.ComboBox cmb_tipoPesquisa 
         Height          =   315
         ItemData        =   "frmProgramaFidelidadeConsultaGerencial.frx":4E992
         Left            =   -72960
         List            =   "frmProgramaFidelidadeConsultaGerencial.frx":4E99F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1005
         Width           =   5685
      End
      Begin VB.TextBox txt_pontos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   -73050
         TabIndex        =   10
         Top             =   2940
         Width           =   2175
      End
      Begin VB.CommandButton cmd_pesquisar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -74400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   7125
      End
      Begin MSFlexGridLib.MSFlexGrid grade_programas 
         Height          =   4125
         Left            =   240
         TabIndex        =   30
         Top             =   2730
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   7276
         _Version        =   393216
         Rows            =   1
         Cols            =   16
         FixedCols       =   0
         BackColor       =   12648447
         BackColorSel    =   12648384
         ForeColorSel    =   0
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "Clientes Cód/Nome"
         Height          =   255
         Left            =   5490
         TabIndex        =   36
         Top             =   465
         Width           =   1515
      End
      Begin VB.Label Label10 
         Caption         =   "ou"
         Height          =   255
         Left            =   4770
         TabIndex        =   35
         Top             =   480
         Width           =   285
      End
      Begin VB.Label Label6 
         Caption         =   "Total de pontos acumulados"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   6960
         Width           =   2265
      End
      Begin VB.Label Label9 
         Caption         =   "Saldo em R$  acumulados"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4710
         TabIndex        =   31
         Top             =   6960
         Width           =   2085
      End
      Begin VB.Label Label7 
         Caption         =   "CPF/CNPJ do Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   450
         Width           =   1605
      End
      Begin VB.Label Label4 
         Caption         =   "Resultado da pesquisa"
         Height          =   225
         Left            =   -74370
         TabIndex        =   14
         Top             =   2580
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Pesquisa"
         Height          =   255
         Left            =   -74400
         TabIndex        =   13
         Top             =   1050
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "PONTOS"
         Height          =   255
         Left            =   -74340
         TabIndex        =   12
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label lbl_resultadoPontos 
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -70710
         TabIndex        =   11
         Top             =   3030
         Width           =   2985
      End
   End
   Begin VB.CommandButton cmd_calendarioDtFim 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4515
      Picture         =   "frmProgramaFidelidadeConsultaGerencial.frx":4E9DB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   30
      Width           =   465
   End
   Begin VB.CommandButton cmd_calendarioDtIni 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      Picture         =   "frmProgramaFidelidadeConsultaGerencial.frx":4F2BD
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   30
      Width           =   465
   End
   Begin MSMask.MaskEdBox Data_Fim 
      Height          =   315
      Left            =   3375
      TabIndex        =   2
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
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
      Left            =   900
      TabIndex        =   0
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
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
   Begin VB.Label Label2 
      Caption         =   "Programa Fidelidade"
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Data De:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Até"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3060
      TabIndex        =   8
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmProgramaFidelidadeConsultaGerencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrPrograma(1000, 2) As String
Dim arrContador As Integer
Dim flagHAMUD As Boolean


Private Sub chk_resgate_Click()
  If chk_resgate.Value = 1 Then
    opt_resgateNaoRecebidos.Enabled = True
    opt_resgateRecebidos.Enabled = True
    opt_resgateTodos.Enabled = True
    chk_resgatePelaDataRecebida.Enabled = True
    chk_compras.Value = 0
  Else
    opt_resgateNaoRecebidos.Enabled = False
    opt_resgateRecebidos.Enabled = False
    opt_resgateTodos.Enabled = False
    chk_resgatePelaDataRecebida.Enabled = False
  End If
End Sub

Private Sub chk_compras_Click()
  If chk_compras.Value = 1 Then
      chk_resgate.Value = 0
  End If
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub cmd_pesqLanc_Click()
On Error GoTo Erro
  Dim iStatus As Integer
  Dim sStatus As String
  Dim arrProg() As String
  Dim sCPF As String
  Dim sCNPJAux() As String
  Dim sCNPJFilial As String
    
  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If cmb_programaFidelidade.Text = "" Then
      MsgBox "Informe qual é o Programa de Fidelidade", vbInformation, "Atenção"
      cmb_programaFidelidade.SetFocus
      Exit Sub
  End If
  
  arrProg = Split(cmb_programaFidelidade.Text, " - ")
  
  grade_programas.Rows = 1
  
  If LTrim(RTrim(txt_cpf.Text)) = "" And cmb_clientes.Text = "" Then
      sCPF = ""
  Else
      If LTrim(RTrim(txt_cpf.Text)) <> "" Then
          sCPF = LTrim(RTrim(txt_cpf.Text))
      Else
          sCPF = LTrim(RTrim(cmb_clientes.Text))
      
          sCNPJAux = Split(sCPF, " *(")
          sCPF = sCNPJAux(1)
      End If
  End If
  
  sCPF = Replace(sCPF, "-", "")
  sCPF = Replace(sCPF, ".", "")
  sCPF = Replace(sCPF, ";", "")
  sCPF = Replace(sCPF, "/", "")
  sCPF = Replace(sCPF, "\", "")
  sCPF = Replace(sCPF, "(", "")
  sCPF = Replace(sCPF, ")", "")
  sCPF = Replace(sCPF, " ", "")

  If LTrim(RTrim(txt_cnpjFilial.Text)) <> "" Then
      sCNPJFilial = LTrim(RTrim(txt_cnpjFilial.Text))
  End If
  
  sCNPJFilial = Replace(sCNPJFilial, "-", "")
  sCNPJFilial = Replace(sCNPJFilial, ".", "")
  sCNPJFilial = Replace(sCNPJFilial, ";", "")
  sCNPJFilial = Replace(sCNPJFilial, "/", "")
  sCNPJFilial = Replace(sCNPJFilial, "\", "")
  sCNPJFilial = Replace(sCNPJFilial, "(", "")
  sCNPJFilial = Replace(sCNPJFilial, ")", "")
  sCNPJFilial = Replace(sCNPJFilial, " ", "")
  
  

  Dim rsPrograma As New ADODB.Recordset
  Dim strSQL As String

  strSQL = "SELECT CNPJ,CPF_CGC_CLIENTE,Cd_programa,Cd_cliente,Dt_criacao,Vl_CompraCliente, "
  strSQL = strSQL & " Nm_PontosAdquiridos,Vl_SaldoEmReais,Tp_lancamento,Cd_operador,Cd_guid_resgate,"
  strSQL = strSQL & " Status_guid_resgate,Dt_recebido_guid_resgate,Cd_operador_recebido_guid_resgate, Cd_SequenciaVenda, Nm_cliente "
  strSQL = strSQL & " FROM ProgramaFidelidade_lancamentos "
  'strSQL = strSQL & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
  
  If LTrim(RTrim(txt_sequencia.Text)) <> "" Or LTrim(RTrim(txt_guidResgate.Text)) <> "" Then
      
      If LTrim(RTrim(txt_sequencia.Text)) <> "" Then
          strSQL = strSQL & " WHERE Cd_SequenciaVenda = " & txt_sequencia.Text
      ElseIf LTrim(RTrim(txt_guidResgate.Text)) <> "" Then
          strSQL = strSQL & " WHERE Cd_guid_resgate = '" & LTrim(RTrim(txt_guidResgate.Text)) & "'"
      End If
  Else
      If chk_resgate.Value = 1 And opt_resgateRecebidos.Value = True And chk_resgatePelaDataRecebida.Value = 1 Then
          strSQL = strSQL & " where convert(NVARCHAR, Dt_recebido_guid_resgate , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
          strSQL = strSQL & " convert(NVARCHAR, Dt_recebido_guid_resgate , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      Else
          strSQL = strSQL & " where convert(NVARCHAR, Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
          strSQL = strSQL & " convert(NVARCHAR, Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      End If
    
      strSQL = strSQL & " Cd_programa = " & arrProg(0)
    
      If sCNPJFilial <> "" Then
          strSQL = strSQL & " AND CNPJ = '" & sCNPJFilial & "' "
      End If
    
      If sCPF <> "" Then
          strSQL = strSQL & " AND CPF_CGC_CLIENTE = '" & sCPF & "' "
      End If
    
      If chk_resgate.Value = 1 Then
          strSQL = strSQL & " AND tp_lancamento = 2 "
    
          If opt_resgateRecebidos.Value = True Then
              strSQL = strSQL & " AND Status_guid_resgate = 1 "
          ElseIf opt_resgateNaoRecebidos.Value = True Then
              strSQL = strSQL & " AND Status_guid_resgate is null "
          End If
      ElseIf chk_compras = 1 Then
          strSQL = strSQL & " AND tp_lancamento = 1 "
      End If
    
      strSQL = strSQL & " ORDER BY Dt_criacao ASC "
  End If

  'Set rsPrograma = db_SQLSERVER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  rsPrograma.Open strSQL, gDB_SQLSERVER

  Dim sSt_RESGATE As String
  Dim sTp_lancamento As String
  Dim lTotalPtsAcumulados As Long
  Dim dTotalReais As Double
  
  lTotalPtsAcumulados = 0
  dTotalReais = 0

  If Not (rsPrograma.EOF And rsPrograma.BOF) Then
    rsPrograma.MoveFirst
  End If
  While Not rsPrograma.EOF
  
      If rsPrograma.Fields("Tp_lancamento").Value = 1 Then
          sTp_lancamento = "COMPRA"
          lTotalPtsAcumulados = lTotalPtsAcumulados + rsPrograma.Fields("Nm_PontosAdquiridos").Value
          dTotalReais = dTotalReais + rsPrograma.Fields("Vl_SaldoEmReais").Value
      Else
          sTp_lancamento = "RESGATE"
          lTotalPtsAcumulados = lTotalPtsAcumulados - rsPrograma.Fields("Nm_PontosAdquiridos").Value
          dTotalReais = dTotalReais - rsPrograma.Fields("Vl_SaldoEmReais").Value
          
          If Not IsNull(rsPrograma.Fields("Status_guid_resgate").Value) And rsPrograma.Fields("Status_guid_resgate").Value = 1 Then
              sSt_RESGATE = "Utilizado pelo Cliente"
          Else
              sSt_RESGATE = "Não utilizado pelo Cliente"
          End If
      End If

      grade_programas.AddItem 0 & vbTab & rsPrograma.Fields("Dt_criacao").Value & vbTab & _
                              rsPrograma.Fields("CNPJ").Value & vbTab & _
                              rsPrograma.Fields("CPF_CGC_CLIENTE").Value & vbTab & _
                              rsPrograma.Fields("Cd_cliente").Value & vbTab & _
                              rsPrograma.Fields("Nm_cliente").Value & vbTab & _
                              Format(rsPrograma.Fields("Vl_CompraCliente").Value, FORMAT_VALUE) & vbTab & _
                              rsPrograma.Fields("Cd_SequenciaVenda").Value & vbTab & _
                              rsPrograma.Fields("Nm_PontosAdquiridos").Value & vbTab & _
                              Format(rsPrograma.Fields("Vl_SaldoEmReais").Value, FORMAT_VALUE) & vbTab & _
                              sTp_lancamento & vbTab & _
                              rsPrograma.Fields("Cd_operador").Value & vbTab & _
                              rsPrograma.Fields("Cd_guid_resgate").Value & vbTab & _
                              sSt_RESGATE & vbTab & _
                              rsPrograma.Fields("Dt_recebido_guid_resgate").Value & vbTab & _
                              rsPrograma.Fields("Cd_operador_recebido_guid_resgate").Value

      rsPrograma.MoveNext
  Wend
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  If chk_resgate.Value = 1 Then
    lTotalPtsAcumulados = lTotalPtsAcumulados * -1
    dTotalReais = dTotalReais * -1
  End If
  
  txt_totalpontos.Text = lTotalPtsAcumulados
  txt_totalReais.Text = FormataValorTexto(dTotalReais, 2)
  
  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da grade...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_pesquisar_Click()
On Error GoTo Erro
    Dim rsPrograma1 As New ADODB.Recordset
    Dim rsPrograma2 As New ADODB.Recordset
    Dim rsPrograma3 As New ADODB.Recordset
    Dim rsPrograma4 As New ADODB.Recordset
    Dim lngPontos As Long
    Dim strSQL As String
    Dim arrAux() As String

  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Fim.SetFocus
    Exit Sub
  End If

  If cmb_programaFidelidade.Text = "" Then
    MsgBox "Selecione um programa de fidelidade", vbInformation, "Atenção"
    cmb_programaFidelidade.SetFocus
    Exit Sub
  End If

  If cmb_tipoPesquisa.Text = "" Then
    MsgBox "Selecione um tipo de pesquisa", vbInformation, "Atenção"
    cmb_tipoPesquisa.SetFocus
    Exit Sub
  End If
  
  Dim cnpjAux As String
  
  arrAux = Split(cmb_programaFidelidade.Text, " - ")
  
  lngPontos = 0
  
  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Then
      cnpjAux = gCNPJ_CPFControleDeLicencaWebApi
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
      cnpjAux = gCNPJProgramaFidelidadeCNPJPrincipal
  End If
  
  If cmb_tipoPesquisa.Text = "PONTOS EM ABERTO" Then
      strSQL = "select SUM(nm_pontosAdquiridos) "
      strSQL = strSQL & " from  ProgramaFidelidade_lancamentos A, ProgramaFidelidade_empresaGrupo B "
      strSQL = strSQL & " where convert(NVARCHAR, A.Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
      strSQL = strSQL & " convert(NVARCHAR, A.Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      strSQL = strSQL & " B.CNPJ_principal = '" & cnpjAux & "' and "
      strSQL = strSQL & " B.CNPJ = A.CNPJ and A.tp_lancamento = 1 and "
      strSQL = strSQL & " A.Cd_programa=" & arrAux(0)

      rsPrograma1.Open strSQL, gDB_SQLSERVER

      If Not (rsPrograma1.EOF And rsPrograma1.BOF) Then
        rsPrograma1.MoveFirst
      End If
      If Not IsNull(rsPrograma1.Fields(0).Value) Then
          lngPontos = rsPrograma1.Fields(0).Value
      End If
      rsPrograma1.Close
      Set rsPrograma1 = Nothing

      strSQL = "select SUM(nm_pontosAdquiridos)"
      strSQL = strSQL & " from  ProgramaFidelidade_lancamentos "
      strSQL = strSQL & " where convert(NVARCHAR, Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
      strSQL = strSQL & " convert(NVARCHAR, Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      strSQL = strSQL & " CNPJ = '" & cnpjAux & "' and tp_lancamento = 1 and "
      strSQL = strSQL & " Cd_programa=" & arrAux(0)
      
      rsPrograma2.Open strSQL, gDB_SQLSERVER

      If Not (rsPrograma2.EOF And rsPrograma2.BOF) Then
        rsPrograma2.MoveFirst
      End If
      If Not IsNull(rsPrograma2.Fields(0).Value) Then
          lngPontos = lngPontos + rsPrograma2.Fields(0).Value
      End If
      rsPrograma2.Close
      Set rsPrograma2 = Nothing

      strSQL = "select SUM(nm_pontosAdquiridos) "
      strSQL = strSQL & " from  ProgramaFidelidade_lancamentos A, ProgramaFidelidade_empresaGrupo B "
      strSQL = strSQL & " where convert(NVARCHAR, A.Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
      strSQL = strSQL & " convert(NVARCHAR, A.Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      strSQL = strSQL & " B.CNPJ_principal = '" & cnpjAux & "' and "
      strSQL = strSQL & " B.CNPJ = A.CNPJ and A.tp_lancamento = 2 And A.Status_guid_resgate <> 1 and "
      strSQL = strSQL & " A.Cd_programa=" & arrAux(0)

      rsPrograma3.Open strSQL, gDB_SQLSERVER

      If Not (rsPrograma3.EOF And rsPrograma3.BOF) Then
        rsPrograma3.MoveFirst
      End If
      If Not IsNull(rsPrograma3.Fields(0).Value) Then
          lngPontos = lngPontos - rsPrograma3.Fields(0).Value
      End If
      rsPrograma3.Close
      Set rsPrograma3 = Nothing
      
      strSQL = "select SUM(nm_pontosAdquiridos)"
      strSQL = strSQL & " from  ProgramaFidelidade_lancamentos "
      strSQL = strSQL & " where convert(NVARCHAR, Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
      strSQL = strSQL & " convert(NVARCHAR, Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      strSQL = strSQL & " CNPJ = '" & cnpjAux & "' and tp_lancamento = 2 And Status_guid_resgate <> 1 and "
      strSQL = strSQL & " Cd_programa=" & arrAux(0)
      
      rsPrograma4.Open strSQL, gDB_SQLSERVER

      If Not (rsPrograma4.EOF And rsPrograma4.BOF) Then
        rsPrograma4.MoveFirst
      End If
      If Not IsNull(rsPrograma4.Fields(0).Value) Then
          lngPontos = lngPontos - rsPrograma4.Fields(0).Value
      End If
      rsPrograma4.Close
      Set rsPrograma4 = Nothing

      txt_pontos.Text = lngPontos
      lbl_resultadoPontos.Caption = cmb_tipoPesquisa.Text

  ElseIf cmb_tipoPesquisa.Text = "PONTOS EM RESGATE" Then
      strSQL = "select SUM(nm_pontosAdquiridos) "
      strSQL = strSQL & " from  ProgramaFidelidade_lancamentos A, ProgramaFidelidade_empresaGrupo B "
      strSQL = strSQL & " where convert(NVARCHAR, A.Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
      strSQL = strSQL & " convert(NVARCHAR, A.Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      strSQL = strSQL & " B.CNPJ_principal = '" & cnpjAux & "' and "
      strSQL = strSQL & " B.CNPJ = A.CNPJ and A.tp_lancamento = 2 And A.Status_guid_resgate is null and "
      strSQL = strSQL & " A.Cd_programa=" & arrAux(0)

      rsPrograma3.Open strSQL, gDB_SQLSERVER

      If Not (rsPrograma3.EOF And rsPrograma3.BOF) Then
        rsPrograma3.MoveFirst
      End If
      If Not IsNull(rsPrograma3.Fields(0).Value) Then
          lngPontos = rsPrograma3.Fields(0).Value
      End If
      rsPrograma3.Close
      Set rsPrograma3 = Nothing

      strSQL = "select SUM(nm_pontosAdquiridos)"
      strSQL = strSQL & " from  ProgramaFidelidade_lancamentos "
      strSQL = strSQL & " where convert(NVARCHAR, Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
      strSQL = strSQL & " convert(NVARCHAR, Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      strSQL = strSQL & " CNPJ = '" & cnpjAux & "' and tp_lancamento = 2 And Status_guid_resgate is null and "
      strSQL = strSQL & " Cd_programa=" & arrAux(0)

      rsPrograma4.Open strSQL, gDB_SQLSERVER

      If Not (rsPrograma4.EOF And rsPrograma4.BOF) Then
        rsPrograma4.MoveFirst
      End If
      If Not IsNull(rsPrograma4.Fields(0).Value) Then
          lngPontos = lngPontos + rsPrograma4.Fields(0).Value
      End If
      rsPrograma4.Close
      Set rsPrograma4 = Nothing
      
      txt_pontos.Text = lngPontos
      lbl_resultadoPontos.Caption = cmb_tipoPesquisa.Text
      
  ElseIf cmb_tipoPesquisa.Text = "PONTOS RESGATADOS" Then
      strSQL = "select SUM(nm_pontosAdquiridos) "
      strSQL = strSQL & " from  ProgramaFidelidade_lancamentos A, ProgramaFidelidade_empresaGrupo B "
      strSQL = strSQL & " where convert(NVARCHAR, A.Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
      strSQL = strSQL & " convert(NVARCHAR, A.Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      strSQL = strSQL & " B.CNPJ_principal = '" & cnpjAux & "' and "
      strSQL = strSQL & " B.CNPJ = A.CNPJ and A.tp_lancamento = 2 And A.Status_guid_resgate = 1 and "
      strSQL = strSQL & " A.Cd_programa=" & arrAux(0)
  
      rsPrograma3.Open strSQL, gDB_SQLSERVER

      If Not (rsPrograma3.EOF And rsPrograma3.BOF) Then
        rsPrograma3.MoveFirst
      End If
      If Not IsNull(rsPrograma3.Fields(0).Value) Then
          lngPontos = rsPrograma3.Fields(0).Value
      End If
      rsPrograma3.Close
      Set rsPrograma3 = Nothing
      
      strSQL = "select SUM(nm_pontosAdquiridos)"
      strSQL = strSQL & " from  ProgramaFidelidade_lancamentos "
      strSQL = strSQL & " where convert(NVARCHAR, Dt_criacao , 112) >= '" & Mid(Data_Ini.Text, 7, 4) & Mid(Data_Ini.Text, 4, 2) & Mid(Data_Ini.Text, 1, 2) & "' and "
      strSQL = strSQL & " convert(NVARCHAR, Dt_criacao , 112) <= '" & Mid(Data_Fim.Text, 7, 4) & Mid(Data_Fim.Text, 4, 2) & Mid(Data_Fim.Text, 1, 2) & "' and "
      strSQL = strSQL & " CNPJ = '" & cnpjAux & "' and tp_lancamento = 2 And Status_guid_resgate = 1 and "
      strSQL = strSQL & " Cd_programa=" & arrAux(0)
      
      rsPrograma4.Open strSQL, gDB_SQLSERVER

      If Not (rsPrograma4.EOF And rsPrograma4.BOF) Then
        rsPrograma4.MoveFirst
      End If
      If Not IsNull(rsPrograma4.Fields(0).Value) Then
          lngPontos = lngPontos + rsPrograma4.Fields(0).Value
      End If
      rsPrograma4.Close
      Set rsPrograma4 = Nothing
      
      txt_pontos.Text = lngPontos
      lbl_resultadoPontos.Caption = cmb_tipoPesquisa.Text
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro no metodo pesquisar...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Sub Form_Load()
On Error GoTo Erro:
  Dim strSQL As String
  Dim iStatus As Integer
  Dim sStatus As String

  grade_programas.ColWidth(0) = 10
  grade_programas.ColWidth(1) = 1550
  grade_programas.ColWidth(2) = 1500
  grade_programas.ColWidth(3) = 1500
  grade_programas.ColWidth(4) = 1000
  grade_programas.ColWidth(5) = 2500
  grade_programas.ColWidth(6) = 1100
  grade_programas.ColWidth(7) = 1200
  grade_programas.ColWidth(8) = 1500
  grade_programas.ColWidth(9) = 1100
  grade_programas.ColWidth(10) = 1280
  grade_programas.ColWidth(11) = 1200
  grade_programas.ColWidth(12) = 2500
  grade_programas.ColWidth(13) = 2200
  grade_programas.ColWidth(14) = 1950
  grade_programas.ColWidth(15) = 1900

  grade_programas.Row = 0
  grade_programas.TextMatrix(0, 1) = "Dt Lançamento"
  grade_programas.TextMatrix(0, 2) = "CNPJ"
  grade_programas.TextMatrix(0, 3) = "CPF/CNPJ Cliente"
  grade_programas.TextMatrix(0, 4) = "Cód Cliente"
  grade_programas.TextMatrix(0, 5) = "Nome Cliente"
  grade_programas.TextMatrix(0, 6) = "Vl Compra"
  grade_programas.TextMatrix(0, 7) = "Nº Sequência"
  grade_programas.TextMatrix(0, 8) = "Pontos Adquiridos"
  grade_programas.TextMatrix(0, 9) = "Vl Ganho"
  grade_programas.TextMatrix(0, 10) = "Tp. lançamento"
  grade_programas.TextMatrix(0, 11) = "Cód.Operador"
  grade_programas.TextMatrix(0, 12) = "Cód RESGATE"
  grade_programas.TextMatrix(0, 13) = "Status RESGATE"
  grade_programas.TextMatrix(0, 14) = "Dt RESGATE Uso Cliente"
  grade_programas.TextMatrix(0, 15) = "Cód RESGATE Operador"

  '***************************************************************
  'Verifica se HAMUD
  If gCNPJ_CPFControleDeLicencaWebApi = "80778855000187" _
      Or gCNPJ_CPFControleDeLicencaWebApi = "06888091000120" _
      Or gCNPJ_CPFControleDeLicencaWebApi = "08518307000190" _
      Or gCNPJ_CPFControleDeLicencaWebApi = "73213944000110" Then

    flagHAMUD = True
  End If
  '***************************************************************

  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER

  'Se HAMUD...então fazer um distinct por nomeCliente nos registros de lançamentos (pois são bases de clientes separados)
  If flagHAMUD = True Then
    Dim rsClientes As New ADODB.Recordset

    strSQL = "SELECT DISTINCT(Nm_cliente),CPF_CGC_CLIENTE, Cd_cliente "
    strSQL = strSQL & " FROM ProgramaFidelidade_lancamentos (nolock) "
    strSQL = strSQL & " WHERE CNPJ in('80778855000187','06888091000120','08518307000190','73213944000110')"

    rsClientes.Open strSQL, gDB_SQLSERVER

    cmb_clientes.AddItem ""

    'Carregar a combo de clientes
    While Not rsClientes.EOF
      cmb_clientes.AddItem rsClientes.Fields(2).Value & " - " & rsClientes.Fields(0).Value & " *(" & rsClientes.Fields(1).Value & ")"

      rsClientes.MoveNext
    Wend
    rsClientes.Close
    Set rsClientes = Nothing

  Else
    Dim rsClientes2 As Recordset
    Set rsClientes2 = db.OpenRecordset("Select Código, Nome, CGC from [Cli_For] order by Nome ", dbOpenDynaset)

    cmb_clientes.AddItem ""

    'Carregar a combo de clientes
    While Not rsClientes2.EOF
      cmb_clientes.AddItem rsClientes2.Fields(0).Value & " - " & rsClientes2.Fields(1).Value & " *(" & rsClientes2.Fields(2).Value & ")"

      rsClientes2.MoveNext
    Wend
    rsClientes2.Close
    Set rsClientes2 = Nothing
  End If

  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER

  Dim rsPrograma As New ADODB.Recordset

  arrContador = 0

  ProgramaFidelidadeEmpresaGrupoValida

  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Then
    strSQL = "SELECT * FROM ProgramaFidelidade_empresa "
    strSQL = strSQL & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
    strSQL = "SELECT * FROM ProgramaFidelidade_empresa "
    strSQL = strSQL & " WHERE CNPJ = '" & gCNPJProgramaFidelidadeCNPJPrincipal & "'"
  End If

  'Set rsPrograma = db_SQLSERVER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  rsPrograma.Open strSQL, gDB_SQLSERVER

  If Not (rsPrograma.EOF And rsPrograma.BOF) Then
    rsPrograma.MoveFirst
  End If
  While Not rsPrograma.EOF

      iStatus = rsPrograma.Fields("Cd_status").Value
      If iStatus = 1 Then
          sStatus = "ATIVO"
      Else
          sStatus = "INATIVO"
      End If

      cmb_programaFidelidade.AddItem rsPrograma.Fields("Cd_programa").Value & " - " & rsPrograma.Fields("Nm_programa").Value & " - " & sStatus
      arrPrograma(arrContador, 0) = rsPrograma.Fields("Cd_programa").Value
      arrPrograma(arrContador, 1) = rsPrograma.Fields("Vl_ProgFidelidadeParaCadaPonto").Value

      arrContador = arrContador + 1
      rsPrograma.MoveNext
  Wend
  rsPrograma.Close
  Set rsPrograma = Nothing

  SSTab1.Tab = 1

  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da tela...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Fechar conexão com o banco de dados SQL SERVER
  'gnCloseDB_SQLSERVER
End Sub


