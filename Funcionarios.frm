VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmFuncionarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"Funcionarios.frx":0000
   ClientHeight    =   8355
   ClientLeft      =   585
   ClientTop       =   870
   ClientWidth     =   12780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1310
   Icon            =   "Funcionarios.frx":00CB
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8355
   ScaleWidth      =   12780
   Begin VB.CheckBox chkAtivo 
      Appearance      =   0  'Flat
      Caption         =   "&Ativo"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11790
      TabIndex        =   5
      ToolTipText     =   "Com esta opção habilitada, o Funcionário estará ativo (visível) para todo sistema."
      Top             =   90
      Width           =   855
   End
   Begin VB.CheckBox Super 
      Appearance      =   0  'Flat
      Caption         =   "&Superusuário"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10350
      TabIndex        =   4
      ToolTipText     =   "Opção ligada atribui a este usuário o status "
      Top             =   90
      Width           =   1365
   End
   Begin VB.CheckBox Liberado 
      Appearance      =   0  'Flat
      Caption         =   "&Liberado"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9135
      TabIndex        =   3
      ToolTipText     =   "Opção ligada libera acesso deste usuário ao sistema."
      Top             =   90
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox Apelido 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2130
      MaxLength       =   10
      TabIndex        =   1
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   4530
      MaxLength       =   35
      TabIndex        =   2
      Top             =   60
      Width           =   3975
   End
   Begin VB.TextBox Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   660
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "Pressione F5 para o Próximo Livre."
      Top             =   60
      Width           =   795
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   7380
      Left            =   75
      TabIndex        =   41
      Top             =   480
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   13018
      _Version        =   393216
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
      TabCaption(0)   =   "&Dados Gerais"
      TabPicture(0)   =   "Funcionarios.frx":4EA25
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label16"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label19"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label20"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label22"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label25"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Shape1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblSupervisor"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblCDRC"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboSupervisor"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "mskPercDesconto"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Nascimento"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Admissão"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Comissão"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "SSFrame1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Senha"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Carteira"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Identidade"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "CPF"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CEP"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Estado"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Cidade"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Endereço"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Filial_Acesso"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Bairro"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Fone"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Cargo"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Frame2"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Comissão_Serviço"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "chkDesconto"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtMargemLimiteCredito"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "chkMarketing"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "datSupervisor"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtNomeSupervisor"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Sexo"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtCDRC"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Frame1"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "chk_mostrarTelaPesquisaTipoFoto"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "chkPrestServ"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).ControlCount=   52
      TabCaption(1)   =   "&Permissões"
      TabPicture(1)   =   "Funcionarios.frx":4EA41
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grade1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DropDown1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Data1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Data2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmd_tabelasDePrecos"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd_DRE_help"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Log Transações"
      TabPicture(2)   =   "Funcionarios.frx":4EA5D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label18"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label24"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label26"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label27"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "grade_log"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Data_Ini"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Data_Fim"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmd_calendarioDtFim"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmd_calendarioDtIni"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmd_pesqLanc"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txt_logParteDados"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cboTituloLog"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.CheckBox chkPrestServ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Prestador de Serviços"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   9930
         TabIndex        =   107
         Top             =   900
         Width           =   2055
      End
      Begin VB.CommandButton cmd_DRE_help 
         Height          =   585
         Left            =   -63240
         Picture         =   "Funcionarios.frx":4EA79
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   390
         Width           =   675
      End
      Begin VB.CheckBox chk_mostrarTelaPesquisaTipoFoto 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar tela de pesquisa de produto 'tipo foto'"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   12300
         TabIndex        =   99
         Top             =   4170
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   4485
         Left            =   120
         TabIndex        =   87
         Top             =   2760
         Width           =   4515
         Begin VB.TextBox Obs 
            BackColor       =   &H00E5E5E5&
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            Left            =   390
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   105
            Top             =   2640
            Width           =   3990
         End
         Begin VB.CheckBox O_Produtos 
            Appearance      =   0  'Flat
            Caption         =   "Acesso aba Cálculos"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   450
            TabIndex        =   90
            Top             =   450
            Width           =   1830
         End
         Begin VB.CheckBox O_Mov_Caixa 
            Appearance      =   0  'Flat
            Caption         =   "Efetuar movimentações manuais"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   450
            TabIndex        =   89
            Top             =   1050
            Width           =   2685
         End
         Begin VB.CheckBox chkSenhaConfirmarCRDiff 
            Appearance      =   0  'Flat
            Caption         =   "Senha do Gerente para efetuar baixas com datas ou valores diferentes do previsto"
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   450
            TabIndex        =   88
            Top             =   1650
            Width           =   3975
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Observações:"
            Height          =   195
            Left            =   180
            TabIndex        =   104
            Top             =   2370
            Width           =   1005
         End
         Begin VB.Label Label34 
            Caption         =   "Contas a Receber:"
            Height          =   195
            Left            =   180
            TabIndex        =   102
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label Label33 
            Caption         =   "Caixa:"
            Height          =   195
            Left            =   180
            TabIndex        =   100
            Top             =   840
            Width           =   585
         End
         Begin VB.Label Label32 
            Caption         =   "Cadastro de Produtos:"
            Height          =   195
            Left            =   180
            TabIndex        =   98
            Top             =   210
            Width           =   1725
         End
      End
      Begin VB.ComboBox cboTituloLog 
         BackColor       =   &H00C0FFFF&
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
         Left            =   -65670
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   390
         Width           =   3225
      End
      Begin VB.TextBox txt_logParteDados 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   -69180
         TabIndex        =   81
         Top             =   405
         Width           =   2925
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
         Left            =   -74910
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   810
         Width           =   12465
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
         Left            =   -73245
         Picture         =   "Funcionarios.frx":5000F
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   367
         Width           =   465
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
         Left            =   -71055
         Picture         =   "Funcionarios.frx":508F1
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   367
         Width           =   465
      End
      Begin VB.CommandButton cmd_tabelasDePrecos 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tabelas de Preços"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -64740
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   6810
         Width           =   2175
      End
      Begin VB.TextBox txtCDRC 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   960
         MaxLength       =   20
         TabIndex        =   70
         Top             =   2340
         Width           =   2415
      End
      Begin VB.ComboBox Sexo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "Funcionarios.frx":511D3
         Left            =   6720
         List            =   "Funcionarios.frx":511DD
         TabIndex        =   9
         Top             =   510
         Width           =   975
      End
      Begin VB.TextBox txtNomeSupervisor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   10155
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   495
         Width           =   2415
      End
      Begin VB.Data datSupervisor 
         Caption         =   "datSupervisor"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   12330
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Supervisores ORDER BY Código"
         Top             =   3810
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CheckBox chkMarketing 
         Appearance      =   0  'Flat
         Caption         =   "Responsável pelo Marketing"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12300
         TabIndex        =   10
         Top             =   3180
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.TextBox txtMargemLimiteCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   12420
         TabIndex        =   27
         Text            =   "0"
         ToolTipText     =   "Recurso utilizado para o sistema de consignação"
         Top             =   3555
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CheckBox chkDesconto 
         Appearance      =   0  'Flat
         Caption         =   "Permite ceder desconto de até (%)"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8775
         TabIndex        =   25
         Top             =   1980
         Width           =   2925
      End
      Begin MSMask.MaskEdBox Comissão_Serviço 
         Height          =   315
         Left            =   11790
         TabIndex        =   23
         Top             =   1260
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cadastro de Clientes"
         Height          =   4485
         Left            =   8760
         TabIndex        =   62
         Top             =   2760
         Width           =   3765
         Begin VB.CheckBox chkContatosEfetuadosLembrarEm 
            Appearance      =   0  'Flat
            Caption         =   "Obrigatório informar data 'Lembrar Em'"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   465
            TabIndex        =   86
            Top             =   2670
            Width           =   3105
         End
         Begin VB.CheckBox O_Serviços 
            Appearance      =   0  'Flat
            Caption         =   "Serviços"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   465
            TabIndex        =   37
            Top             =   1305
            Width           =   1020
         End
         Begin VB.CheckBox O_Conta 
            Appearance      =   0  'Flat
            Caption         =   "Conta do Cliente"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   465
            TabIndex        =   40
            Top             =   2055
            Width           =   1620
         End
         Begin VB.CheckBox O_Outras 
            Appearance      =   0  'Flat
            Caption         =   "Outras Informações"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   465
            TabIndex        =   39
            Top             =   1800
            Width           =   1770
         End
         Begin VB.CheckBox O_Cheques 
            Appearance      =   0  'Flat
            Caption         =   "Cheques e Cartões"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   465
            TabIndex        =   38
            Top             =   1560
            Width           =   1740
         End
         Begin VB.CheckBox O_Receber 
            Appearance      =   0  'Flat
            Caption         =   "Contas a Receber"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   465
            TabIndex        =   36
            Top             =   1050
            Width           =   1710
         End
         Begin VB.CheckBox O_Pagar 
            Appearance      =   0  'Flat
            Caption         =   "Contas a Pagar"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   465
            TabIndex        =   35
            Top             =   795
            Width           =   1470
         End
         Begin VB.CheckBox O_Compras 
            Appearance      =   0  'Flat
            Caption         =   "Compras / Vendas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   465
            TabIndex        =   34
            Top             =   540
            Width           =   1680
         End
         Begin VB.Label Label31 
            Caption         =   "Registro de Eventos e Contatos:"
            Height          =   195
            Left            =   240
            TabIndex        =   97
            Top             =   2430
            Width           =   2535
         End
         Begin VB.Label Label30 
            Caption         =   "Informações a serem visualizadas:"
            Height          =   195
            Left            =   240
            TabIndex        =   96
            Top             =   300
            Width           =   2595
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
         Height          =   375
         Left            =   -72000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ZZZProgramas"
         Top             =   4020
         Visible         =   0   'False
         Width           =   2415
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
         Height          =   375
         Left            =   -72000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3660
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Cargo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   960
         MaxLength       =   20
         TabIndex        =   20
         Top             =   1965
         Width           =   2415
      End
      Begin VB.TextBox Fone 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   960
         MaxLength       =   30
         TabIndex        =   17
         Top             =   1605
         Width           =   2415
      End
      Begin VB.TextBox Bairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4680
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1245
         Width           =   1575
      End
      Begin VB.TextBox Filial_Acesso 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   8790
         MaxLength       =   2
         TabIndex        =   22
         Top             =   2340
         Width           =   930
      End
      Begin VB.TextBox Endereço 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   960
         MaxLength       =   50
         TabIndex        =   11
         Top             =   870
         Width           =   5295
      End
      Begin VB.TextBox Cidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   960
         MaxLength       =   25
         TabIndex        =   13
         Top             =   1245
         Width           =   2415
      End
      Begin VB.TextBox Estado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1245
         Width           =   375
      End
      Begin VB.TextBox CEP 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6720
         MaxLength       =   9
         TabIndex        =   16
         Top             =   1245
         Width           =   975
      End
      Begin VB.TextBox CPF 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   18
         Top             =   1605
         Width           =   1455
      End
      Begin VB.TextBox Identidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   19
         Top             =   1605
         Width           =   1455
      End
      Begin VB.TextBox Carteira 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   21
         Top             =   1965
         Width           =   1455
      End
      Begin VB.TextBox Senha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   8
         PasswordChar    =   "•"
         TabIndex        =   6
         ToolTipText     =   "Seqüência de até 8 caracteres alfanuméricos (A caixa da letra é significativa)"
         Top             =   495
         Width           =   990
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   4485
         Left            =   4740
         TabIndex        =   54
         Top             =   2760
         Width           =   3915
         _Version        =   65536
         _ExtentX        =   6906
         _ExtentY        =   7911
         _StockProps     =   14
         Caption         =   "Permissões para Vendas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox chkLucroMinimoPermitido 
            Appearance      =   0  'Flat
            Caption         =   "Permite desconto maior que a margem de lucro do produto"
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   390
            TabIndex        =   103
            Top             =   1140
            Width           =   3375
         End
         Begin VB.CheckBox chkImprimirTicket 
            Appearance      =   0  'Flat
            Caption         =   "Obrigatório a impressão de tickets"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   390
            TabIndex        =   101
            Top             =   840
            Width           =   2805
         End
         Begin VB.CheckBox O_Recebe_Saídas 
            Appearance      =   0  'Flat
            Caption         =   "Recebimento"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   390
            TabIndex        =   95
            Top             =   3990
            Width           =   1425
         End
         Begin VB.CheckBox chkAllowDescProd 
            Appearance      =   0  'Flat
            Caption         =   "Desconto para produto não habilitado"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   390
            TabIndex        =   92
            Top             =   3390
            Width           =   3090
         End
         Begin VB.CheckBox chkPermiteAcharVenda 
            Appearance      =   0  'Flat
            Caption         =   "Função Achar Vendas"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   390
            TabIndex        =   91
            Top             =   570
            Width           =   2025
         End
         Begin VB.CheckBox chk_usuarioAcessoApenasTelaVendaRapida 
            Appearance      =   0  'Flat
            Caption         =   "Apenas funções mínimas"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   390
            TabIndex        =   85
            Top             =   3120
            Width           =   2085
         End
         Begin VB.CheckBox chkSenhaClear 
            Appearance      =   0  'Flat
            Caption         =   "Não precisa senha p/ limpar campos da tela"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   390
            TabIndex        =   33
            Top             =   300
            Width           =   3465
         End
         Begin VB.CheckBox chkVR_PermiteVisualizarLimiteCredito 
            Appearance      =   0  'Flat
            Caption         =   "Visualizar o limite de crédito"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   390
            TabIndex        =   32
            Top             =   2850
            Width           =   2340
         End
         Begin VB.CheckBox chkVRVisualizarEstoque 
            Appearance      =   0  'Flat
            Caption         =   "Visualizar estoque"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   390
            TabIndex        =   31
            Top             =   2610
            Width           =   1635
         End
         Begin VB.CheckBox chkVRVisualizarPreco 
            Appearance      =   0  'Flat
            Caption         =   "Visualizar preços"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   390
            TabIndex        =   30
            Top             =   2355
            Width           =   1515
         End
         Begin VB.CheckBox Recebimento 
            Appearance      =   0  'Flat
            Caption         =   "Recebimento"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   390
            TabIndex        =   28
            Top             =   1845
            Width           =   1275
         End
         Begin VB.CheckBox Clientes 
            Appearance      =   0  'Flat
            Caption         =   "Cadastrar novos clientes"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   390
            TabIndex        =   29
            Top             =   2100
            Width           =   2175
         End
         Begin VB.Label Label29 
            Caption         =   "Saídas:"
            Height          =   195
            Left            =   150
            TabIndex        =   94
            Top             =   3780
            Width           =   585
         End
         Begin VB.Label Label28 
            Caption         =   "Rápida:"
            Height          =   195
            Left            =   150
            TabIndex        =   93
            Top             =   1620
            Width           =   585
         End
      End
      Begin SSDataWidgets_B.SSDBDropDown DropDown1 
         Bindings        =   "Funcionarios.frx":511E7
         Height          =   2085
         Left            =   -72480
         TabIndex        =   61
         Top             =   1140
         Width           =   4695
         DataFieldList   =   "Descrição"
         ListAutoValidate=   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseExactRowCount=   0   'False
         ForeColorEven   =   0
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   53
         Columns.Count   =   3
         Columns(0).Width=   8017
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   6033
         Columns(1).Caption=   "Nome Programa"
         Columns(1).Name =   "Nome Programa"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome Programa"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "Número"
         Columns(2).Name =   "Número"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Número"
         Columns(2).DataType=   3
         Columns(2).FieldLen=   256
         _ExtentX        =   8281
         _ExtentY        =   3678
         _StockProps     =   77
      End
      Begin SSDataWidgets_B.SSDBGrid Grade1 
         Bindings        =   "Funcionarios.frx":511FB
         Height          =   5730
         Left            =   -74820
         TabIndex        =   60
         Top             =   1035
         Width           =   12270
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         ForeColorEven   =   0
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   53
         Columns(0).Width=   3200
         UseDefaults     =   0   'False
         _ExtentX        =   21643
         _ExtentY        =   10107
         _StockProps     =   79
         Caption         =   "Telas / Programas"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox Comissão 
         Height          =   315
         Left            =   11790
         TabIndex        =   24
         Top             =   1620
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Admissão 
         Height          =   315
         Left            =   2820
         TabIndex        =   7
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   510
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
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
      Begin MSMask.MaskEdBox Nascimento 
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   510
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   15066597
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
      Begin MSMask.MaskEdBox mskPercDesconto 
         Height          =   315
         Left            =   11790
         TabIndex        =   26
         Top             =   1980
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin SSDataWidgets_B.SSDBCombo cboSupervisor 
         Bindings        =   "Funcionarios.frx":5120F
         Height          =   315
         Left            =   9315
         TabIndex        =   12
         Top             =   510
         Width           =   795
         DataFieldList   =   "Código"
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
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Código"
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   -72375
         TabIndex        =   75
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   -74610
         TabIndex        =   76
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid grade_log 
         Height          =   5595
         Left            =   -74910
         TabIndex        =   80
         Top             =   1350
         Width           =   12465
         _ExtentX        =   21987
         _ExtentY        =   9869
         _Version        =   393216
         Rows            =   1
         Cols            =   4
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
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         Caption         =   "Título"
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
         Left            =   -66210
         TabIndex        =   83
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         Caption         =   "Parte do Texto"
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
         Left            =   -70470
         TabIndex        =   82
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
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
         Left            =   -72720
         TabIndex        =   78
         Top             =   450
         Width           =   300
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         Caption         =   "De"
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
         Left            =   -74880
         TabIndex        =   77
         Top             =   450
         Width           =   300
      End
      Begin VB.Label lblCDRC 
         AutoSize        =   -1  'True
         Caption         =   "CDRC"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   2415
         Width           =   420
      End
      Begin VB.Label lblSupervisor 
         AutoSize        =   -1  'True
         Caption         =   "Supervisor"
         Height          =   195
         Left            =   8445
         TabIndex        =   69
         Top             =   570
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00404040&
         Height          =   390
         Left            =   6240
         Top             =   2325
         Width           =   6285
      End
      Begin VB.Label Label25 
         Caption         =   "Margem excedente ao limite de crédito (%)"
         Height          =   225
         Left            =   12330
         TabIndex        =   67
         Top             =   3420
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label22 
         Caption         =   "Comissão Serviços (%)"
         Height          =   255
         Left            =   9990
         TabIndex        =   63
         Top             =   1290
         Width           =   1665
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nascimento"
         Height          =   195
         Left            =   4110
         TabIndex        =   58
         Top             =   570
         Width           =   825
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
         Height          =   195
         Left            =   6330
         TabIndex        =   57
         Top             =   570
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   4200
         TabIndex        =   55
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Admissão"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   53
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Endereço"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Cidade"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1290
         Width           =   585
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "UF"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3420
         TabIndex        =   50
         Top             =   1320
         Width           =   195
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6330
         TabIndex        =   49
         Top             =   1320
         Width           =   285
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3420
         TabIndex        =   48
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Identidade"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5400
         TabIndex        =   47
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cart. Trabalho"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5130
         TabIndex        =   46
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   570
         Width           =   450
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Permite acesso somente a filial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6450
         TabIndex        =   44
         Top             =   2415
         Width           =   2190
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "use 0 (zero) para todas as filiais"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9930
         TabIndex        =   43
         Top             =   2400
         Width           =   2355
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         Caption         =   "Comissão Vendas (%)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9990
         TabIndex        =   42
         Top             =   1650
         Width           =   1845
      End
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Nome Completo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3360
      TabIndex        =   66
      Top             =   105
      Width           =   1125
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Apelido"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      TabIndex        =   65
      Top             =   105
      Width           =   525
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   64
      Top             =   105
      Width           =   495
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   120
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "Funcionarios.frx":5122B
   End
End
Attribute VB_Name = "frmFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsParametros As Recordset
Private Num_Registro    As Variant
Private rsFuncionarios  As Recordset
Private rsAcessos       As Recordset
Private rsProgramas     As Recordset
Private rsComissões     As Recordset
Private gbLoading       As Boolean

Private gbNavigating    As Boolean
'11/03/2004 - Daniel
'Flag de indicação que é o cliente
'F. Linhares
Private m_blnFLinhares  As Boolean
'30/07/2004 - Daniel
'Esta variável monitora o uso de Supervisores
'Case: Inicialmente STC
Private m_blnSupervisor As Boolean

'18/06/2007 - Anderson
'Variável utilizada para validar o campo conceder desconto no momento da gravação
Private bolPercDesconto As Boolean

Private Sub ShowRecord()

  On Error GoTo ErrHandler

  Código.Text = rsFuncionarios("Código")
  Nome.Text = rsFuncionarios("Nome")
  Apelido.Text = rsFuncionarios("Apelido")
  Admissão.Text = gsFormatDate(rsFuncionarios("Admissão"))
  If IsDate(rsFuncionarios("Nascimento")) Then
    Nascimento.Text = CDate(rsFuncionarios("Nascimento"))
  Else
    Nascimento.Mask = ""
    Nascimento.Text = ""
    Nascimento.Mask = "##/##/####"
  End If
  Sexo.Text = Trim(rsFuncionarios("Sexo")) & ""
  Endereço.Text = Trim(rsFuncionarios("Endereço")) & ""
  '30/07/2004 - Daniel
  'Tratamento para o campo Supervisor
  'Mostrará apenas quando estiver habilitado m_blnSupervisor
  If m_blnSupervisor Then
    cboSupervisor.Text = rsFuncionarios.Fields("Supervisor").Value
    cboSupervisor_LostFocus
  End If
  '---------------------------------------------------------------
  Cidade.Text = Trim(rsFuncionarios("Cidade")) & ""
  Estado.Text = Trim(rsFuncionarios("Estado")) & ""
  CEP.Text = Trim(rsFuncionarios("CEP")) & ""
  Bairro.Text = Trim(rsFuncionarios("Bairro")) & ""
  Fone.Text = Trim(rsFuncionarios("Telefone")) & ""
  Carteira.Text = Trim(rsFuncionarios("Carteira Trabalho")) & ""
  CPF.Text = Trim(rsFuncionarios("CPF")) & ""
  Identidade.Text = Trim(rsFuncionarios("Identidade")) & ""
  Senha.Text = String(8, "*")
  Liberado.Value = -rsFuncionarios("Liberado")
  '06/06/2005 - Daniel
  'Tratamento para o campo Ativo
  'Finalidade: Capacitar ou não o user para uso do Quick Store
  chkAtivo.Value = IIf(rsFuncionarios("Ativo").Value, vbChecked, vbUnchecked)
  '-----------------------------
  O_Mov_Caixa.Value = -rsFuncionarios("Movimentar Caixa")
  Super.Value = -rsFuncionarios("Superusuário")
  Recebimento.Value = -rsFuncionarios("Recebimento")
  Clientes.Value = -rsFuncionarios("Clientes")
  Filial_Acesso.Text = rsFuncionarios("Filial Acesso")
  Comissão.Text = rsFuncionarios("Comissão") & ""
  '13/06/2004 - Daniel
  'Case........: TV Shopping
  'Finalidade..: Monitorar Funcionário Responsável pelo Departamento de Marketing.
  'A partir do login deste user o sistema atualizará a curva ABC de clientes
'''  If rsFuncionarios("Marketing").Value Then
'''    chkMarketing.Value = vbChecked
'''  Else
'''    chkMarketing.Value = vbUnchecked
'''  End If
  '-------------------------------------------------------------------------------
  
  
  If rsFuncionarios("bMostrarTelaPesquisaProdutoTipoFoto").Value Then
      chk_mostrarTelaPesquisaTipoFoto.Value = vbChecked
  Else
      chk_mostrarTelaPesquisaTipoFoto.Value = vbUnchecked
  End If
  
  If rsFuncionarios("bUsuarioAcessoApenasTelaVendaRapida").Value Then
      chk_usuarioAcessoApenasTelaVendaRapida.Value = vbChecked
  Else
      chk_usuarioAcessoApenasTelaVendaRapida.Value = vbUnchecked
  End If

  
  Comissão_Serviço.Text = rsFuncionarios("Comissão Serviço") & ""
  Cargo.Text = Trim(rsFuncionarios("Cargo")) & ""
  
  '---[ 02/06/2003 - Maikel ]---'
  ' Case      : Projeto Deluken
  ' Descrição : Adicionado o campo abaixo para guardar a margem excedente ao limite de crédito
    txtMargemLimiteCredito.Text = IIf(IsNumeric(rsFuncionarios.Fields("MargemLimiteCredito")), rsFuncionarios.Fields("MargemLimiteCredito"), 0)
  '-----------------------------'
  
  chkPrestServ.Value = IIf(rsFuncionarios("isPrestServ"), vbChecked, vbUnchecked)
  chkDesconto.Value = IIf(rsFuncionarios("bPermiteDesconto"), vbChecked, vbUnchecked)
  chkVRVisualizarEstoque.Value = -rsFuncionarios("VRVisualizarEstoque")
  chkVRVisualizarPreco.Value = -rsFuncionarios("VRVisualizarPreco")
  chkVR_PermiteVisualizarLimiteCredito.Value = -rsFuncionarios("VR_PermiteVisualizarLimiteCredito")
  '21/06/2004 - Daniel
  'Adicionado campo SenhaClear
  chkSenhaClear.Value = IIf(rsFuncionarios("SenhaClear"), vbChecked, vbUnchecked)
  
  mskPercDesconto.Text = IIf(IsNull(rsFuncionarios("nPercDesconto")), 0, rsFuncionarios("nPercDesconto"))
  
  O_Compras.Value = -rsFuncionarios("Pasta Compras")
  O_Pagar.Value = -rsFuncionarios("Pasta Pagar")
  O_Receber.Value = -rsFuncionarios("Pasta Receber")
  O_Cheques.Value = -rsFuncionarios("Pasta Cheques")
  O_Outras.Value = -rsFuncionarios("Pasta Outras")
  O_Conta.Value = -rsFuncionarios("Pasta Conta")
  O_Serviços.Value = -rsFuncionarios("Pasta Serviços")
  
  O_Produtos.Value = -rsFuncionarios("Custo Produtos")
  
  O_Recebe_Saídas.Value = -rsFuncionarios("Recebimento Saídas")
  
  
  '29/08/2003 - mpdea
  'Permissão para Achar Venda
  chkPermiteAcharVenda.Value = IIf(rsFuncionarios.Fields("PermiteAcharVenda").Value, vbChecked, vbUnchecked)
  
  '02/06/2005 - Daniel
  'Tratamento para o campo AllowDescProd
  'Finalidade explicada em AlteraDB
  chkAllowDescProd.Value = IIf(rsFuncionarios.Fields("AllowDescProd").Value, vbChecked, vbUnchecked)
  
  '29/12/2003 - Daniel
  'Senha para efetuar baixas com datas ou valores
  'diferentes dos previstos
  chkSenhaConfirmarCRDiff.Value = IIf(rsFuncionarios.Fields("SenhaConfirmarCRDiff").Value, vbChecked, vbUnchecked)
  
  '11/03/2004 - Daniel
  'Tratamento da Impressão de Tickets por funcionário
  'restrito para F. Linhares
  If m_blnFLinhares Then
    chkImprimirTicket.Value = IIf(rsFuncionarios.Fields("ImprimirTicket").Value, vbChecked, vbUnchecked)
  End If
  
  '19/10/2007 - Anderson
  'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
  chkLucroMinimoPermitido = IIf(rsFuncionarios.Fields("LucroMinimoPermitido").Value, vbChecked, vbUnchecked)
  
  Obs.Text = Trim(rsFuncionarios("Observação")) & ""
  
  '18/07/2007 - Anderson
  'Exibir campo CDRC para SadigWeb
  txtCDRC.Text = "" & rsFuncionarios("SadigWeb_CDRC").Value
  
  '19/07/2007 - Anderson
  'Implementação do campo de obrigatoriedade da data Lembrar Em
  chkContatosEfetuadosLembrarEm.Value = IIf(rsFuncionarios.Fields("ContatosEfetuadosLembrarEm").Value, vbChecked, vbUnchecked)
  
  If Tab1.Tab = 1 Then
    Call Monta_Grade
  End If
  Tab1.TabEnabled(1) = True
  
  Num_Registro = rsFuncionarios.Bookmark
  gbNavigating = False
  
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao apresentar registro de Funcionários."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)

End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpFirst"
      Call MoveFirst
    Case "miOpPrevious"
      Call MovePrevious
    Case "miOpNext"
      Call MoveNext
    Case "miOpLast"
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
    Case "miComplCopiaPermissoes"
      Call CopiaPermissoes
    Case "miComplLiberacaoTotal"
      Call LiberacaoTotal
    Case "miComplRestricaoTotal"
      Call RestricaoTotal
    Case "miComplRepFuncGeral"
      '06/06/2005 - Daniel
      'Criado validação para ver se o usuário tem
      'a permissão de acessar o Rel. de Usuários/Funcionários
      If blnRelFunc Then
        frmRelFuncGeral.Show
      Else
        MsgBox "Funcionário " & "[ " & gnUserCode & " ]" & " não possui permissão de uso deste relatório.", vbExclamation, "Quick Store"
      End If
  End Select
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case 0
          Set rsFuncionarios = db.OpenRecordset("SELECT * FROM Funcionários ORDER BY Código", dbOpenDynaset)
        Case 1
          Set rsFuncionarios = db.OpenRecordset("SELECT * FROM Funcionários ORDER BY Nome", dbOpenDynaset)
      End Select
  End Select
End Sub

Private Sub Admissão_LostFocus()
  Admissão.Text = Ajusta_Data(Admissão.Text)
End Sub

Private Sub Admissão_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Admissão.Text = frmCalendario.gsDateCalender(Admissão.Text)
  End Select
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  gbNavigating = True
  With rsFuncionarios
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  gbNavigating = True
  With rsFuncionarios
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  gbNavigating = True
  With rsFuncionarios
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With
End Sub

Private Sub MoveNext()
  On Error Resume Next
  gbNavigating = True
  With rsFuncionarios
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Private Sub RestricaoTotal()
  Dim sSql As String
  sSql = "Delete * From Acessos Where Acessos.Usuário = " & Código.Text
  Call db.Execute(sSql, dbFailOnError)
  Call Monta_Grade
End Sub

Private Sub DeleteRecord()
  Dim Resposta As Integer
  Dim Aux_Código As Double
  Dim Aux_Str As String
  Dim sSql As String

  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(238)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  If rsFuncionarios("Código") = 1 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(239)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  
 ' Set rsComissões = db.OpenRecordset("Comissões")

  Call StatusMsg("")
  Aux_Str = LoadResString(240)
  Aux_Str = Aux_Str + Chr(13)
  Aux_Str = Aux_Str & vbCrLf & "Comissões"
  Aux_Str = Aux_Str & vbCrLf & "Contas a Receber por Vendedor"
  Aux_Str = Aux_Str & vbCrLf & "Vendas por Vendedor"
  Aux_Str = Aux_Str & vbCrLf & "Tela de Saídas / Venda Rápida"
  Aux_Str = Aux_Str & vbCrLf & Chr(13)
  Aux_Str = Aux_Str + LoadResString(241)
  gsTitle = LoadResString(201)
  gsMsg = Aux_Str
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  If gnResponse = vbYes Then
    
    gbNavigating = True
    
    Call StatusMsg("Apagando acessos...")
    Call ws.BeginTrans
    
    sSql = "DELETE * FROM Acessos WHERE Usuário = " & CStr(rsFuncionarios("Código"))
    Call db.Execute(sSql, dbFailOnError)
    
    rsFuncionarios.Delete
    Num_Registro = Null
    
    Call ws.CommitTrans
    
    Call ClearScreen
    
    gbNavigating = False
  
  End If

End Sub

Private Sub CopiaPermissoes()

  Call StatusMsg("")
  gsCodigoFrom = Código.Text
  
  frmCopiaAcesso.Show vbModal
  
  If gbCopyPermissoes = False Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(247)
    gnStyle = vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Call Monta_Grade
 
End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  Dim sCPF As String
  Dim sTexto As String
  Dim nI As Integer
  Dim sCh As String
  
  Call StatusMsg("")
  
  '18/06/2007 - Anderson
  'Força a validação do campo conceder desconto
  bolPercDesconto = False
  Call mskPercDesconto_Validate(False)
  If Not bolPercDesconto Then
    Exit Sub
  End If
  
  Código.Text = gsHandleNull(Código.Text)
  If Not IsNumeric(Código.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(249)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Código.SetFocus
    Exit Sub
  End If
  
  If CInt(Código.Text) < 0 Or CInt(Código.Text) > 9999 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(249)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Código.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(Nome.Text & "")) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(250)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Nome.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(Apelido.Text & "")) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(234)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Apelido.SetFocus
    Exit Sub
  End If
  
  For nI = 1 To Len(Apelido.Text)
    sCh = UCase(Mid(Apelido.Text, nI, 1))
    If (sCh < "0" Or sCh > "9") And (sCh < "A" Or sCh > "Z") Then
      gsTitle = LoadResString(201)
      gsMsg = "Apelido do usuário possui caracter inválido. Reentre."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Apelido.SetFocus
      Exit Sub
    End If
  Next nI
  
  If Not IsDate(Admissão.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(231)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Admissão.SetFocus
    Exit Sub
  End If
  
  Comissão.Text = gsHandleNull(Comissão.Text)
  If CSng(Comissão.Text) < 0 Or CSng(Comissão.Text) > 100 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(232)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Comissão.SetFocus
    Exit Sub
  End If
  
  Comissão_Serviço.Text = gsHandleNull(Comissão_Serviço.Text)
  If CSng(Comissão_Serviço.Text) < 0 Or CSng(Comissão_Serviço.Text) > 100 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(232)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Comissão_Serviço.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(Senha.Text & "")) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(233)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Senha.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(CPF.Text)) > 0 Then
    sCPF = gsRetiraSpecialChars(CPF.Text)
    
    '15/07/2005 - Daniel
    'Adicionado regra para não permitir gravação de CPF com menos de 11 caracteres
    'ou com caracteres alfanuméricos
    If Len(sCPF) < 11 Then
      MsgBox "CPF com menos de 11 caracteres, verifique.", vbExclamation, "Erro no CPF"
      CPF.SetFocus
      Exit Sub
    End If
    '
    If Not IsNumeric(sCPF) Then
      MsgBox "O CPF contém 'letras', verifique.", vbExclamation, "Erro no CPF"
      CPF.SetFocus
      Exit Sub
    End If
    
    If bCheckCPF(sCPF) = False Then
      gsTitle = LoadResString(201)
      gsMsg = LoadResString(248)
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    End If
  End If
  
  '30/07/2004 - Daniel
  'Tratamento para o campo Supervisor
  If m_blnSupervisor Then
    If Len(txtNomeSupervisor.Text) <= 0 Then
      MsgBox "Supervisor inválido, verifique.", vbExclamation, "Quick Store"
      cboSupervisor.SetFocus
      Exit Sub
    End If
  End If
  '-------------------------------------------------------------------------
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("Gravando ...")
  
  With rsFuncionarios
  
    If IsNull(Num_Registro) Then
      .AddNew
      sTexto = "Registro inserido "
      .Fields("Código") = Código.Text
    Else
      .Edit
      sTexto = "Registro alterado "
    End If
    
    .Fields("Nome") = Nome.Text
    .Fields("Apelido") = Apelido.Text
    .Fields("Admissão") = CDate(Admissão.Text)
    .Fields("Endereço") = Endereço.Text
    '30/07/2004 - Daniel
    'Tratamento para o campo Supervisor
    If m_blnSupervisor Then
      .Fields("Supervisor").Value = CInt(cboSupervisor.Text)
    Else
      .Fields("Supervisor").Value = 0
    End If
    '-------------------------------------------------------
    .Fields("Cidade") = Cidade.Text & ""
    .Fields("Estado") = Estado.Text & ""
    .Fields("CEP") = CEP.Text & ""
    .Fields("Carteira Trabalho") = Carteira.Text & ""
    .Fields("CPF") = CPF.Text & ""
    .Fields("Identidade") = Identidade.Text & ""
    If Senha.Text <> String(8, "*") Then
      .Fields("ValorP") = CStr(CriptografaSenha(Senha.Text))
      .Fields("Senha") = Format(Date, "yyyymmdd")
    End If

    .Fields("Liberado") = Liberado.Value
    '06/06/2005 - Daniel
    'Tratamento para o campo Ativo
    'Finalidade: Capacitar ou não o user para uso do Quick Store
    .Fields("Ativo").Value = IIf(chkAtivo.Value = vbChecked, True, False)
    '-----------------------------
    .Fields("Superusuário") = Super.Value
    .Fields("Recebimento") = Recebimento.Value
    .Fields("Clientes") = Clientes.Value
    .Fields("Filial Acesso") = Val(Filial_Acesso.Text)
    '13/06/2004 - Daniel
    'Adicionado para a TV Shopping o campo Responsável pelo Marketing
'''    If chkMarketing.Value Then
'''      .Fields("Marketing").Value = True
'''    Else
'''      .Fields("Marketing").Value = False
'''    End If
    '----------------------------------------------------------------
    .Fields("Comissão") = CDbl(Comissão.Text)
    .Fields("Comissão Serviço") = CDbl(Comissão_Serviço.Text)
    .Fields("Nascimento") = Nascimento.Text
    .Fields("Telefone") = Fone.Text & ""
    .Fields("Bairro") = Bairro.Text & ""
    .Fields("Sexo") = Sexo.Text & ""
    .Fields("Cargo") = Cargo.Text & ""
    .Fields("Movimentar Caixa") = O_Mov_Caixa.Value
    .Fields("Observação") = Obs.Text & " "
    .Fields("Pasta Compras") = False
    
  '---[ 02/06/2003 - Maikel ]---'
  ' Case      : Projeto Deluken
  ' Descrição : Adicionado o campo abaixo para guardar a margem excedente ao limite de crédito
    .Fields("MargemLimiteCredito") = IIf(IsNumeric(txtMargemLimiteCredito.Text), txtMargemLimiteCredito.Text, 0)
  '-----------------------------'
    
    If chk_mostrarTelaPesquisaTipoFoto.Value = vbChecked Then
        .Fields("bMostrarTelaPesquisaProdutoTipoFoto") = True
    Else
        .Fields("bMostrarTelaPesquisaProdutoTipoFoto") = False
    End If
    
    If chk_usuarioAcessoApenasTelaVendaRapida.Value = vbChecked Then
        .Fields("bUsuarioAcessoApenasTelaVendaRapida") = True
    Else
        .Fields("bUsuarioAcessoApenasTelaVendaRapida") = False
    End If
    
    
    If chkDesconto.Value = vbChecked Then
      .Fields("nPercDesconto") = CSng(mskPercDesconto.Text)
      .Fields("bPermiteDesconto") = True
    Else
      .Fields("nPercDesconto").Value = 0
      .Fields("bPermiteDesconto") = False
    End If
    
    If chkPrestServ.Value = vbChecked Then
      .Fields("isPrestServ") = True
    Else
      .Fields("isPrestServ") = False
    End If
    
    .Fields("VRVisualizarEstoque") = -chkVRVisualizarEstoque.Value
    .Fields("VRVisualizarPreco") = -chkVRVisualizarPreco.Value
    .Fields("VR_PermiteVisualizarLimiteCredito") = -chkVR_PermiteVisualizarLimiteCredito.Value
    
    '21/06/2004 - Daniel
    'Adicionado o campo SenhaClear
    If chkSenhaClear.Value Then
      .Fields("SenhaClear").Value = True
    Else
      .Fields("SenhaClear").Value = False
    End If
    
    If O_Compras.Value = 1 Then .Fields("Pasta Compras") = True
    
    .Fields("Pasta Pagar") = False
    If O_Pagar.Value = 1 Then .Fields("Pasta Pagar") = True
    
    .Fields("Pasta Receber") = False
    If O_Receber.Value = 1 Then .Fields("Pasta Receber") = True
    
    .Fields("Pasta Cheques") = False
    If O_Cheques.Value = 1 Then .Fields("Pasta Cheques") = True
    
    .Fields("Pasta Outras") = False
    If O_Outras.Value = 1 Then .Fields("Pasta Outras") = True
    
    .Fields("Pasta Conta") = False
    If O_Conta.Value = 1 Then .Fields("Pasta Conta") = True
    
    .Fields("Pasta Serviços") = False
    If O_Serviços.Value = 1 Then .Fields("Pasta Serviços") = True
    
    .Fields("Custo Produtos") = False
    If O_Produtos.Value = 1 Then .Fields("Custo Produtos") = True
    
    .Fields("Recebimento Saídas") = False
    If O_Recebe_Saídas.Value = 1 Then .Fields("Recebimento Saídas") = True
    
  
    '29/08/2003 - mpdea
    'Permissão para Achar Venda
    .Fields("PermiteAcharVenda").Value = chkPermiteAcharVenda.Value = vbChecked
    
    '02/06/2005 - Daniel
    'Tratamento para o campo AllowDescProd
    'Finalidade explicada em AlteraDB
    If chkAllowDescProd.Value = vbChecked Then
      .Fields("AllowDescProd").Value = True
    Else
      .Fields("AllowDescProd").Value = False
    End If
    
    '29/12/2003 - Daniel
    'Senha para efetuar baixas com datas ou valores
    'diferentes dos previstos
    .Fields("SenhaConfirmarCRDiff") = False
    If chkSenhaConfirmarCRDiff.Value = 1 Then .Fields("SenhaConfirmarCRDiff") = True
    
    '11/03/2004 - Daniel
    'Tratamento da Impressão de Tickets por funcionário
    'restrito para F. Linhares
    If m_blnFLinhares Then
      If chkImprimirTicket.Value = 1 Then
        .Fields("ImprimirTicket") = True
      Else
        .Fields("ImprimirTicket") = False
      End If
    Else 'Caso não seja F. Linhares inputado False
      .Fields("ImprimirTicket").Value = False
    End If
    
    '19/10/2007 - Anderson
    'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
    If chkLucroMinimoPermitido.Value = 1 Then
      .Fields("LucroMinimoPermitido") = True
    Else
      .Fields("LucroMinimoPermitido") = False
    End If
    
    '18/07/2007 - Anderson
    'Grava as informações do CDRC para SadigWeb
    .Fields("SadigWeb_CDRC").Value = "" & txtCDRC.Text
    
    '19/07/2007 - Anderson
    'Grava as informações do campo de obrigatoriedade do campo Lembrar Em, na tela de Contatos Efetuados de Clientes / Forncedores
    If chkContatosEfetuadosLembrarEm.Value = 1 Then
      .Fields("ContatosEfetuadosLembrarEm") = True
    Else
      .Fields("ContatosEfetuadosLembrarEm") = False
    End If
    
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
  
  End With
  
  If gnUserCode = CInt(Código.Text) Then
    gbSuperUser = Super.Value
    gsUserName = UCase(Apelido.Text)
    Call StatusMsg("Reabilitando Menus ...")
    
    '27/01/2009 - mpdea
    'Adaptado para o novo menu
    'Key: Q7MENU
    frmMain.CommandBars.StatusBar.FindPane(ID_STATUSBAR_USUARIO).Text = "Usuário: " & CStr(gnUserCode) & "-" & gsUserName
    SetMenuAcesso
    
    Call StatusMsg("")
  End If
  
  StatusMsg sTexto & "com sucesso."
  Tab1.TabEnabled(1) = True
  
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
' Joga dados do funcionário para o banco do GestoPDV por causa do PAF
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
   If frmParametros.VerificaPAF = True Then
     Dim GestoBD As Database
     Dim Tecnico As Recordset
     Dim Usuarios As Recordset
 
     Set rsParametros = db.OpenRecordset("Select * from [Parâmetros Filial] Where Filial = " & gnCodFilial & ";")
                          
     Dim fso As New FileSystemObject
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists(rsParametros("BancoPDV").Value & "\Gesto.mde") Then

     Set GestoBD = OpenDatabase(rsParametros("BancoPDV").Value & "\Gesto.mde", False, False)
     Set Usuarios = GestoBD.OpenRecordset("Select * from Usuarios where Cod_Tecnico = " & Código.Text & ";")
     If Usuarios.EOF Then
       Usuarios.AddNew
       Usuarios!Nome_completo_usuario = Nome.Text
       Usuarios!Apelido_usuario = Apelido.Text
       Usuarios!Senha_usuario = Senha.Text
       Usuarios!Cod_Tecnico = Código.Text
       Usuarios!Ind_Resp_BK = 0
       Usuarios!indAltMo = -1
       Usuarios!Acrescimo_venda = -1
       Usuarios!bloqueio_exclusao = 0
       Usuarios!ind_auto_desc_venda = -1
       Usuarios!baixa_estoque = -1
       Usuarios!reintegra_estoque = -1
       Usuarios!ind_baixa_estoque_os = -1
       Usuarios!ind_estorna_estoque_os = -1
       Usuarios!bloqVisualizaMargem = 0
       Usuarios!ind_resp_estq = 0
       Usuarios.Update
     Else
       Usuarios.Edit
       Usuarios!Nome_completo_usuario = Nome.Text
       Usuarios!Apelido_usuario = Apelido.Text
       Usuarios!Senha_usuario = Senha.Text
       Usuarios.Update
     End If
     Set Tecnico = GestoBD.OpenRecordset("Select * from Tecnico where CODIGO_TECNICO = " & Código.Text & "")
     If Tecnico.EOF Then
       Tecnico.AddNew
       Tecnico!CODIGO_TECNICO = Código.Text
       Tecnico!usuario = Apelido.Text
       Tecnico!Nome = Nome.Text
       If chkAtivo.Value = 1 Then
         Tecnico!Ativo = "S"
       Else
         Tecnico!Ativo = "N"
       End If
       Tecnico!Vendedor = "S"
       Tecnico!Tecnico = "S"
       Tecnico!Atendente = "S"
       Tecnico!Ind_Retira = True
       Tecnico!Ind_Entrega = True
       Tecnico!Indpesquisa = True
       Tecnico!Geral = True
       Tecnico.Update
     Else
       Tecnico.Edit
       Tecnico!CODIGO_TECNICO = Código.Text
       Tecnico!usuario = Apelido.Text
       Tecnico!Nome = Nome.Text
       If chkAtivo.Value = 1 Then
         Tecnico!Ativo = "S"
       Else
         Tecnico!Ativo = "N"
       End If
       Tecnico!Vendedor = "S"
       Tecnico!Tecnico = "S"
       Tecnico!Atendente = "S"
       Tecnico!Ind_Retira = True
       Tecnico!Ind_Entrega = True
       Tecnico!Indpesquisa = True
       Tecnico!Geral = True
       Tecnico.Update
     End If
 
    End If
   End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
  
  
  Exit Sub

ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar atualizar cadastro de usuários."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  On Error Resume Next

End Sub


Public Sub ClearScreen()
  Call StatusMsg("")
  gbNavigating = True
  Código.Text = ""
  Nome.Text = ""
  Apelido.Text = ""
  Admissão.Mask = ""
  Admissão.Text = ""
  Admissão.Mask = "##/##/####"
  Endereço.Text = ""
  '30/07/2004 - Daniel
  'Campo Supervisor
  cboSupervisor.Text = ""
  txtNomeSupervisor.Text = ""
  '--------------------------
  Cidade.Text = ""
  Estado.Text = ""
  CEP.Text = ""
  Carteira.Text = ""
  CPF.Text = ""
  Identidade.Text = ""
  Senha.Text = ""
  Filial_Acesso.Text = "0"
  Liberado.Value = 0
  '06/06/2005 - Daniel
  'Tratamento para o campo Ativo
  'Finalidade: Capacitar ou não o user para uso do Quick Store
  chkAtivo.Value = vbUnchecked
  '-----------------------------
  Super.Value = 0
  Recebimento.Value = 0
  Clientes.Value = 0
  O_Mov_Caixa.Value = 0
  Comissão.Text = 0
  '13/06/2004 - Daniel
  'Adicionado para a TV Shopping o campo Responsável pelo Marketing
'''  chkMarketing.Value = vbUnchecked
  '----------------------------------------------------------------
  Comissão_Serviço.Text = 0
  Nascimento.Mask = ""
  Nascimento.Text = ""
  Nascimento.Mask = "##/##/####"
  Fone.Text = ""
  Bairro.Text = ""
  Sexo.Text = ""
  Cargo.Text = ""
  Obs.Text = ""
  
  '---[ 02/06/2003 - Maikel ]---'
  ' Case      : Projeto Deluken
  ' Descrição : Adicionado o campo abaixo para guardar a margem excedente ao limite de crédito
    txtMargemLimiteCredito.Text = 0
  '-----------------------------'
  
  
  chkDesconto.Value = vbUnchecked
  chkPrestServ.Value = vbUnchecked
  chkVRVisualizarEstoque.Value = vbChecked
  chkVRVisualizarPreco.Value = vbChecked
  chkVR_PermiteVisualizarLimiteCredito.Value = vbUnchecked
  '21/06/2004 - Daniel
  'Adicionado campo SenhaClear
  chkSenhaClear.Value = vbUnchecked
  
  mskPercDesconto.Text = "0"
  
  Tab1.TabEnabled(1) = False
  
  O_Compras.Value = 1
  O_Pagar.Value = 1
  O_Receber.Value = 1
  O_Cheques.Value = 1
  O_Outras.Value = 1
  O_Conta.Value = 1
  O_Serviços.Value = 1
  
  O_Produtos.Value = 1
   
  O_Recebe_Saídas.Value = 0
  
  
  '29/08/2003 - mpdea
  'Permissão para Achar Venda
  chkPermiteAcharVenda.Value = vbUnchecked
  
  '02/06/2005 - Daniel
  'Tratamento para o campo AllowDescProd
  'Finalidade explicada em AlteraDB
  chkAllowDescProd.Value = vbUnchecked
  
  '26/12/2003 - Daniel
  'Senha para efetuar baixas com datas ou valores
  'diferentes dos previstos
  chkSenhaConfirmarCRDiff.Value = vbChecked 'Marcado
  
  '11/03/2004 - Daniel
  'Tratamento da Impressão de Tickets por funcionário
  'restrito para F. Linhares
  If m_blnFLinhares Then
    chkImprimirTicket.Value = vbChecked 'Marcado
  End If
  
  '19/10/2007 - Anderson
  'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
  chkLucroMinimoPermitido.Value = vbUnchecked
  
  '18/07/2007 - Anderson
  'Limpa o campo CDRC para SadigWeb
  txtCDRC.Text = ""
  
  '19/07/2007 - Anderson
  'Limpa informações do campo de obrigatoriedade do Lembrar em, em contatos efetuados na tela de cadastro de clientes / fornecedores
  chkContatosEfetuadosLembrarEm.Value = vbUnchecked

  If Not rsFuncionarios.EOF Then
    On Error Resume Next
    rsFuncionarios.MoveFirst
    rsFuncionarios.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  
  Tab1.Tab = 0
  
  Código.SetFocus
  
End Sub

Private Sub LiberacaoTotal()
  '01/06/2005 - Daniel
  'Adicionado tratamento de erro e tratamento para transação
  Dim sSql             As String
  Dim Aux_Nome         As String
  Dim blnInTransaction As Boolean
  
  On Error GoTo TratarErro
  
  Call ws.BeginTrans
  blnInTransaction = True
  
  sSql = "Delete * From Acessos Where Usuário = " & Código.Text
  Call db.Execute(sSql, dbFailOnError)
  
  rsProgramas.MoveFirst
  Do While Not rsProgramas.EOF
    With rsAcessos
      .AddNew
      .Fields("Numero") = rsProgramas("Número")
      .Fields("Programa") = rsProgramas("Nome Programa")
      .Fields("Usuário") = Código.Text
      .Fields("Gravar") = True
      .Fields("Apagar") = True
      .Update
    End With
    rsProgramas.MoveNext
  Loop
  
  Call ws.CommitTrans
  blnInTransaction = False
  
  Call Monta_Grade
  
  Exit Sub

TratarErro:
  If blnInTransaction Then Call ws.Rollback
  blnInTransaction = False
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Atenção"
  
End Sub

Private Sub cboSupervisor_CloseUp()
  cboSupervisor.Text = cboSupervisor.Columns(0).Text
  cboSupervisor_LostFocus
End Sub

Private Sub cboSupervisor_LostFocus()
  Dim rstSupervisor As Recordset

  txtNomeSupervisor.Text = ""
  
  If Not IsNumeric(cboSupervisor.Text) Then Exit Sub

  Set rstSupervisor = db.OpenRecordset("SELECT Código, Nome FROM Supervisores WHERE Código = " & CInt(cboSupervisor.Text), dbOpenDynaset)

  With rstSupervisor
    If Not (.BOF And .EOF) Then
      txtNomeSupervisor.Text = .Fields("Nome") & ""
    End If
  End With

  rstSupervisor.Close
  Set rstSupervisor = Nothing

End Sub

Private Sub chkDesconto_Click()
  mskPercDesconto.Enabled = chkDesconto.Value
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub cmd_DRE_help_Click()
    Dim strfile As String
    Dim objHelp As clsGeral
    Set objHelp = New clsGeral
    strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
    'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
    'Call objHelp.Show(strfile, "QuickStore10Help")
    Call objHelp.Show(strfile, "QuickStore10Help", 10059)
    Set objHelp = Nothing
End Sub

Private Sub cmd_pesqLanc_Click()
On Error GoTo Erro

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
  
  If CDate(Data_Fim.Text) - CDate(Data_Ini.Text) > 15 Then
    DisplayMsg "Escolha um período de até 15 dias"
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  grade_log.Rows = 1
    
  Dim rsLog As Recordset
  Dim strSQL As String

  strSQL = "SELECT * "
  strSQL = strSQL & " FROM ZZZLog "
  strSQL = strSQL & " WHERE (Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY 00:00:00") & "# "
  strSQL = strSQL & "AND #" & Format(Data_Fim.Text, "MM/DD/YYYY 23:59:59") & "#) "
  
  If Trim(cboTituloLog.Text) <> "" Then
      strSQL = strSQL & "AND tipo = '" & cboTituloLog.Text & "' "
  End If
  
  If txt_logParteDados.Text <> "" Then
      strSQL = strSQL & "AND texto like '*" & txt_logParteDados.Text & "*' "
  End If
  
  strSQL = strSQL & " Order by Data"
  
  Set rsLog = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    
  If Not (rsLog.EOF And rsLog.BOF) Then
    rsLog.MoveFirst
  End If
  While Not rsLog.EOF

      grade_log.AddItem 0 & vbTab & rsLog.Fields("Data").Value & vbTab & _
                              rsLog.Fields("tipo").Value & vbTab & _
                              rsLog.Fields("Texto").Value

      rsLog.MoveNext
  Wend
  rsLog.Close
  Set rsLog = Nothing
  
  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da grade...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
    
End Sub

Private Sub cmd_tabelasDePrecos_Click()
  
  Dim bm As Variant
  Dim obj_formPreco As Form
  
  Set obj_formPreco = New frmAcessosTabelasDePrecosProdutos
    
  obj_formPreco.sUsuario = Código.Text
  obj_formPreco.Show vbModal
  
  Set obj_formPreco = Nothing
End Sub

Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
'    Call O_Código_Click
    Call GetNewCode(Me, rsFuncionarios, 9999)
  End If
End Sub

Private Sub Código_LostFocus()

  If IsNull(Código.Text) Then Exit Sub
  If Not IsNumeric(Código.Text) Then Exit Sub
  If Val(Código.Text) < 1 Then Exit Sub
  If Val(Código.Text) > 9999 Then Exit Sub

  With rsFuncionarios
    .FindFirst "Código = " & CInt(Código.Text)
    If Not .NoMatch Then
      Num_Registro = rsFuncionarios.Bookmark
      gbNavigating = True
      Call ShowRecord
    End If
  End With

End Sub

Private Sub Comissão_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub Comissão_Serviço_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub DropDown1_CloseUp()
  Dim Colu As Integer
  
  Colu = Grade1.Col
  Grade1.Columns(Colu).Text = DropDown1.Columns(1).Text
  Grade1.Columns(5).Text = DropDown1.Columns(2).Text
  
End Sub

Private Sub Filial_Acesso_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
      'Call objHelp.Show(strfile, "QuickStore10Help")
      Call objHelp.Show(strfile, "QuickStore10Help", 10009)
      Set objHelp = Nothing
  Else
      Call HandleKeyDown(KeyCode, Shift)
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
  If KeyAscii = 13 Then
     SendKeys "{Tab}"
     KeyAscii = 0
 End If
End Sub

Private Sub Form_Load()

  gbLoading = True
  
  Screen.MousePointer = vbHourglass
  
  KeyPreview = True
  
  Call CenterForm(Me)
  
  ActiveBar1.Tools("miOpOrdem").CBList.Clear
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 0, "Por Código"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 1, "Por Nome"
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  
  DoEvents

  ActiveBar1.Tools("miComplCopiaPermissoes").Enabled = False
  ActiveBar1.Tools("miComplLiberacaoTotal").Enabled = False
  ActiveBar1.Tools("miComplRestricaoTotal").Enabled = False
  ActiveBar1.Tools("miOpSearch").Enabled = False
  ActiveBar1.Refresh
  
  Dim rsZZZLog As Recordset
  Set rsZZZLog = db.OpenRecordset("SELECT distinct(Tipo) FROM ZZZLog", dbOpenDynaset)
  
  If rsZZZLog.RecordCount > 0 Then
      rsZZZLog.MoveFirst
      
      cboTituloLog.AddItem ""
      While Not rsZZZLog.EOF
          cboTituloLog.AddItem rsZZZLog.Fields(0).Value
          rsZZZLog.MoveNext
      Wend
  End If
  rsZZZLog.Close
  Set rsZZZLog = Nothing
  
  
  Set rsFuncionarios = db.OpenRecordset("SELECT * FROM Funcionários ORDER BY Código", dbOpenDynaset)
  Set rsAcessos = db.OpenRecordset("SELECT * FROM Acessos ORDER BY Numero, Usuário", dbOpenDynaset)
  Set rsProgramas = db.OpenRecordset("SELECT * FROM ZZZProgramas ORDER BY Número", dbOpenDynaset)
  
  Tab1.TabEnabled(1) = False
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  '29/07/2004 - Daniel
  'Adicionado data para o campo Supervisor
  datSupervisor.DatabaseName = gsQuickDBFileName
  
  '-----------------------------------------------------------------
  '16/01/2003 - mpdea
  'Modo limitado do Quick Store
  'Opções permitidas somente em modo full
  Super.Visible = gblnQuickFull
  Tab1.TabVisible(1) = gblnQuickFull
  Label15.Visible = gblnQuickFull: Comissão.Visible = gblnQuickFull
  Label22.Visible = gblnQuickFull: Comissão_Serviço.Visible = gblnQuickFull
  '-----------------------------------------------------------------
  
  '11/03/2004 - Daniel
  'Tratamento de impressão de Tickets no Manutenção
  'do Contas a Receber obrigatório ou não para os funcionários
  'Case: F. Linhares
  m_blnFLinhares = CheckSerialCaseMod("QS37818-990")
  chkImprimirTicket.Visible = m_blnFLinhares
  '------------------------------------------------------------
  
  '19/10/2007 - Anderson
  'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
  chkLucroMinimoPermitido.Visible = g_bolLucroMinimoClasse
  
  
  '13/07/2004 - Daniel
  'Caso seja TV Shopping o chkMarketing ficará visível
'''  chkMarketing.Visible = CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959")
  '-------------------------------------------------------------------------------------
  
  '30/07/2004 - Daniel
  'Habilitar ou não o relacionamento entre funcionários
  'e seus respectivos supervisores
  m_blnSupervisor = CheckSerialCaseMod("QS39823-684")
  
  'Visibilidade dos Objetos
  lblSupervisor.Visible = m_blnSupervisor
  cboSupervisor.Visible = m_blnSupervisor
  txtNomeSupervisor.Visible = m_blnSupervisor
  '----------------------------------------------------
  
  '18/07/2007 - Anderson
  'Verifica se o cliente está habilitado para utilizar sistema da SadigWeb
  txtCDRC.Visible = g_blnSadigWeb
  lblCDRC.Visible = g_blnSadigWeb
  
  Me.Show
  DoEvents
  
  Call ActiveBarLoadToolTips(Me)

  Call ClearScreen
  
  gbLoading = False
  
  grade_log.ColWidth(0) = 10
  grade_log.ColWidth(1) = 1670
  grade_log.ColWidth(2) = 1670
  grade_log.ColWidth(3) = 9500

  grade_log.Row = 0
  grade_log.TextMatrix(0, 1) = "Dt Log"
  grade_log.TextMatrix(0, 2) = "Tipo"
  grade_log.TextMatrix(0, 3) = "Texto"
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call StatusMsg("")
  rsFuncionarios.Close
  rsAcessos.Close
  rsProgramas.Close
  Set rsFuncionarios = Nothing
  Set rsAcessos = Nothing
  Set rsProgramas = Nothing
End Sub

Private Sub Grade1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant

  Grade1.Columns(0).Text = Val(Código.Text)
  
  If ColIndex = 1 Then
    Aux = Grade1.Columns(5).Text
    If IsNull(Aux) Then
      Cancel = True
      Exit Sub
    End If
  
    If Aux = "" Then
      Cancel = True
      Exit Sub
    End If
  
    rsAcessos.FindFirst "Numero = " & Aux & " And Usuário = " & Val(Código.Text)
    If Not rsAcessos.NoMatch Then
      Beep
      gsTitle = LoadResString(201)
      gsMsg = LoadResString(235)
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Cancel = True
      Exit Sub
    End If
  End If
  
End Sub

Private Sub Grade1_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  If Not bGridBeforeDelete() Then
    Cancel = True
  End If
End Sub


Private Sub Liberado_Click()
  Dim rs As Recordset
  Dim sSql As String
  If Liberado.Value = 0 And Not gbNavigating Then
    sSql = "SELECT Count(*) as nCount FROM Funcionários WHERE Liberado = True"
    Set rs = db.OpenRecordset(sSql, dbOpenDynaset)
    If rs("nCount").Value <= 1 And Not gbLoading Then
      gsTitle = LoadResString(201)
      gsMsg = "Nenhum usuário restará como Liberado. Este usuário continuará como Liberado."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Liberado.Value = 1
    End If
    rs.Close
    Set rs = Nothing
  End If
End Sub

Private Sub mskPercDesconto_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub mskPercDesconto_Validate(Cancel As Boolean)
  Dim nMaxDesc As Single
  Dim nFilial As Byte
  
  'Verifica a qual filial deve ser estimado o máx. de desconto
  If IsNumeric(Filial_Acesso.Text) And Filial_Acesso.Text <> "0" Then
    nFilial = CByte(Filial_Acesso.Text)
  Else
    nFilial = gnCodFilial
  End If
  nMaxDesc = gvGetValueInTable("[Parâmetros Filial]", "[VR Desconto]", ftNumero, "Filial", ftNumero, nFilial)
  If CSng(mskPercDesconto.Text) > nMaxDesc Then
    Call DisplayMsg("Desconto superior aos " & Format(nMaxDesc / 100, "#.#%") & "permitido em Parâmetros da Filial nº " & nFilial & ".")
    Call SelectAllText(mskPercDesconto, True)
    Cancel = True
    '18/06/2007 - Anderson
    'Variável utilizada para validar o campo no momento da gravação do registro.
    bolPercDesconto = False
  Else
    '18/06/2007 - Anderson
    'Variável utilizada para validar o campo no momento da gravação do registro.
    bolPercDesconto = True
  End If
  
End Sub

Private Sub Nascimento_LostFocus()
  Nascimento.Text = Ajusta_Data(Nascimento.Text)
End Sub

Private Sub Nascimento_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Nascimento.Text = frmCalendario.gsDateCalender(Nascimento.Text)
  End Select
End Sub


Private Sub Sexo_LostFocus()
  If IsNull(Sexo.Text) Then Exit Sub
  If Sexo.Text = "" Then Exit Sub
  Sexo = UCase$(Sexo.Text)
End Sub


Private Sub Super_Click()
  Dim rs As Recordset
  Dim sSql As String
  
  If Super.Value = 0 And Not gbNavigating Then
    sSql = "SELECT Count(*) as nCount FROM Funcionários WHERE Superusuário = True"
    Set rs = db.OpenRecordset(sSql, dbOpenDynaset)
    If rs("nCount").Value <= 1 Then
      gsTitle = LoadResString(201)
      gsMsg = "Nenhum usuário restará como Superusuário. Este usuário continuará como Superusuário."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Super.Value = 1
    End If
    rs.Close
    Set rs = Nothing
  End If
  
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
  If Tab1.Tab = 1 Then
    ActiveBar1.Tools("miComplCopiaPermissoes").Enabled = True
    ActiveBar1.Tools("miComplLiberacaoTotal").Enabled = True
    ActiveBar1.Tools("miComplRestricaoTotal").Enabled = True
    Call Monta_Grade
  Else
    ActiveBar1.Tools("miComplCopiaPermissoes").Enabled = False
    ActiveBar1.Tools("miComplLiberacaoTotal").Enabled = False
    ActiveBar1.Tools("miComplRestricaoTotal").Enabled = False
  End If
End Sub

'01/06/2005 - mpdea
'Incluído tratamento de erro
Sub Monta_Grade()
  Dim rsAcessosAux As Recordset
  Dim sSql As String
  
  '01/06/2005 - mpdea
  'Desativa o desenho do grid
  Grade1.Redraw = False
  
  sSql = "SELECT Usuário, Programa, ZZZProgramas.Descrição, Gravar, Apagar, Numero FROM [Acessos]"
  sSql = sSql + " INNER JOIN ZZZProgramas On Acessos.Programa = ZZZProgramas.[Nome Programa]"
  sSql = sSql + " WHERE Usuário = " + Código.Text
  sSql = sSql + " ORDER BY Programa"
'  If O1_Data.Value = True Then sSql = sSql + " ORDER By Dia"
'  If O1_Produto.Value = True Then sSql = sSql + " ORDER By Produtos.Nome"
  
  Set rsAcessosAux = db.OpenRecordset(sSql, dbOpenDynaset)

  On Error Resume Next
  Grade1.DataMode = 1

  Set Data1.Recordset = rsAcessosAux
  On Error GoTo 0

  Grade1.DataMode = 0

  Grade1.ReBind
  Grade1.Columns(1).DropDownHwnd = DropDown1.hwnd
  Grade1.Columns(0).Visible = False
  Grade1.Columns(1).Width = 3600
  Grade1.Columns(2).Width = 4600
  Grade1.Columns(3).Style = 2
  Grade1.Columns(3).Width = 700
  Grade1.Columns(4).Style = 2
  Grade1.Columns(4).Width = 700
  Grade1.Columns(5).Visible = False
  Grade1.Columns(2).Locked = True
  
  '01/06/2005 - mpdea
  'Ativa o desenho do grid
  Grade1.Redraw = True
  
  Exit Sub
  
ErrHandler:
  Grade1.Redraw = True
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Function blnRelFunc() As Boolean
  '06/06/2005 - Daniel
  'Criado validação para ver se o usuário tem
  'a permissão de acessar o Rel. de Usuários/Funcionários
  Dim rstFunc As Recordset
  
  On Error GoTo TratarErro

  Set rstFunc = db.OpenRecordset("SELECT * FROM Acessos WHERE Programa = '" & "RELATÓRIO DE USUÁRIOS/FUNCIONÁRIOS" & "'" & " AND Usuário = " & gnUserCode & " AND Gravar AND Apagar", dbOpenDynaset)

  If rstFunc.RecordCount = 0 Then
    blnRelFunc = False
  Else
    blnRelFunc = True
  End If

  rstFunc.Close
  Set rstFunc = Nothing

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function
