VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCliFor 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFA324&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Clientes e Fornecedores"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   330
   ClientWidth     =   15030
   FillColor       =   &H00F7F7F7&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   20
   Icon            =   "CliFor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   9330
   ScaleWidth      =   15030
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   12930
      ScaleHeight     =   645
      ScaleWidth      =   1905
      TabIndex        =   198
      Top             =   8910
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Fax 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   13365
      MaxLength       =   15
      TabIndex        =   188
      Top             =   435
      Width           =   1605
   End
   Begin VB.TextBox Fone2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   13365
      MaxLength       =   43
      TabIndex        =   187
      Top             =   75
      Width           =   1605
   End
   Begin VB.TextBox Fone1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   11745
      MaxLength       =   43
      TabIndex        =   186
      Top             =   75
      Width           =   1605
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   7980
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   870
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   14076
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      BackColor       =   16753444
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Cadastrais"
      TabPicture(0)   =   "CliFor.frx":4E95A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label20"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label22"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label25"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "L_CGC"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "L_Inscrição"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label37"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label39"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label36"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label41"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label42"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Nome_Transp"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label46"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Nome_Vendedor"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label26"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblTitle(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblTitle(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblTitle(71)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblTitle(70)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblTitle(69)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblIndicadorIE"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label58"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Combo_Vendedor"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "mskDatNasc"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Endereço"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Complemento"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Bairro"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "CGC"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Insc"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "CEP"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "B_Expande"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "email"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "home_page"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Comentários"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Combo_Transp"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "grdContatos"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Frame4"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cboTabelaPrecos"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Frame2"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Frame3"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "frameSex"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "chkAgencia"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Estado"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtEnderecoNumero"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtPais"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtCNAE"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtInscricaoSuframa"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtInscricaoMunicipal"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cboIndicadorIE"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cboSadigWebTipo"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cbo_cidades"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "cmd_carregaEnderecoCEP"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).ControlCount=   54
      TabCaption(1)   =   "&Eventos"
      TabPicture(1)   =   "CliFor.frx":4E976
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label56"
      Tab(1).Control(1)=   "Label57"
      Tab(1).Control(2)=   "txtContatosEfetuadosLembrarEm"
      Tab(1).Control(3)=   "grdContatosEfetuados"
      Tab(1).Control(4)=   "cmdContatoEfetuadoNovo"
      Tab(1).Control(5)=   "cmdContatosEfetuadosSalvar"
      Tab(1).Control(6)=   "chkContatoEfetuadoPendencia"
      Tab(1).Control(7)=   "txtContatosEfetuadosData"
      Tab(1).Control(8)=   "txtContatosEfetuadosDescricao"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Mala Direta"
      TabPicture(2)   =   "CliFor.frx":4E992
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "Frame7"
      Tab(2).Control(3)=   "Data7"
      Tab(2).Control(4)=   "fraIPI"
      Tab(2).Control(5)=   "fraMsg"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Bens &Numeráveis"
      TabPicture(3)   =   "CliFor.frx":4E9AE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdNumeravel"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "C&rédito"
      TabPicture(4)   =   "CliFor.frx":4E9CA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fraSacadorAvalista"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "txtDiaBaseConsignacao"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "C_Empresa"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "C_Salário"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "C_Cargo"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "C_Contratação"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "C_Telefone"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "C_C_Empresa"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "C_C_Salário"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "C_C_Cargo"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "C_C_Contratação"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "C_C_Telefone"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "C_C_Nome"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "C_C_RG"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "C_Obs"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "Desconto"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "mskDatAbertura"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "mskPercentualLimite"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "Label44"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "Label34"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "Label48"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "Label47"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "Label38"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "Label1"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "Label2"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "Label3"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "Label4"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "Label5"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "Label6"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "Label7"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).Control(31)=   "Label8"
      Tab(4).Control(31).Enabled=   0   'False
      Tab(4).Control(32)=   "Label9"
      Tab(4).Control(32).Enabled=   0   'False
      Tab(4).Control(33)=   "Label10"
      Tab(4).Control(33).Enabled=   0   'False
      Tab(4).Control(34)=   "Label11"
      Tab(4).Control(34).Enabled=   0   'False
      Tab(4).Control(35)=   "Label12"
      Tab(4).Control(35).Enabled=   0   'False
      Tab(4).Control(36)=   "Label13"
      Tab(4).Control(36).Enabled=   0   'False
      Tab(4).ControlCount=   37
      TabCaption(5)   =   "Outros &Dados"
      TabPicture(5)   =   "CliFor.frx":4E9E6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grdCaract"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Loja Virtual"
      TabPicture(6)   =   "CliFor.frx":4EA02
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblTitle(2)"
      Tab(6).Control(1)=   "lblTitle(1)"
      Tab(6).Control(2)=   "lblTitle(0)"
      Tab(6).Control(3)=   "lblTitle(3)"
      Tab(6).Control(4)=   "lblTitle(4)"
      Tab(6).Control(5)=   "lblTitle(5)"
      Tab(6).Control(6)=   "sspWebOrigem"
      Tab(6).Control(7)=   "sspWebEMailLoja"
      Tab(6).Control(8)=   "sspWebEMailMerco"
      Tab(6).Control(9)=   "sspWebBonus"
      Tab(6).Control(10)=   "sspWebDataCadastro"
      Tab(6).Control(11)=   "sspWebShopperID"
      Tab(6).ControlCount=   12
      TabCaption(7)   =   "Cla&ssificação"
      TabPicture(7)   =   "CliFor.frx":4EA1E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame8"
      Tab(7).ControlCount=   1
      Begin VB.CommandButton cmd_carregaEnderecoCEP 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Endereço CEP"
         Height          =   435
         Left            =   13830
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1260
         Width           =   795
      End
      Begin VB.Frame Frame5 
         Caption         =   "Limite de Crédito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74700
         TabIndex        =   191
         Top             =   390
         Width           =   7365
         Begin VB.CommandButton cmdDocLimCred 
            Caption         =   "Gerar Documento de Limite de Crédito"
            Height          =   645
            Left            =   5250
            TabIndex        =   199
            Top             =   990
            Width           =   2025
         End
         Begin MSMask.MaskEdBox Limite 
            Height          =   375
            Left            =   1920
            TabIndex        =   192
            Top             =   330
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskEdit_limiteCreditoDisponivel 
            Height          =   375
            Left            =   1920
            TabIndex        =   193
            Top             =   1230
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648384
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskEdit_limiteCreditoUtilizado 
            Height          =   375
            Left            =   1920
            TabIndex        =   196
            Top             =   780
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12632319
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Utilizado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   690
            TabIndex        =   197
            Top             =   810
            Width           =   720
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Concedido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   690
            TabIndex        =   195
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Disponível"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   690
            TabIndex        =   194
            Top             =   1260
            Width           =   855
         End
      End
      Begin VB.ComboBox cbo_cidades 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   6090
         TabIndex        =   24
         Text            =   "cbo_cidades"
         Top             =   1305
         Width           =   4155
      End
      Begin VB.ComboBox cboSadigWebTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "CliFor.frx":4EA3A
         Left            =   12450
         List            =   "CliFor.frx":4EA5C
         TabIndex        =   43
         Top             =   2970
         Width           =   2160
      End
      Begin VB.ComboBox cboIndicadorIE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "CliFor.frx":4EB4E
         Left            =   10440
         List            =   "CliFor.frx":4EB5B
         TabIndex        =   35
         Text            =   "cboIndicadorIE"
         Top             =   2130
         Width           =   4170
      End
      Begin VB.TextBox txtInscricaoMunicipal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4995
         MaxLength       =   20
         TabIndex        =   33
         Top             =   2130
         Width           =   1605
      End
      Begin VB.TextBox txtInscricaoSuframa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1215
         MaxLength       =   9
         TabIndex        =   40
         Top             =   2970
         Width           =   2175
      End
      Begin VB.TextBox txtCNAE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   3975
         MaxLength       =   10
         TabIndex        =   41
         Top             =   2970
         Width           =   795
      End
      Begin VB.TextBox txtPais 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   10650
         MaxLength       =   64
         TabIndex        =   25
         Top             =   1305
         Width           =   1305
      End
      Begin VB.TextBox txtEnderecoNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7650
         MaxLength       =   10
         TabIndex        =   20
         Top             =   900
         Width           =   1110
      End
      Begin VB.TextBox txtContatosEfetuadosDescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   -69975
         MultiLine       =   -1  'True
         TabIndex        =   175
         Top             =   480
         Width           =   9645
      End
      Begin MSMask.MaskEdBox txtContatosEfetuadosData 
         Height          =   300
         Left            =   -71355
         TabIndex        =   174
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   480
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
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
      Begin VB.CheckBox chkContatoEfetuadoPendencia 
         Appearance      =   0  'Flat
         Caption         =   "Pendencia"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71355
         TabIndex        =   178
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdContatosEfetuadosSalvar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Salvar"
         Height          =   420
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   960
         Width           =   2400
      End
      Begin VB.CommandButton cmdContatoEfetuadoNovo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Novo"
         Height          =   420
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   480
         Width           =   2400
      End
      Begin VB.Frame fraSacadorAvalista 
         Caption         =   "Sacador / Avalista"
         Height          =   810
         Left            =   -74700
         TabIndex        =   169
         ToolTipText     =   "Estes campos possuem finalidade para correspondências com o aplicativo CNAB"
         Top             =   6990
         Width           =   14160
         Begin VB.OptionButton optCNPJ 
            Caption         =   "CNPJ"
            Height          =   255
            Left            =   6360
            TabIndex        =   82
            Top             =   450
            Width           =   735
         End
         Begin VB.OptionButton optCPF 
            Caption         =   "CPF"
            Height          =   195
            Left            =   6360
            TabIndex        =   81
            Top             =   255
            Width           =   615
         End
         Begin VB.TextBox txtCPFSacadorAvalista 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   7200
            MaxLength       =   20
            TabIndex        =   83
            Top             =   300
            Width           =   2055
         End
         Begin VB.TextBox txtNomeSacadorAvalista 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   900
            MaxLength       =   40
            TabIndex        =   80
            Top             =   330
            Width           =   5340
         End
         Begin VB.Label lblNomeSacadorAvalista 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   390
            TabIndex        =   170
            Top             =   390
            Width           =   405
         End
      End
      Begin VB.ComboBox Estado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "CliFor.frx":4EC42
         Left            =   4590
         List            =   "CliFor.frx":4EC9A
         TabIndex        =   23
         ToolTipText     =   "Caso o Estado do Cli/For seja fora do Brasil, favor escolher a última opção 'EX'"
         Top             =   1320
         Width           =   870
      End
      Begin VB.CheckBox chkAgencia 
         Appearance      =   0  'Flat
         Caption         =   "Agência Publicidade"
         ForeColor       =   &H00666666&
         Height          =   330
         Left            =   4935
         TabIndex        =   46
         Top             =   2970
         Width           =   1860
      End
      Begin VB.Frame Frame8 
         Caption         =   "Classificação do Cliente "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6165
         Left            =   -74880
         TabIndex        =   161
         Top             =   525
         Width           =   14610
         Begin VB.TextBox txtGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2655
            Locked          =   -1  'True
            TabIndex        =   158
            Top             =   3030
            Width           =   3870
         End
         Begin MSMask.MaskEdBox mskDinheiroBoletos 
            Height          =   315
            Left            =   2655
            TabIndex        =   154
            Top             =   1050
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,###,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCheques 
            Height          =   315
            Left            =   2655
            TabIndex        =   155
            Top             =   1545
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,###,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCartoes 
            Height          =   315
            Left            =   2655
            TabIndex        =   156
            Top             =   2040
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,###,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskTotRecebido 
            Height          =   315
            Left            =   2655
            TabIndex        =   157
            Top             =   2535
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,###,###,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Pertence ao Grupo"
            Height          =   195
            Left            =   240
            TabIndex        =   168
            Top             =   3075
            Width           =   1350
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Total Recebido"
            Height          =   195
            Left            =   240
            TabIndex        =   167
            Top             =   2580
            Width           =   1065
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Recebimentos efetuados até"
            Height          =   195
            Left            =   240
            TabIndex        =   166
            Top             =   480
            Width           =   2070
         End
         Begin VB.Label lblDataAtual 
            AutoSize        =   -1  'True
            Caption         =   "Data_Atual"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2655
            TabIndex        =   165
            Top             =   480
            Width           =   810
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Total em Cartões"
            Height          =   195
            Left            =   240
            TabIndex        =   164
            Top             =   2085
            Width           =   1230
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Total em Cheques"
            Height          =   195
            Left            =   240
            TabIndex        =   163
            Top             =   1590
            Width           =   1290
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Total em Dinheiro ou Boletos"
            Height          =   195
            Left            =   240
            TabIndex        =   162
            Top             =   1095
            Width           =   2040
         End
      End
      Begin VB.Frame fraMsg 
         Appearance      =   0  'Flat
         Caption         =   "Obs. referente Isenção IPI"
         ForeColor       =   &H00000000&
         Height          =   2100
         Left            =   -74865
         TabIndex        =   146
         Top             =   5640
         Width           =   6795
         Begin VB.TextBox txtMensagemIPI 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   165
            MaxLength       =   70
            TabIndex        =   60
            ToolTipText     =   "O máximo de caracteres para este campo é 70."
            Top             =   1260
            Width           =   6375
         End
         Begin VB.Label lblMsg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Caso haja Isenção de IPI deste Cliente / Fornecedor, crie uma mensagem personalizada para impressão de NF."
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   180
            TabIndex        =   147
            Top             =   690
            Width           =   6375
         End
      End
      Begin VB.Frame fraIPI 
         Appearance      =   0  'Flat
         Caption         =   "Isenção de IPI"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   -71400
         TabIndex        =   145
         Top             =   4725
         Width           =   3345
         Begin VB.CheckBox chkIsencaoIPI 
            Appearance      =   0  'Flat
            Caption         =   "Isento"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   570
            TabIndex        =   59
            Top             =   330
            Width           =   855
         End
      End
      Begin VB.Frame frameSex 
         Appearance      =   0  'Flat
         Caption         =   "Sexo"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   11055
         TabIndex        =   16
         Top             =   320
         Width           =   3555
         Begin VB.OptionButton optFeminino 
            Appearance      =   0  'Flat
            Caption         =   "Feminino"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2010
            TabIndex        =   18
            Top             =   210
            Width           =   1020
         End
         Begin VB.OptionButton optMasculino 
            Appearance      =   0  'Flat
            Caption         =   "Masculino"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   450
            TabIndex        =   17
            Top             =   210
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Pessoa Jurídica / Física"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   7155
         TabIndex        =   13
         Top             =   320
         Width           =   3870
         Begin VB.OptionButton O_Jurídica 
            Appearance      =   0  'Flat
            Caption         =   "Jurídica"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   210
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton O_Física 
            Appearance      =   0  'Flat
            Caption         =   "Física"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2205
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   210
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   165
         TabIndex        =   8
         Top             =   315
         Width           =   6960
         Begin VB.OptionButton O_Outros 
            Appearance      =   0  'Flat
            Caption         =   "Outros"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5490
            TabIndex        =   12
            Top             =   210
            Width           =   855
         End
         Begin VB.OptionButton O_Revendedor 
            Appearance      =   0  'Flat
            Caption         =   "Revenda"
            CausesValidation=   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3915
            TabIndex        =   11
            Top             =   210
            Width           =   1215
         End
         Begin VB.OptionButton O_Fornecedor 
            Appearance      =   0  'Flat
            Caption         =   "Fornecedor"
            CausesValidation=   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2250
            TabIndex        =   10
            Top             =   210
            Width           =   1215
         End
         Begin VB.OptionButton O_Cliente 
            Appearance      =   0  'Flat
            Caption         =   "Cliente"
            CausesValidation=   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   405
            TabIndex        =   9
            Top             =   210
            Value           =   -1  'True
            Width           =   820
         End
         Begin VB.Label lblClienteWEB 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WEB"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00666666&
            Height          =   225
            Left            =   1305
            TabIndex        =   123
            Top             =   225
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.ComboBox cboTabelaPrecos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "CliFor.frx":4ED0E
         Left            =   11235
         List            =   "CliFor.frx":4ED10
         TabIndex        =   39
         Top             =   2580
         Width           =   3375
      End
      Begin VB.Data Data7 
         Caption         =   "Data7"
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
         Left            =   -69510
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Conta"
         Top             =   495
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         Caption         =   "Informações sobre cobrança"
         ForeColor       =   &H00000000&
         Height          =   4185
         Left            =   -74880
         TabIndex        =   133
         Top             =   405
         Width           =   6795
         Begin VB.TextBox CEP_Cob 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   5310
            MaxLength       =   9
            TabIndex        =   57
            Top             =   3315
            Width           =   1230
         End
         Begin VB.TextBox Endereço_Cob 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1290
            MaxLength       =   50
            TabIndex        =   52
            Top             =   2325
            Width           =   5250
         End
         Begin VB.TextBox Complemento_Cob 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   53
            Top             =   2820
            Width           =   1935
         End
         Begin VB.TextBox Bairro_Cob 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   4095
            MaxLength       =   20
            TabIndex        =   54
            Top             =   2820
            Width           =   2445
         End
         Begin VB.TextBox Estado_Cob 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   4095
            MaxLength       =   2
            TabIndex        =   56
            Top             =   3315
            Width           =   705
         End
         Begin VB.TextBox Cidade_Cob 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   55
            Top             =   3315
            Width           =   1935
         End
         Begin VB.TextBox Boleto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2055
            MaxLength       =   50
            TabIndex        =   51
            Top             =   945
            Width           =   4485
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Cobrança 
            Bindings        =   "CliFor.frx":4ED12
            DataSource      =   "Data7"
            Height          =   285
            Left            =   2070
            TabIndex        =   50
            Top             =   450
            Width           =   1215
            DataFieldList   =   "Descrição"
            ListAutoValidate=   0   'False
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
            Columns.Count   =   3
            Columns(0).Width=   7832
            Columns(0).Caption=   "Descrição"
            Columns(0).Name =   "Descrição"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Descrição"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Conta"
            Columns(1).Name =   "Conta"
            Columns(1).CaptionAlignment=   0
            Columns(1).DataField=   "Conta"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   1773
            Columns(2).Caption=   "Código"
            Columns(2).Name =   "Código"
            Columns(2).Alignment=   1
            Columns(2).CaptionAlignment=   1
            Columns(2).DataField=   "Código"
            Columns(2).DataType=   2
            Columns(2).FieldLen=   256
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483630
            BackColor       =   12648447
         End
         Begin VB.Label Label40 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Mensagem para Boleto"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   210
            TabIndex        =   143
            Top             =   945
            Width           =   1815
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   210
            TabIndex        =   142
            Top             =   2340
            Width           =   855
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3495
            TabIndex        =   141
            Top             =   2835
            Width           =   495
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   210
            TabIndex        =   140
            Top             =   2835
            Width           =   1095
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3450
            TabIndex        =   139
            Top             =   3330
            Width           =   615
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   210
            TabIndex        =   138
            Top             =   3330
            Width           =   855
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            Caption         =   $"CliFor.frx":4ED26
            ForeColor       =   &H00000000&
            Height          =   630
            Left            =   210
            TabIndex        =   137
            Top             =   1440
            Width           =   6330
         End
         Begin VB.Label Label43 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conta para Cobrança"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   210
            TabIndex        =   136
            Top             =   450
            Width           =   1695
         End
         Begin VB.Label Nome_Cobrança 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3375
            TabIndex        =   135
            Top             =   450
            Width           =   3165
         End
         Begin VB.Label Label50 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cep"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4905
            TabIndex        =   134
            Top             =   3330
            Width           =   330
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "Mala Direta"
         ForeColor       =   &H00000000&
         Height          =   7380
         Left            =   -67935
         TabIndex        =   129
         Top             =   405
         Width           =   7620
         Begin SSDataWidgets_B.SSDBDropDown ddwGruposMalaDireta 
            Bindings        =   "CliFor.frx":4EDAF
            Height          =   1905
            Left            =   705
            TabIndex        =   132
            Top             =   2160
            Width           =   4920
            DataFieldList   =   "Nome"
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
            BevelColorFrame =   15066597
            BevelColorHighlight=   15066597
            BevelColorShadow=   10066329
            BevelColorFace  =   15066597
            ForeColorEven   =   6710886
            ForeColorOdd    =   6710886
            BackColorEven   =   16250871
            BackColorOdd    =   12648447
            RowHeight       =   423
            ExtraHeight     =   79
            Columns.Count   =   2
            Columns(0).Width=   5768
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Nome"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   2223
            Columns(1).Caption=   "Código"
            Columns(1).Name =   "Código"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Código"
            Columns(1).DataType=   3
            Columns(1).FieldLen=   256
            _ExtentX        =   8678
            _ExtentY        =   3360
            _StockProps     =   77
            ForeColor       =   6710886
            BackColor       =   -2147483648
         End
         Begin VB.CheckBox Sem_Mala 
            Appearance      =   0  'Flat
            Caption         =   "Não incluir na mala direta"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   405
            TabIndex        =   61
            Top             =   450
            Width           =   2655
         End
         Begin VB.Data Data4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            Caption         =   "Data4"
            Connect         =   "Access 2000;"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00666666&
            Height          =   345
            Left            =   1440
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Grupos Interesse"
            Top             =   4200
            Visible         =   0   'False
            Width           =   2175
         End
         Begin SSDataWidgets_B.SSDBGrid grdGruposMalaDireta 
            Height          =   5325
            Left            =   405
            TabIndex        =   131
            Top             =   1650
            Width           =   6975
            _Version        =   196617
            DataMode        =   2
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Col.Count       =   2
            BevelColorFrame =   15066597
            BevelColorHighlight=   15066597
            BevelColorShadow=   10066329
            BevelColorFace  =   15066597
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            AllowGroupSizing=   0   'False
            AllowColumnSizing=   0   'False
            AllowGroupMoving=   0   'False
            AllowGroupSwapping=   0   'False
            AllowGroupShrinking=   0   'False
            AllowColumnShrinking=   0   'False
            AllowDragDrop   =   0   'False
            ForeColorEven   =   6710886
            ForeColorOdd    =   6710886
            BackColorEven   =   16250871
            BackColorOdd    =   12648447
            RowHeight       =   423
            ExtraHeight     =   79
            Columns.Count   =   2
            Columns(0).Width=   2328
            Columns(0).Caption=   "Grupo"
            Columns(0).Name =   "Grupo"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   5292
            Columns(1).Caption=   "Nome"
            Columns(1).Name =   "Nome"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            UseDefaults     =   0   'False
            _ExtentX        =   12303
            _ExtentY        =   9393
            _StockProps     =   79
            ForeColor       =   6710886
            BackColor       =   -2147483648
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "WeblySleek UI Semibold"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "WeblySleek UI Semibold"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label33 
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"CliFor.frx":4EDC3
            ForeColor       =   &H00000000&
            Height          =   660
            Left            =   405
            TabIndex        =   130
            Top             =   810
            Width           =   6975
         End
      End
      Begin VB.TextBox txtDiaBaseConsignacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -63090
         TabIndex        =   74
         Text            =   "01"
         Top             =   4260
         Width           =   855
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frete"
         ForeColor       =   &H00666666&
         Height          =   510
         Left            =   7470
         TabIndex        =   38
         Top             =   2430
         Width           =   2265
         Begin VB.OptionButton O_Frete_Pagar 
            Caption         =   "&A Pagar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1035
            TabIndex        =   48
            Top             =   210
            Width           =   1005
         End
         Begin VB.OptionButton O_Frete_Pago 
            Caption         =   "&Pago"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   255
            TabIndex        =   47
            Top             =   195
            Value           =   -1  'True
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Contabilidade"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   -74880
         TabIndex        =   120
         Top             =   4740
         Width           =   3390
         Begin VB.TextBox Contabilidade 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   58
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   480
            TabIndex        =   121
            Top             =   330
            Width           =   495
         End
      End
      Begin SSDataWidgets_B.SSDBGrid grdCaract 
         Height          =   7305
         Left            =   -74880
         TabIndex        =   84
         Top             =   450
         Width           =   14565
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   3
         AllowAddNew     =   -1  'True
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   238
         Columns.Count   =   3
         Columns(0).Width=   10239
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descricao"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   7461
         Columns(1).Caption=   "Valor"
         Columns(1).Name =   "Valor"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   255
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "Codigo"
         Columns(2).Name =   "Codigo"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   25691
         _ExtentY        =   12885
         _StockProps     =   79
         Caption         =   "Outros Dados"
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid grdContatos 
         Height          =   2535
         Left            =   135
         TabIndex        =   49
         Top             =   3375
         Width           =   14460
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Col.Count       =   6
         UseGroups       =   -1  'True
         DividerStyle    =   2
         BeveColorScheme =   1
         BevelColorFrame =   15066597
         BevelColorHighlight=   15066597
         BevelColorShadow=   15066597
         BevelColorFace  =   15066597
         CheckBox3D      =   0   'False
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         MultiLine       =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   16250871
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   185
         Groups(0).Width =   24262
         Groups(0).AllowSizing=   0   'False
         Groups(0).HasHeadBackColor=   -1  'True
         Groups(0).HeadBackColor=   15724527
         Groups(0).Columns.Count=   6
         Groups(0).Columns(0).Width=   7144
         Groups(0).Columns(0).Caption=   "Nome"
         Groups(0).Columns(0).Name=   "Contato"
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).DataType=   8
         Groups(0).Columns(0).FieldLen=   30
         Groups(0).Columns(0).BackColor=   15724527
         Groups(0).Columns(1).Width=   5318
         Groups(0).Columns(1).Caption=   "Vínculo"
         Groups(0).Columns(1).Name=   "Cargo"
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).DataType=   8
         Groups(0).Columns(1).FieldLen=   20
         Groups(0).Columns(2).Width=   1349
         Groups(0).Columns(2).Caption=   "Dia Aniv"
         Groups(0).Columns(2).Name=   "DiaAniv"
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   8
         Groups(0).Columns(2).FieldLen=   2
         Groups(0).Columns(3).Width=   1323
         Groups(0).Columns(3).Caption=   "Mês Aniv"
         Groups(0).Columns(3).Name=   "MesAniv"
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   8
         Groups(0).Columns(3).FieldLen=   3
         Groups(0).Columns(4).Width=   2540
         Groups(0).Columns(4).Caption=   "WhatsApp"
         Groups(0).Columns(4).Name=   "Ramal"
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).FieldLen=   9
         Groups(0).Columns(5).Width=   6588
         Groups(0).Columns(5).Caption=   "e-mail"
         Groups(0).Columns(5).Name=   "e-mail"
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   8
         Groups(0).Columns(5).FieldLen=   40
         _ExtentX        =   25506
         _ExtentY        =   4471
         _StockProps     =   79
         Caption         =   "Referências Pessoais"
         BackColor       =   15066597
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Transp 
         Bindings        =   "CliFor.frx":4EE66
         DataSource      =   "Data6"
         Height          =   315
         Left            =   1395
         TabIndex        =   36
         Top             =   2565
         Width           =   1275
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
         BevelColorFrame =   15066597
         BevelColorHighlight=   15066597
         BevelColorShadow=   10066329
         BevelColorFace  =   15066597
         ForeColorEven   =   0
         BackColorEven   =   15066597
         BackColorOdd    =   12640511
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   8520
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2170
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   2249
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483630
         BackColor       =   12648447
      End
      Begin VB.TextBox Comentários 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   6195
         Width           =   14160
      End
      Begin VB.TextBox C_Empresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   50
         TabIndex        =   63
         Top             =   2145
         Width           =   5535
      End
      Begin VB.TextBox C_Salário 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   20
         TabIndex        =   65
         Top             =   2475
         Width           =   855
      End
      Begin VB.TextBox C_Cargo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   100
         TabIndex        =   67
         Top             =   2805
         Width           =   2895
      End
      Begin VB.TextBox C_Contratação 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   15
         TabIndex        =   68
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox C_Telefone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -69405
         MaxLength       =   20
         TabIndex        =   69
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox C_C_Empresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   50
         TabIndex        =   73
         Top             =   4260
         Width           =   5535
      End
      Begin VB.TextBox C_C_Salário 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   20
         TabIndex        =   75
         Top             =   4575
         Width           =   855
      End
      Begin VB.TextBox C_C_Cargo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -70250
         MaxLength       =   20
         TabIndex        =   76
         Top             =   4575
         Width           =   2895
      End
      Begin VB.TextBox C_C_Contratação 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   15
         TabIndex        =   77
         Top             =   4905
         Width           =   1215
      End
      Begin VB.TextBox C_C_Telefone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -69405
         MaxLength       =   20
         TabIndex        =   78
         Top             =   4905
         Width           =   2055
      End
      Begin VB.TextBox C_C_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   50
         TabIndex        =   70
         Top             =   3615
         Width           =   5535
      End
      Begin VB.TextBox C_C_RG 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -72885
         MaxLength       =   30
         TabIndex        =   72
         Top             =   3945
         Width           =   3975
      End
      Begin VB.TextBox C_Obs 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -74700
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   79
         Top             =   5655
         Width           =   14115
      End
      Begin VB.TextBox home_page 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7890
         MaxLength       =   100
         TabIndex        =   42
         Top             =   2970
         Width           =   2730
      End
      Begin VB.TextBox email 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   915
         MaxLength       =   100
         TabIndex        =   28
         Top             =   1710
         Width           =   4485
      End
      Begin VB.CommandButton B_Expande 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   14355
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Expandir a caixa de comentários"
         Top             =   6195
         Width           =   255
      End
      Begin VB.TextBox CEP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   12420
         MaxLength       =   15
         TabIndex        =   26
         Top             =   1320
         Width           =   1380
      End
      Begin VB.TextBox Insc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7785
         MaxLength       =   23
         TabIndex        =   34
         Top             =   2130
         Width           =   1605
      End
      Begin VB.TextBox CGC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   915
         MaxLength       =   18
         TabIndex        =   32
         Top             =   2130
         Width           =   2580
      End
      Begin VB.TextBox Bairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   915
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1305
         Width           =   3270
      End
      Begin VB.TextBox Complemento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   10425
         MaxLength       =   50
         TabIndex        =   21
         Top             =   900
         Width           =   4185
      End
      Begin VB.TextBox Endereço 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   915
         MaxLength       =   211
         TabIndex        =   19
         Top             =   900
         Width           =   6150
      End
      Begin SSDataWidgets_B.SSDBGrid grdContatosEfetuados 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   179
         Top             =   1590
         Width           =   14580
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   5
         UseGroups       =   -1  'True
         BevelColorFrame =   15066597
         BevelColorHighlight=   15066597
         BevelColorShadow=   10066329
         BevelColorFace  =   15066597
         CheckBox3D      =   0   'False
         AllowDelete     =   -1  'True
         AllowUpdate     =   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         MaxSelectedRows =   1
         ForeColorEven   =   6710886
         ForeColorOdd    =   6710886
         BackColorEven   =   16250871
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   291
         Groups(0).Width =   24342
         Groups(0).Caption=   "Contatos"
         Groups(0).AllowSizing=   0   'False
         Groups(0).Columns.Count=   5
         Groups(0).Columns(0).Width=   3200
         Groups(0).Columns(0).Visible=   0   'False
         Groups(0).Columns(0).Caption=   "Sequência"
         Groups(0).Columns(0).Name=   "Sequência"
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).DataType=   8
         Groups(0).Columns(0).FieldLen=   256
         Groups(0).Columns(1).Width=   1958
         Groups(0).Columns(1).Caption=   "Data"
         Groups(0).Columns(1).Name=   "Data"
         Groups(0).Columns(1).Alignment=   1
         Groups(0).Columns(1).CaptionAlignment=   2
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).DataType=   8
         Groups(0).Columns(1).FieldLen=   10
         Groups(0).Columns(1).Locked=   -1  'True
         Groups(0).Columns(1).Mask=   "##/##/####"
         Groups(0).Columns(1).PromptInclude=   -1  'True
         Groups(0).Columns(1).PromptChar=   32
         Groups(0).Columns(2).Width=   16404
         Groups(0).Columns(2).Caption=   "Descrição"
         Groups(0).Columns(2).Name=   "Descricao"
         Groups(0).Columns(2).CaptionAlignment=   0
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   8
         Groups(0).Columns(2).FieldLen=   1000
         Groups(0).Columns(2).VertScrollBar=   -1  'True
         Groups(0).Columns(2).Locked=   -1  'True
         Groups(0).Columns(3).Width=   2355
         Groups(0).Columns(3).Caption=   "Pendência"
         Groups(0).Columns(3).Name=   "Pendencia"
         Groups(0).Columns(3).Alignment=   1
         Groups(0).Columns(3).CaptionAlignment=   2
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   11
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(3).Locked=   -1  'True
         Groups(0).Columns(3).Style=   2
         Groups(0).Columns(3).HasForeColor=   -1  'True
         Groups(0).Columns(3).ForeColor=   255
         Groups(0).Columns(4).Width=   3625
         Groups(0).Columns(4).Caption=   "Lembrar Em"
         Groups(0).Columns(4).Name=   "DataAviso"
         Groups(0).Columns(4).Alignment=   1
         Groups(0).Columns(4).CaptionAlignment=   2
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).FieldLen=   10
         Groups(0).Columns(4).Locked=   -1  'True
         Groups(0).Columns(4).Mask=   "##/##/####"
         Groups(0).Columns(4).PromptInclude=   -1  'True
         Groups(0).Columns(4).PromptChar=   32
         UseDefaults     =   0   'False
         _ExtentX        =   25717
         _ExtentY        =   11033
         _StockProps     =   79
         ForeColor       =   6710886
         BackColor       =   15066597
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox Desconto 
         Height          =   285
         Left            =   -65835
         TabIndex        =   64
         Top             =   2145
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox mskDatNasc 
         Height          =   330
         Left            =   6540
         TabIndex        =   29
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1710
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         ForeColor       =   0
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDatAbertura 
         Height          =   345
         Left            =   -65340
         TabIndex        =   124
         Top             =   555
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648384
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
         Bindings        =   "CliFor.frx":4EE7A
         DataSource      =   "Data5"
         Height          =   315
         Left            =   8745
         TabIndex        =   30
         Top             =   1725
         Width           =   990
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
         BevelColorFrame =   15066597
         BevelColorHighlight=   15066597
         BevelColorShadow=   10066329
         BevelColorFace  =   15066597
         ForeColorEven   =   0
         BackColorEven   =   15066597
         BackColorOdd    =   12640511
         Columns(0).Width=   3200
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin SSDataWidgets_B.SSDBGrid grdNumeravel 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   62
         Top             =   450
         Width           =   14610
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   5
         BevelColorFrame =   6710886
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   5
         Columns(0).Width=   5133
         Columns(0).Caption=   "Identificação"
         Columns(0).Name =   "Id"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   25
         Columns(1).Width=   5212
         Columns(1).Caption=   "Código do Produto"
         Columns(1).Name =   "Codigo"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   20
         Columns(2).Width=   2117
         Columns(2).Caption=   "Data 1"
         Columns(2).Name =   "Data1"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Mask =   "##/##/####"
         Columns(2).PromptInclude=   -1  'True
         Columns(2).PromptChar=   32
         Columns(3).Width=   2117
         Columns(3).Caption=   "Data 2"
         Columns(3).Name =   "Data2"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Mask =   "##/##/####"
         Columns(3).PromptInclude=   -1  'True
         Columns(3).PromptChar=   32
         Columns(4).Width=   4524
         Columns(4).Caption=   "Ref Doc"
         Columns(4).Name =   "RefDoc"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   20
         _ExtentX        =   25770
         _ExtentY        =   12726
         _StockProps     =   79
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspWebShopperID 
         Height          =   315
         Left            =   -74760
         TabIndex        =   85
         Top             =   840
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspWebDataCadastro 
         Height          =   315
         Left            =   -74760
         TabIndex        =   86
         Top             =   1560
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspWebBonus 
         Height          =   315
         Left            =   -74760
         TabIndex        =   87
         Top             =   2280
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspWebEMailMerco 
         Height          =   315
         Left            =   -74760
         TabIndex        =   150
         Top             =   3000
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspWebEMailLoja 
         Height          =   315
         Left            =   -74760
         TabIndex        =   151
         Top             =   3720
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspWebOrigem 
         Height          =   315
         Left            =   -74760
         TabIndex        =   153
         Top             =   4440
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin MSMask.MaskEdBox mskPercentualLimite 
         Height          =   285
         Left            =   -70800
         TabIndex        =   66
         Top             =   2475
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtContatosEfetuadosLembrarEm 
         Height          =   300
         Left            =   -71355
         TabIndex        =   176
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   840
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
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
      Begin VB.Label Label58 
         Caption         =   "Tp Cliente Sadig Web"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10695
         TabIndex        =   190
         Top             =   3015
         Width           =   1785
      End
      Begin VB.Label lblIndicadorIE 
         Caption         =   "Indicador IE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9450
         TabIndex        =   185
         Top             =   2175
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Insc. Municipal"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   69
         Left            =   3825
         TabIndex        =   184
         Top             =   2190
         Width           =   1050
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Insc. Suframa"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   70
         Left            =   165
         TabIndex        =   183
         Top             =   3030
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "CNAE"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   71
         Left            =   3495
         TabIndex        =   182
         Top             =   3030
         Width           =   405
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "País"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   10350
         TabIndex        =   181
         Top             =   1365
         Width           =   285
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   7410
         TabIndex        =   180
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label57 
         Caption         =   "Lembrar Em"
         Height          =   255
         Left            =   -72405
         TabIndex        =   177
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label56 
         Caption         =   "Data"
         Height          =   255
         Left            =   -71925
         TabIndex        =   173
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "<< Preenchendo este campo e dando 'TAB' será feito o cálculo automático para o Limite."
         Height          =   195
         Left            =   -69960
         TabIndex        =   160
         Top             =   2520
         Width           =   6375
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Percentual"
         Height          =   195
         Left            =   -71760
         TabIndex        =   159
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Origem"
         Height          =   195
         Index           =   5
         Left            =   -74760
         TabIndex        =   152
         Top             =   4200
         Width           =   510
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Gostaria de receber E-mail informando sobre lançamentos e promoções da Loja Virtual."
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   149
         Top             =   3480
         Width           =   6225
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Gostaria de receber E-mail informando sobre novas lojas e eventos do CIAShop."
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   148
         Top             =   2760
         Width           =   5775
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tabela de Preços"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9825
         TabIndex        =   144
         Top             =   2610
         Width           =   1395
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Código de identificação do cliente na Loja Virtual"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   128
         Top             =   600
         Width           =   3465
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Data de Cadastro na Loja Virtual "
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   127
         Top             =   1320
         Width           =   2385
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Total de Bônus em compras"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   126
         Top             =   2040
         Width           =   1965
      End
      Begin VB.Label Label48 
         Caption         =   "Dia base para acerto de consignação"
         Height          =   255
         Left            =   -66090
         TabIndex        =   125
         Top             =   4275
         Width           =   2835
      End
      Begin VB.Label Nome_Vendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   9825
         TabIndex        =   31
         Top             =   1725
         Width           =   4785
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Data de abertura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -66990
         TabIndex        =   71
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label46 
         Caption         =   "Nascimento"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5565
         TabIndex        =   122
         Top             =   1755
         Width           =   945
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Desconto %"
         Height          =   195
         Left            =   -66840
         TabIndex        =   117
         Top             =   2175
         Width           =   885
      End
      Begin VB.Label Nome_Transp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2715
         TabIndex        =   37
         Top             =   2550
         Width           =   4650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Cônjuge"
         Height          =   195
         Left            =   -74700
         TabIndex        =   113
         Top             =   3630
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empresa onde trabalha"
         Height          =   195
         Left            =   -74700
         TabIndex        =   112
         Top             =   2175
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Salário Atual"
         Height          =   195
         Left            =   -74700
         TabIndex        =   111
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   195
         Left            =   -74670
         TabIndex        =   110
         Top             =   2850
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data Contratação"
         Height          =   195
         Left            =   -74700
         TabIndex        =   109
         Top             =   3165
         Width           =   1290
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefone da Empresa"
         Height          =   195
         Left            =   -71025
         TabIndex        =   108
         Top             =   3165
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa onde trabalha"
         Height          =   195
         Left            =   -74700
         TabIndex        =   107
         Top             =   4290
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Salário Atual"
         Height          =   195
         Left            =   -74700
         TabIndex        =   106
         Top             =   4620
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Data Contratação"
         Height          =   195
         Left            =   -74700
         TabIndex        =   105
         Top             =   4950
         Width           =   1290
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefone da Empresa"
         Height          =   195
         Left            =   -71025
         TabIndex        =   104
         Top             =   4950
         Width           =   1515
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "CPF / RG"
         Height          =   195
         Left            =   -74700
         TabIndex        =   103
         Top             =   3960
         Width           =   645
      End
      Begin VB.Label Label12 
         Caption         =   "Consultas SPC, filiação, bens e outras informações"
         Height          =   255
         Left            =   -74700
         TabIndex        =   102
         Top             =   5370
         Width           =   4095
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   195
         Left            =   -70755
         TabIndex        =   101
         Top             =   4620
         Width           =   435
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         Caption         =   "Home Page"
         ForeColor       =   &H00666666&
         Height          =   210
         Left            =   6945
         TabIndex        =   100
         Top             =   3030
         Width           =   975
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         Caption         =   "Email"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   99
         Top             =   1755
         Width           =   540
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         Caption         =   "Comentários"
         ForeColor       =   &H00666666&
         Height          =   255
         Left            =   135
         TabIndex        =   98
         Top             =   5955
         Width           =   1095
      End
      Begin VB.Label Label39 
         Caption         =   "Vendedor"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7875
         TabIndex        =   97
         Top             =   1755
         Width           =   765
      End
      Begin VB.Label Label37 
         Caption         =   "Transportadora"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   96
         Top             =   2610
         Width           =   1215
      End
      Begin VB.Label L_Inscrição 
         Appearance      =   0  'Flat
         Caption         =   "Insc. Estadual"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   6705
         TabIndex        =   95
         Top             =   2190
         Width           =   1065
      End
      Begin VB.Label L_CGC 
         Appearance      =   0  'Flat
         Caption         =   "CNPJ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   94
         Top             =   2175
         Width           =   420
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12105
         TabIndex        =   93
         Top             =   1365
         Width           =   285
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "UF"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4305
         TabIndex        =   92
         Top             =   1365
         Width           =   195
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         Caption         =   "Cidade"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5580
         TabIndex        =   91
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         Caption         =   "Bairro"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   90
         Top             =   1350
         Width           =   495
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         Caption         =   "Complemento"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9270
         TabIndex        =   89
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   88
         Top             =   945
         Width           =   855
      End
   End
   Begin VB.CheckBox Bloqueado 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Bloqueado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8685
      TabIndex        =   6
      Top             =   465
      Width           =   1155
   End
   Begin VB.CheckBox Inativo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Inativo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8700
      TabIndex        =   4
      Top             =   105
      Width           =   855
   End
   Begin VB.CheckBox Faturado 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Crediário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9885
      TabIndex        =   7
      Top             =   105
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.CheckBox Conta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      Caption         =   "Conta de Caderno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9900
      TabIndex        =   5
      Top             =   480
      Width           =   1800
   End
   Begin SSDataWidgets_B.SSDBCombo cboCodigo 
      Bindings        =   "CliFor.frx":4EE8E
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      ToolTipText     =   "Pressione F5 para o Próximo Livre."
      Top             =   75
      Width           =   1260
      DataFieldList   =   "Código"
      MaxDropDownItems=   16
      BevelType       =   0
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
      BevelColorFrame =   15066597
      BevelColorHighlight=   15066597
      BevelColorShadow=   15066597
      BevelColorFace  =   15066597
      ForeColorEven   =   6710886
      ForeColorOdd    =   6710886
      BackColorEven   =   16250871
      BackColorOdd    =   12640511
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   2910
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Código"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   10134
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Nome"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1005
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   2222
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483630
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo cboFantasia 
      Bindings        =   "CliFor.frx":4EEA9
      Height          =   315
      Left            =   3345
      TabIndex        =   3
      Top             =   435
      Width           =   5280
      DataFieldList   =   "Fantasia"
      MaxDropDownItems=   16
      BevelType       =   0
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
      BevelColorFrame =   15066597
      BevelColorHighlight=   15066597
      BevelColorShadow=   15066597
      BevelColorFace  =   15066597
      ForeColorEven   =   6710886
      ForeColorOdd    =   6710886
      BackColorEven   =   16250871
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   8599
      Columns(0).Caption=   "Fantasia"
      Columns(0).Name =   "Fantasia"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Fantasia"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   873
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   9313
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483630
      BackColor       =   12648447
      DataFieldToDisplay=   "Fantasia"
   End
   Begin VB.Data datFantasia 
      Caption         =   "datFantasia"
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
      Height          =   360
      Left            =   4170
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Fantasia, Código, Tipo FROM Cli_For WHERE Código <> 0 AND Fantasia <> ''  ORDER BY Fantasia"
      Top             =   8745
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data datCliForNome 
      Caption         =   "datCliForNome"
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
      Height          =   360
      Left            =   2190
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código, Tipo FROM Cli_For WHERE Código <> 0  ORDER BY Nome"
      Top             =   8790
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data datCliForCod 
      Caption         =   "datCliForCod"
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
      Height          =   360
      Left            =   225
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome, Tipo FROM Cli_For WHERE Código <> 0  ORDER BY Código"
      Top             =   8790
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Left            =   4185
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contatos"
      Top             =   8940
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Left            =   2250
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE"
      Top             =   8940
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Left            =   195
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Transportadora"
      Top             =   8940
      Visible         =   0   'False
      Width           =   2010
   End
   Begin SSDataWidgets_B.SSDBCombo cboNome 
      Bindings        =   "CliFor.frx":4EEC3
      Height          =   315
      Left            =   3345
      TabIndex        =   2
      Top             =   75
      Width           =   5280
      DataFieldList   =   "Nome"
      MaxDropDownItems=   16
      BevelType       =   0
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
      BevelColorFrame =   15066597
      BevelColorHighlight=   15066597
      BevelColorShadow=   15066597
      BevelColorFace  =   15066597
      ForeColorEven   =   6710886
      ForeColorOdd    =   6710886
      BackColorEven   =   16250871
      BackColorOdd    =   12640511
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   10134
      Columns(0).Caption=   "Nome ou Razão Social"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2910
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   847
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   9313
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483630
      BackColor       =   12648447
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Fones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11205
      TabIndex        =   189
      Top             =   105
      Width           =   465
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Última Compra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   119
      Top             =   465
      Width           =   1185
   End
   Begin VB.Label Última_Compra 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1290
      TabIndex        =   118
      Top             =   435
      Width           =   1260
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fantasia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2595
      TabIndex        =   116
      Top             =   495
      Width           =   690
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2595
      TabIndex        =   115
      Top             =   105
      Width           =   750
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   114
      Top             =   105
      Width           =   705
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   60
      Top             =   8610
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
      Bands           =   "CliFor.frx":4EEDF
   End
End
Attribute VB_Name = "frmCliFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsCliFor As Recordset
Private rsCli_Cred As Recordset
Private Rec_Mala_Direta As Recordset
Private rsMala As Recordset
Private Num_Registro As Variant
Private rsMala_Tempo As Recordset
Private rsVendedores As Recordset
Private rsConta_Cliente As Recordset
Private rsContatos As Recordset
Private rsContatos_Efetuados As Recordset
Private rsEmprestimos As Recordset
Private rsResumo_Clientes As Recordset
Private rsTransportadoras As Recordset
Private rsContas As Recordset
Private rsCliFor2 As Recordset
Private rsGrupos As Recordset
Private rsCliForMaterial As Recordset
Private gbComboDropped As Boolean

'11/12/2007 - Celso
Private rsParametros As Recordset

'21/12/2006 - Anderson - Alteração na tela de contatos efetuados
Dim lngContatosEfetuadosSequencia As Long
Dim strContatosEfetuadosData As String
Dim bolCalendario As Boolean

'18/02/2004 - Daniel
'Var para controle da formatação da máscara
'Case: TV Shopping
Dim m_blnFlag As Boolean

'17/05/2004 - Daniel
'Este flag setará se a empresa tem necessidade de cadastrar
'o mesmo CNPJ/CPF para n clientes
'Cases: Embalavi
Dim m_blnNcnpjcpf As Boolean

'09/11/2004 - Daniel
'Esta var é utilizada para identificar o cliente
'Teknika onde o cadastro da Inscrição Estadual não
'deverá ficar em branco
Dim m_blnTeknika As Boolean
'----------------------------------------------------------------------------------
'04/08/2006 - Andrea
'Esta variável é utilizada para identificar o cliente Idéia Intima, para que
'quando entrar na tela Form_Load, marque a opção pessoa física ao invés de
'Jurídica (que é o Padrão)
Dim m_blnIdeiaIntima As Boolean
'--------------------------------------------------------------------------------

'13/07/2007 - Anderson
'Informa as configurações do usuário atual para o contatos efetuados
Dim m_strContatosEfetuados As String

'19/01/2007 - Anderson
'Variavel utilizada para saber se o código do vendedor foi alterado e se exige senha do gerente
Dim strVendedor As String

Dim bolCaseSupriPrint As Boolean '19/07/2007 - Anderson - Verifica se o case é pra Supri Print

Dim sEstadoInicioFoco As String

Public Sub ClearScreen(Optional Encontra As Integer)
  Dim sSql As String
  Dim rstTabelaPrecos As Recordset
  
  '05/04/2007 - Anderson
  'Criado para limpar a tela de contatos
  Call cmdContatoEfetuadoNovo_Click
  
  Call StatusMsg("")
  
  cboIndicadorIE.Text = "9=Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS"
  lblClienteWEB.Visible = False
  cboCodigo.Text = ""
  gsCodCliente = ""
  cboNome.Text = ""
  cboFantasia.Text = ""
  Endereço.Text = ""
  '26/07/2004 - Daniel
  'Adicionado campo AgenciaPublicidade
  'Case: STC
  chkAgencia.Value = vbUnchecked
  Complemento.Text = ""
  Bairro.Text = ""
  'If ActiveBar1.Tools("miOpFreezeCidade").Checked = False Then Cidade.Text = ""
  If ActiveBar1.Tools("miOpFreezeCidade").Checked = False Then cbo_cidades.Clear
  If ActiveBar1.Tools("miOpFreezeEstado").Checked = False Then Estado.Text = ""
  CEP.Text = ""
  Fone1.Text = ""
  Fone2.Text = ""
  
  '27/08/2009 - mpdea
  'Novos campos para NFe
  txtEnderecoNumero.Text = ""
  txtPais.Text = ""
  txtInscricaoMunicipal.Text = ""
  txtInscricaoSuframa.Text = ""
  txtCNAE.Text = ""
  
  mskDatNasc.Mask = "": mskDatNasc.Text = "": mskDatNasc.Mask = "##/##/####"
  mskDatAbertura.Mask = "": mskDatAbertura.Text = "": mskDatAbertura.Mask = "##/##/####"
  optMasculino.Value = True
  
  '---[ 02/06/2003 - Maikel ]---'
  ' Case      : Projeto Deluken
  ' Descrição : Adicionado o campo abaixo para guardar o dia base de acerto de consignações
    txtDiaBaseConsignacao.Text = "25"
  '-----------------------------'
  
  C_Cargo.Text = ""
  C_Contratação.Text = ""
  C_Telefone.Text = ""
  C_C_Nome.Text = ""
  C_C_RG.Text = ""
  C_C_Empresa.Text = ""
  C_C_Salário.Text = ""
  '30/06/2004 - Daniel
  'Adicionado campo [Cli_For - Crédito].PercentualLimite
  mskPercentualLimite.Mask = ""
  mskPercentualLimite.Text = 0
  '---------------------------------
  C_C_Cargo.Text = ""
  C_C_Contratação.Text = ""
  C_C_Telefone.Text = ""
  C_Obs.Text = ""
  
  '15/06/2005 - Daniel
  'Tratamento para os campos NomeSacadorAvalista, CPFSacadorAvalista e CPFCNPJSacadorAvalista
  txtNomeSacadorAvalista.Text = ""
  txtCPFSacadorAvalista.Text = ""
  optCPF.Value = False
  optCNPJ.Value = False
  
  Fax.Text = ""
  CGC.Text = ""
  Insc.Text = ""
  
  Inativo.Value = 0
  Bloqueado.Value = 0
  Faturado.Value = 1
  Conta.Value = 0
  Sem_Mala.Value = 0
  Desconto.Text = ""
  Última_Compra.Caption = ""
  
  
  Comentários.Text = ""
  Combo_Transp.Text = ""
  Nome_Transp.Caption = ""
  Combo_Vendedor.Text = ""
  Nome_Vendedor.Caption = ""
  
  email.Text = ""
  home_page.Text = ""
  
  Limite.Text = ""
  
  Contabilidade.Text = ""
  
  Combo_Cobrança.Text = 0
  Nome_Cobrança.Caption = ""
  
  Endereço_Cob.Text = ""
  Complemento_Cob.Text = ""
  Bairro_Cob.Text = ""
  Cidade_Cob.Text = ""
  Estado_Cob.Text = ""
  CEP_Cob.Text = ""
  '05/05/2004 - Daniel
  'Adicionado o campo IsentoIPI
  chkIsencaoIPI.Value = vbUnchecked
  '06/05/2004 - Daniel
  'Adicionado campo ObsIsentoIPI
  txtMensagemIPI.Text = ""
    
  If ActiveBar1.Tools("miOpFreezeMsg").Checked = False Then Boleto.Text = ""
  
  C_Empresa.Text = ""
  C_Salário.Text = ""
  C_Cargo.Text = ""
  C_Contratação.Text = ""
  C_Telefone.Text = ""
  C_C_Nome.Text = ""
  C_C_RG.Text = ""
  C_C_Empresa.Text = ""
  C_C_Salário.Text = ""
  C_C_Cargo.Text = ""
  C_C_Contratação.Text = ""
  C_C_Telefone.Text = ""
  C_Obs.Text = ""
  
  Limite.Text = ""
  
  sspWebShopperID.Caption = ""
  sspWebBonus.Caption = ""
  sspWebDataCadastro.Caption = ""
  
  '23/06/2004 - mpdea
  'Novos campos
  sspWebEMailMerco.Caption = ""
  sspWebEMailLoja = ""
  sspWebOrigem.Caption = ""
  
  If Encontra <> 1 Then cboCodigo.SetFocus
  
  Tab1.TabEnabled(1) = False
  Tab1.TabEnabled(2) = False
  Tab1.TabEnabled(3) = False
  Tab1.TabEnabled(4) = False
  Tab1.TabEnabled(5) = False
  Tab1.TabEnabled(6) = False
  '14/07/2004 - Daniel
  'Tab de Classificação
  Tab1.TabEnabled(7) = False
  
  If Not rsCliFor.EOF Then
    On Error Resume Next
    rsCliFor.MoveFirst
    rsCliFor.MovePrevious
    On Error GoTo 0
  End If
  
  Call ResetAllGrids
  
  '22/10/2003 - Maikel
  '---[ Carrega a lista de tabela de preços ]---'
    cboTabelaPrecos.Text = ""
    cboTabelaPrecos.Clear
    
    Set rstTabelaPrecos = db.OpenRecordset("SELECT Distinct TABELA FROM Preços", dbOpenSnapshot)
    
    With rstTabelaPrecos
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do While Not .EOF
          cboTabelaPrecos.AddItem .Fields("Tabela") & ""
          .MoveNext
        Loop
      End If
      
      .Close
      Set rstTabelaPrecos = Nothing
    End With
  '---[ Carrega a lista de tabela de preços ]---'
  
  O_Jurídica_Click
  Num_Registro = Null
  Tab1.Tab = 0
  
  '18/01/2007 - Anderson
  'Solicitação senha do gerente ao alterar o vendedor relacionado ao cliente
  If gbVendedorSenhaGerente Then
    Combo_Vendedor.Text = gnUserCode
    strVendedor = gnUserCode
    Call Combo_Vendedor_LostFocus
  End If
  
  '18/07/2007 - Anderson
  'Limpa campo SadigWeb
  cboSadigWebTipo.Text = "(Nenhum)"
  
  '19/07/2007 - Anderson
  'Exibe os possíveis objetos ocultos conforme customização da Supri Print
  txtContatosEfetuadosLembrarEm.Visible = True
  Label57.Visible = True
  chkContatoEfetuadoPendencia.Visible = True
  
End Sub

Private Sub LoadGridCaract()
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim Aux_Prod As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim sDescricao As String
  Dim sUnidVenda As String
  Dim sTipo As String
  Dim sCriteria As String
  Dim sSql As String
  Dim rsCliForCaract As Recordset
  Dim rsTabCaractCliFor As Recordset
  Dim sCod As String
  
  On Error GoTo ErrCaract
  
  grdCaract.CancelUpdate
  
  bAllow = grdCaract.AllowAddNew
  grdCaract.AllowAddNew = True
  grdCaract.AllowUpdate = True
  
  If O_Cliente.Value = True Then
    sTipo = "C"
  Else
    If O_Fornecedor.Value = True Then
      sTipo = "F"
    Else
      If O_Revendedor.Value = True Then
        sTipo = "R"
      Else
        sTipo = "O"
      End If
    End If
  End If
  
  If Len(Trim(cboCodigo.Text)) > 0 Then
    sCod = cboCodigo.Text
  Else
    sCod = "-1"
  End If
  
  sSql = "SELECT * FROM CliForCaract "
  sSql = sSql & " WHERE CodCliCaract = " & sCod
  sSql = sSql & " And TipoCliCaract = '" & sTipo & "' "
  sSql = sSql & " ORDER BY CodCaract "
  Set rsCliForCaract = db.OpenRecordset(sSql, dbOpenDynaset)
  
  sSql = "SELECT * FROM TabCaractCliFor WHERE TipoCliCaract = '" & sTipo & "' "
  sSql = sSql & " ORDER BY CodCaract "
  Set rsTabCaractCliFor = db.OpenRecordset(sSql, dbOpenDynaset)
  
  grdCaract.Redraw = False
  grdCaract.RemoveAll
  
  With rsTabCaractCliFor
    Do While Not .EOF
      sCriteria = "CodCaract = " & .Fields("CodCaract")
      rsCliForCaract.FindFirst sCriteria
      
      If Not rsCliForCaract.NoMatch Then
        sRecord = .Fields("DescCaract") & vbTab & _
                  rsCliForCaract.Fields("ValCaract") & vbTab & _
                  .Fields("CodCaract")
      Else
        sRecord = .Fields("DescCaract") & vbTab & _
                  "" & vbTab & _
                  .Fields("CodCaract")
      End If
      grdCaract.AddItem sRecord
      .MoveNext
    Loop
  End With
  
  grdCaract.Scroll -99, -99
  grdCaract.Redraw = True
  grdCaract.AllowAddNew = bAllow
  grdCaract.AllowUpdate = bAllow

  rsTabCaractCliFor.Close
  Set rsTabCaractCliFor = Nothing
  rsCliForCaract.Close
  Set rsCliForCaract = Nothing
  
  Exit Sub
ErrCaract:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao carregar Outros Dados."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Private Sub WriteGridCaract()
  Dim sSql As String
  Dim bm As Variant
  Dim nRow As Long
  Dim sTipo As String
  Dim rsCliForCaract As Recordset
  
  On Error GoTo ErrHandler
  
  grdCaract.Update
  
  If O_Cliente.Value = True Then
    sTipo = "C"
  Else
    If O_Fornecedor.Value = True Then
      sTipo = "F"
    Else
      If O_Revendedor.Value = True Then
        sTipo = "R"
      Else
        sTipo = "O"
      End If
    End If
  End If
  
  Call ws.BeginTrans
  
  sSql = "SELECT * FROM CliForCaract "
  sSql = sSql & " WHERE CodCliCaract = " & cboCodigo.Text
 ' sSql = sSql & " And TipoCliCaract = '" & sTipo & "' "
  sSql = sSql & " ORDER BY CodCaract "
  Set rsCliForCaract = db.OpenRecordset(sSql, dbOpenDynaset)
  
  On Error GoTo 0
  
  With rsCliForCaract
    
    If Not .EOF Then
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
    End If
    
    For nRow = 0 To grdCaract.Rows - 1
      bm = grdCaract.AddItemBookmark(nRow)
      If Len(Trim(grdCaract.Columns("Valor").CellText(bm))) > 0 Then
        .AddNew
        .Fields("CodCliCaract") = cboCodigo.Text
        .Fields("TipoCliCaract") = sTipo
        .Fields("CodCaract") = CInt(grdCaract.Columns("Codigo").CellValue(bm))
        .Fields("ValCaract") = UCase(grdCaract.Columns("Valor").CellText(bm))
        .Update
      End If
    Next nRow
  
  End With
  
  rsCliForCaract.Close
  Set rsCliForCaract = Nothing
  
  Call ws.CommitTrans
  Exit Sub
  
ErrHandler:
  ws.Rollback
  
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Atualizar Valores para Características."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Private Sub LoadGridNumeraveis()
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim Aux_Prod As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim sDescricao As String
  Dim sUnidVenda As String
  Dim sTipo As String
  Dim sCriteria As String
  Dim sSql As String
  Dim sData1 As String
  Dim sData2 As String
  Dim rsCliForNumeravel As Recordset
  Dim sCod As String
  
  On Error GoTo ErrCaract
  
  grdNumeravel.CancelUpdate
  
  bAllow = grdNumeravel.AllowAddNew
  grdNumeravel.AllowAddNew = True
  grdNumeravel.AllowUpdate = True
  
  If O_Cliente.Value = True Then
    sTipo = "C"
  Else
    If O_Fornecedor.Value = True Then
      sTipo = "F"
    Else
      If O_Revendedor.Value = True Then
        sTipo = "R"
      Else
        sTipo = "O"
      End If
    End If
  End If
  
  If Len(Trim(cboCodigo.Text)) > 0 Then
    sCod = cboCodigo.Text
  Else
    sCod = "-1"
  End If
  
  sSql = "SELECT * FROM CliForNumeravel "
  sSql = sSql & " WHERE CodCliNumer = " & sCod
  sSql = sSql & " And TipoCliNumer = '" & sTipo & "' "
  sSql = sSql & " ORDER BY CodNumer "
  Set rsCliForNumeravel = db.OpenRecordset(sSql, dbOpenDynaset)
  
  grdNumeravel.Redraw = False
  grdNumeravel.RemoveAll
  
  With rsCliForNumeravel
    
    Do While Not .EOF
      
      If IsDate(.Fields("Data1Numer")) Then
        sData1 = .Fields("Data1Numer")
      Else
        sData1 = ""
      End If
      
      If IsDate(.Fields("Data2Numer")) Then
        sData2 = .Fields("Data2Numer")
      Else
        sData2 = ""
      End If
      
      sRecord = .Fields("CodNumer") & vbTab & _
                .Fields("CodProdNumer") & vbTab & _
                sData1 & vbTab & _
                sData2 & vbTab & _
                .Fields("CodRefDocNumer")
      grdNumeravel.AddItem sRecord
      .MoveNext
    
    Loop
    grdNumeravel.Scroll -99, -99
    
  End With
  
  grdNumeravel.Redraw = True
  grdNumeravel.AllowAddNew = bAllow
  grdNumeravel.AllowUpdate = bAllow

  rsCliForNumeravel.Close
  Set rsCliForNumeravel = Nothing
  
  Exit Sub
  
ErrCaract:
  Exit Sub

End Sub

Private Sub WriteGridNumeraveis()
  Dim sSql As String
  Dim bm As Variant
  Dim nRow As Long
  Dim sTipo As String
  Dim rsCliForNumeravel As Recordset
  
  On Error GoTo ErrHandler
  
  grdNumeravel.Update
  
  If O_Cliente.Value = True Then
    sTipo = "C"
  Else
    If O_Fornecedor.Value = True Then
      sTipo = "F"
    Else
      If O_Revendedor.Value = True Then
        sTipo = "R"
      Else
        sTipo = "O"
      End If
    End If
  End If
  
  Call ws.BeginTrans
  
  sSql = "SELECT * FROM CliForNumeravel "
  sSql = sSql & " WHERE CodCliNumer = " & cboCodigo.Text
  'sSql = sSql & " And TipoCliNumer = '" & sTipo & "' "
  sSql = sSql & " ORDER BY CodNumer "
  Set rsCliForNumeravel = db.OpenRecordset(sSql, dbOpenDynaset)
  
  With rsCliForNumeravel
    
    If Not .EOF Then
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
    End If
    
    For nRow = 0 To grdNumeravel.Rows - 1
      bm = grdNumeravel.AddItemBookmark(nRow)
      If Len(Trim(grdNumeravel.Columns("Id").CellText(bm))) > 0 Then
        .AddNew
        .Fields("CodCliNumer") = cboCodigo.Text
        .Fields("TipoCliNumer") = sTipo
        .Fields("CodNumer") = UCase(grdNumeravel.Columns("Id").CellText(bm))
        .Fields("CodProdNumer") = UCase(grdNumeravel.Columns("Codigo").CellText(bm) & "")
        If IsDate(grdNumeravel.Columns("Data1").CellText(bm)) Then
          .Fields("Data1Numer") = grdNumeravel.Columns("Data1").CellText(bm)
        End If
        If IsDate(grdNumeravel.Columns("Data2").CellText(bm)) Then
          .Fields("Data2Numer") = grdNumeravel.Columns("Data2").CellText(bm)
        End If
        .Fields("CodRefDocNumer") = UCase(grdNumeravel.Columns("RefDoc").CellText(bm) & "")
        .Update
      End If
    Next nRow
  End With
  
  rsCliForNumeravel.Close
  Set rsCliForNumeravel = Nothing
  
  Call ws.CommitTrans
  Exit Sub
  
ErrHandler:
  ws.Rollback
  
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Atualizar Valores para Numeráveis."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Private Sub LoadGridContatos()
  Dim rsContatos As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim Aux_Erro As Integer
  Dim sDescricao As String
  Dim sUnidVenda As String
  Dim sCod As String
  
  On Error GoTo ErrContatos
  
  grdContatos.CancelUpdate
  
  bAllow = grdContatos.AllowAddNew
  grdContatos.AllowAddNew = True
  grdContatos.AllowUpdate = True
  
  If Len(Trim(cboCodigo.Text)) > 0 Then
    sCod = cboCodigo.Text
  Else
    sCod = "-1"
  End If
  
  Set rsContatos = db.OpenRecordset("SELECT * FROM Contatos WHERE Cliente = " & sCod & " ORDER BY Seqüência", dbOpenDynaset)

  grdContatos.RemoveAll
  grdContatos.Redraw = False
  
  If Not rsContatos.EOF Then
    With rsContatos
      .MoveFirst
      Do While Not .EOF
        sRecord = .Fields("Contato") & vbTab & _
          .Fields("Cargo") & vbTab & _
          .Fields("Dia Aniversário") & vbTab & _
          .Fields("Mês Aniversário") & vbTab & _
          .Fields("Ramal") & vbTab & _
          .Fields("email")
        grdContatos.AddItem sRecord
        .MoveNext
      Loop
      .MoveFirst
    End With
    grdContatos.Scroll -99, -99
  End If

  grdContatos.Redraw = True
  grdContatos.AllowAddNew = bAllow
  grdContatos.AllowUpdate = bAllow

  rsContatos.Close
  Set rsContatos = Nothing
  Exit Sub
  
ErrContatos:
  Exit Sub


End Sub

Private Sub WriteGridContatos()
  Dim sSql As String
  Dim bm As Variant
  Dim nRow As Long
  Dim rsContatos As Recordset
  
  On Error GoTo ErrHandler
  
  grdContatos.Update
  
  Call ws.BeginTrans
  
  Set rsContatos = db.OpenRecordset("SELECT * FROM Contatos WHERE Cliente = " & cboCodigo.Text, dbOpenDynaset)
  
  With rsContatos
    
    If Not .EOF Then
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
    End If
    
    For nRow = 0 To grdContatos.Rows - 1
      bm = grdContatos.AddItemBookmark(nRow)
      If Len(grdContatos.Columns("Contato").CellText(bm)) > 0 Then
        .AddNew
        .Fields("Cliente") = cboCodigo.Text
        .Fields("Seqüência") = nRow + 1
        .Fields("Contato") = grdContatos.Columns("Contato").CellText(bm)
        .Fields("Cargo") = grdContatos.Columns("Cargo").CellText(bm)
        .Fields("Dia Aniversário") = CInt(gsHandleNull(grdContatos.Columns("DiaAniv").CellValue(bm) & ""))
        .Fields("Mês Aniversário") = grdContatos.Columns("MesAniv").CellValue(bm) & ""
        .Fields("Ramal") = grdContatos.Columns("Ramal").CellValue(bm) & ""
        .Fields("email") = grdContatos.Columns("e-mail").CellValue(bm) & ""
        .Update
      End If
    Next nRow
  
  End With
  
  rsContatos.Close
  Set rsContatos = Nothing
  
  Call ws.CommitTrans
  Exit Sub
  
ErrHandler:
  ws.Rollback

  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Atualizar Contatos."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Private Sub LoadGridGruposMalaDireta()
  Dim rsGruposMalaDireta As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sSql As String
  Dim sCod As String
  
  On Error GoTo ErrHandler
  
  grdGruposMalaDireta.CancelUpdate
  
  bAllow = grdGruposMalaDireta.AllowAddNew
  grdGruposMalaDireta.AllowAddNew = True
  grdGruposMalaDireta.AllowUpdate = True
  
  If Len(Trim(cboCodigo.Text)) > 0 Then
    sCod = cboCodigo.Text
  Else
    sCod = "-1"
  End If
  
  sSql = "SELECT Cliente, Grupo, [Grupos Interesse].Nome As Nome FROM [Mala Direta]"
  sSql = sSql + " INNER JOIN [Grupos Interesse] ON [Mala Direta].Grupo = [Grupos Interesse].Código  WHERE Cliente = " & sCod
  Set rsGruposMalaDireta = db.OpenRecordset(sSql, dbOpenDynaset)

  grdGruposMalaDireta.RemoveAll
  grdGruposMalaDireta.Redraw = False
  
  If Not rsGruposMalaDireta.EOF Then
    With rsGruposMalaDireta
      .MoveFirst
      Do While Not .EOF
        sRecord = .Fields("Grupo") & vbTab & _
          .Fields("Nome")
        grdGruposMalaDireta.AddItem sRecord
        .MoveNext
      Loop
      .MoveFirst
    End With
    grdGruposMalaDireta.Scroll -99, -99
  End If

  grdGruposMalaDireta.Redraw = True
  grdGruposMalaDireta.AllowAddNew = bAllow
  grdGruposMalaDireta.AllowUpdate = bAllow

  rsGruposMalaDireta.Close
  Set rsGruposMalaDireta = Nothing
  Exit Sub
  
ErrHandler:
  Exit Sub

End Sub

Private Sub WriteGridGruposMalaDireta()
  Dim sSql As String
  Dim bm As Variant
  Dim nRow As Long
  Dim rsGruposMalaDireta As Recordset
  
  On Error GoTo ErrHandler
  
  grdGruposMalaDireta.Update
  
  Call ws.BeginTrans
  
  Set rsGruposMalaDireta = db.OpenRecordset("SELECT * FROM [Mala Direta] WHERE Cliente = " & cboCodigo.Text, dbOpenDynaset)
  
  With rsGruposMalaDireta
    
    If Not .EOF Then
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
    End If
    
    For nRow = 0 To grdGruposMalaDireta.Rows - 1
      bm = grdGruposMalaDireta.AddItemBookmark(nRow)
      If Len(grdGruposMalaDireta.Columns("Grupo").CellText(bm)) > 0 Then
        .AddNew
        .Fields("Cliente") = cboCodigo.Text
        .Fields("Grupo") = grdGruposMalaDireta.Columns("Grupo").CellText(bm)
        .Update
      End If
    Next nRow
  
  End With
  
  rsGruposMalaDireta.Close
  Set rsGruposMalaDireta = Nothing
  
  Call ws.CommitTrans
  Exit Sub
  
ErrHandler:
  ws.Rollback
  
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Atualizar Grupos de Interesses para Mala Direta."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Private Sub LoadGridContatosEfetuados()
  Dim rsContatosEfetuados As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim Aux_Prod As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor As Integer
  Dim Aux_Edição As Long
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim sDescricao As String
  Dim sUnidVenda As String
  Dim sCod As String
  
  On Error GoTo ErrContatos
  
  grdContatosEfetuados.CancelUpdate
  
  '21/12/2006 - Anderson - linhas retiradas devido as alterações na tela de contatos efetuados
  'grdContatosEfetuados.AllowAddNew = True
  'grdContatosEfetuados.AllowUpdate = True
  'bAllow = grdContatosEfetuados.AllowAddNew
  grdContatosEfetuados.AllowAddNew = False
  grdContatosEfetuados.AllowUpdate = False
   
  If Len(Trim(cboCodigo.Text)) > 0 Then
    sCod = cboCodigo.Text
  Else
    sCod = "-1"
  End If
  
  Set rsContatosEfetuados = db.OpenRecordset("SELECT * FROM [Contatos Efetuados] WHERE Cliente = " & sCod & " ORDER BY Data, Seqüência", dbOpenDynaset)

  grdContatosEfetuados.RemoveAll
  grdContatosEfetuados.Redraw = False
  
  If Not rsContatosEfetuados.EOF Then
    With rsContatosEfetuados
      .MoveFirst
      Do While Not .EOF
        sRecord = .Fields("Seqüência") & vbTab & _
          .Fields("Data") & vbTab & _
          .Fields("Descrição") & vbTab & _
          .Fields("Pendência") & vbTab & _
          .Fields("Data Aviso")
        grdContatosEfetuados.AddItem sRecord
        .MoveNext
      Loop
      .MoveFirst
    End With
    grdContatosEfetuados.Scroll -99, -99
  End If

  '20/12/2006 - Anderson - Supri Print - Movimentar registro para o último contato
  'grdContatosEfetuados.MoveFirst
  grdContatosEfetuados.MoveLast
  If grdContatosEfetuados.Rows <= 8 Then
    grdContatosEfetuados.Scroll 0, -8
  Else
    grdContatosEfetuados.Scroll 0, -7
    grdContatosEfetuados.Scroll 0, 7
  End If
  
  grdContatosEfetuados.Redraw = True
  
  '21/12/2006 - Anderson - Alterações na tela de de contatos efetuados
  'grdContatosEfetuados.AllowAddNew = bAllow
  'grdContatosEfetuados.AllowUpdate = bAllow

  rsContatosEfetuados.Close
  Set rsContatosEfetuados = Nothing
  Exit Sub
  
ErrContatos:
  Exit Sub

End Sub

Private Sub WriteGridContatosEfetuados()
  Dim sSql As String
  Dim bm As Variant
  Dim nRow As Long
  Dim rsContatosEfetuados As Recordset
  
  On Error GoTo ErrHandler
  
  grdContatosEfetuados.Update
  
  Call ws.BeginTrans
  
  Set rsContatosEfetuados = db.OpenRecordset("SELECT * FROM [Contatos Efetuados] WHERE Cliente = " & cboCodigo.Text, dbOpenDynaset)
  
  With rsContatosEfetuados
    
    If Not .EOF Then
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
    End If
    
    For nRow = 0 To grdContatosEfetuados.Rows - 1
      bm = grdContatosEfetuados.AddItemBookmark(nRow)
      If IsDate(grdContatosEfetuados.Columns("Data").CellText(bm)) Then
        .AddNew
        .Fields("Cliente") = cboCodigo.Text
        .Fields("Data") = grdContatosEfetuados.Columns("Data").CellText(bm)
        .Fields("Descrição") = grdContatosEfetuados.Columns("Descricao").CellText(bm) & ""
        If grdContatosEfetuados.Columns("Pendencia").CellText(bm) = "" Then
          .Fields("Pendência") = False
        Else
          If grdContatosEfetuados.Columns("Pendencia").CellText(bm) = False Then
            .Fields("Pendência") = False
          Else
            .Fields("Pendência") = True
          End If
        End If
        If IsDate(grdContatosEfetuados.Columns("DataAviso").CellText(bm)) Then
          .Fields("Data Aviso") = grdContatosEfetuados.Columns("DataAviso").CellText(bm)
        End If
        .Update
      End If
    Next nRow
  
  End With
  
  rsContatosEfetuados.Close
  Set rsContatosEfetuados = Nothing
  
  Call ws.CommitTrans
  Exit Sub
  
ErrHandler:
  ws.Rollback

  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Atualizar Contatos Efetuados."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

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
    Case "miOpSearch"
      Call SearchRecord
    Case "miOpPassword"
    Case "miComplInfo"
      Call GetInformation
    Case "miComplMalaDireta"
      Call AddToMalaDireta
    Case "miVendaRapida"
      If Len(Trim(cboCodigo.Text)) > 0 Then
        Dim rsFuncionarios As Recordset
        Dim rsCaixas As Recordset
        
        Set rsFuncionarios = db.OpenRecordset("SELECT Código, Nome FROM Funcionários WHERE Código = " & gnUserCode, dbOpenSnapshot)
        Set rsCaixas = db.OpenRecordset("SELECT * FROM [Caixas em Uso] WHERE Caixa =  1", dbOpenSnapshot)
        
        '18/01/2006 - mpdea
        'Escolha da tela de Venda Rápida a ser utilizada
        If g_frmVendaRapida Is Nothing Then
          If g_bln_VendaRapidaCheckOut(gnCodFilial) Then
            'Minimiza a tela principal para evitar exibições incorretas por foco
            frmMain.WindowState = vbMinimized
            'Utilização da tela de Venda Rápida em tela cheia
            Set g_frmVendaRapida = frmVendaRap2_CheckOut
          Else
            'Tela Padrão
            Set g_frmVendaRapida = frmVendaRap2
          End If
        End If
        'Setado funcionário antes de carregar o form
        Funcionario = rsFuncionarios!Código
        'Adicionado Load
        Load g_frmVendaRapida
        'Alterado objeto frmVendaRap2 -> g_frmVendaRapida
        With g_frmVendaRapida
          If Not (rsFuncionarios.BOF And rsFuncionarios.EOF) Then
            .Nome_Operador.Caption = rsFuncionarios!Nome         'Vendedor.Caption
            .Cod_Operador.Caption = rsFuncionarios!Código        'Cod_Operador.Caption
            .Combo_Vendedor.Text = Trim(Combo_Vendedor.Text)     'Cod_Operador.Caption
            .Nome_Vendedor.Caption = Trim(Nome_Vendedor.Caption) 'Cod_Operador.Caption
            .Combo_Cliente.Text = Trim(cboCodigo.Text)
            .Nome_Cliente.Caption = LTrim(RTrim((cboNome.Text)))
            .Cod_Caixa = 1
            '13/07/2006 - mpdea
            'Incluído exibição do código do caixa (RT-13 ao gravar)
            .cboCaixa.Text = 1
            .Nome_Caixa.Caption = IIf(Not rsCaixas.EOF, rsCaixas!Descrição, "-")
            .Combo_Preço.Text = cboTabelaPrecos.Text & ""
            
            .Show
            '.CheckMovimentacao
            .Grade1.SetFocus
          End If
        End With
        rsFuncionarios.Close
        rsCaixas.Close
        
        Set rsFuncionarios = Nothing
        Set rsCaixas = Nothing
      Else
        MsgBox "Encontre o registro de um cliente antes", vbInformation, "Quick Store"
      End If
    
    'Opções
    
    Case "miOpFreezeCidade"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Clientes", "Mantem Cidade", Tool.Checked)
    
    Case "miOpFreezeEstado"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Clientes", "Mantem Estado", Tool.Checked)
    
    Case "miOpFreezeMsg"
      Tool.Checked = Not Tool.Checked
      Call UpdateArqConfig("Clientes", "Mantem Boleto", Tool.Checked)
      
    'Imprissão de Ficha Cadastral
    Case "nme_fichaCadastral"
      ImprimirFichaCadastral
    
  End Select

End Sub

Private Sub ImprimirFichaCadastral()
  On Error GoTo Erro
  
  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome As String
  Dim strNomeLPT As String
  Dim strPortaLPT As String
  Dim intX As Integer
  Dim i As Integer
  
  If Trim(cboCodigo.Text) = "" Or Trim(cboNome.Text) = "" Then
      MsgBox "Selecione um Cliente/Fornecedor válido", vbInformation, "Atenção"
      Exit Sub
  End If
  
  strNome = "REL"
  strNomeLPT = "NOME IMPRESSORA REL"
  strPortaLPT = "PORTA IMPRESSORA REL"

  strImpressora = GetSetting("QuickStore", "ConfigLPT", strNomeLPT, "")
  strPorta = GetSetting("QuickStore", "ConfigLPT", strPortaLPT, "")
      
  If Len(Trim(strImpressora)) > 0 And Len(Trim(strPorta)) > 0 Then
      For Each objPrinter In Printers
        If objPrinter.DeviceName = strImpressora And objPrinter.Port = strPorta Then
            Set Printer = objPrinter
            Exit For
        End If
      Next objPrinter
  End If

  Dim sAUX As String
  Dim sLinha As String
  
  'sDataAtual = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  
  Printer.Font = "LUCIDA CONSOLE"

  ' 116 COLUNAS
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""
  Printer.Print "                                                   FICHA CADASTRAL"
  Printer.Print ""
      
  If Picture1.Picture <> 0 Then
      Printer.PaintPicture Picture1, 9000, 500, 2300, 1000
  End If
  
  sLinha = "   Código     : " & Trim(cboCodigo.Text)
  Printer.Print sLinha
  
  sLinha = "   Nome       : " & Trim(cboNome.Text)
  Printer.Print sLinha
 
  sLinha = "   CFP/CNPJ   : " & CGC.Text
  Printer.Print sLinha
  
  sLinha = "   RG         : " & Insc.Text
  Printer.Print sLinha
    
  If mskDatNasc.Text = "  /  /    " Then
      sLinha = "   Nascimento : __/__/____"
  Else
      sLinha = "   Nascimento : " & mskDatNasc.Text
  End If
  Printer.Print sLinha
  
  ' Endereço e Numero
  sAUX = Trim(Endereço.Text)
  If Len(sAUX) > 71 Then
      sLinha = "   Endereço   : " & Mid(sAUX, 1, 70)
  Else
      sLinha = "   Endereço   : " & sAUX
  End If
  sLinha = sLinha & "   Nº : " & Trim(txtEnderecoNumero.Text)
  Printer.Print sLinha
  
  sLinha = "   Bairro     : " & Trim(Bairro.Text)
  Printer.Print sLinha
  
  ' Cidade e Estado
  sLinha = "   Cidade     : " & Trim(cbo_cidades.Text)
  sLinha = sLinha & "   Estado : " & Estado.Text
  Printer.Print sLinha
  
  sLinha = "   País       : " & Trim(txtPais.Text)
  Printer.Print sLinha
  
  sLinha = "   E-mail     : " & Trim(email.Text)
  Printer.Print sLinha

  Printer.Print ""
  Printer.Print ""

  Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
  Printer.Print "   Dados de Crédito"
  Printer.Print ""

  sLinha = "   Data de Abertura      : " & mskDatAbertura.Text
  Printer.Print sLinha
  
  sLinha = "   Empresa onde Trabalha : " & C_Empresa.Text
  Printer.Print sLinha

  If Trim(C_Salário.Text) = "" Then
      sLinha = "   Salário Atual         : 0,00   "
  Else
      sLinha = "   Salário Atual         : " & C_Salário.Text
  End If
  Printer.Print sLinha
  
  sLinha = "   Cargo                 : " & C_Cargo.Text
  Printer.Print sLinha
  
  If Trim(C_Contratação.Text) = "" Then
      sLinha = "   Data da Contratação   : __/__/____"
  Else
      sLinha = "   Data da Contratação   : " & C_Contratação.Text
  End If
  Printer.Print sLinha
  
  sLinha = "   Telefone da Empresa   : " & C_Telefone.Text
  Printer.Print sLinha
    
  ' dados do conjuge
  Printer.Print ""
  Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
  Printer.Print "   Dados de Crédito - Conjuge"
  Printer.Print ""
  sLinha = "   Nome do Conjuge       : " & C_C_Nome.Text
  Printer.Print sLinha
  
  sLinha = "   RG/CPF                : " & C_C_RG.Text
  Printer.Print sLinha

  sLinha = "   Empresa onde Trabalha : " & C_C_Empresa.Text
  Printer.Print sLinha
    
  If Trim(C_C_Salário.Text) = "" Then
      sLinha = "   Salário Atual         : 0,00   "
  Else
      sLinha = "   Salário Atual         : " & C_C_Salário.Text
  End If
  
  sLinha = "   Cargo                 : " & C_C_Cargo.Text
  Printer.Print sLinha
    
  If Trim(C_C_Contratação.Text) = "" Then
      sLinha = "   Data da Contratação   : __/__/____"
  Else
      sLinha = "   Data da Contratação   : " & C_C_Contratação.Text
  End If
  Printer.Print sLinha
  
  sLinha = "   Telefone da Empresa   : " & C_C_Telefone.Text
  Printer.Print sLinha
  Printer.Print ""
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""
  Printer.Print "   _________________________________________________________________________________________________________________"
  Printer.Print ""
  Printer.Print "   _________________________________________________________________________________________________________________"
    
  Printer.EndDoc
  
  Exit Sub
Erro:
    MsgBox "Erro na impressão da Ficha Cadastral " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case 0
          Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY Código", dbOpenDynaset)
        Case 1
          Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY Nome", dbOpenDynaset)
        Case 2
          Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY Fantasia", dbOpenDynaset)
        Case 3
          Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY [Fone 1]", dbOpenDynaset)
        Case 4
          Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código<> 0 ORDER BY Cidade, Código", dbOpenDynaset)
        Case 5
          Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY CGC", dbOpenDynaset)
      End Select
  End Select
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsCliFor
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
      Call ResetAllGrids
    End If
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  With rsCliFor
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
      Call ResetAllGrids
    End If
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  With rsCliFor
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
      Call ResetAllGrids
    Else
      Beep
      .MoveNext
    End If
  End With
End Sub

Private Sub MoveNext()
  On Error Resume Next
  With rsCliFor
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
      Call ResetAllGrids
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Private Sub ResetAllGrids()

  Screen.MousePointer = vbHourglass
  Call LoadGridContatos
  Call LoadGridContatosEfetuados
  Call LoadGridGruposMalaDireta
  Call LoadGridCaract
  Call LoadGridNumeraveis
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub DeleteRecord()
  Dim Resposta As Integer
  Dim Aux_Código As Double
  Dim Aux_Str As String
'  Dim F As Form
  
  If IsNull(Num_Registro) Or (rsCliFor.EOF And rsCliFor.BOF) Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "Não existe registro para apagar !"
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

'  Set F = New frmGerente
'  F.Show vbModal
'  Set F = Nothing
'  If gsRetornoDoc <> "OK" Then
'    Exit Sub
'  End If
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  
  Call StatusMsg("")
  Aux_Str = "Ao apagar clientes/fornecedores que já tenham sido usados nas tabelas abaixo e os relatórios que delas dependem ficarão incorretos."
  Aux_Str = Aux_Str & vbCrLf & vbCrLf & "Comissões"
  Aux_Str = Aux_Str & vbCrLf & "Contas do Cliente"
  Aux_Str = Aux_Str & vbCrLf & "Contas a Pagar / Contas a Receber"
  Aux_Str = Aux_Str & vbCrLf & "Contatos / Contatos Efetuados"
  Aux_Str = Aux_Str & vbCrLf & "Controle de Cartões e Cheques Pré-Datados"
  Aux_Str = Aux_Str & vbCrLf & "Empréstimos"
  Aux_Str = Aux_Str & vbCrLf & "Vendas e Compras"
  Aux_Str = Aux_Str & vbCrLf & vbCrLf & "Qualquer relatório que use uma das tabelas acima terá valores e totais incorretos caso você prossiga."
  Aux_Str = Aux_Str & vbCrLf & vbCrLf & "É ALTAMENTE DESCONSELHÁVEL APAGAR CLIENTES/FORNECEDORES."
  Aux_Str = Aux_Str & vbCrLf & "AO INVÉS DE APAGAR CONSIDERE A OPÇÃO DE APENAS DEIXÁ-LO INATIVO."
  Aux_Str = Aux_Str & vbCrLf & Chr(13)
  Aux_Str = Aux_Str + "Deseja realmente apagar este cliente/fornecedor ?"
  Resposta = MsgBox(Aux_Str, 20, "ATENÇÃO!!")
  
  If Resposta = 6 Then
    rsCli_Cred.Index = "Código"
    rsCli_Cred.Seek "=", rsCliFor("Código")
    If Not rsCli_Cred.NoMatch Then rsCli_Cred.Delete
    
    Call StatusMsg("Apagando conta do cliente...")
    Set rsConta_Cliente = db.OpenRecordset("Conta Cliente")
    rsConta_Cliente.Index = "Cliente"
Lp_Conta_Cliente:
    rsConta_Cliente.Seek ">", rsCliFor("Código"), 0
    If Not rsConta_Cliente.NoMatch Then
      If rsConta_Cliente("Cliente") = rsCliFor("Código") Then
        rsConta_Cliente.Delete
        GoTo Lp_Conta_Cliente
      End If
    End If
     
    Call StatusMsg("Apagando contatos...")
    Set rsContatos = db.OpenRecordset("Contatos")
    rsContatos.Index = "Cliente"
LP_Contatos:
    rsContatos.Seek ">", rsCliFor("Código"), 0
    If Not rsContatos.NoMatch Then
      If rsContatos("Cliente") = rsCliFor("Código") Then
        rsContatos.Delete
        GoTo LP_Contatos
      End If
    End If
     
    Call StatusMsg("Apagando contatos efetuados...")
    '21/12/2006 - Anderson - A abertura do recordset ficou registrada no Form Load para a utilização na nova tela de contatos efetuados
    'Set rsContatos_Efetuados = db.OpenRecordset("Contatos Efetuados")
    rsContatos_Efetuados.Index = "Cliente"
LP_Contatos_Efetuados:
    rsContatos_Efetuados.Seek ">", rsCliFor("Código"), CDate("01/01/2050"), 0
    If Not rsContatos_Efetuados.NoMatch Then
      If rsContatos_Efetuados("Cliente") = rsCliFor("Código") Then
        rsContatos_Efetuados.Delete
        GoTo LP_Contatos_Efetuados
      End If
    End If
    
    Call StatusMsg("Apagando empréstimos...")
    Set rsEmprestimos = db.OpenRecordset("Consignação Saída")
    rsEmprestimos.Index = "Cliente2"
LP_Empréstimos:
    rsEmprestimos.Seek ">", rsCliFor("Código"), 0, 0, "", 0, 0, 0, 0
    If Not rsEmprestimos.NoMatch Then
      If rsEmprestimos("Cliente") = rsCliFor("Código") Then
        rsEmprestimos.Delete
        GoTo LP_Empréstimos
      End If
    End If
    
    Call StatusMsg("Apagando mala direta...")
    rsMala.Index = "Cliente"
LP_Mala:
    rsMala.Seek ">", rsCliFor("Código"), 0
    If Not rsMala.NoMatch Then
      If rsMala("Cliente") = rsCliFor("Código") Then
        rsMala.Delete
        GoTo LP_Mala
      End If
    End If
    
    
    Call StatusMsg("Apagando resumo...")
    Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
    rsResumo_Clientes.Index = "Cliente"
LP_Resumo_Clientes:
    rsResumo_Clientes.Seek ">", rsCliFor("Código"), CDate("01/01/2050"), 0, 0, 0, 0
    If Not rsResumo_Clientes.NoMatch Then
      If rsResumo_Clientes("Cliente") = rsCliFor("Código") Then
        rsResumo_Clientes.Delete
        GoTo LP_Resumo_Clientes
      End If
    End If
    
    rsCliFor.Delete
    Num_Registro = Null
    
    datCliForCod.Refresh
    datCliForNome.Refresh
    datFantasia.Refresh
    
    Call ClearScreen
  End If

End Sub

Private Sub SearchRecord()
  frmPesquisaCliFor.Show
End Sub

Private Sub B_Expande_Click()
  frmComenta.Comenta.Text = Comentários.Text
  cboCodigo.Enabled = False
  frmComenta.Show vbModal
  DoEvents
  Comentários = gsBuffer
  cboCodigo.Enabled = True
End Sub

'30/04/2003 - mpdea
'Incluído tratamento na execução das transações com a variável blnInTransaction
Private Sub UpdateRecord()
  Dim strAux As String
  'Dim nCtErr As Integer
  'Variávels de Tratamento de Erro
  Dim intRepeatUpdateLocked As Integer
  Dim blnInTransaction    As Boolean
  Dim rstTabelaPrecos     As Recordset
  
  
  On Error GoTo ErrHandler
'  nCtErr = 0
  
  If IsNull(Num_Registro) And gbDemoVersion Then
    rsCliFor.MoveLast
    rsCliFor.MoveFirst
    If rsCliFor.RecordCount >= NMAXREGDEMO Then
      gsTitle = LoadResString(201)
      gsMsg = LoadResString(13)
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Exit Sub
    End If
  End If
  
  cboCodigo.Text = gsHandleNull(cboCodigo.Text & "")
  If CLng(cboCodigo.Text) <= 0 Then
    DisplayMsg "Código incorreto, verifique."
    cboCodigo.SetFocus
    Exit Sub
  End If
'  On Error GoTo 0
  
  If Len(Trim(cboNome.Text & "")) = 0 Then
    DisplayMsg "Nome incorreto, verifique."
    cboNome.SetFocus
    Exit Sub
  End If
  cboNome.Value = cboNome.Text

  '10/07/2006 - mpdea
  'Verifica alteração do campo nome
  If Not IsNull(Num_Registro) Then
    If rsCliFor.Fields("Nome").Value <> cboNome.Text Then
      strAux = "Confirma alteração do nome do cliente?"
      strAux = strAux & vbCrLf & vbCrLf & "De: " & rsCliFor.Fields("Nome").Value
      strAux = strAux & vbCrLf & "Para: " & cboNome.Text & "?"
      Select Case MsgBox(strAux, vbExclamation + vbYesNoCancel, "Atenção")
        Case vbNo
          cboNome.Text = rsCliFor.Fields("Nome").Value
        Case vbCancel
          Exit Sub
      End Select
    End If
  End If
  
  
  If IsNull(Contabilidade.Text) Then Contabilidade.Text = 0
  If Not IsNumeric(Contabilidade.Text) Then Contabilidade.Text = 0
  If Val(Contabilidade.Text) < 0 Then Contabilidade.Text = 0
  
  If Not IsNumeric(Desconto.Text) Then Desconto.Text = 0
  If CDbl(Desconto.Text) < 0 Then Desconto.Text = 0
  If CDbl(Desconto.Text) > 100 Then Desconto.Text = 0
  If IsNull(mskDatNasc.Text) Then mskDatNasc.Text = "00/00/00"
  If IsNull(mskDatAbertura.Text) Then mskDatAbertura.Text = "00/00/00"
'  If Not IsNull(mskDatNasc.Text) And Not IsDate(mskDatNasc.Text) Then
'     DisplayMsg "Data de Nascimento inválida, verifique."
'     mskDatNasc.SetFocus
'     Exit Sub
'  End If
  
  If O_Cliente.Value = True And O_Física.Value = True Then
     If optMasculino.Value = False And optFeminino.Value = False Then
        DisplayMsg "Sexo incorreto, verifique."
        optMasculino.SetFocus
        Exit Sub
      End If
  End If
  
  
  '17/04/2003 - mpdea
  'Verifica se possui configuração personalizada
  'Preenchimento obrigatório do campo CNPJ-CPF
  If g_blnCheckChangePersonalized("CNPJ-CPF") Then
    If CGC.Text = "" Then
      If O_Jurídica.Value Then
        DisplayMsg "Preencher CNPJ."
      Else
        DisplayMsg "Preencher CPF."
      End If
      CGC.SetFocus
      Exit Sub
    End If
  End If
  
  '---------------------------------------------------------
  ' Data.......: 20/01/2004
  ' Dev........: Maikel
  ' Case.......: TV Shopping
  ' Finalidade.: Exigir que o usuário preencha o endereço
    If CheckSerialCaseMod("QS39945-043") Then
      If Len(Trim(Endereço.Text)) <= 0 Then
        MsgBox "Favor preencher o campo Endereço !", vbCritical, "Quick Store"
        Exit Sub
      End If
    End If
  '---------------------------------------------------------
  ' Data.......: 20/01/2004
  ' Dev........: Maikel
  ' Case.......: TV Shopping
  ' Finalidade.: Exigir que o usuário preencha o cep
    If CheckSerialCaseMod("QS39945-043") Then
      If Len(Trim(CEP.Text)) <= 0 Then
        MsgBox "Favor preencher o campo CEP !", vbCritical, "Quick Store"
        Exit Sub
      End If
    End If
  '---------------------------------------------------------
  
  '09/11/2004 - Daniel
  'Case: Teknika
  'I.E. preenchimento obrigatório
  If m_blnTeknika Then
    If Len(Insc.Text) <= 0 Then
      MsgBox "Inscrição Estadual vazia, verifique.", vbCritical, "Quick Store"
      Insc.SetFocus
      Exit Sub
    End If
  End If
  
  '22/10/2003 - Maikel
  '---[ Verifica se a tabela de preços indicada é válida ]---'
    If Len(Trim(cboTabelaPrecos.Text)) > 0 Then
      Set rstTabelaPrecos = db.OpenRecordset("SELECT Distinct TABELA FROM Preços WHERE Tabela = '" & cboTabelaPrecos.Text & "'", dbOpenSnapshot)
      
      With rstTabelaPrecos
        If (.BOF And .EOF) Then
          MsgBox "Tabela de preços inválida, verifique !", vbCritical, "Quick Store"
          Exit Sub
        End If
        .Close
        Set rstTabelaPrecos = Nothing
      End With
    End If
  '---[ Verifica se a tabela de preços indicada é válida ]---'
  
  '30/04/2003 - mpdea
  'Verifica se possui configuração personalizada
  'Preenchimento obrigatório do campo CEP
  If g_blnCheckChangePersonalized("CEP") Then
    If CEP.Text = "" Then
      DisplayMsg "Preencher CEP."
      CEP.SetFocus
      Exit Sub
    End If
  End If
  
  
  '---------------------------------------------------------
  ' Data.......: 18/02/2004
  ' Dev........: Daniel
  ' Case.......: TV Shopping x Aproveitado idéia
  ' Finalidade.: Evitar duplicidade de CNPJ/CPF
  If Not m_blnFlag Then
    Call CGC_LostFocus
  End If
  '---------------------------------------------------------
  
  '17/06/2005 - Daniel
  '
  'Validação para a frame "Sacador/Avalista"
  'Caso o campo Nome do Sacador/Avalista esteja preenchido,
  'o sistema exigirá que o CPF ou CNPJ seja informado
  If Len(txtNomeSacadorAvalista.Text) > 0 Then  'Foi preenchido
    If Len(txtCPFSacadorAvalista.Text) <= 0 Then
      MsgBox "Favor informar o CPF/CNPJ do Sacador/Avalista.", vbExclamation, "Atenção"
      txtCPFSacadorAvalista.SetFocus
      Exit Sub
    End If
  End If
  '
  'Validação para caso o usuário não informe
  'se é o CPF ou CNPJ do Sacador/Avalista
  If Len(txtCPFSacadorAvalista.Text) > 0 Then 'Está preenchido
    If optCPF.Value = False And optCNPJ.Value = False Then
      MsgBox "Você está informando o CPF ou CNPJ do Sacador/Avalista ?", vbExclamation, "Atenção"
      optCPF.ForeColor = &HFF&
      optCNPJ.ForeColor = &HFF&
      Exit Sub
    End If
  End If
  
  '18/01/2007 - Anderson
  'Solicitação senha do gerente ao alterar o vendedor relacionado ao cliente
  If gbVendedorSenhaGerente Then
    If Combo_Vendedor.Text <> strVendedor Then
      If MsgBox("Você não tem permissão para alterar o campo vendedor. Suas alterações não serão gravadas. A senha do gerente será necessária para concluir a gravação do cliente." & Chr(13) & "Deseja continuar assim mesmo?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
        If Not frmGerente.gbSenhaGerente Then
          Exit Sub
        End If
      Else
        Exit Sub
      End If
    End If
  End If
  
  '18/07/2007 - Anderson
  'Validação do tipo de cliente SadigWeb
  If cboSadigWebTipo.Text <> "(Nenhum)" And _
     cboSadigWebTipo.Text <> "AG - AGROPECUARIA" And _
     cboSadigWebTipo.Text <> "CC - CRIADOR DE CAES" And _
     cboSadigWebTipo.Text <> "CG - CRIADOR DE GATOS" And _
     cboSadigWebTipo.Text <> "CL - CLINICA VETERINARIA" And _
     cboSadigWebTipo.Text <> "CP - CLINICA VETERINARIA COM PETSHOP" And _
     cboSadigWebTipo.Text <> "CR - CRIADOR DE CAES E GATOS" And _
     cboSadigWebTipo.Text <> "PC - PETSHOP COM CLINICA VETERINARIA" And _
     cboSadigWebTipo.Text <> "PE - PETSHOP" And _
     cboSadigWebTipo.Text <> "VE - VETERINARIO" Then
     
    MsgBox "Tipo de cliente utilizado para exportar dados no sistema Sadig Web, está inválido!", vbCritical, "Quick Store"
    cboSadigWebTipo.SetFocus
    Exit Sub
          
  End If

  Call ConsultarLimiteCreditoCliente
  
  '---------------------------------------------------------
  ' GRAVAÇÃO
  '---------------------------------------------------------
  
  Call StatusMsg("Gravando ...")
  
  Screen.MousePointer = vbHourglass
    
  Call ws.BeginTrans
  blnInTransaction = True
    
   'Solicitado por Petrys para exigir senha do gerente quando o cliente é desbloqueado
   'Eduardo - 06/02/2015
    'Dim QuickBD As Database
    Dim rsBloqueado As Recordset
    If CheckSerialCaseMod("QS40073-020") Then
      Set rsBloqueado = db.OpenRecordset("Select Bloqueado from Cli_For Where Código = " & cboCodigo & "")
      If Not rsBloqueado.EOF Then
        If Bloqueado.Value = 0 And rsBloqueado!Bloqueado = True Then
          MsgBox "Cliente bloqueado, favor inserir senha do gerente.", vbInformation, "Cliente bloqueado!"
            If Not frmGerente.gbSenhaGerente Then
              Exit Sub
            End If
        End If
      End If
    End If
    
  '18/01/2007 - Anderson
  'Solicitação senha do gerente ao alterar o vendedor relacionado ao cliente
  strVendedor = Combo_Vendedor.Text
  
  With rsCliFor
  
    If IsNull(Num_Registro) Then
      .AddNew
      .Fields("Código") = Val(cboCodigo.Text)
      strAux = "Registro inserido"
    Else
      .Edit
      strAux = "Registro alterado"
    End If
    
    If O_Cliente.Value = True Then .Fields("Tipo") = "C"
    If O_Fornecedor.Value = True Then .Fields("Tipo") = "F"
    If O_Revendedor.Value = True Then .Fields("Tipo") = "R"
    If O_Outros.Value = True Then .Fields("Tipo") = "O"
    
    If O_Jurídica.Value = True Then .Fields("Física_Jurídica") = "J"
    If O_Física.Value = True Then .Fields("Física_Jurídica") = "F"
    
    '---[ 02/06/2003 - Maikel ]---'
    ' Case      : Projeto Deluken
    ' Descrição : Adicionado o campo abaixo para guardar o dia base de acerto de consignações
      .Fields("DiaBaseConsignacao") = IIf(IsNumeric(txtDiaBaseConsignacao.Text), txtDiaBaseConsignacao.Text, 25)
    '-----------------------------'
    
    .Fields("Nome") = LTrim(RTrim(cboNome.Text))
    
    If Len(cboFantasia.Text) > 0 Then
        .Fields("Fantasia") = Trim(cboFantasia.Text) & ""
    Else
        .Fields("Fantasia") = cboFantasia.Text & ""
    End If
    
    .Fields("Endereço") = Endereço.Text & ""
    '26/07/2004 - Daniel
    'Adicionado campo AgenciaPublicidade
    'Case: STC
    If O_Fornecedor.Value Then
      If chkAgencia.Value = vbChecked Then
        .Fields("AgenciaPublicidade").Value = True
      Else
        .Fields("AgenciaPublicidade").Value = False
      End If
    Else
      .Fields("AgenciaPublicidade").Value = False
    End If
    
    .Fields("Complemento") = Complemento.Text & ""
    .Fields("Bairro") = Bairro.Text & ""
    '.Fields("Cidade") = Cidade.Text & ""
    .Fields("Cidade") = cbo_cidades.Text & ""
    .Fields("Estado") = UCase(Estado.Text) & ""
    .Fields("CEP") = CEP.Text & ""
    .Fields("CGC") = CGC.Text & ""
    '09/11/2004 - Daniel
    If m_blnTeknika Then
      .Fields("Inscrição") = UCase(Insc.Text) & ""
    Else
      .Fields("Inscrição") = Insc.Text & ""
    End If
    
    '27/08/2009 - mpdea
    'Novos campos para NFe
    Dim pab_tempEnderecoNumero As String
    If Trim(txtEnderecoNumero.Text) = "" Then
      pab_tempEnderecoNumero = "S/N"
    Else
      pab_tempEnderecoNumero = Trim(txtEnderecoNumero.Text)
    End If
    .Fields("Endereço Número").Value = pab_tempEnderecoNumero
    
    If Trim(txtPais.Text) = "" Then
        .Fields("Pais").Value = "Brasil"
    Else
        .Fields("Pais").Value = Trim(txtPais.Text)
    End If
    .Fields("InscricaoMunicipal").Value = txtInscricaoMunicipal.Text
    .Fields("InscricaoSuframa").Value = txtInscricaoSuframa.Text
    .Fields("CNAE").Value = txtCNAE.Text
    
    .Fields("Fone 1") = Fone1.Text & ""
    .Fields("Fone 2") = Fone2.Text & ""
    .Fields("Fax") = Fax.Text & ""
    .Fields("DataNascimento") = IIf(IsDate(mskDatNasc.Text), mskDatNasc.Text, Null)
    .Fields("datAberturaCadastro") = IIf(IsDate(mskDatAbertura.Text), mskDatAbertura.Text, Null)
    .Fields("Sexo") = IIf(optMasculino.Value, "M", "F")
    
    .Fields("Inativo") = False
    If Inativo.Value = 1 Then .Fields("Inativo") = True
    
    .Fields("Bloqueado") = False
    If Bloqueado.Value = 1 Then .Fields("Bloqueado") = True
    
    .Fields("Faturado") = False
    If Faturado.Value = 1 Then .Fields("Faturado") = True
    
    .Fields("Tem Conta") = False
    If Conta.Value = 1 Then .Fields("Tem Conta") = True
    
    .Fields("Sem Mala Direta") = False
    If Sem_Mala.Value = 1 Then .Fields("Sem Mala Direta") = True
    
    .Fields("Desconto") = CDbl(Desconto.Text)
    
    .Fields("Comentários") = Comentários.Text
    
    If Nome_Transp.Caption <> "" Then
      .Fields("Transportadora") = Combo_Transp.Text
    Else
      .Fields("Transportadora") = 0
    End If
    
    .Fields("Contabilidade") = Contabilidade.Text
    
    '05/05/2004 - Daniel
    'Adicionado o campo IsentoIPI
    If chkIsencaoIPI.Value = vbChecked Then
      .Fields("IsentoIPI").Value = True
    Else
      .Fields("IsentoIPI").Value = False
    End If
    '06/05/2004 - Daniel
    'Adicionado o campo ObsIsentoIPI
    .Fields("ObsIsentoIPI").Value = CStr(txtMensagemIPI.Text)
    
    '13/07/2004 - Daniel
    'Foram adicionados nesta data, 05 campos em (Cli_For) para uso exclusivo
    'da TV Shopping otimização da Curva "A B C D". Para outros clientes do Quick
    'serão preenchidos estes campos com valores iguais a zero
    If Not CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959") Then
      .Fields("CodGrupo").Value = 0
      .Fields("TotDinheiroBoletos").Value = 0
      .Fields("TotCheques").Value = 0
      .Fields("TotCartoes").Value = 0
      .Fields("TotRecebido").Value = 0
    End If
    '--------------------------------------------------------------------
    
    If Not IsNumeric(Combo_Vendedor.Text) Then Combo_Vendedor.Text = 0
    If Val(Combo_Vendedor.Text) < 0 Or Val(Combo_Vendedor.Text) > 9999 Then Combo_Vendedor.Text = 0
    .Fields("Vendedor") = Val(Combo_Vendedor.Text)
    .Fields("TabelaPrecoPadrao") = Trim(cboTabelaPrecos.Text)
    .Fields("email") = email.Text
    .Fields("Home Page") = home_page.Text
     
    
    .Fields("Conta Cobrança") = 0
    If Nome_Cobrança.Caption <> "" Then
      .Fields("Conta Cobrança") = Combo_Cobrança.Text
    End If
    
    If O_Frete_Pagar.Value = True Then
      .Fields("CodTipoFrete") = "1"
    Else
      .Fields("CodTipoFrete") = "2"
    End If
    
    .Fields("Endereço Cob") = Endereço_Cob.Text
    .Fields("Complemento Cob") = Complemento_Cob.Text
    .Fields("Bairro Cob") = Bairro_Cob.Text
    .Fields("Cidade Cob") = Cidade_Cob.Text
    .Fields("Estado Cob") = Estado_Cob.Text
    .Fields("CEP Cob") = CEP_Cob.Text
    .Fields("Mensagem Boleto") = Trim(Boleto.Text & "")
    .Fields("Data Alteração") = Format(Date, "dd/mm/yyyy")
    
    If IsNull(Limite.Text) Or Limite.Text = "" Then Limite.Text = 0
    .Fields("Limite Crédito") = CDbl(Limite.Text)
    
    '15/06/2005 - Daniel
    'Tratamento para os campos NomeSacadorAvalista, CPFSacadorAvalista e CPFCNPJSacadorAvalista
    .Fields("NomeSacadorAvalista").Value = txtNomeSacadorAvalista.Text & ""
    .Fields("CPFSacadorAvalista").Value = txtCPFSacadorAvalista.Text & ""
    
    If optCPF.Value Then .Fields("CPFCNPJSacadorAvalista").Value = "CPF"
    If optCNPJ.Value Then .Fields("CPFCNPJSacadorAvalista").Value = "CNPJ"
    '------------------------------------------------------------------------------------------
    
    '18/07/2007 - Anderson
    'Grava informação do tipo de cliente SadigWeb
    .Fields("SadigWeb_Tipo").Value = cboSadigWebTipo.Text
    
    .Fields("IndicadorIE").Value = Left(cboIndicadorIE.Text, 1)
    
    If Trim(txtPais.Text) = "" Then
        txtPais.Text = "Brasil"
    End If
    
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
    
  End With
  
  'Grava crédito
  With rsCli_Cred
    .Index = "Código"
    .Seek "=", cboCodigo.Text
    If .NoMatch Then
      .AddNew
      .Fields("Código") = cboCodigo.Text
    Else
      .Edit
    End If
    .Fields("Empresa") = C_Empresa.Text
    .Fields("Salário") = C_Salário.Text
    '30/06/2004 - Daniel
    'Adicionado campo [Cli_For - Crédito].PercentualLimite
    If Len(mskPercentualLimite.Text) <= 0 Then
      .Fields("PercentualLimite").Value = 0
    Else
      .Fields("PercentualLimite").Value = CDbl(Format(mskPercentualLimite.Text, "###,###,##0.00"))
    End If
    '----------------------------------------------------------------------------------------------
    .Fields("Cargo") = C_Cargo.Text
    .Fields("Contratação") = C_Contratação.Text
    .Fields("Telefone") = C_Telefone.Text
    .Fields("Conjugê") = C_C_Nome.Text
    .Fields("C_CPF") = C_C_RG.Text
    .Fields("C_Empresa") = C_C_Empresa.Text
    .Fields("C_Salário") = C_C_Salário.Text
    .Fields("C_Cargo") = C_C_Cargo.Text
    .Fields("C_Contratação") = C_C_Contratação.Text
    .Fields("C_Telefone") = C_C_Telefone.Text
    .Fields("OBS") = C_Obs.Text
    .Update
  End With
  
    
  Call WriteGridContatos
  
  Call WriteGridContatosEfetuados
  
  Call WriteGridGruposMalaDireta
  
  Call WriteGridCaract
  
  Call WriteGridNumeraveis
  '19/07/2007 - Anderson
  'Não habilita o campo Lembrar em quando o cliente está Inativo
  'Solicitante: Supri Print
  If bolCaseSupriPrint Then
    If rsCliFor("Inativo") Then
    'Atualiza as pendências - Case Supri Print
      db.Execute "UPDATE [Contatos Efetuados] " & _
                 "Set Pendência = 0, [Data Aviso]='" & Null & "' " & _
                 "Where Cliente = " & cboCodigo.Text
      txtContatosEfetuadosLembrarEm.Visible = False
      Label57.Visible = False
      chkContatoEfetuadoPendencia.Visible = False
    Else
      txtContatosEfetuadosLembrarEm.Visible = True
      Label57.Visible = True
      chkContatoEfetuadoPendencia.Visible = True
    End If
  End If
  
  Call ws.CommitTrans
  blnInTransaction = False
  
  Call LoadGridContatosEfetuados
  
  If Tab1.TabEnabled(1) = False Then
    Tab1.TabEnabled(1) = True
    Tab1.TabEnabled(2) = True
    Tab1.TabEnabled(3) = True
    Tab1.TabEnabled(4) = True
    Tab1.TabEnabled(5) = True
    
    '27/05/2004 - mpdea
    'Web
    Tab1.TabEnabled(6) = rsCliFor("Web").Value
    
  End If
  
  Tab1.Tab = 0
  
  Screen.MousePointer = vbDefault
  
  Call StatusMsg(strAux & " com sucesso.")


'-----------------------------------------------------------------------------------------------------------------------------------------------------------
' Joga dados do cliente para o banco do GestoPDV por causa do PAF
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

   If O_Cliente = True Then
     Dim GestoBD As Database
     Dim Consumidor As Recordset
     Dim CodConsumidor As Double
     CodConsumidor = CDbl(cboCodigo.Text)
     Dim Nome_cli As String
     Nome_cli = Left(LTrim(RTrim(cboNome.Text)), 40)
     If frmParametros.VerificaPAF = True Then
       Set rsParametros = db.OpenRecordset("Select * from [Parâmetros Filial] Where Filial = " & gnCodFilial & ";")
                     
       Dim fso As New FileSystemObject
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FileExists(rsParametros("BancoPDV").Value & "\Gesto.mde") Then

       Set GestoBD = OpenDatabase(rsParametros("BancoPDV").Value & "\Gesto.mde", False, False)
       Set Consumidor = GestoBD.OpenRecordset("Select * from Consumidor Where CODIGO_CONSUMIDOR = " & CodConsumidor & ";")
       If Consumidor.EOF Then
         Consumidor.AddNew
         Consumidor!CODIGO_CONSUMIDOR = CodConsumidor
         Consumidor!Nome_Consumidor = Nome_cli
         Consumidor!razao_Social = LTrim(RTrim(cboNome.Text))
'         If Cidade.Text <> "" Then
'           Consumidor!Cidade = Cidade.Text
'         End If
         If cbo_cidades.Text <> "" Then
           Consumidor!Cidade = cbo_cidades.Text
         End If
         If Endereço.Text <> "" Then
           Consumidor!Nome_Rua = Endereço.Text
         End If
         If Trim(txtEnderecoNumero.Text) <> "" Then
           Consumidor!Num_Rua = Trim(txtEnderecoNumero.Text)
         End If
         If Complemento.Text <> "" Then
           Consumidor!Complemento = Complemento.Text
         End If
         If Bairro.Text <> "" Then
           Consumidor!Bairro = Bairro.Text
         End If
         If CEP.Text <> "" Then
           Consumidor!CEP = CEP.Text
         End If
         If Estado.Text <> "" Then
           Consumidor!UF = Estado.Text
         End If
         If Fone1.Text <> "" Then
           Consumidor!DDD = Left(Fone1.Text, 3)
           Consumidor!FONE_RESIDENCIAL = Right(Fone1.Text, 14)
         End If
         If Fone2.Text <> "" Then
           Consumidor!FONE_CELULAR = Right(Fone2.Text, 14)
         End If
         If Fax.Text <> "" Then
           Consumidor!Fax = Fax.Text
         End If
         If CGC.Text <> "" Then
           Consumidor!CGC_CPF = CGC.Text
         End If
         If O_Jurídica = True Then
           If Insc.Text <> "" Then
             Consumidor!INSCR_ESTADUAL = Insc.Text
           End If
           Consumidor!Tipo_Pessoa = "J"
         Else
           Consumidor!Tipo_Pessoa = "F"
         End If
         If Inativo = 0 Then
           Consumidor!Ativo = True
         Else
           Consumidor!Ativo = False
         End If
         Consumidor!Cod_Procedencia = 7
         Consumidor!CODIGO_TECNICO = 99
         Consumidor.Update
       Else
         Consumidor.Edit
         Consumidor!Nome_Consumidor = Nome_cli
         Consumidor!razao_Social = RTrim(LTrim(cboNome.Text))
'         If Cidade.Text <> "" Then
'           Consumidor!Cidade = Cidade.Text
'         End If
         If cbo_cidades.Text <> "" Then
           Consumidor!Cidade = cbo_cidades.Text
         End If
         If Endereço.Text <> "" Then
           Consumidor!Nome_Rua = Endereço.Text
         End If
         If Trim(txtEnderecoNumero.Text) <> "" Then
           Consumidor!Num_Rua = Trim(txtEnderecoNumero.Text)
         End If
         If Complemento.Text <> "" Then
           Consumidor!Complemento = Complemento.Text
         End If
         If Bairro.Text <> "" Then
           Consumidor!Bairro = Bairro.Text
         End If
         If CEP.Text <> "" Then
           Consumidor!CEP = CEP.Text
         End If
         If Estado.Text <> "" Then
           Consumidor!UF = Estado.Text
         End If
         If Fone1.Text <> "" Then
           Consumidor!DDD = Left(Fone1.Text, 3)
           Consumidor!FONE_RESIDENCIAL = Right(Fone1.Text, 14)
         End If
         If Fone2.Text <> "" Then
           Consumidor!FONE_CELULAR = Right(Fone2.Text, 14)
         End If
         If Fax.Text <> "" Then
           Consumidor!Fax = Fax.Text
         End If
         If CGC.Text <> "" Then
           Consumidor!CGC_CPF = CGC.Text
         End If
         If O_Jurídica = True Then
           If Insc.Text <> "" Then
             Consumidor!INSCR_ESTADUAL = Insc.Text
           End If
           Consumidor!Tipo_Pessoa = "J"
         Else
           Consumidor!Tipo_Pessoa = "F"
         End If
         If Inativo = 0 Then
           Consumidor!Ativo = True
         Else
           Consumidor!Ativo = False
         End If
         Consumidor.Update
       End If
      End If
     End If
   End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

  Exit Sub

ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        Call WaitSeconds(1)  'Aguarda um segundo
        Screen.MousePointer = vbHourglass
        Resume
      Else
        
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Screen.MousePointer = vbHourglass
          Resume
        Else
          'Cancelamento da transação
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
        
'        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
'          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Clientes e Fornecedores - Gravar") = vbOK Then
'          intRepeatUpdateLocked = 0
'          Resume
'        End If
      End If
    Case Else
      'Cancelamento da transação
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      MsgBox "Erro em Clientes e Fornecedores - Gravar: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
      Exit Sub
    
'      'Outros Erros
'      Select Case frmErro.gnShowErr(Err.Number, "Clientes e Fornecedores - Gravar")
'        Case 0 'Repetir
'          Resume
'        Case 1 'Prosseguir
'          Resume Next
'        Case 2 'Sair
'          If blnInTransaction Then ws.Rollback
'          Exit Sub
'        Case 3 'Encerrar
'          If blnInTransaction Then ws.Rollback
'          End
'      End Select
  End Select
  
'ErrHandler:
'  If Err.Number = 3260 Then
'    nCtErr = nCtErr + 1
'    If nCtErr <= 50 Then
'      Call WaitSeconds(1)
'      Err.Clear
'      Resume
'    End If
'    gsMsg = "Não foi possível atualizar neste momento o Cadastro de Clientes/Fornecedores."
'  Else
'    gsMsg = "Não foi possível atualizar o Cadastro de Clientes/Fornecedores."
'  End If
'  Screen.MousePointer = vbDefault
'  gsTitle = LoadResString(201)
'  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
'  gnStyle = vbOKOnly + vbCritical
'  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'  On Error Resume Next
'  Exit Sub
  
End Sub

Private Sub GetInformation()
  On Error GoTo ErrHandler
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
   DisplayMsg "Encontre um registro antes."
   Exit Sub
  End If
  
  gsCodCliente = rsCliFor("Código").Value
  
  '20/04/2006 - mpdea
  'Modificado exibição do form de informações do cliente
  'para poder ser acessível de diversas maneiras (Ex.: VR CheckOut)
  frmInformacoes.Show , frmMain
  
  Exit Sub

ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apresentar o registro em Informações."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
 
End Sub

'Private Sub B_Limite_Click()
Private Sub ConsultarLimiteCreditoCliente()
  Dim Usado As Double
  Dim Limi As Double
  Dim Erro As Boolean
  Dim Resto As Double
  
  Erro = False
  
  mskEdit_limiteCreditoDisponivel.Text = "0,00"
  mskEdit_limiteCreditoUtilizado.Text = "0,00"
  
  If IsNull(Limite.Text) Then Erro = True
  If Erro = False Then If Limite.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Limite.Text) Then Erro = True
  If Erro = True Then
    Limi = 0
  Else
    Limi = CDbl(Limite.Text)
  End If
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Cadastre ou encontre um cliente antes."
    Exit Sub
  End If
  
  If Limi = 0 Then
    'DisplayMsg "Este cliente não tem crédito limitado."
    Exit Sub
  End If
  
  Usado = Pega_Limite_Usado(rsCliFor("Código"))
  
  Resto = Limi - Usado
  
  If Usado = 0 Then
    'DisplayMsg "Este cliente tem todo o limite disponível."
    mskEdit_limiteCreditoDisponivel.Text = "0,00"
    mskEdit_limiteCreditoUtilizado.Text = "0,00"
  Else
    'DisplayMsg "Valor de crédito já utilizado " + Format(Usado, "###,###,##0.00") + " , disponível " + Format(Resto, "###,###,##0.00")
    mskEdit_limiteCreditoDisponivel.Text = Format(Resto, "###,###,##0.00")
    mskEdit_limiteCreditoUtilizado.Text = Format(Usado, "###,###,##0.00")
  End If
End Sub

Private Sub AddToMalaDireta()
  Call StatusMsg("")
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre um cliente antes."
    Exit Sub
  End If
  
  rsMala_Tempo.AddNew
    rsMala_Tempo("Cliente") = rsCliFor("Código")
    rsMala_Tempo("Nome") = ""
  rsMala_Tempo.Update
  
  DisplayMsg "Adicionado à mala direta."
  
  
End Sub

Private Sub cboFantasia_Click()
  If Not rsCliFor.EOF And Not rsCliFor.BOF Then
    rsCliFor.FindFirst "Código = " & CLng(cboFantasia.Columns(1).Text)
    Call ShowRecord
    Call ResetAllGrids
  End If
End Sub

Private Sub cboFantasia_CloseUp()
  If Not rsCliFor.EOF Then
   If Not (cboFantasia.Columns(1).Text) = "" Then
      rsCliFor.FindFirst "Código = " & CLng(cboFantasia.Columns(1).Text)
      If Not rsCliFor.EOF Then
        Call ShowRecord
        Call ResetAllGrids
      End If
   End If
  End If
End Sub

Private Sub cboFantasia_GotFocus()
  cboFantasia.SelStart = 0
  cboFantasia.SelLength = Len(cboFantasia.Text)
End Sub

Private Sub cboFantasia_KeyPress(KeyAscii As Integer)
  If Len(cboFantasia.Text) >= rsCliFor("Fantasia").Size Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub cboFantasia_LostFocus()
  gbComboDropped = False
  If Not rsCliFor.EOF Then
    If Len(Trim(cboFantasia.Text)) > 0 Then
      If Mid(cboFantasia.Text, Len(cboFantasia.Text), 1) = "*" Then
        rsCliFor.FindFirst "Fantasia Like '*" & cboFantasia.Text & "'"
        If Not rsCliFor.NoMatch Then
          Call ShowRecord
          Call ResetAllGrids
        End If
      End If
    End If
  End If
 
 If IsNull(cboFantasia.Text) Then Exit Sub
 If cboFantasia.Text = "" Then Exit Sub
 If IsNull(Num_Registro) Then
   rsCliFor2.Index = "Fantasia"
   rsCliFor2.Seek ">", cboFantasia.Text, 0
   If rsCliFor2.NoMatch Then Exit Sub
   If UCase(cboFantasia.Text) = UCase(rsCliFor2("Fantasia")) Then
     DisplayMsg "Atenção, já existe um cliente/fornecedor com este Nome Fantasia. Código " + str(rsCliFor2("Código"))
   End If
 End If
End Sub

Private Sub cboIndicadorIE_Change()
    Insc.Enabled = (Left(cboIndicadorIE.Text, 1) <> 2)
    Insc.Refresh
End Sub

Private Sub cboIndicadorIE_LostFocus()
    Insc.Enabled = (Left(cboIndicadorIE.Text, 1) <> 2)
    Insc.Refresh
End Sub

Private Sub cboNome_GotFocus()
  cboNome.SelStart = 0
  cboNome.SelLength = Len(cboNome.Text)
End Sub

Private Sub CEP_Cob_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then Exit Sub   'Back
 If KeyAscii = 45 Then Exit Sub  ' -
 If (KeyAscii < 45 Or KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub CEP_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then Exit Sub   'Back
 If KeyAscii = 45 Then Exit Sub  ' -
 If (KeyAscii < 45 Or KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub CEP_LostFocus()
  Call StatusMsg("")
End Sub

Private Sub CGC_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 45 Or KeyAscii > 57 Then KeyAscii = 0
 
 Rem aceita . - /
End Sub

Private Sub CGC_LostFocus()
  Dim Aux As Variant
  Dim i, F As Integer
  Dim STR_CGC As String
  
  Call StatusMsg("")

  If Not rsParametros.Fields("NaoPermiteDuplicarCNPJ").Value Then
     If IsNull(CGC.Text) Then Exit Sub
     If CGC.Text = "" Then Exit Sub
  End If
  
  '18/02/2004 - Daniel
  'Case.......: TV Shopping
  'Finalidade.: Mascarar o CPF/CNPJ
  If CheckSerialCaseMod("QS39945-043", "QS39944-959") Then 'Mascarar o CPF/CNPJ
    Dim strCGC  As String
    If Len(Trim(CGC.Text)) > 0 Then 'Há conteúdo
      'Se estiver mascarado sai fora, evitando bagunçar o conteúdo
      If Mid(CGC.Text, 4, 1) = "." Or Mid(CGC.Text, 3, 1) = "." Then
        Exit Sub
      Else
        If Len(Trim(CGC.Text)) = 11 Then   '---> CPF
          strCGC = Mid((CGC.Text), 1, 3) & "." & Mid((CGC.Text), 4, 3) & "." & Mid((CGC.Text), 7, 3) & "-" & Mid((CGC.Text), 10, 2)
          CGC.Text = strCGC
          m_blnFlag = True
        Else
          If Len(Trim(CGC.Text)) = 14 Then '---> CNPJ
            strCGC = Mid((CGC.Text), 1, 2) & "." & Mid((CGC.Text), 3, 3) & "." & Mid((CGC.Text), 6, 3) & "/" & Mid((CGC.Text), 9, 4) & "-" & Mid((CGC.Text), 13, 2)
            CGC.Text = strCGC
            m_blnFlag = True
          End If
        End If
      End If
    End If
  End If 'If CheckSerialCaseMod
  
  
  '11/11/2008 - mpdea
  'Aplica/refaz máscara ao CPF/CNPJ para evitar problemas de duplicidade de cadastro
  Dim str_doc As String
  str_doc = g_str_SomenteNumero(CGC.Text)
  If Len(str_doc) = 14 Then 'CNPJ
    str_doc = Format(str_doc, "@@.@@@.@@@/@@@@-@@")
  ElseIf Len(str_doc) = 11 Then 'CPF
    str_doc = Format(str_doc, "@@@.@@@.@@@-@@")
  End If
  CGC.Text = str_doc
  
  
  'If IsNull(Num_Registro) Then   '18/02/2004 - Daniel
'    rsCliFor2.Index = "CGC"
'    rsCliFor2.Seek ">", CGC.Text, 0
'    If rsCliFor2.NoMatch Then Exit Sub
    
      '23/03/2004 - Daniel
      'Condição se realmente é outro record
      '26/04/2004 - Daniel
      'Primeiro verificar se realmente o cboCodigo.Text está preenchido
      'para não dar Type Mismatch na condição If CLng(cboCodigo.Text) <> rsCliFor2.Fields("Código").Value Then
'      If Len(cboCodigo.Text) > 0 Then
          
'          If Not m_blnNcnpjcpf Then '17/05/2004 - Daniel - A var m_blnNcnpjcpf controla se
'                                    'a empresa tem necessidade de cadastrar n clientes com o mesmo CNPJ/CPF
'            If cboCodigo.Text <> rsCliFor2.Fields("Código").Value Then
'              If UCase(CGC.Text) = UCase(rsCliFor2("CGC")) Then
'                '27/07/2004 - Daniel
'                'Alteração - Permitir que o usuário continue ou não
'                'com a gravação de um CNPJ/CPF já existente
'                If MsgBox("O Cli/For com Código " & str(rsCliFor2("Código")) & " já possui este CNPJ/CPF, Deseja continuar?", vbQuestion + vbYesNo, "Duplicidade de CNPJ/CPF") = vbYes Then
'                  home_page.SetFocus
'                Else
'                  'Evita a gravação com CPF/CNPJ já existente
'                  CGC.Text = ""
'                  CGC.SetFocus
'                  Exit Sub
'                End If
'
'              End If
'            End If 'If CLng(cboCodigo.Text)...
'          End If
'      End If 'Len(cboCodigo.Text) > 0
    
  'End If                        '18/02/2004 - Daniel
 
  Aux = CGC.Text
'  If IsNull(Aux) Then
'     If Not rsParametros.Fields("NaoPermiteDuplicarCNPJ").Value Then
'        Exit Sub
'     Else
'        Beep
'        DisplayMsg "CNPJ / CPF inválido"
'        CGC.Text = ""
'        CGC.SetFocus
'        Exit Sub
'     End If
'  End If
'
'  i = Len(Aux)
'  If i = 0 Then
'     If Not rsParametros.Fields("NaoPermiteDuplicarCNPJ").Value Then
'        Exit Sub
'     Else
'        Beep
'        DisplayMsg "CNPJ / CPF inválido"
'        CGC.Text = ""
'        CGC.SetFocus
'        Exit Sub
'     End If
'  End If
  
  STR_CGC = Aux
  STR_CGC = LTrim(STR_CGC)
  
  i = True
  F = True
  
  i = Valid_CPF(STR_CGC)
  F = Valid_CGC(STR_CGC)
  
  '24/03/2004 - Daniel
  'Ficou a critério do usuário escolher se irá
  'cadastrar ou não o CPF/CNPJ incorreto
  If i = False And F = False Then
    If rsParametros.Fields("NaoPermiteDuplicarCNPJ").Value Then
      DisplayMsg "CPF/CNPJ incorreto."
      CGC.Text = ""
      CGC.SetFocus
      Exit Sub
    Else
      'DisplayMsg "CPF/CNPJ Incorreto." (Antigo código)
      If MsgBox("O CPF/CNPJ é incorreto, deseja cadastrar assim mesmo?", vbExclamation + vbYesNo) = vbYes Then
        Exit Sub
      End If
    End If
  End If
  
  '11/12/2007 - Celso
  'Conforme parâmetro, não permitir duplicidade de CNPJ / CPF
  
  If rsParametros.Fields("NaoPermiteDuplicarCNPJ").Value Then
     If IsNull(CGC.Text) Then Exit Sub
     If CGC.Text = "" Then Exit Sub
     rsCliFor2.Index = "CGC"
     rsCliFor2.Seek ">", CGC.Text, 0
     If Not rsCliFor2.NoMatch Then
        If cboCodigo.Text <> rsCliFor2.Fields("Código").Value Then
           If UCase(CGC.Text) = UCase(rsCliFor2("CGC")) Then
              '17/05/2023 - Pablo
              'apenas impedito o mesmo CPF/CNPJ se o tipo for igual ao já cadastrado
              Dim pab_Tipo As String
              pab_Tipo = "F"
              If O_Cliente.Value = True Then pab_Tipo = "C"
              If O_Fornecedor.Value = True Then pab_Tipo = "F"
              If O_Revendedor.Value = True Then pab_Tipo = "R"
              If O_Outros.Value = True Then pab_Tipo = "O"

              If pab_Tipo = Trim(CStr(UCase(rsCliFor2("Tipo")))) Then
                'Não aceita CPF/CNPJ já existente com o mesmo tipo
                DisplayMsg "CPF/CNPJ já em uso para o Código: " & str(rsCliFor2("Código")) & vbCrLf & "Você poderá cadastrar este mesmo CPF/CNPJ com outro Tipo"
                CGC.Text = ""
                CGC.SetFocus
                Exit Sub
              End If
           End If
        End If
     End If
  End If

End Sub

Private Sub cboCodigo_Click()
  If Not rsCliFor.EOF Then
    If Len(Trim(cboCodigo.Text)) > 0 Then
      rsCliFor.FindFirst "Código = " & CLng(cboCodigo.Text)
      If Not rsCliFor.NoMatch Then
        Call ShowRecord
        Call ResetAllGrids
      End If
    End If
  End If
End Sub

Private Sub cboCodigo_CloseUp()
  gbComboDropped = False
End Sub

Private Sub cboCodigo_DropDown()
  gbComboDropped = True
End Sub

Private Sub cboCodigo_GotFocus()
  cboCodigo.SelStart = 0
  cboCodigo.SelLength = Len(cboCodigo.Text)
End Sub

Private Sub cboCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  If (Shift And vbAltMask) > 0 Then
    If KeyCode = 40 Or KeyCode = 38 Then
      Call cboCodigo_DropDown
    End If
    Exit Sub
  Else
    If Not gbComboDropped Then
      Call Form_KeyDown(KeyCode, Shift)
      KeyCode = 0
    End If
  End If
End Sub

Private Sub cboCodigo_KeyPress(KeyAscii As Integer)
  If Len(cboCodigo.Text) >= 8 Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If
  
  If Not IsNull(Num_Registro) Then
    Beep
    KeyAscii = 0
    Exit Sub
  End If
  
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub cboCodigo_LostFocus()
  gbComboDropped = False
  cboCodigo.Text = UCase(cboCodigo.Text & "")
  If Not rsCliFor.EOF Then
    If Len(Trim(cboCodigo.Text)) > 0 Then
      If Mid(cboCodigo.Text, Len(cboCodigo.Text), 1) = "*" Then
        rsCliFor.FindFirst "Código >= " & CLng(cboCodigo.Text)
        If Not rsCliFor.NoMatch Then
          Call ShowRecord
          Call ResetAllGrids
        End If
      End If
      rsCliFor.FindFirst "Código = " & CLng(cboCodigo.Text)
      If Not rsCliFor.NoMatch Then
        Call ShowRecord
        Call ResetAllGrids
      End If
    End If
  End If
End Sub

Private Sub cmd_carregaEnderecoCEP_Click()
  On Error GoTo Erro
  
  Dim strCodigoCep As String
  Dim rsCep As Recordset
  Dim arrayCidadeEstado() As String
  Dim strCidade As String
  Dim strEstado As String
  
  If Trim(CEP.Text) = "" Then
      MsgBox "Informe um CEP válido caso deseje carregar o endereço por aqui.", vbInformation, "Atenção"
      CEP.SetFocus
      Exit Sub
  End If
  
  strCodigoCep = Trim(CEP.Text)
  
  strCodigoCep = Replace(strCodigoCep, ".", "")
  strCodigoCep = Replace(strCodigoCep, "-", "")
  strCodigoCep = Replace(strCodigoCep, "/", "")
  strCodigoCep = Replace(strCodigoCep, "\", "")
  strCodigoCep = Replace(strCodigoCep, "_", "")
  strCodigoCep = Replace(strCodigoCep, ";", "")
  strCodigoCep = Replace(strCodigoCep, " ", "")
  strCodigoCep = Replace(strCodigoCep, ",", "")
  strCodigoCep = Replace(strCodigoCep, ":", "")
  
  Set rsCep = dbTemp.OpenRecordset("Select * from Cep where Codigo = " & strCodigoCep, dbReadOnly)
  
  If rsCep.EOF And rsCep.BOF Then
      MsgBox "Endereço não encontrado pelo Código do CEP.", vbInformation, "Atenção"
      Exit Sub
  Else
      Endereço.Text = ""
      Bairro.Text = ""
      Estado.Text = ""
      cbo_cidades.Text = ""
      
      If Not IsNull(rsCep.Fields("Logradouro")) Then
          Endereço.Text = rsCep.Fields("Logradouro")
      End If
      
      If Not IsNull(rsCep.Fields("Bairro")) Then
          Bairro.Text = rsCep.Fields("Bairro")
      End If
      
      If Not IsNull(rsCep.Fields("Cidade_Estado")) Then
          strCidade = rsCep.Fields("Cidade_Estado")
          
          arrayCidadeEstado() = Split(strCidade, "/")
          strEstado = arrayCidadeEstado(1)
          strCidade = arrayCidadeEstado(0)
          Estado.Text = strEstado
          Estado_LostFocus
          cbo_cidades.Text = strCidade
      End If
      
  End If
  txtPais.Text = "Brasil"
  
  rsCep.Close
  Set rsCep = Nothing
  
  Exit Sub
Erro:
  MsgBox "Erro na função achar endereço pelo CEP informado " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

'21/12/2006 - Anderson - Alterações na tela de contatos efetuados
Private Sub cmdContatoEfetuadoNovo_Click()

  'Limpa a tela e atualiza as variáveis
  lngContatosEfetuadosSequencia = vbEmpty
  strContatosEfetuadosData = vbEmpty
  txtContatosEfetuadosData.Text = "  /  /    "
  txtContatosEfetuadosLembrarEm.Text = "  /  /    "
  txtContatosEfetuadosDescricao.Text = ""
  chkContatoEfetuadoPendencia.Value = False
  
  grdContatosEfetuados.SelectByCell = False
  grdContatosEfetuados.SelectTypeRow = ssSelectionTypeNone
  grdContatosEfetuados.SelectTypeRow = ssSelectionTypeSingleSelect
  grdContatosEfetuados.MoveLast
  
  txtContatosEfetuadosData.SetFocus

End Sub

'21/12/2006 - Anderson - Alteração na tela de contatos efetuados
Private Sub cmdContatosEfetuadosSalvar_Click()

  Dim bolLembrarEm As Boolean '19/07/2007 - Anderson - Informa se o campo Lembrar Em deve ser obrigatório.
  
  'Verifica se a data foi preenchida corretamente
  If txtContatosEfetuadosData.Text = "  /  /    " Then
    MsgBox "Não é possível salvar o registro. Verifique se o campo data foi preenchido corretamente.", vbOKOnly + vbInformation, "Atenção"
    Exit Sub
  End If
  
  bolLembrarEm = gbGetUserFunction(gnUserCode, "ContatosEfetuadosLembrarEm")

  '13/07/2007 - Anderson
  'Verifica se o cadastro está com data retroativa
  'Solicitado por Supriprint
  If bolCaseSupriPrint Then
  
    '20/07/2007 - Anderson
    'A verificação deverá ser realizada somente para quem não é superusuário e quando for modo de edição .quando
    If (lngContatosEfetuadosSequencia > 0 And Not gbSuperUser) Or lngContatosEfetuadosSequencia = 0 Then

      If txtContatosEfetuadosData.Text < Date Then
        MsgBox "Não é possível salvar contatos efetuados com data retroativa ou em branco. Favor verificar se a data está correta.", vbOKOnly + vbInformation, "Atenção"
        Exit Sub
      End If
  
      If Weekday(txtContatosEfetuadosData.Text, vbMonday) >= 6 Then
        MsgBox "Não é possível salvar contatos efetuados aos sábados e domingos. Favor verificar se a data está correta.", vbOKOnly + vbInformation, "Atenção"
        Exit Sub
      End If
      
    End If
        
  End If
  
  'Verifica se o campo descrição foi preenchido corretamente
  If Trim(txtContatosEfetuadosDescricao.Text) = "" Then
    MsgBox "Não é possível salvar o registro. Verifique se o campo descrição foi preenchido corretamente.", vbOKOnly + vbInformation, "Atenção"
    Exit Sub
  End If
  
  
  '19/07/2007 - Anderson
  'Verifica se o cliente é Supri Print para desbloquear o campo Lembrar Em quando o cliente é inativo
  If bolCaseSupriPrint Then
    '19/07/2007 - Anderson
    'Alterações para evitar a digitação da data Lembrar em retroativa
    'Solicitado por Supri Print
    If bolLembrarEm And Not rsCliFor("Inativo") Then
      If txtContatosEfetuadosLembrarEm.Text = "  /  /    " Then
        MsgBox "Não é possível salvar o registro. O campo Lembrar Em é obrigatório. Verifique se o campo Lembrar Em foi preenchido corretamente.", vbOKOnly + vbInformation, "Atenção"
        Exit Sub
      End If
      
      If CDate(txtContatosEfetuadosLembrarEm.Text) < Date Then
        MsgBox "Não é possível salvar contatos efetuados com data Lembrar Em retroativa ou em branco. Favor verificar se a data Lembrar Em está correta.", vbOKOnly + vbInformation, "Atenção"
        Exit Sub
      End If
      
      If Weekday(txtContatosEfetuadosLembrarEm.Text, vbMonday) >= 6 Then
        MsgBox "Não é possível salvar contatos efetuados com data Lembrar Em aos sábados e domingos. Favor verificar se a data Lembrar Em está correta.", vbOKOnly + vbInformation, "Atenção"
        Exit Sub
      End If
    End If
  Else
    '19/07/2007 - Anderson
    'Alterações para evitar a digitação da data Lembrar em retroativa
    'Solicitado por Supri Print
    If bolLembrarEm Then
      If txtContatosEfetuadosLembrarEm.Text = "  /  /    " Then
        MsgBox "Não é possível salvar o registro. O campo Lembrar Em é obrigatório. Verifique se o campo Lembrar Em foi preenchido corretamente.", vbOKOnly + vbInformation, "Atenção"
        Exit Sub
      End If
    End If
  End If
  
  '19/07/2007 - Anderson
  'Verifica se a data Lembrar em está digitada e coloca o check da pendencia
  If bolCaseSupriPrint Then txtContatosEfetuadosLembrarEm_LostFocus

  
  '04/05/2007 - Anderson
  'Utilizado para evitar o erro de perder o código do cliente quando clica na aba Contatos efetuados.
  If gsCodCliente = "" And cboCodigo.Text = "" Then
      MsgBox "Limpe a tela ou feche e abra-a novamente. Então selecione o cliente desejado", vbInformation, "Atenção"
      Exit Sub
  End If
  
  If Not lngContatosEfetuadosSequencia > 0 Then
    'Adiciona um novo contato
    With rsContatos_Efetuados
      .AddNew
      lngContatosEfetuadosSequencia = .Fields("Seqüência")
      '.Fields("Cliente") = CInt(cboCodigo.Text)
      .Fields("Cliente") = CLng(Trim(cboCodigo.Text))
      If txtContatosEfetuadosData.Text <> "  /  /    " Then .Fields("Data") = txtContatosEfetuadosData.Text
      If txtContatosEfetuadosLembrarEm.Text <> "  /  /    " Then .Fields("Data Aviso") = txtContatosEfetuadosLembrarEm.Text
      .Fields("Descrição") = txtContatosEfetuadosDescricao.Text
      .Fields("Pendência") = -chkContatoEfetuadoPendencia.Value
      .Update
    End With
    
  Else
    'Altera um contato já existente
    With rsContatos_Efetuados
      .Index = "Cliente"
      .Seek "=", cboCodigo.Text, strContatosEfetuadosData, lngContatosEfetuadosSequencia
      If Not .NoMatch Then
        With rsContatos_Efetuados
          .Edit
          .Fields("Data") = IIf(IsDate(txtContatosEfetuadosData.Text), txtContatosEfetuadosData.Text, Null)
          .Fields("Data Aviso") = IIf(IsDate(txtContatosEfetuadosLembrarEm.Text), txtContatosEfetuadosLembrarEm.Text, Null)
          .Fields("Descrição") = txtContatosEfetuadosDescricao.Text
          .Fields("Pendência") = -chkContatoEfetuadoPendencia.Value
          .Update
        End With
      End If
    End With
  End If
  
  'Atualiza as pendências - Case Supri Print
  If bolCaseSupriPrint Then
    '21/09/2007 - Anderson
    'Correção para Supri-Print
    'O sistema estava colocando os contatos efetuados retroativos como pendente
    If IsDate(txtContatosEfetuadosData.Text) Then
      If txtContatosEfetuadosData.Text >= Date Then
        db.Execute "UPDATE [Contatos Efetuados] " & _
                   "Set Pendência = 0, [Data Aviso]='" & Null & "' " & _
                   "Where Cliente = " & cboCodigo.Text & " AND Seqüência <> " & lngContatosEfetuadosSequencia
      End If
    End If
  End If

  Call cmdContatoEfetuadoNovo_Click
  Call LoadGridContatosEfetuados

End Sub

Private Sub cmdDocLimCred_Click()
  'Call ImprimirLimiteCreditoExcel
  Call ImprimirLimiteCreditoRpt
End Sub

Private Sub ImprimirLimiteCreditoRpt()
  Dim strSQL As String
  
  Dim C_Filial As String
  C_Filial = CStr(rsParametros.Fields("Filial"))
  
  Dim C_Codigo As String
  C_Codigo = cboCodigo.Text
  
  Dim C_Disponivel As String
  C_Disponivel = mskEdit_limiteCreditoDisponivel.Text
  
  'dados da instituição e do cliente
  strSQL = "SELECT Con_LimiteCredito.F_Filial, Con_LimiteCredito.F_Nome, Con_LimiteCredito.F_RazaoSocial, Con_LimiteCredito.F_CNPJ, Con_LimiteCredito.F_IE, Con_LimiteCredito.F_Logradouro, Con_LimiteCredito.F_Numero, Con_LimiteCredito.F_Complemento, "
  strSQL = strSQL & "Con_LimiteCredito.F_Cidade, Con_LimiteCredito.F_UF, Con_LimiteCredito.F_CEP, Con_LimiteCredito.F_Fone, Con_LimiteCredito.F_Email, Con_LimiteCredito.C_Codigo, Con_LimiteCredito.C_Nome, Con_LimiteCredito.C_CNPJ_CPF, Con_LimiteCredito.C_IE_RG, "
  strSQL = strSQL & "Con_LimiteCredito.C_Logradouro, Con_LimiteCredito.C_Numero, Con_LimiteCredito.C_Complemento, Con_LimiteCredito.C_Bairro, Con_LimiteCredito.C_Cidade, Con_LimiteCredito.C_UF, Con_LimiteCredito.C_CEP, Con_LimiteCredito.C_Fone1, Con_LimiteCredito.C_Fone2, "
  strSQL = strSQL & "Con_LimiteCredito.C_Email, Con_LimiteCredito.C_Trabalho, Con_LimiteCredito.C_TrabalhoFone, Con_LimiteCredito.C_Cargo, Con_LimiteCredito.C_Salario, '" & C_Disponivel & "' AS C_Disponivel "
  strSQL = strSQL & "FROM Con_LimiteCredito "
  strSQL = strSQL & "WHERE (((Con_LimiteCredito.F_Filial)=" & C_Filial & ") AND ((Con_LimiteCredito.C_Codigo)=" & C_Codigo & "));"
  ConsultaRpt strSQL
  
  With rsRdp
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para a Filial: " & C_Filial, vbExclamation, "Atenção"
      .Close
      Exit Sub
    Else
      Set rptLimiteCredito.DataSource = Nothing
      rptLimiteCredito.DataMember = ""
      Set rptLimiteCredito.DataSource = .DataSource
      Dim sTmp As String
      Dim obj As Variant
      
      'Título
      obj = .Fields("F_Nome")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = obj & "  -  LIMITE DE CRÉDITO"
      rptLimiteCredito.Sections("ReportHeader").Controls("lblTitulo").Caption = sTmp
      'Sub-Título
      obj = .Fields("F_Logradouro")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = obj
      obj = .Fields("F_Numero")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = sTmp & IIf(obj = "", "", ", " & obj)
      obj = .Fields("F_Fone")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = sTmp & IIf(obj = "", "", " - FONE: " & obj)
      obj = .Fields("F_Email")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = sTmp & IIf(obj = "", "", " - EMAIL:  " & obj)
      obj = .Fields("F_Cidade")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = sTmp & IIf(obj = "", "", " - " & obj)
      obj = .Fields("F_UF")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = sTmp & IIf(obj = "", "", "/" & obj)
      obj = .Fields("F_CEP")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = sTmp & IIf(obj = "", "", " - CEP: " & obj)
      sTmp = sTmp & Chr(10)
      obj = .Fields("F_CNPJ")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = sTmp & IIf(obj = "", "", "CNPJ: " & obj)
      obj = .Fields("F_IE")
      obj = CStr(IIf(IsNull(obj), "", obj))
      sTmp = sTmp & IIf(obj = "", "", " - Inscrição Estadual: " & obj)
      rptLimiteCredito.Sections("ReportHeader").Controls("lblSubtitulo").Caption = sTmp
      'Formulário Cliente
      rptLimiteCredito.Sections("Detail").Controls("txtCod").DataMember = ""
      rptLimiteCredito.Sections("Detail").Controls("txtCod").DataField = .Fields("C_Codigo").Name
      rptLimiteCredito.Sections("Detail").Controls("txtNome").DataMember = ""
      rptLimiteCredito.Sections("Detail").Controls("txtNome").DataField = .Fields("C_Nome").Name
      rptLimiteCredito.Sections("Detail").Controls("txtCPF").DataMember = ""
      rptLimiteCredito.Sections("Detail").Controls("txtCPF").DataField = .Fields("C_CNPJ_CPF").Name
      rptLimiteCredito.Sections("Detail").Controls("txtRG").DataMember = ""
      rptLimiteCredito.Sections("Detail").Controls("txtRG").DataField = .Fields("C_IE_RG").Name
      
      rptLimiteCredito.Show
    End If
  End With
End Sub

Private Sub ImprimirLimiteCreditoExcel()
  Dim aplic As New Excel.Application
  Dim xls As Excel.Workbook
  Set xls = aplic.Workbooks.Open(App.Path & "\LimiteCredito.xls")
  xls.Sheets(2).Select
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  
  Dim Filial As Integer
  Filial = rsParametros.Fields("Filial")
  
  'dados da instituição
  strSQL = "SELECT [Parâmetros Filial].[CGC], [Parâmetros Filial].[Razão Social], [Parâmetros Filial].[Nome], [Parâmetros Filial].[Inscrição], [Parâmetros Filial].[Endereço], [Parâmetros Filial].[EnderecoNumero], [Parâmetros Filial].[EnderecoComplemento], [Parâmetros Filial].[Cidade], [Parâmetros Filial].[Estado], [Parâmetros Filial].[CEP], [Parâmetros Filial].[Fone], [Email].[EmailRemetente] "
  strSQL = strSQL & "FROM [Parâmetros Filial] LEFT JOIN [Email] ON [Parâmetros Filial].[Filial] = [Email].[Filial] "
  strSQL = strSQL & "WHERE [Parâmetros Filial].[Filial] = " & Filial
  Set rsTemp = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsTemp
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para a Filial: " & Filial, vbExclamation, "Atenção"
      .Close
      Exit Sub
    Else
      xls.Application.Cells(1, 2).Value = .Fields("Nome").Value
      xls.Application.Cells(2, 2).Value = .Fields("Razão Social").Value
      xls.Application.Cells(3, 2).Value = .Fields("CGC").Value
      xls.Application.Cells(4, 2).Value = .Fields("Inscrição").Value
      xls.Application.Cells(5, 2).Value = .Fields("Endereço").Value
      xls.Application.Cells(6, 2).Value = .Fields("EnderecoNumero").Value
      xls.Application.Cells(7, 2).Value = .Fields("EnderecoComplemento").Value
      xls.Application.Cells(8, 2).Value = .Fields("Cidade").Value
      xls.Application.Cells(9, 2).Value = .Fields("Estado").Value
      xls.Application.Cells(10, 2).Value = .Fields("CEP").Value
      xls.Application.Cells(11, 2).Value = .Fields("Fone").Value
      xls.Application.Cells(12, 2).Value = .Fields("EmailRemetente").Value
    End If
  End With
  rsTemp.Close
  Set rsTemp = Nothing
  
  'dados do cliente
  strSQL = "SELECT Cli_For.Código, Cli_For.Nome, Cli_For.CGC, Cli_For.Inscrição, Cli_For.Endereço, Cli_For.[Endereço Número], Cli_For.Complemento, Cli_For.Bairro, Cli_For.Cidade, Cli_For.Estado, Cli_For.CEP, Cli_For.[Fone 1], Cli_For.[Fone 2], [Cli_For - Crédito].Empresa, [Cli_For - Crédito].Telefone, [Cli_For - Crédito].Cargo, Cli_For.email, [Cli_For - Crédito].Salário "
  strSQL = strSQL & "FROM Cli_For INNER JOIN [Cli_For - Crédito] ON Cli_For.Código = [Cli_For - Crédito].Código "
  strSQL = strSQL & "WHERE Cli_For.Código = " & cboCodigo.Text & ";"
  Set rsTemp = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsTemp
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para o Cliente", vbExclamation, "Atenção"
      .Close
      Exit Sub
    Else
      xls.Application.Cells(14, 2).Value = .Fields("Código").Value
      xls.Application.Cells(15, 2).Value = .Fields("Nome").Value
      xls.Application.Cells(16, 2).Value = .Fields("CGC").Value
      xls.Application.Cells(17, 2).Value = .Fields("Inscrição").Value
      xls.Application.Cells(18, 2).Value = .Fields("Endereço").Value
      xls.Application.Cells(19, 2).Value = .Fields("Endereço Número").Value
      xls.Application.Cells(20, 2).Value = .Fields("Complemento").Value
      xls.Application.Cells(21, 2).Value = .Fields("Bairro").Value
      xls.Application.Cells(22, 2).Value = .Fields("Cidade").Value
      xls.Application.Cells(23, 2).Value = .Fields("Estado").Value
      xls.Application.Cells(24, 2).Value = .Fields("CEP").Value
      xls.Application.Cells(25, 2).Value = .Fields("Fone 1").Value
      xls.Application.Cells(26, 2).Value = .Fields("Fone 2").Value
      xls.Application.Cells(27, 2).Value = .Fields("Empresa").Value
      xls.Application.Cells(28, 2).Value = .Fields("Telefone").Value
      xls.Application.Cells(29, 2).Value = .Fields("Cargo").Value
      xls.Application.Cells(30, 2).Value = .Fields("email").Value
      xls.Application.Cells(31, 2).Value = .Fields("Salário").Value
    End If
  End With
  rsTemp.Close
  
  xls.Application.Cells(32, 2).Value = mskEdit_limiteCreditoDisponivel.Text
  
  xls.Sheets(1).Select
  xls.Sheets(1).PrintOut
  MsgBox "Aguarde a impressão e depois clique em OK", vbOKOnly, "Atenção"
  
  xls.Close False
  Set xls = Nothing
  Set aplic = Nothing
  Set rsTemp = Nothing
End Sub

Private Sub Combo_Cobrança_CloseUp()
  Combo_Cobrança.Text = Combo_Cobrança.Columns(2).Text
  Combo_Cobrança_LostFocus
End Sub

Private Sub Combo_Cobrança_LostFocus()

 Nome_Cobrança.Caption = ""
  
 If IsNull(Combo_Cobrança.Text) Then Exit Sub
 If Combo_Cobrança.Text = "" Then Exit Sub
 If Not IsNumeric(Combo_Cobrança.Text) Then Exit Sub
 If Val(Combo_Cobrança.Text) < 1 Then Exit Sub
 '28/11/2006 - Anderson
 'Alteração do número de contas bancárias de 99 para 255
 'Solicitado por: 2227883 - SANTA FÉ DO ARAGUAIA PREFEITURA MUNICIPAL
 If Val(Combo_Cobrança.Text) > 255 Then Exit Sub

 rsContas.Index = "Código"
 rsContas.Seek "=", Val(Combo_Cobrança.Text)
 If rsContas.NoMatch Then Exit Sub
 Nome_Cobrança.Caption = rsContas("Descrição") & ""
 
End Sub

Private Sub Combo_Transp_CloseUp()
  
  Combo_Transp.Text = Combo_Transp.Columns(1).Text
  Combo_Transp_LostFocus

End Sub

Private Sub Combo_Transp_LostFocus()
  Nome_Transp.Caption = ""
  If IsNull(Combo_Transp.Text) Then Exit Sub
  If Combo_Transp.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Transp.Text) Then Exit Sub
  If Val(Combo_Transp.Text) < 1 Then Exit Sub
  If Val(Combo_Transp.Text) > 9999 Then Exit Sub
  
  rsTransportadoras.Index = "Código"
  rsTransportadoras.Seek "=", Val(Combo_Transp.Text)
  If rsTransportadoras.NoMatch Then Exit Sub
  Nome_Transp.Caption = rsTransportadoras("Nome") & ""

End Sub

Private Sub Combo_Vendedor_CloseUp()

    

  Combo_Vendedor.Text = Combo_Vendedor.Columns(0).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_LostFocus()
  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Combo_Vendedor.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) < 1 Then Exit Sub
  If Val(Combo_Vendedor.Text) > 9999 Then Exit Sub
  
  rsVendedores.Index = "Código"
  rsVendedores.Seek "=", Val(Combo_Vendedor.Text)
  If rsVendedores.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsVendedores("Nome")
  
End Sub

Private Sub ddwGruposMalaDireta_CloseUp()
  grdGruposMalaDireta.Columns("Grupo").Text = ddwGruposMalaDireta.Columns("Código").Text
  grdGruposMalaDireta.Columns("Nome").Text = ddwGruposMalaDireta.Columns("Nome").Text
End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Estado_GotFocus()
  sEstadoInicioFoco = Estado.Text
End Sub

Private Sub Estado_LostFocus()
  Dim sAUX As String
  
  Call StatusMsg("")
  
  sAUX = UCase(Estado.Text)
  Estado.Text = sAUX
  If IsNull(sAUX) Then
    Exit Sub
  ElseIf sAUX = "" Then
    Exit Sub
  Else
    Select Case sAUX
      '03/09/2003 - Maikel
      '             Adicionado o tipo EX para exportação.
      Case "AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", _
           "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", _
           "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO", "EX"
         Estado.Text = sAUX
      Case Else
        cbo_cidades.Clear
        DisplayMsg "Estado incorreto."
        Exit Sub
    End Select
  End If
  
  If sEstadoInicioFoco <> Estado.Text Then
    CarregarCidadesDoEstadoSelecionado Estado.Text
    sEstadoInicioFoco = ""
  End If
  
End Sub

Private Sub CarregarCidadesDoEstadoSelecionado(sEstado As String)
    Dim rsCidadesEstado As Recordset
  
    cbo_cidades.Clear
    Set rsCidadesEstado = db.OpenRecordset("Select * from TerritorioMunicipio where uf='" & sEstado & "' ", dbOpenDynaset)

    While Not rsCidadesEstado.EOF
        cbo_cidades.AddItem rsCidadesEstado.Fields("Nome").Value
        rsCidadesEstado.MoveNext
    Wend

    rsCidadesEstado.Close
    Set rsCidadesEstado = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If gbComboDropped Then
    Call cboCodigo_KeyDown(KeyCode, Shift)
  Else
    If KeyCode = vbKeyF5 Then
      Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY Código", dbOpenDynaset)
      Call GetNewCode2(Me, cboCodigo, rsCliFor, 99999999)
      '04/05/2007 - Anderson
      'Correção do erro ao iniciar um novo registro através da tecla F5. O problema ocorria porque ao registrar um contato efetuado, o sistema estava perdendo o código novo.
      'cboCodigo.Text = Format(cboCodigo.Text, String(8, "#") & "0")
      gsCodCliente = Format(cboCodigo.Text, String(8, "#") & "0")
      cboCodigo.Text = gsCodCliente
    Else
      Call HandleKeyDown(KeyCode, Shift)
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
On Error GoTo Erro

  Dim Resp As Integer
  Dim sResult As String
  
  ' 11/12/2007 - Celso
  ' Carga Parâmetros
  Set rsParametros = db.OpenRecordset("Parâmetros Filial")
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then
     MsgBox "Filial não encontrada", vbCritical, "Erro"
     Exit Sub
  End If
  
  On Error Resume Next
  Picture1.Picture = LoadPicture(App.Path & "\Imagens\logotipo.bmp")
  
  gbComboDropped = False
  Call CenterForm(Me)
  
  ActiveBar1.Tools("miOpOrdem").CBList.Clear
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 0, "Por Código"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 1, "Por Nome/Razão Social"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 2, "Por Fantasia"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 3, "Por Telefone"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 4, "Por Cidade e Código"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 5, "Por CNPJ"
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  
  DoEvents

  Num_Registro = Null
  Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY Código", dbOpenDynaset)
  Set rsMala_Tempo = db.OpenRecordset("Mala Direta - Tempo")
  Set rsCli_Cred = db.OpenRecordset("Cli_For - Crédito")
  Set rsVendedores = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsMala = db.OpenRecordset("Mala Direta")
  Set rsTransportadoras = db.OpenRecordset("Transportadoras", , dbReadOnly)
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsCliFor2 = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsGrupos = db.OpenRecordset("SELECT * FROM [Grupos Interesse] ORDER BY Código", dbOpenDynaset)
  
  '21/12/2006 - Anderson - Alterações para a tela de contatos efetuados
  Set rsContatos_Efetuados = db.OpenRecordset("Contatos Efetuados")
  lngContatosEfetuadosSequencia = vbEmpty
  strContatosEfetuadosData = vbEmpty
  bolCalendario = False

  Data1.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  Data6.DatabaseName = gsQuickDBFileName
  Data7.DatabaseName = gsQuickDBFileName
  datCliForCod.DatabaseName = gsQuickDBFileName
  datCliForNome.DatabaseName = gsQuickDBFileName
  datFantasia.DatabaseName = gsQuickDBFileName
  
  Tab1.TabEnabled(1) = False
  Tab1.TabEnabled(2) = False
  Tab1.TabEnabled(3) = False
  Tab1.TabEnabled(4) = False
  Tab1.TabEnabled(5) = False
  Tab1.TabEnabled(6) = False
  '09/07/2004 - Daniel
  'Adição da Tab de Classificação de Clientes
  'Case: TV Shopping
  Tab1.TabEnabled(7) = False
  
  '31/07/2002 - mpdea
  'Loja Virtual (Web)
  Tab1.TabVisible(6) = gblnWorkWeb
  
  '09/07/2004 - Daniel
  'Adição da Tab de Classificação de Clientes
  'Case: TV Shopping
  Tab1.TabVisible(7) = CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959")
  
  '09/08/2005 - Daniel
  'Inclusão da Function blnHabilitarCampoVendedor que fará validação em cima
  'campo [Parâmetros Filial].AlterVendedorCliFor
  '
  'Finalidade : Apenas o Superusuário poderá alterar o campo
  '             Vendedor no cadastro Cli / For
  '
  'Solicitante: Konrad Comercial Ltda
  Combo_Vendedor.Enabled = blnHabilitarCampoVendedor
  
  sResult = GetSetting("QuickStore", "Clientes", "Mantem Cidade", False)
  ActiveBar1.Tools("miOpFreezeCidade").Checked = CBool(sResult)
  
  sResult = GetSetting("QuickStore", "Clientes", "Mantem Estado", False)
  ActiveBar1.Tools("miOpFreezeEstado").Checked = CBool(sResult)
  
  sResult = GetSetting("QuickStore", "Clientes", "Mantem Boleto", False)
  ActiveBar1.Tools("miOpFreezeMsg").Checked = CBool(sResult)
  
  Set Data4.Recordset = db.OpenRecordset("SELECT DISTINCTROW [Grupos Interesse].Nome, [Grupos Interesse].Código From [Grupos Interesse] ORDER BY [Grupos Interesse].Nome;", dbOpenDynaset)
  Data4.Refresh
  
  Rem Desabilita os OCX que estão em outros TABS
  Rem  feito para consertar o cursor que some
  CEP_Cob.TabStop = False 'tab3
  
  Call ActiveBarLoadToolTips(Me)


  '23/01/2003 - mpdea
  'Quick em modo limitado
  If Not gblnQuickFull Then
    'Transportadora
    Label37.Visible = False
    Combo_Transp.Visible = False
    Nome_Transp.Visible = False
    Frame4.Visible = False
    'Cobrança
    Label43.Visible = False
    Combo_Cobrança.Visible = False
    Nome_Cobrança.Visible = False
  End If

  '17/05/2004 - Daniel
  'Este flag setará se a empresa tem necessidade de cadastrar
  'o mesmo CNPJ/CPF para n clientes
  'Cases: Embalavi => "QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162"
  m_blnNcnpjcpf = CheckSerialCaseMod("QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162")

  '09/11/2004 - Daniel
  'Case: Cliente Teknika
  m_blnTeknika = CheckSerialCaseMod("QS40966-243")

  '26/07/2004 - Daniel
  'Campo Agência de Publicidade, será mostrado apenas quando for Fornecedor e cliente STC
  'Inicia como invisível
  chkAgencia.Visible = False
  
  '--------------------------------------------------------------------------
  '04/08/2006 - Andrea
  'Este flag indicará se o numero de série é do cliente Idéia Intima
  m_blnIdeiaIntima = CheckSerialCaseMod("QS71251-786", "QS71256-206", "QS71257-290")
  If m_blnIdeiaIntima Then
    O_Física.Value = True
    O_Jurídica.Value = False
  End If
  '--------------------------------------------------------------------------
  
  '13/07/2007 - Anderson
  'Verfica as permissões do usuário na tela de cadastro de clientes/fornecedores
  'Solicitado por: Supri Print
  If Not gbSuperUser Then
    m_strContatosEfetuados = gbGetUserPermition(gnUserCode, 176)
    Tab1.TabVisible(1) = (Mid(m_strContatosEfetuados, 1, 1) = "1")
    cmdContatoEfetuadoNovo.Enabled = (Mid(m_strContatosEfetuados, 2, 1) = "1")
    cmdContatosEfetuadosSalvar.Enabled = (Mid(m_strContatosEfetuados, 2, 1) = "1")
  End If
  
  '18/07/2007 - Anderson
  cboSadigWebTipo.Text = "(Nenhum)"
  
  '19/07/2007 - Anderson
  'Verifica se o Case é da Supri Print
  bolCaseSupriPrint = CheckSerialCaseMod("QS38926-932")

  Me.Show
  DoEvents
  
  Call ClearScreen
  
  Exit Sub
Erro:
  MsgBox "Erro na carga da tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  gbComboDropped = (cboCodigo.DroppedDown = True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call StatusMsg("")
  
  On Error Resume Next
  
  rsCliFor.Close
  rsMala_Tempo.Close
  rsCli_Cred.Close
  rsVendedores.Close
  rsMala.Close
  rsTransportadoras.Close
  rsContas.Close
  rsCliFor2.Close
  rsGrupos.Close
  
  '21/12/2006 - Anderson
  rsContatos_Efetuados.Close
  
  Set rsCliFor = Nothing
  Set rsMala_Tempo = Nothing
  Set rsCli_Cred = Nothing
  Set rsVendedores = Nothing
  Set rsMala = Nothing
  Set rsTransportadoras = Nothing
  Set rsContas = Nothing
  Set rsCliFor2 = Nothing
  Set rsGrupos = Nothing
  '21/12/2006 - Anderson
  Set rsContatos_Efetuados = Nothing
  
  EncerraRpt
End Sub

Sub ShowRecord()

  On Error GoTo ErrHandler

  Call StatusMsg("")
  
  '21/12/2006 - Anderson
  Call cmdContatoEfetuadoNovo_Click

  gsCodCliente = rsCliFor("Código")
  cboCodigo.Text = gsCodCliente
  cboFantasia.Text = rsCliFor("Fantasia") & ""
  
  Select Case UCase(rsCliFor("Tipo").Value)
    Case "C"
      O_Cliente.Value = True
    Case "F"
      O_Fornecedor.Value = True
    Case "R"
      O_Revendedor.Value = True
    Case "O"
      O_Outros.Value = True
  End Select
  
  Select Case rsCliFor("IndicadorIE").Value
    Case "1"
      cboIndicadorIE.Text = "1=Contribuinte ICMS (informar a IE do destinatário)"
    Case "2"
      cboIndicadorIE.Text = "2=Contribuinte isento de Inscrição no cadastro de Contribuintes do ICMS"
    Case "9"
      cboIndicadorIE.Text = "9=Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS"
    Case Else
      cboIndicadorIE.Text = "9=Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS"
  End Select
  
  '27/05/2004 - mpdea
  'Web
  lblClienteWEB.Visible = rsCliFor("Web").Value
  Tab1.TabEnabled(6) = rsCliFor("Web").Value
  
  cboNome.Text = rsCliFor("Nome")
  
  'Select Case Tab1.Tab
  
    'Case 0
  
      If rsCliFor("Física_Jurídica") = "F" Then
         O_Física.Value = True
         O_Física_Click
      Else
         O_Jurídica.Value = True
         O_Jurídica_Click
      End If
        
      Endereço.Text = rsCliFor("Endereço") & ""
      '26/07/2004 - Daniel
      'Adicionado o campo AgenciaPublicidade
      'Case: STC
      If CheckSerialCaseMod("QS39823-684") Then
        If O_Fornecedor.Value Then
          If rsCliFor("AgenciaPublicidade").Value Then chkAgencia.Value = vbChecked
          Else: chkAgencia.Value = vbUnchecked
        End If
      End If
      '----------------------------------------------------------------------------
      Complemento.Text = rsCliFor("Complemento") & ""
      Bairro.Text = rsCliFor("Bairro") & ""
      'Cidade.Text = rsCliFor("Cidade") & ""
      
      Estado.Text = rsCliFor("Estado") & ""
      
      If Estado.Text <> "" Then
        CarregarCidadesDoEstadoSelecionado Estado.Text
      End If
      
      cbo_cidades.Text = rsCliFor("Cidade") & ""
      CEP.Text = rsCliFor("CEP") & ""
      Fone1.Text = rsCliFor("Fone 1") & ""
      Fone2.Text = rsCliFor("Fone 2") & ""
      Fax.Text = rsCliFor("Fax") & ""
      CGC.Text = rsCliFor("CGC") & ""
      Insc.Text = rsCliFor("Inscrição") & ""
      
      '27/08/2009 - mpdea
      'Novos campos para NFe
      txtEnderecoNumero.Text = rsCliFor.Fields("Endereço Número").Value & ""
      txtPais.Text = rsCliFor.Fields("Pais").Value & ""
      txtInscricaoMunicipal.Text = rsCliFor.Fields("InscricaoMunicipal").Value & ""
      txtInscricaoSuframa.Text = rsCliFor.Fields("InscricaoSuframa").Value & ""
      txtCNAE.Text = rsCliFor.Fields("CNAE").Value & ""
      
      mskDatNasc.Mask = ""
      mskDatNasc.Text = rsCliFor("DataNascimento") & ""
      mskDatNasc.Mask = "##/##/####"
      
      mskDatAbertura.Mask = ""
      mskDatAbertura.Text = rsCliFor("datAberturaCadastro") & ""
      mskDatAbertura.Mask = "##/##/####"
      
      
      If rsCliFor("Sexo") = "M" Then
         optMasculino.Value = True
      ElseIf rsCliFor("Sexo") = "F" Then
         optFeminino.Value = True
      ElseIf rsCliFor("Sexo") = "" Or IsNull(rsCliFor("Sexo")) Then
         optMasculino.Value = False
         optFeminino.Value = False
      End If
      
      If rsCliFor("Inativo") = False Then Inativo.Value = 0
      If rsCliFor("Inativo") = True Then Inativo.Value = 1
    
      If rsCliFor("Bloqueado") = False Then Bloqueado.Value = 0
      If rsCliFor("Bloqueado") = True Then Bloqueado.Value = 1
    
      If rsCliFor("Faturado") = False Then Faturado.Value = 0
      If rsCliFor("Faturado") = True Then Faturado.Value = 1
      
      If rsCliFor("Tem Conta") = False Then Conta.Value = 0
      If rsCliFor("Tem Conta") = True Then Conta.Value = 1
      
      Desconto.Text = Format(rsCliFor("Desconto"), "00.00")
      Última_Compra.Caption = rsCliFor("Última Compra") & ""
      If IsDate(rsCliFor("Última Compra")) Then
        Última_Compra.Caption = Format(rsCliFor("Última Compra"), "dd/mm/yyyy")
      End If
      
      Comentários.Text = rsCliFor("Comentários") & ""
      
      Combo_Transp.Text = rsCliFor("Transportadora") & ""
      Combo_Transp_LostFocus
      
      If IsNull(rsCliFor("CodTipoFrete")) Then
        O_Frete_Pagar.Value = "1"
        O_Frete_Pago.Value = "2"
      Else
        O_Frete_Pagar.Value = rsCliFor("CodTipoFrete") = "1"
        O_Frete_Pago.Value = rsCliFor("CodTipoFrete") = "2"
      End If
      
      Contabilidade.Text = rsCliFor("Contabilidade")
      '05/05/2004 - Daniel
      'Adicionado campo IsentoIPI
      If rsCliFor("IsentoIPI").Value = True Then
        chkIsencaoIPI.Value = vbChecked
      Else
        chkIsencaoIPI.Value = vbUnchecked
      End If
      '06/05/2004 - Daniel
      'Adicionado campo ObsIsentoIPI
      txtMensagemIPI = rsCliFor("ObsIsentoIPI").Value & ""
      
      Combo_Vendedor.Text = rsCliFor("Vendedor") & ""
      cboTabelaPrecos.Text = rsCliFor("TabelaPrecoPadrao") & ""
      Combo_Vendedor_LostFocus
      
      '18/01/2007 - Anderson
      'Solicitação senha do gerente ao alterar o vendedor relacionado ao cliente
      strVendedor = rsCliFor("Vendedor") & ""
        
      email.Text = rsCliFor("Email") & ""
      home_page.Text = rsCliFor("Home Page") & ""
      
      Call LoadGridContatos
      
'    Case 1
      Call LoadGridContatosEfetuados
'
    'Case 2
      '---[ 02/06/2003 - Maikel ]---'
      ' Case      : Projeto Deluken
      ' Descrição : Adicionado o campo abaixo para guardar o dia base de acerto de consignações
        txtDiaBaseConsignacao.Text = rsCliFor.Fields("DiaBaseConsignacao") & ""
      '--------------------------------------------------------
      
      Combo_Cobrança.Text = rsCliFor("Conta Cobrança")
      Combo_Cobrança_LostFocus
        
      Endereço_Cob.Text = rsCliFor("Endereço Cob") & ""
      Complemento_Cob.Text = rsCliFor("Complemento Cob") & ""
      Bairro_Cob.Text = rsCliFor("Bairro Cob") & ""
      Cidade_Cob.Text = rsCliFor("Cidade Cob") & ""
      Estado_Cob.Text = rsCliFor("Estado Cob") & ""
      CEP_Cob.Text = rsCliFor("CEP Cob") & ""
      Boleto.Text = Trim(rsCliFor("Mensagem Boleto")) & ""
      
    'Case 3
      If rsCliFor("Sem Mala Direta") = False Then Sem_Mala.Value = 0
      If rsCliFor("Sem Mala Direta") = True Then Sem_Mala.Value = 1
      
    'Case 4
  
      rsCli_Cred.Index = "Código"
      rsCli_Cred.Seek "=", rsCliFor("Código")
      If rsCli_Cred.NoMatch Then
        C_Empresa.Text = ""
        C_Salário.Text = ""
        '30/06/2004 - Daniel
        'Adicionado campo PercentualLimite
        mskPercentualLimite.Mask = ""
        mskPercentualLimite.Text = ""
        '---------------------------------
        C_Cargo.Text = ""
        C_Contratação.Text = ""
        C_Telefone.Text = ""
        C_C_Nome.Text = ""
        C_C_RG.Text = ""
        C_C_Empresa.Text = ""
        C_C_Salário.Text = ""
        C_C_Cargo.Text = ""
        C_C_Contratação.Text = ""
        C_C_Telefone.Text = ""
        C_Obs.Text = ""
      Else
        C_Empresa.Text = rsCli_Cred("Empresa") & ""
        C_Salário.Text = rsCli_Cred("Salário") & ""
        '30/06/2004 - Daniel
        'Adicionado campo PercentualLimite
        mskPercentualLimite.Text = rsCli_Cred("PercentualLimite").Value
        '--------------------------------------------------------------
        C_Cargo.Text = rsCli_Cred("Cargo") & ""
        C_Contratação.Text = rsCli_Cred("Contratação") & ""
        C_Telefone.Text = rsCli_Cred("Telefone") & ""
        C_C_Nome.Text = rsCli_Cred("Conjugê") & ""
        C_C_RG.Text = rsCli_Cred("C_CPF") & ""
        C_C_Empresa.Text = rsCli_Cred("C_Empresa") & ""
        C_C_Salário.Text = rsCli_Cred("C_Salário") & ""
        C_C_Cargo.Text = rsCli_Cred("C_Cargo") & ""
        C_C_Contratação.Text = rsCli_Cred("C_Contratação") & ""
        C_C_Telefone.Text = rsCli_Cred("C_Telefone") & ""
        C_Obs.Text = rsCli_Cred("OBS") & ""
      End If
      
      '15/06/2005 - Daniel
      'Tratamento para os campos NomeSacadorAvalista, CPFSacadorAvalista e CPFCNPJSacadorAvalista
      txtNomeSacadorAvalista.Text = rsCliFor.Fields("NomeSacadorAvalista").Value & ""
      txtCPFSacadorAvalista.Text = rsCliFor.Fields("CPFSacadorAvalista").Value & ""
      
      If rsCliFor.Fields("CPFCNPJSacadorAvalista").Value = "CPF" Then optCPF.Value = True
      If rsCliFor.Fields("CPFCNPJSacadorAvalista").Value = "CNPJ" Then optCNPJ.Value = True
      '------------------------------------------------------------------------------------------
      
      Limite.Text = rsCliFor("Limite Crédito")
            
      'tab Loja Virtual
      sspWebShopperID.Caption = rsCliFor.Fields("WebShopperID").Value & ""
      sspWebBonus.Caption = rsCliFor.Fields("WebBonus").Value & ""
      sspWebDataCadastro.Caption = rsCliFor.Fields("WebDataCadastro").Value & ""
      
      '23/06/2004 - mpdea
      'Novos campos
      sspWebEMailMerco.Caption = IIf(rsCliFor.Fields("WebEMailMerco").Value, "SIM", "NÃO")
      sspWebEMailLoja = IIf(rsCliFor.Fields("WebEMailLoja").Value, "SIM", "NÃO")
      sspWebOrigem.Caption = gstrWEB_GetDescOrigem(rsCliFor.Fields("WebOrigem").Value & "")
      
      '14/07/2004 - Daniel
      'Campos da guia Classificação de Clientes
      'Serão vistos somente para a TV Shopping
      If CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959") Then
        If IsNumeric(rsCliFor.Fields("TotDinheiroBoletos").Value) Then mskDinheiroBoletos.Text = rsCliFor.Fields("TotDinheiroBoletos").Value
        If IsNumeric(rsCliFor.Fields("TotCheques").Value) Then mskCheques.Text = rsCliFor.Fields("TotCheques").Value
        If IsNumeric(rsCliFor.Fields("TotCartoes").Value) Then mskCartoes.Text = rsCliFor.Fields("TotCartoes").Value
        If IsNumeric(rsCliFor.Fields("TotRecebido").Value) Then mskTotRecebido.Text = rsCliFor.Fields("TotRecebido").Value
      
        txtGrupo.Text = "" 'Será preenchido somente quando for cliente
      
        If rsCliFor.Fields("Tipo").Value = "C" Then 'Cliente, tratamento para mostrar o código e o nome do grupo de classificação
          Dim rstGrupos As Recordset
          Dim strQuery  As String
          
          strQuery = "SELECT * FROM GruposDeClientes "
          strQuery = strQuery & " WHERE Filial = " & gnCodFilial
          
          Set rstGrupos = db.OpenRecordset(strQuery, dbOpenDynaset)
          
          If rstGrupos.RecordCount <> 0 Then 'Não está vazio
            With rstGrupos
              If Not (.BOF And .EOF) Then
                .MoveFirst
                
                Select Case (rsCliFor.Fields("CodGrupo").Value)
                  Case .Fields("CodigoG1").Value
                    txtGrupo.Text = .Fields("CodigoG1").Value & " - " & .Fields("NomeG1").Value & ""
                  Case .Fields("CodigoG2").Value
                    txtGrupo.Text = .Fields("CodigoG2").Value & " - " & .Fields("NomeG2").Value & ""
                  Case .Fields("CodigoG3").Value
                    txtGrupo.Text = .Fields("CodigoG3").Value & " - " & .Fields("NomeG3").Value & ""
                  Case .Fields("CodigoG4").Value
                    txtGrupo.Text = .Fields("CodigoG4").Value & " - " & .Fields("NomeG4").Value & ""
                End Select
              End If
            End With
          End If
          
          rstGrupos.Close
          Set rstGrupos = Nothing
        End If
      End If
      
      
  'End Select
  
  
  If Tab1.TabEnabled(1) = False Then
    Tab1.TabEnabled(1) = True
    Tab1.TabEnabled(2) = True
    Tab1.TabEnabled(3) = True
    Tab1.TabEnabled(4) = True
    Tab1.TabEnabled(5) = True
    '27/05/2004 - mpdea
    'Web
    Tab1.TabEnabled(6) = rsCliFor("Web").Value
    '09/07/2004 - Daniel
    'Adição da Tab de Classificação de Clientes
    'Case: TV Shopping
    Tab1.TabEnabled(7) = CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959")
    
    Frame1.Refresh
    Frame2.Refresh
    Frame3.Refresh
  End If
  
  '18/07/2007 - Anderson
  'Exibe campo SadigWeb
  Select Case "" & rsCliFor("SadigWeb_Tipo")
    Case "AG - AGROPECUARIA", "CC - CRIADOR DE CAES", "CG - CRIADOR DE GATOS", "CL - CLINICA VETERINARIA", "CP - CLINICA VETERINARIA COM PETSHOP", "CR - CRIADOR DE CAES E GATOS", "PC - PETSHOP COM CLINICA VETERINARIA", "PE -PETSHOP", "VE -VETERINARIO"
      cboSadigWebTipo.Text = "" & rsCliFor("SadigWeb_Tipo").Value
    Case Else
      cboSadigWebTipo.Text = "(Nenhum)"
  End Select
  
  '19/07/2007 - Anderson
  'Não habilita o campo Lembrar em quando o cliente está Inativo
  'Solicitante: Supri Print
  If bolCaseSupriPrint Then
    If rsCliFor("Inativo") Then
      txtContatosEfetuadosLembrarEm.Visible = False
      Label57.Visible = False
      chkContatoEfetuadoPendencia.Visible = False
    Else
      txtContatosEfetuadosLembrarEm.Visible = True
      Label57.Visible = True
      chkContatoEfetuadoPendencia.Visible = True
    End If
  End If
  
  Num_Registro = rsCliFor.Bookmark
  
  ConsultarLimiteCreditoCliente
  
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apresentar o registro em Clientes/Fornecedores."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

'Private Sub grdCaract_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
'  If Len(Trim(grdCaract.Columns("Valor").Text)) > 0 Then
'    If Len(Trim(grdCaract.Columns("Descricao").Text)) = 0 Then
'      Beep
'      Cancel = True
'    End If
'  End If
'End Sub
'
Private Sub grdCaract_KeyPress(KeyAscii As Integer)
  Dim sTipo As String
  If Len(Trim(grdCaract.Columns("Descricao").Text)) = 0 Then
    If O_Cliente.Value = True Then
      sTipo = "Cliente"
    Else
      If O_Fornecedor.Value = True Then
        sTipo = "Fornecedor"
      Else
        If O_Revendedor.Value = True Then
          sTipo = "Revendedor"
        Else
          sTipo = "Outros"
        End If
      End If
    End If
    gsTitle = LoadResString(201)
    gsMsg = "Defina antes características para o tipo """ & sTipo & """ no menu ""Cadastro/Características..."" para inserir valores aqui."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    KeyAscii = 0
  End If
End Sub

Private Sub grdContatos_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim sDia As String
  Dim sMes As String
  
  Call StatusMsg("")
  If ColIndex = 2 Then
    sDia = gsHandleNull(grdContatos.Columns("DiaAniv").Text)
    If Not IsNumeric(sDia) Or Val(sDia) < 0 Or Val(sDia) > 31 Then
      DisplayMsg "Dia de aniversário inválido."
      Cancel = True
      grdContatos.SetFocus
    End If
  ElseIf ColIndex = 3 Then
    grdContatos.Columns("MesAniv").Text = UCase(grdContatos.Columns("MesAniv").Text)
    sMes = grdContatos.Columns("MesAniv").Text
    Select Case sMes
      Case "", "JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"
        'OK
      Case Else
        DisplayMsg "Mês incorreto, use JAN,FEV,MAR,ABR,MAI,JUN,JUL,AGO,SET,OUT,NOV,DEZ ou deixe em branco."
        Cancel = True
        grdContatos.SetFocus
    End Select
  End If
End Sub

Private Sub grdContatos_GotFocus()
  If Len(Trim(cboNome.Text)) > 0 And O_Física.Value = True Then
    If Len(Trim(grdContatos.Columns("Contato").Text)) = 0 Then
      grdContatos.Columns("Contato").Text = Mid(LTrim(RTrim(cboNome.Text)), 1, 30)
    End If
  End If
End Sub

Private Sub grdContatos_LostFocus()
 If grdContatos.RowChanged = True Then
   grdContatos.Update
 End If
End Sub

Private Sub grdContatos_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  If Len(Trim(grdContatos.ActiveCell.Text)) = 0 Then
    If bGridBeforeDelete() = True Then
      Cancel = False
    Else
      Cancel = True
    End If
  Else
    Cancel = True
  End If
End Sub

Private Sub grdContatosEfetuados_AfterDelete(RtnDispErrMsg As Integer)
  '21/12/2006 - Anderson
  Call cmdContatoEfetuadoNovo_Click
End Sub

Private Sub grdContatosEfetuados_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  If grdContatosEfetuados.Columns(ColIndex).Name = "Data" Or _
     grdContatosEfetuados.Columns(ColIndex).Name = "DataAviso" Then
    If Len(Trim(grdContatosEfetuados.Columns(ColIndex).Text)) > 0 Then
      grdContatosEfetuados.Columns(ColIndex).Text = Format(grdContatosEfetuados.Columns(ColIndex).Text, "dd/mm/yyyy")
      If Not IsDate(grdContatosEfetuados.Columns(ColIndex).Text) Then
        DisplayMsg "Data inválida. Reentre."
        Cancel = True
      End If
    End If
  End If
End Sub

Private Sub grdContatosEfetuados_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)

  '13/07/2007 - Anderson
  'Verifica se o usuário possui permissão para excluir contatos efetuados.
  If (Mid(m_strContatosEfetuados, 3, 1) = "0") Then
    gsTitle = LoadResString(201)
    gsMsg = "Você não tem permissão para excluir contatos efetuados. Favor verificar as permissões para este usuário."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Cancel = True
    Exit Sub
  End If
  
  DispPromptMsg = False
  Call StatusMsg("")
  gsTitle = LoadResString(201)
  gsMsg = "Deseja apagar os Contatos selecionados?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    DisplayMsg "Contatos selecionados não apagados."
    Cancel = True
  End If
  
End Sub

Private Sub grdContatosEfetuados_Click()
  grdContatosEfetuados.SelectByCell = True
End Sub

Private Sub grdContatosEfetuados_GotFocus()

  '30/06/2005 - Daniel
  'Adicionado tratamento de erros
  On Error GoTo TratarErro
  
  If grdContatosEfetuados.Col = 0 Then
    grdContatosEfetuados.ActiveCell.SelStart = 0
    grdContatosEfetuados.ActiveCell.SelLength = 10
  End If
  
  Exit Sub

TratarErro:
  MsgBox "Erro: " & Err.Number & Err.Description, vbExclamation, "Quick Store"
  Err.Clear
  Exit Sub
  
End Sub

Private Sub grdContatosEfetuados_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF2
      If grdContatosEfetuados.Col <> -1 Then
        With grdContatosEfetuados.Columns(grdContatosEfetuados.Col)
          If .Name = "Data" Or .Name = "DataAviso" Then
            .Text = frmCalendario.gsDateCalender(.Text)
          End If
        End With
      End If
  End Select
End Sub

'21/12/2006 - Anderson - Exibe o registro selecionado nos campos da tela de contatos efetuados
Private Sub grdContatosEfetuados_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

  With rsContatos_Efetuados
    .Index = "Cliente"
    .Seek "=", cboCodigo.Value, grdContatosEfetuados.Columns(1).Value, grdContatosEfetuados.Columns(0).Value
    If Not .NoMatch Then
      lngContatosEfetuadosSequencia = .Fields("Seqüência")
      strContatosEfetuadosData = .Fields("Data")
      '
      'txtContatosEfetuadosData.Text = IIf(IsNull(.Fields("Data")), "  /  /    ", Format(.Fields("Data"), "dd/mm/yyyy"))
      'txtContatosEfetuadosLembrarEm.Text = IIf(IsNull(.Fields("Data Aviso")), "  /  /    ", Format(.Fields("Data Aviso"), "dd/mm/yyyy"))
      txtContatosEfetuadosData.Text = IIf(IsNull(.Fields("Data")), "  /  /    ", Format(.Fields("Data"), "dd/mm/yyyy"))
      txtContatosEfetuadosLembrarEm.Text = IIf(IsNull(.Fields("Data Aviso")), "  /  /    ", Format(.Fields("Data Aviso"), "dd/mm/yyyy"))
      txtContatosEfetuadosDescricao.Text = .Fields("Descrição")
      chkContatoEfetuadoPendencia.Value = -.Fields("Pendência")
    End If
  End With

End Sub

Private Sub grdGruposMalaDireta_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  If grdGruposMalaDireta.Columns(ColIndex).Name = "Grupo" Then
    If Len(Trim(grdGruposMalaDireta.Columns(ColIndex).Value)) > 0 Then
      rsGrupos.FindFirst "Código = " & Val(grdGruposMalaDireta.Columns(ColIndex).Value)
      If Not rsGrupos.NoMatch Then
        grdGruposMalaDireta.Columns("Nome").Value = rsGrupos("Nome").Value
      Else
        grdGruposMalaDireta.Columns("Nome").Value = ""
      End If
    End If
  End If
End Sub

Private Sub grdGruposMalaDireta_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  gsTitle = LoadResString(201)
  gsMsg = "Deseja apagar os Grupos selecionados?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    DisplayMsg "Grupos selecionados não apagados."
    Cancel = True
  End If
End Sub

Private Sub grdGruposMalaDireta_InitColumnProps()
  grdGruposMalaDireta.Columns(0).DropDownHwnd = ddwGruposMalaDireta.hwnd
End Sub

Private Sub grdNumeravel_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  If grdNumeravel.Columns(ColIndex).Name = "Data1" Or _
     grdNumeravel.Columns(ColIndex).Name = "Data2" Then
    If Len(Trim(grdNumeravel.Columns(ColIndex).Text)) > 0 Then
      grdNumeravel.Columns(ColIndex).Text = Format(grdNumeravel.Columns(ColIndex).Text, "dd/mm/yyyy")
      If Not IsDate(grdNumeravel.Columns(ColIndex).Text) Then
        DisplayMsg "Data inválida. Reentre."
        Cancel = True
      End If
    End If
  End If
End Sub

Private Sub grdNumeravel_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF2
      If grdNumeravel.Col <> -1 Then
        With grdNumeravel.Columns(grdNumeravel.Col)
          If .Name = "Data1" Or .Name = "Data2" Then
            .Text = frmCalendario.gsDateCalender(.Text)
          End If
        End With
      End If
  End Select
End Sub

Private Sub Insc_LostFocus()
  If Len(Insc.Text) > 0 Then Insc.Text = UCase(Insc.Text)
End Sub

Private Sub Limite_LostFocus()
  If IsNull(Limite.Text) Then Limite.Text = 0
  If Limite.Text = "" Then Limite.Text = 0
  If Not IsNumeric(Limite.Text) Then Limite.Text = 0
  If CDbl(Limite.Text) < 0 Then Limite.Text = 0
'  IsNull(Limite.Text)GoSub ConstLimite.Text = 0
  Call ConsultarLimiteCreditoCliente
End Sub

Private Sub cboNome_CloseUp()
  If Not rsCliFor.EOF Then
    rsCliFor.FindFirst "Código = " & CLng(cboNome.Columns(1).Text)
    If Not rsCliFor.EOF Then
      Call ShowRecord
      Call ResetAllGrids
    End If
  End If
End Sub

Private Sub cboNome_DropDown()
  gbComboDropped = True
End Sub

Private Sub cboNome_KeyPress(KeyAscii As Integer)
  If Len(cboNome.Text) >= rsCliFor("Nome").Size Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub cboNome_LostFocus()
  gbComboDropped = False
  If Not rsCliFor.EOF Then
    If Len(Trim(cboNome.Text)) > 0 Then
      If Mid(cboNome.Text, Len(cboNome.Text), 1) = "*" Then
        rsCliFor.FindFirst "Nome Like '*" & cboNome.Text & "'"
        If Not rsCliFor.NoMatch Then
          Call ShowRecord
          Call ResetAllGrids
        End If
      End If
    End If
  End If

 If IsNull(cboNome.Text) Then Exit Sub
 If cboNome.Text = "" Then Exit Sub

 If IsNull(Num_Registro) Then
   rsCliFor2.Index = "Nome"
   rsCliFor2.Seek ">", cboNome.Text, 0
   If rsCliFor2.NoMatch Then Exit Sub
   If UCase(cboNome.Text) = UCase(rsCliFor2("Nome")) Then
     DisplayMsg "Atenção, já existe um cliente/fornecedor com este Nome. Código " + str(rsCliFor2("Código"))
   End If
 End If
 
End Sub

Private Sub mskDatAbertura_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      mskDatAbertura.Text = frmCalendario.gsDateCalender(mskDatAbertura.Text)
  End Select
End Sub

Private Sub mskDatAbertura_LostFocus()
 mskDatAbertura.Text = Ajusta_Data(mskDatAbertura.Text)
 If Not IsDate(mskDatAbertura.Text) And mskDatAbertura.Text <> "  /  /    " Then
    DisplayMsg ("Data de abertura inválida. Verifique")
    mskDatAbertura.SelStart = 0
    mskDatAbertura.SelLength = Len(mskDatAbertura.Text)
    mskDatAbertura.SetFocus
 End If
End Sub

Private Sub mskDatNasc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      mskDatNasc.Text = frmCalendario.gsDateCalender(mskDatNasc.Text)
  End Select
End Sub

Private Sub mskDatNasc_LostFocus()
 mskDatNasc.Text = Ajusta_Data(mskDatNasc.Text)
 If Not IsDate(mskDatNasc.Text) And mskDatNasc.Text <> "  /  /    " Then
    DisplayMsg ("Data de Nascimento inválida. Verifique")
    mskDatNasc.SelStart = 0
    mskDatNasc.SelLength = Len(mskDatNasc.Text)
'    mskDatNasc.SetFocus
 End If
End Sub

Private Sub mskPercentualLimite_LostFocus()
  '30/06/2004 - Daniel
  'Case: Nazareno
  'Criado rotina para calcular automaticamente após o Lost do campo PercentualLimite
  'o cálculo é o percentual sobre o salário, este valor será lançado no campo limite
  If Len(C_Salário.Text) > 0 Then
    If IsNumeric(mskPercentualLimite.Text) Then
      Dim dblLimite As Double
      
      dblLimite = CDbl(C_Salário.Text) * CDbl(mskPercentualLimite.Text) / 100
      Limite.Text = ""
      Limite.Text = Format(dblLimite, "###,###,##0.00")
      C_Cargo.SetFocus
    End If
  End If
End Sub

Private Sub O_Cliente_Click()
  '26/07/2004 - Daniel
  'Não mostrar o campo Agência de Publicidade
  'Case: STC
  chkAgencia.Visible = False
  
  If IsNull(Num_Registro) Then
    Exit Sub
  End If
  If rsCliFor("Tipo") <> "C" Then
    Call ResetGridCaract
  End If
  
End Sub

Private Sub O_Física_Click()
  If O_Física.Value = True Then
    L_CGC.Caption = "CPF :"
    L_Inscrição.Caption = "RG :"
    
    frameSex.Visible = True
    Label46.Visible = True
    mskDatNasc.Visible = True
  End If
  
 '15/06/2005 - Daniel
 'Tratamento para os campos da frame "fraSacadorAvalista"
 'que estarão habilitados somente para Pessoa "J"
 fraSacadorAvalista.Enabled = False
 lblNomeSacadorAvalista.Enabled = False
 optCNPJ.Enabled = False
 optCPF.Enabled = False
 txtNomeSacadorAvalista.Enabled = False
 txtCPFSacadorAvalista.Enabled = False
 txtNomeSacadorAvalista.Text = ""
 txtCPFSacadorAvalista.Text = ""

  '27/08/2009 - mpdea
  'Novos campos para NFe
  txtInscricaoMunicipal.Visible = False
  txtInscricaoSuframa.Visible = False
  txtCNAE.Visible = False
  
  Insc.Enabled = O_Física.Value

End Sub

Private Sub O_Fornecedor_Click()
  '26/07/2004 - Daniel
  'Carregar o campo Agência de Publicidade
  'Case: STC
  If CheckSerialCaseMod("QS39823-684") Then
    chkAgencia.Visible = True
  End If

  If IsNull(Num_Registro) Then
    Exit Sub
  End If
  
  If rsCliFor("Tipo") <> "F" Then
    Call ResetGridCaract
  End If
  
End Sub

Private Sub O_Jurídica_Click()
 If O_Jurídica.Value = True Then
   L_CGC.Caption = "CNPJ:"
   L_Inscrição.Caption = "Insc.Estadual :"
   
   Label46.Visible = False
   mskDatNasc.Visible = False
   frameSex.Visible = False
 End If
 
 '15/06/2005 - Daniel
 'Tratamento para os campos da frame "fraSacadorAvalista"
 'que estarão habilitados somente para Pessoa "J"
 fraSacadorAvalista.Enabled = True
 lblNomeSacadorAvalista.Enabled = True
 optCNPJ.Enabled = True
 optCPF.Enabled = True
 txtNomeSacadorAvalista.Enabled = True
 txtCPFSacadorAvalista.Enabled = True
 
  '27/08/2009 - mpdea
  'Novos campos para NFe
  txtInscricaoMunicipal.Visible = True
  txtInscricaoSuframa.Visible = True
  txtCNAE.Visible = True
 
End Sub

Private Sub O_Outros_Click()
  If IsNull(Num_Registro) Then
    Exit Sub
  End If
  If rsCliFor("Tipo") <> "O" Then
    Call ResetGridCaract
  End If
End Sub

Private Sub O_Revendedor_Click()
  If IsNull(Num_Registro) Then
    Exit Sub
  End If
  If rsCliFor("Tipo") <> "R" Then
    Call ResetGridCaract
  End If
End Sub

Private Function ResetGridCaract() As Boolean
  Dim bm As Variant
  Dim nRow As Long
  Dim nCt As Long
  grdCaract.Update
  If grdCaract.Rows > 0 Then
    nCt = 0
    For nRow = 0 To grdCaract.Rows - 1
      bm = grdCaract.AddItemBookmark(nRow)
      If Len(Trim(grdCaract.Columns("Valor").CellText(bm))) > 0 Then
        nCt = nCt + 1
      End If
    Next nRow
    If nCt > 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "Ao trocar de tipo, dados existentes na tabela ""Outros Dados"" foram apagados. Caso não queira perder estes dados não salve a operação atual sobre este registro."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      grdCaract.RemoveAll
    End If
  End If
End Function

Private Sub optCNPJ_Click()
  optCPF.ForeColor = &H0&
  optCNPJ.ForeColor = &H0&
End Sub

Private Sub optCPF_Click()
  optCPF.ForeColor = &H0&
  optCNPJ.ForeColor = &H0&
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
  If Not IsDate(mskDatAbertura.Text) Then mskDatAbertura.Text = Data_Atual
  
  '20/07/2004 - Daniel
  'Personalização TV Shopping
  Dim strData As String
  
  'Chamará somente se for TV Shopping
  If CheckSerialCaseMod("QS39945-043", "QS39944-959", "QS40449-276") Then
    
    '20/07/2004 - Daniel
    'Última Classificação de Clientes
    Call DataCurva(strData)
    
    lblDataAtual.Caption = strData
  End If
  
End Sub

'Private Sub Tab1_Click(PreviousTab As Integer)
'  Dim sSql As String
'  Dim Cód As Long
'
'  Rem Desabilita os OCX que estão em outros TABS
'  Rem  feito para consertar o cursor que some
'
'  If Tab1.Tab = 0 Then
'    CEP.TabStop = True       'tab 0
'    Desconto.TabStop = True  'tab 0
'    CEP_Cob.TabStop = False 'tab3
'  End If
'  If Tab1.Tab = 2 Then
'    CEP.TabStop = False       'tab 0
'    Desconto.TabStop = False  'tab 0
'    CEP_Cob.TabStop = False   'tab 3
'  End If
'  If Tab1.Tab = 4 Then
'    CEP.TabStop = False       'tab 0
'    Desconto.TabStop = False  'tab 0
'    CEP_Cob.TabStop = False    'tab 3
'  End If
'
'  If IsNull(Num_Registro) Then
'    Exit Sub
'  End If
'
'  Call ShowRecord
'
'End Sub
'
'Private Sub GetNewCode(ByRef F As Form, ByRef rs As Recordset, ByVal nMaxCod As Long)
'  Dim rsTemp As Recordset
'  Dim nCod As Long
'
'  nCod = 1
'  Set rsTemp = rs.Clone
'  With rsTemp
'    If Not .EOF Then
'      .MoveLast
'      nCod = .Fields(0) + 1
'      If nCod > nMaxCod Then
'        For nCod = 1 To nMaxCod
'          DoEvents
'          .FindFirst .Fields(0).Name & " = " & nCod
'          If .NoMatch Then
'            F.cboCodigo.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
'            Exit Sub
'          End If
'        Next nCod
'        gsTitle = LoadResString(201)
'        gsMsg = "Nenhum Código livre disponível para o intervalo."
'        gnStyle = vbOKOnly + vbExclamation
'        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'      Else
'        F.cboCodigo.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
'      End If
'    Else
'      F.cboCodigo.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
'    End If
'    .Close
'  End With
'
'  Set rsTemp = Nothing
'
'End Sub
'
'

Private Sub Tab1_DblClick()

End Sub

Private Sub DataCurva(ByRef strData As String)
  Dim rstDataCurva As Recordset

  On Error GoTo TratarErro
  
  Set rstDataCurva = dbTemp.OpenRecordset("SELECT * FROM DataCurva", dbOpenDynaset)
  
  If rstDataCurva.RecordCount = 0 Then
    strData = ""
    rstDataCurva.Close
    Set rstDataCurva = Nothing
    Exit Sub
  Else
    With rstDataCurva
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        strData = CStr(Format((.Fields("Ultima").Value), "dd/mm/yyyy"))
      End If
      .Close
    End With
    
    Set rstDataCurva = Nothing
    
  End If

  Exit Sub

TratarErro:
  MsgBox "Ocorrência de erro em Private <DataCurva>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Err.Clear
  
End Sub

Private Function blnHabilitarCampoVendedor() As Boolean
  Dim rstParametros As Recordset
  Dim strQuery      As String
  Dim blnVerificar  As Boolean
  
  On Error GoTo TratarErro

  strQuery = ""
  strQuery = "SELECT AlterVendedorCliFor FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial
  
  Set rstParametros = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      .MoveFirst
      blnVerificar = .Fields("AlterVendedorCliFor").Value
    End If
    .Close
  End With
  
  Set rstParametros = Nothing
  
  If Not blnVerificar Then
    blnHabilitarCampoVendedor = True
    
    Exit Function
  Else
    Dim rstFuncionarios As Recordset
    
    strQuery = ""
    strQuery = "SELECT Superusuário FROM Funcionários WHERE Código = " & gnUserCode
  
    Set rstFuncionarios = db.OpenRecordset(strQuery, dbOpenDynaset)
  
    With rstFuncionarios
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        If .Fields("Superusuário").Value Then
          blnHabilitarCampoVendedor = True
        Else
          blnHabilitarCampoVendedor = False
        End If
        
      End If
      .Close
    End With
  
    Set rstFuncionarios = Nothing
  
  End If
  
  Exit Function

TratarErro:
  MsgBox "Ocorrência de erro em Function <blnHabilitarCampoVendedor>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Err.Clear

End Function

'21/12/2006 - Anderson - Exibe calendário para a tela de contatos efetuados
Private Sub txtContatosEfetuadosData_KeyDown(KeyCode As Integer, Shift As Integer)

  '09/05/2007 - Anderson
  'Customização para Supriprint para obrigar a selecção da data pelo calendário. Isto evita erros do tipo digitar datas 09/05/0207
  If bolCaseSupriPrint Then
    KeyCode = vbKeyF2
  End If

  Select Case KeyCode
    Case vbKeyF2
      bolCalendario = True
      txtContatosEfetuadosData.Text = frmCalendario.gsDateCalender(txtContatosEfetuadosData.Text)
      bolCalendario = False
  End Select
End Sub

'21/12/2006 - Anderson - Verifica se a data digitada está correta
Private Sub txtContatosEfetuadosData_LostFocus()
 txtContatosEfetuadosData.Text = Ajusta_Data(txtContatosEfetuadosData.Text)
 If Not IsDate(txtContatosEfetuadosData.Text) And txtContatosEfetuadosData.Text <> "  /  /    " And bolCalendario = False Then
    DisplayMsg ("Data inválida. Favor verificar!")
    txtContatosEfetuadosData.SetFocus
    txtContatosEfetuadosData.SelStart = 0
    txtContatosEfetuadosData.SelLength = Len(txtContatosEfetuadosData.Text)
  End If
End Sub

'21/12/2006 - Anderson - Exibe calendário para a tela de contatos efetuados
Private Sub txtContatosEfetuadosLembrarEm_KeyDown(KeyCode As Integer, Shift As Integer)

  '09/05/2007 - Anderson
  'Customização para Supriprint para obrigar a selecção da data pelo calendário. Isto evita erros do tipo digitar datas 09/05/0207
  If bolCaseSupriPrint Then
    KeyCode = vbKeyF2
  End If

  Select Case KeyCode
    Case vbKeyF2
      bolCalendario = True
      txtContatosEfetuadosLembrarEm.Text = frmCalendario.gsDateCalender(txtContatosEfetuadosLembrarEm.Text)
      bolCalendario = False
  End Select
End Sub

'21/12/2006 - Anderson - Verifica se a data digitada está correta
Private Sub txtContatosEfetuadosLembrarEm_LostFocus()
 txtContatosEfetuadosLembrarEm.Text = Ajusta_Data(txtContatosEfetuadosLembrarEm.Text)
 If bolCalendario = False Then
  If Not IsDate(txtContatosEfetuadosLembrarEm.Text) And txtContatosEfetuadosLembrarEm.Text <> "  /  /    " And bolCalendario = False Then
     DisplayMsg ("Data inválida. Favor verificar!")
     txtContatosEfetuadosLembrarEm.SetFocus
     txtContatosEfetuadosLembrarEm.SelStart = 0
     txtContatosEfetuadosLembrarEm.SelLength = Len(txtContatosEfetuadosLembrarEm.Text)
     chkContatoEfetuadoPendencia.Value = vbUnchecked
   Else
     If txtContatosEfetuadosLembrarEm.Text <> "  /  /    " Then
       chkContatoEfetuadoPendencia.Value = vbChecked
     ElseIf txtContatosEfetuadosLembrarEm.Text = "  /  /    " Then
       chkContatoEfetuadoPendencia.Value = vbUnchecked
     End If
  End If
  'Case Supri print - Não permite que a data digitada seja um final de semana
  If bolCaseSupriPrint Then
    If IsDate(txtContatosEfetuadosLembrarEm.Text) Then
      If (Weekday(CDate(txtContatosEfetuadosLembrarEm.Text)) = 1 Or Weekday(CDate(txtContatosEfetuadosLembrarEm.Text)) = 7) Then
        MsgBox "Não é possível salvar o registro. Verifique se a data digitada é um dia útil da semana.", vbOKOnly + vbInformation, "Atenção"
        txtContatosEfetuadosLembrarEm.SetFocus
        txtContatosEfetuadosLembrarEm.SelStart = 0
        txtContatosEfetuadosLembrarEm.SelLength = Len(txtContatosEfetuadosLembrarEm.Text)
        chkContatoEfetuadoPendencia.Value = vbUnchecked
        Exit Sub
      End If
    End If
  End If

  End If
End Sub

