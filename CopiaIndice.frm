VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosCopiaIndice 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Copia Tabela de Pre�os"
   ClientHeight    =   4965
   ClientLeft      =   4005
   ClientTop       =   1065
   ClientWidth     =   9990
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
   HelpContextID   =   1660
   Icon            =   "CopiaIndice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4965
   ScaleWidth      =   9990
   Begin VB.CheckBox chkContaClientes 
      Appearance      =   0  'Flat
      Caption         =   "&Refletir altera��o tamb�m na Conta de Clientes"
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
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5250
      TabIndex        =   14
      Top             =   3810
      Width           =   4620
   End
   Begin VB.Data datPrecos 
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
      Left            =   3315
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Pre�os ORDER BY Tabela"
      Top             =   7575
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CheckBox Pre�o_Zero 
      Appearance      =   0  'Flat
      Caption         =   "&N�o copiar para produtos com pre�o original igual a 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   105
      TabIndex        =   13
      Top             =   3750
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   75
      TabIndex        =   24
      Top             =   2820
      Width           =   9780
      Begin VB.OptionButton Arredonda_1000 
         Appearance      =   0  'Flat
         Caption         =   "10.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7710
         TabIndex        =   12
         Top             =   570
         Width           =   1035
      End
      Begin VB.OptionButton Arredonda_500 
         Appearance      =   0  'Flat
         Caption         =   "5.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6240
         TabIndex        =   11
         Top             =   570
         Width           =   735
      End
      Begin VB.OptionButton Arredonda_100 
         Appearance      =   0  'Flat
         Caption         =   "1.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2850
         TabIndex        =   10
         Top             =   570
         Width           =   735
      End
      Begin VB.OptionButton Arredonda_050 
         Appearance      =   0  'Flat
         Caption         =   "0.50"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7710
         TabIndex        =   9
         Top             =   210
         Width           =   735
      End
      Begin VB.OptionButton Arredonda_010 
         Appearance      =   0  'Flat
         Caption         =   "0.10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6240
         TabIndex        =   8
         Top             =   210
         Width           =   735
      End
      Begin VB.OptionButton Arredonda_005 
         Appearance      =   0  'Flat
         Caption         =   "Arredondar para 0.05"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2850
         TabIndex        =   7
         Top             =   210
         Width           =   2295
      End
      Begin VB.OptionButton O_Sem_Arredondamento 
         Appearance      =   0  'Flat
         Caption         =   "N�o arrendondar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   195
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   1995
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
      Left            =   1680
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Sub_Classe"
      Top             =   7470
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Sobre 
      Appearance      =   0  'Flat
      Caption         =   "&Sobrep�e pre�os existentes na tabela destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   840
      Width           =   4785
   End
   Begin VB.CommandButton B_Calcula 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Copiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4395
      Width           =   9780
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
      Left            =   75
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1725
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Sub 
      Bindings        =   "CopiaIndice.frx":4E95A
      DataSource      =   "Data2"
      Height          =   405
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Use 0 para todas as Subclasses"
      Top             =   1725
      Width           =   1035
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
      Columns(0).Width=   8202
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2090
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1826
      _ExtentY        =   714
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "CopiaIndice.frx":4E96E
      DataSource      =   "Data1"
      Height          =   405
      Left            =   1320
      TabIndex        =   3
      Top             =   1215
      Width           =   1035
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
      Columns(0).Width=   9208
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1852
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1826
      _ExtentY        =   714
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox Multiplicador 
      Height          =   405
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "##0.00"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo cboTabOrig 
      Bindings        =   "CopiaIndice.frx":4E982
      Height          =   375
      Left            =   1710
      TabIndex        =   0
      Top             =   180
      Width           =   3165
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      Columns(0).Width=   3200
      _ExtentX        =   5583
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo cboTabDest 
      Bindings        =   "CopiaIndice.frx":4E99A
      Height          =   375
      Left            =   6780
      TabIndex        =   1
      Top             =   180
      Width           =   3105
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      Columns(0).Width=   3200
      _ExtentX        =   5477
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Tabela"
   End
   Begin VB.Label Nome_Sub 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2400
      TabIndex        =   23
      Top             =   1725
      Width           =   7455
   End
   Begin VB.Label Label4 
      Caption         =   "Subclasse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   22
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Classe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   21
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Nome_Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2400
      TabIndex        =   20
      Top             =   1215
      Width           =   7455
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Multiplicador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   19
      Top             =   2355
      Width           =   1185
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      Caption         =   "Para manter os pre�os use 1,00. Para aumentar os pre�os 10% use 1,10. Para diminuir 10% use 0,9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   2400
      TabIndex        =   18
      Top             =   2220
      Width           =   7455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Tabela DESTINO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5250
      TabIndex        =   17
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label Tabela 
      Appearance      =   0  'Flat
      Caption         =   "Tabela ORIGINAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   16
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrecosCopiaIndice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rsPre�os As Recordset
Dim rsPre�os2 As Recordset
Dim rsClasses As Recordset
Dim rsSubclasses As Recordset
Dim rsProdutos As Recordset
Dim rsTabelas As Recordset
Private rsConta_Cli As Recordset
 
Private Sub cboTabDest_CloseUp()
  chkContaClientes.Enabled = True
  chkContaClientes.Value = vbUnchecked
  If cboTabDest.Text = cboTabOrig.Text And cboTabDest.Text <> "" Then
    chkContaClientes.Value = vbChecked
  Else
    chkContaClientes.Enabled = False
  End If
End Sub

Private Sub cboTabOrig_CloseUp()
  chkContaClientes.Enabled = True
  chkContaClientes.Value = vbUnchecked
  If cboTabDest.Text = cboTabOrig.Text And cboTabOrig.Text <> "" Then
    chkContaClientes.Value = vbChecked
  Else
    chkContaClientes.Enabled = False
  End If
End Sub

Private Sub chkContaClientes_Click()
  If chkContaClientes.Value = vbChecked Then
    If Not frmGerente.gbSenhaGerente Then
      chkContaClientes.Value = vbUnchecked
      Exit Sub
    End If
  End If
End Sub

Private Sub Combo_Classe_CloseUp()
  Combo_Classe.Text = Combo_Classe.Columns(1).Text
  Combo_Classe_LostFocus
End Sub

'-----------------------------------------------------------------------------------
'08/07/2002 - mpdea
'Implementado o suporte a transa��o com tratamento a erro
'Implementado a atualiza��o de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub B_Calcula_Click()
  Dim Produto As Variant
  Dim Pre�o As Variant
  Dim Copiados As Long
  Dim Aux As Integer
  Dim i As Integer
  Dim nTempCopiados As Long
  
  Dim Str_Arredonda As String
  Dim Novo_Pre�o As Double
  
  Dim blnOnTransaction As Boolean
  
  On Error GoTo ErrHandler
  
  Copiados = 0
  Produto = 0

  Call StatusMsg("")
  
  If IsNull(cboTabOrig.Text) Or cboTabOrig.Text = "" Then
    DisplayMsg "Tabela de Origem inv�lida!"
    cboTabOrig.SetFocus
    Exit Sub
  End If

  If IsNull(cboTabDest.Text) Or cboTabDest.Text = "" Then
    DisplayMsg "Tabela Destino inv�lida!"
    cboTabDest.SetFocus
    Exit Sub
  End If
  
  cboTabDest.Text = Trim(cboTabDest.Text)

  If IsNull(Multiplicador.Text) Then
    DisplayMsg "Digite o multiplicador."
    Multiplicador.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(Multiplicador.Text) Then
    DisplayMsg "Digite o multiplicador."
    Multiplicador.SetFocus
    Exit Sub
  End If
  

  If cboTabDest.Text = cboTabOrig.Text Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja efetuar as altera��es na mesma tabela de pre�os?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
'      DisplayMsg "Tabela n�o alterada."
      Exit Sub
    End If
    Sobre.Value = 1
  End If

  '28/02/2005 - Daniel
  '
  'Solicita��o: Consultora Marineida
  '
  'Criado valida��o para n�o prosseguir caso n�o seja o gerente
  If Trim(cboTabOrig.Text) = Trim(cboTabDest.Text) Then
    If Not frmGerente.gbSenhaGerente Then Exit Sub
  End If
  '------------------------------------------------------------

  Screen.MousePointer = vbHourglass
  ws.BeginTrans
  blnOnTransaction = True
  
  If CDbl(Multiplicador.Text) <= 0 Then Multiplicador.Text = 1

  If IsNull(Nome_Classe.Caption) Or Nome_Classe.Caption = "" Then Combo_Classe.Text = 0
  If IsNull(Nome_Sub.Caption) Or Nome_Sub.Caption = "" Then Combo_Sub.Text = 0


  Rem Come�a a copiar as tabelas
  rsProdutos.Index = "C�digo"
  rsPre�os.Index = "Tabela"
  rsPre�os2.Index = "Tabela"
  
  Str_Arredonda = "000"
  If Arredonda_005.Value = True Then Str_Arredonda = "005"
  If Arredonda_010.Value = True Then Str_Arredonda = "010"
  If Arredonda_050.Value = True Then Str_Arredonda = "050"
  If Arredonda_100.Value = True Then Str_Arredonda = "100"
  If Arredonda_500.Value = True Then Str_Arredonda = "500"
  If Arredonda_1000.Value = True Then Str_Arredonda = "1000"
  

Lp1:
  If nTempCopiados <> Copiados Then
    nTempCopiados = Copiados
    Call StatusMsg("Foram copiados " & Copiados & " registros.")
  End If
  rsPre�os.Seek ">", cboTabOrig.Text, Produto
  If rsPre�os.NoMatch Then
    Aux = 1
    GoTo Fim
  End If
  If rsPre�os("Tabela") <> cboTabOrig.Text Then
    Aux = 2
    GoTo Fim
  End If

  Produto = rsPre�os("Produto")
  
  rsProdutos.Seek "=", Produto
  If rsProdutos.NoMatch Then GoTo Lp1

  If Pre�o_Zero.Value = 1 Then
    If rsPre�os("Pre�o") = 0 Then
      GoTo Lp1
    End If
  End If


  Rem Verifica se e' da classe desejada
  If Val(Combo_Classe.Text) <> 0 Then
     If rsProdutos("Classe") <> Val(Combo_Classe.Text) Then GoTo Lp1
  End If

  Rem Verifica se e' da sub classe desejada
  If Val(Combo_Sub.Text) <> 0 Then
     If rsProdutos("Sub Classe") <> Val(Combo_Sub.Text) Then GoTo Lp1
  End If


  Novo_Pre�o = rsPre�os("Pre�o") * CDbl(Multiplicador.Text)
  Novo_Pre�o = Arredonda_Valor(Novo_Pre�o, Str_Arredonda)


  rsPre�os2.Seek "=", cboTabDest.Text, rsPre�os("Produto")
  If Not rsPre�os2.NoMatch Then
    If Sobre.Value = 0 Then
      GoTo Lp1
    End If

    rsPre�os2.Edit
    rsPre�os2("Pre�o") = Format(Novo_Pre�o, "#############0.00")
    rsPre�os2("Data Altera��o") = Format(Date, "dd/mm/yyyy")
    rsPre�os2.Update
    
    If chkContaClientes.Value = vbChecked Then
      Call UpdateContaClientes(cboTabDest.Text, rsPre�os2("Produto").Value, Novo_Pre�o)
    End If
  
    'Atualiza o sincronismo para o produto WEB alterado
    Call WEB_SynchronizeProduct(rsPre�os("Produto").Value)
    
    Copiados = Copiados + 1
    GoTo Lp1
  End If


  rsPre�os2.AddNew
  
  rsPre�os2("Tabela") = cboTabDest.Text
  rsPre�os2("Produto") = rsPre�os("Produto")
  rsPre�os2("Pre�o") = Format(Novo_Pre�o, "#############0.00")
  rsPre�os2("Data Altera��o") = Format(Date, "dd/mm/yyyy")
  
  rsPre�os2.Update

  If chkContaClientes.Value = vbChecked Then
    Call UpdateContaClientes(cboTabDest.Text, rsPre�os2("Produto").Value, Novo_Pre�o)
  End If
  
  'Atualiza o sincronismo para o produto WEB alterado
  Call WEB_SynchronizeProduct(rsPre�os("Produto").Value)
  
  Copiados = Copiados + 1

  GoTo Lp1

Fim:
 
  'Cria configura��o da tabela
  Call CheckConfigTablePrice(cboTabDest.Text)
  
  ' 10/12/2007 - Celso
  '---[ Gera Log do usu�rio ]---'
      g_GravaLog CDate(CStr(Data_Atual) & " " & Format(Now, "hh:mm:ss")), _
                 "Usr: " & gnUserCode & " - " & gsUserName & _
                 " -Tb org: " & cboTabOrig.Text & " -Tb dst: " & cboTabDest.Text & _
                 " -Ind: " & Multiplicador.Text _
                 , "C�pia Tbl apl.�ndice"
                             
  '---[ Gera Log do usu�rio ]---'
 
  
  ws.CommitTrans
  blnOnTransaction = False
  
  datPrecos.Refresh
  cboTabOrig.Refresh
  cboTabDest.Refresh
  
  cboTabDest.Text = ""
  
  Screen.MousePointer = vbDefault
  
  DisplayMsg "Final de processo. Copiados " & Copiados & " registros."
  
  Call StatusMsg("")

  Exit Sub

ErrHandler:
  Screen.MousePointer = vbDefault
  If blnOnTransaction Then ws.Rollback
  MsgBox "Erro [" & Err.Number & "] - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Combo_Classe_LostFocus()
  Nome_Classe.Caption = ""
  If IsNull(Combo_Classe.Text) Then Exit Sub
  If Not IsNumeric(Combo_Classe.Text) Then Exit Sub

  rsClasses.Index = "C�digo"
  rsClasses.Seek "=", Combo_Classe.Text
  If Not rsClasses.NoMatch Then
     Nome_Classe.Caption = rsClasses("Nome")
  Else
     Combo_Classe.Text = 0
  End If

End Sub

Private Sub Combo_Sub_CloseUp()
 Combo_Sub.Text = Combo_Sub.Columns(1).Text
 Combo_Sub_LostFocus

End Sub

Private Sub Combo_Sub_LostFocus()
  Nome_Sub.Caption = ""
  If IsNull(Combo_Sub.Text) Then Exit Sub
  If Not IsNumeric(Combo_Sub.Text) Then Exit Sub

  rsSubclasses.Index = "C�digo"
  rsSubclasses.Seek "=", Combo_Sub.Text
  If Not rsSubclasses.NoMatch Then
     Nome_Sub.Caption = rsSubclasses("Nome")
  Else
     Combo_Sub.Text = 0
  End If

End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsPre�os = db.OpenRecordset("Pre�os")
  Set rsPre�os2 = db.OpenRecordset("Pre�os")
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsTabelas = db.OpenRecordset("Tabela de Pre�os")
  Set rsConta_Cli = db.OpenRecordset("SELECT * FROM [Conta Cliente]", dbOpenDynaset)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  datPrecos.DatabaseName = gsQuickDBFileName

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsPre�os.Close
  rsPre�os2.Close
  rsClasses.Close
  rsSubclasses.Close
  rsProdutos.Close
  rsTabelas.Close
  rsConta_Cli.Close
  Set rsPre�os = Nothing
  Set rsPre�os2 = Nothing
  Set rsClasses = Nothing
  Set rsSubclasses = Nothing
  Set rsProdutos = Nothing
  Set rsTabelas = Nothing
  Set rsConta_Cli = Nothing
End Sub

Private Sub Multiplicador_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub cboTabDest_KeyPress(KeyAscii As Integer)
  KeyAscii = gnLimitKeyPress(cboTabDest, 15, KeyAscii)
  If KeyAscii <> 0 Then
    KeyAscii = gnTypeValidKey(KeyAscii)
  End If
End Sub

Private Sub cboTabDest_LostFocus()
  If IsNull(cboTabDest.Text) Then Exit Sub
  cboTabDest.Text = UCase$(cboTabDest.Text)
'  If cboTabDest.Text = cboTabOrig.Text And Len(cboTabOrig.Text) > 0 Then
'    DisplayMsg "Aviso: As altera��es ser�o realizadas na mesma tabela e n�o existe um desfaz autom�tico."
'  End If
End Sub
