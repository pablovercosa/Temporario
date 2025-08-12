VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmGrafico5 
   Caption         =   " Maiores Fornecedores no período"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrafico5.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   13920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_detalharForn2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Detalhar o movimento de UM fornecedor da grade (Produtos da Data de Compra)"
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
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7515
      Width           =   6800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5505
      Left            =   90
      TabIndex        =   12
      Top             =   1935
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   9710
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   15066597
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Maiores Fornecedores do período"
      TabPicture(0)   =   "frmGrafico5.frx":4E95A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gridForn"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalhamento nível 01"
      TabPicture(1)   =   "frmGrafico5.frx":4E976
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_nmForn01"
      Tab(1).Control(1)=   "txt_codForn01"
      Tab(1).Control(2)=   "txt_qtde01"
      Tab(1).Control(3)=   "gridFornNivel01"
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "lbl_qtde01"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Detalhamento nível 02"
      TabPicture(2)   =   "frmGrafico5.frx":4E992
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txt_nmForn01 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
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
         Height          =   360
         Left            =   -70185
         TabIndex        =   27
         Top             =   630
         Width           =   8835
      End
      Begin VB.TextBox txt_codForn01 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
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
         Height          =   360
         Left            =   -72345
         TabIndex        =   25
         Top             =   630
         Width           =   2130
      End
      Begin VB.TextBox txt_qtde01 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
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
         Height          =   360
         Left            =   -74880
         TabIndex        =   23
         Top             =   630
         Width           =   2130
      End
      Begin MSFlexGridLib.MSFlexGrid gridForn 
         Height          =   5010
         Left            =   90
         TabIndex        =   13
         Top             =   405
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   8837
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   8454143
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16250871
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridFornNivel01 
         Height          =   4365
         Left            =   -74910
         TabIndex        =   18
         Top             =   1050
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   7699
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   16112179
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16250871
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   -70185
         TabIndex        =   26
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Código Fornecedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -72345
         TabIndex        =   24
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lbl_qtde01 
         Caption         =   "Qtde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74910
         TabIndex        =   22
         Top             =   360
         Width           =   2475
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   13740
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
         Left            =   6975
         Picture         =   "frmGrafico5.frx":4E9AE
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   202
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
         Left            =   9000
         Picture         =   "frmGrafico5.frx":4F290
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   202
         Width           =   465
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Por grandeza de ""Valor (R$) pago"""
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   20
         Top             =   765
         Width           =   3240
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Por grandeza de ""Número de itens comprados"""
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Top             =   765
         Value           =   -1  'True
         Width           =   4260
      End
      Begin VB.ComboBox cmb_numForn 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmGrafico5.frx":4FB72
         Left            =   10665
         List            =   "frmGrafico5.frx":4FB9A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   255
         Width           =   1230
      End
      Begin VB.ComboBox Combo_Filial 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmGrafico5.frx":4FBD7
         Left            =   630
         List            =   "frmGrafico5.frx":4FBD9
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   870
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   7830
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   255
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
         Left            =   5820
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   270
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ou"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6705
         TabIndex        =   21
         Top             =   780
         Width           =   210
      End
      Begin VB.Label Label8 
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
         Left            =   7515
         TabIndex        =   11
         Top             =   285
         Width           =   300
      End
      Begin VB.Label Label1 
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
         Left            =   5505
         TabIndex        =   10
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label7 
         Caption         =   "Filial :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   285
         Width           =   465
      End
      Begin VB.Label Nome_Filial 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1515
         TabIndex        =   1
         Top             =   255
         Width           =   3480
      End
      Begin VB.Label Label2 
         Caption         =   "Visualizar (os)                            maiores fornecedores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9540
         TabIndex        =   8
         Top             =   285
         Width           =   4110
      End
   End
   Begin VB.CommandButton cmd_detalharForn 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Detalhar o movimento de UM fornecedor da grade (Datas de Compras)"
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7515
      Width           =   6800
   End
   Begin VB.CommandButton cmd_pesquisar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar Maiores fornecedores do período"
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1350
      Width           =   13755
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9315
      TabIndex        =   17
      Top             =   1890
      Width           =   4425
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4770
      TabIndex        =   15
      Top             =   1890
      Width           =   4425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   14
      Top             =   1890
      Width           =   4425
   End
End
Attribute VB_Name = "frmGrafico5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrayFiliais(20, 2) As String
Dim iConta As Integer

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub cmd_detalharForn_Click()
On Error GoTo Erro
 
  Dim rsEntrada As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
 
  If gridForn.RowSel < 1 Then
    MsgBox "Selecione um Fornecedor na grade da ABA AMARELA.", vbInformation
    SSTab1.Tab = 0
    Exit Sub
  End If

  lbl_qtde01.Caption = gridForn.TextMatrix(0, 1)
  txt_qtde01.Text = gridForn.TextMatrix(gridForn.RowSel, 1)
  txt_codForn01.Text = gridForn.TextMatrix(gridForn.RowSel, 2)
  txt_nmForn01.Text = gridForn.TextMatrix(gridForn.RowSel, 3)
   
  gridFornNivel01.Rows = 1
  gridFornNivel01.Row = 0
  
  If lbl_qtde01.Caption = "Qtde Unid. comprada" Then
      strSQL = "SELECT Sum(EP.Qtde) as SomaDeProdutos,"
      
      gridFornNivel01.Row = 0
      gridFornNivel01.TextMatrix(0, 2) = "Qtde Unid. comprada"
  Else
      strSQL = "SELECT Sum(EP.[Preço Final]) as SomaPreco,"
      
      gridFornNivel01.Row = 0
      gridFornNivel01.TextMatrix(0, 2) = "Qtde R$ pago"
  End If

  strSQL = strSQL & " E.Data , E.Fornecedor, c.Nome "
  strSQL = strSQL & " From Entradas E, [Entradas - Produtos] EP, Cli_for C "
  strSQL = strSQL & " where E.data > CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " E.data < CDATE('" & Data_Fim.Text & " 00:00:00') and "
  strSQL = strSQL & " E.Sequência=EP.Sequência and "
  strSQL = strSQL & " E.Filial= " & Combo_Filial.Text & " and "
  strSQL = strSQL & " E.Fornecedor = c.Código and "
  strSQL = strSQL & " E.Fornecedor = " & txt_codForn01.Text
  strSQL = strSQL & " GROUP BY E.data, E.Fornecedor, C.Nome "
  strSQL = strSQL & " ORDER BY 2 DESC "

  Screen.MousePointer = vbHourglass
  
  Set rsEntrada = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  lngContadorRegGrid = 1
  
  If Not (rsEntrada.EOF And rsEntrada.BOF) Then
    rsEntrada.MoveFirst
  End If
  While Not rsEntrada.EOF
  
      If lbl_qtde01.Caption = "Qtde Unid. comprada" Then
          gridFornNivel01.AddItem lngContadorRegGrid & vbTab & rsEntrada.Fields(1).Value & vbTab & _
                          rsEntrada.Fields(0).Value
      Else
          gridFornNivel01.AddItem lngContadorRegGrid & vbTab & rsEntrada.Fields(1).Value & vbTab & _
                          FormataValorTexto(rsEntrada.Fields(0).Value, 2)
      End If
      
      rsEntrada.MoveNext
      lngContadorRegGrid = lngContadorRegGrid + 1
  Wend
  rsEntrada.Close
  Set rsEntrada = Nothing

  SSTab1.Tab = 1

  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub


Private Sub cmd_detalharForn2_Click()
  SSTab1.Tab = 2
End Sub

Private Sub cmd_pesquisar_Click()
On Error GoTo Erro
 
  Dim rsEntrada As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
 
  If Nome_Filial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    Combo_Filial.SetFocus
    Exit Sub
  End If
 
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
   
  gridForn.Rows = 1
  gridForn.Row = 0
  
  If opt1.Value = True Then
      If cmb_numForn.Text = "TODOS" Then
        strSQL = "SELECT Sum(EP.Qtde) as SomaDeProdutos,"
      Else
        strSQL = "SELECT top " & cmb_numForn.Text & " Sum(EP.Qtde) as SomaDeProdutos,"
      End If
      
      gridForn.Row = 0
      gridForn.TextMatrix(0, 1) = "Qtde Unid. comprada"
  Else
      If cmb_numForn.Text = "TODOS" Then
        strSQL = "SELECT Sum(EP.[Preço Final]) as SomaPreco,"
      Else
        strSQL = "SELECT top " & cmb_numForn.Text & " Sum(EP.[Preço Final]) as SomaPreco,"
      End If
      
      gridForn.Row = 0
      gridForn.TextMatrix(0, 1) = "Qtde R$ pago"
  End If
      
  strSQL = strSQL & " E.Fornecedor , c.Nome "
  strSQL = strSQL & " From Entradas E, [Entradas - Produtos] EP, Cli_For C, [Operações Entrada] OP "
  strSQL = strSQL & " where E.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " E.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
  strSQL = strSQL & " E.Operação=OP.Código and "
  strSQL = strSQL & " not OP.Tipo in('D','A','E') and "  ' Diferente de D=Devolucao, A=Ajuste de Entrada e E=Recebimento de Emprestimo
  strSQL = strSQL & " E.Sequência=EP.Sequência and "
  strSQL = strSQL & " E.Filial= " & Combo_Filial.Text & " and "
  strSQL = strSQL & " E.Fornecedor = c.Código "
  strSQL = strSQL & " GROUP BY E.Fornecedor, C.Nome "
  strSQL = strSQL & " ORDER BY 1 DESC "

  Screen.MousePointer = vbHourglass
  
  Set rsEntrada = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  lngContadorRegGrid = 1
  
  If Not (rsEntrada.EOF And rsEntrada.BOF) Then
    rsEntrada.MoveFirst
  End If
  While Not rsEntrada.EOF
  
      If opt2.Value = True Then
          gridForn.AddItem lngContadorRegGrid & vbTab & FormataValorTexto(rsEntrada.Fields(0).Value, 2) & vbTab & _
                          rsEntrada.Fields(1).Value & vbTab & _
                          rsEntrada.Fields(2).Value
      Else
          gridForn.AddItem lngContadorRegGrid & vbTab & rsEntrada.Fields(0).Value & vbTab & _
                          rsEntrada.Fields(1).Value & vbTab & _
                          rsEntrada.Fields(2).Value
      End If
      
      rsEntrada.MoveNext
      lngContadorRegGrid = lngContadorRegGrid + 1
  Wend
  rsEntrada.Close
  Set rsEntrada = Nothing

  SSTab1.Tab = 0

  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Sub Combo_Filial_LostFocus()
  Dim i As Integer
  Nome_Filial.Caption = ""
  
  If Combo_Filial.Text <> "" Then
      For i = 0 To iConta
        If Combo_Filial.Text = arrayFiliais(i, 0) Then
          Nome_Filial.Caption = arrayFiliais(i, 1)
          Exit For
        End If
      Next
  End If
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

Private Sub Form_Load()
  Dim rsParametros As Recordset
  Set rsParametros = db.OpenRecordset("select Filial, Nome from [Parâmetros Filial]", dbOpenDynaset)
  
  iConta = 0
  While Not rsParametros.EOF
      arrayFiliais(iConta, 0) = rsParametros.Fields(0).Value
      arrayFiliais(iConta, 1) = rsParametros.Fields(1).Value
      Combo_Filial.AddItem rsParametros.Fields(0).Value, iConta
      iConta = iConta + 1
      rsParametros.MoveNext
  Wend
  rsParametros.Close
  Set rsParametros = Nothing

  
' Grade de Erros e Críticas do XML
  gridForn.ColWidth(0) = 600
  gridForn.ColWidth(1) = 2000
  gridForn.ColWidth(2) = 2500
  gridForn.ColWidth(3) = 7900
   
  gridForn.Row = 0
  gridForn.TextMatrix(0, 1) = "Quantidade comprada"
  gridForn.TextMatrix(0, 2) = "Código do Forncedor"
  gridForn.TextMatrix(0, 3) = "Nome do fornecedor"
  
  ' Grade de Erros e Críticas do XML
  gridFornNivel01.ColWidth(0) = 600
  gridFornNivel01.ColWidth(1) = 6200
  gridFornNivel01.ColWidth(2) = 6200
  
  gridFornNivel01.Row = 0
  gridFornNivel01.TextMatrix(0, 1) = "Data da compra"
  gridFornNivel01.TextMatrix(0, 2) = "Qtde da compra"
  
  Data_Fim.Text = Format(Now, "dd/mm/yyyy")
  Data_Ini.Text = Format(Now - 90, "dd/mm/yyyy")
  cmb_numForn.ListIndex = 5
  SSTab1.Tab = 0
  
  Combo_Filial.ListIndex = 0
  Combo_Filial_LostFocus
  
End Sub

Private Sub gridForn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridForn.Redraw = False
End Sub

Private Sub gridForn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridForn.RowSel = gridForn.Row
  gridForn.Redraw = True
End Sub


Private Sub sAbas_Click()

End Sub
