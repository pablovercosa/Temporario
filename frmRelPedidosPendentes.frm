VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelPedidosPendentes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Controle de Entregas"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frmRelPedidosPendentes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   8520
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
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
      Height          =   465
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   8445
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   8445
      Begin VB.Frame Frame1 
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2880
         TabIndex        =   9
         Top             =   300
         Width           =   5385
         Begin MSMask.MaskEdBox txtDataFinal 
            Height          =   315
            Left            =   3540
            TabIndex        =   4
            Top             =   300
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
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
         Begin MSMask.MaskEdBox txtDataInicial 
            Height          =   315
            Left            =   780
            TabIndex        =   3
            Top             =   300
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
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
         Begin VB.Label Label3 
            Caption         =   "Inicial"
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
            Left            =   240
            TabIndex        =   11
            Top             =   330
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Final"
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
            Left            =   3090
            TabIndex        =   10
            Top             =   330
            Width           =   375
         End
      End
      Begin VB.TextBox txtSequencia 
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
         Height          =   375
         Left            =   150
         TabIndex        =   0
         Top             =   510
         Width           =   2415
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   375
         Left            =   150
         TabIndex        =   1
         Top             =   1230
         Width           =   2415
      End
      Begin VB.TextBox txtProduto 
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
         Height          =   375
         Left            =   150
         TabIndex        =   2
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Sequência"
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
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
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
         Left            =   150
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Produto"
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
         Left            =   150
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   300
      Top             =   2910
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRelPedidosPendentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'08/02/2006 - mpdea
'Incluído tratamento de erro
'Modificado comparações a string "__/__/____" para "  /  /    "
Private Sub cmdImprimir_Click()
  Dim sSelection As String
  
  On Error GoTo ErrHandler
  
  If (txtDataInicial.Text <> "  /  /    " And _
      txtDataFinal.Text <> "  /  /    ") Then
    
    If (Not (IsDate(txtDataInicial.Text) _
        Or (IsDate(txtDataFinal.Text)))) Or _
       (CDate(txtDataInicial.Text) > _
        CDate(txtDataFinal.Text)) Then
      MsgBox "Data inválida, verifique "
      Exit Sub
    End If
    
  End If
  
  
  '08/02/2006 - mpdea
  'Corrigido referência ao diretório padrão da base de dados
  CR1.DataFiles(0) = gsQuickDBFileName
  CR1.DataFiles(1) = gsQuickDBFileName
  CR1.DataFiles(2) = gsQuickDBFileName
  CR1.DataFiles(3) = gsQuickDBFileName
  CR1.DataFiles(4) = gsQuickDBFileName
  
  '08/02/2006 - mpdea
  'Corrigido referência ao diretório padrão de relatórios
  CR1.ReportFileName = gsReportPath & "PedidosPendentes.rpt"
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 CR1
  
  '18/06/2007 - Anderson
  'Alterado para exibir o relatório de entregas
  'sSelection = " {Saídas - Produtos.Qtde} > {Saídas - Produtos.QtdeEntregue} "
  
  '08/02/2006 - mpdea
  'Corrigido seleção de operações com controle de entrega
  sSelection = sSelection & " {Operações Saída.ControleEntregas} = TRUE "
  
  If Len(txtSequencia.Text) > 0 Then
    sSelection = sSelection & " AND {Saídas - Produtos.Sequência} = " & Val(txtSequencia.Text)
  End If
  
  If Len(txtProduto.Text) > 0 Then
    sSelection = sSelection & " AND {Saídas - Produtos.Código} = '" & txtProduto.Text & "' "
  End If
  
  If Len(txtCliente.Text) > 0 Then
    sSelection = sSelection & " AND {Saídas.Cliente} = " & txtCliente.Text
  End If
  
  If (txtDataInicial.Text <> "  /  /    " And _
      txtDataFinal.Text <> "  /  /    ") Then
    sSelection = sSelection & _
                 " AND {Saídas.Data} >= date" & Format$(txtDataInicial.Text, "(yyyy,mm,dd)") & _
                 " AND {Saídas.Data} <= date" & Format$(txtDataFinal.Text, "(yyyy,mm,dd)")
  End If
  
  CR1.SelectionFormula = sSelection
  
  CR1.Destination = crptToWindow
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", CR1)
  
  CR1.Action = 1
  
  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

Private Sub txtDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    txtDataFinal.Text = frmCalendario.gsDateCalender(txtDataFinal.Text)
  End If
End Sub

Private Sub txtDataFinal_LostFocus()
  txtDataFinal.Text = Ajusta_Data(txtDataFinal.Text)
End Sub

Private Sub txtDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    txtDataInicial.Text = frmCalendario.gsDateCalender(txtDataInicial.Text)
  End If
End Sub

Private Sub txtDataInicial_LostFocus()
  txtDataInicial.Text = Ajusta_Data(txtDataInicial.Text)
End Sub

Private Sub txtSequencia_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub
