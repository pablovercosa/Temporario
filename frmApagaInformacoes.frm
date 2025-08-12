VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmApagaInformacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apaga informações"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmApagaInformacoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B.SSDBCombo cboCodigoCliente 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1335
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   2355
      _ExtentY        =   503
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4455
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdApagar 
         Caption         =   "Apagar"
         Default         =   -1  'True
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDataInicial 
         Height          =   255
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Data final"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Data inicial"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label lblNomeCliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmApagaInformacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCodigoCliente_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cmdApagar_Click()
  Dim sSql As String
  
  If Not IsDate(mskDataInicial.Text) Then
    MsgBox "Data inicial inválida, verifique . . . ", vbCritical, "Quick Store"
    mskDataInicial.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(mskDataFinal.Text) Then
    MsgBox "Data final inválida, verifique . . . ", vbCritical, "Quick Store"
    mskDataFinal.SetFocus
    Exit Sub
  Else
    If CDate(mskDataInicial.Text) > CDate(mskDataFinal.Text) Then
      MsgBox "A data inicial não pode ser maior que a data final, verifique . . . ", vbCritical, "Quick Store"
      mskDataInicial.SetFocus
      Exit Sub
    End If
  End If

  If MsgBox("ATENÇÃO" & vbCrLf & vbCrLf & _
           "As informações sobre compras/vendas no período indicado serão apagadas definitivamente" & vbCrLf & vbCrLf & _
           "Deseja prosseguir ?", vbQuestion + vbYesNo, "Quick Store") = vbYes Then
    sSql = "DELETE * FROM [Resumo Clientes] WHERE Dia >= #" & Format(mskDataInicial.Text, "MM/dd/yyyy") & "#" & _
                                            " AND Dia <= #" & Format(mskDataFinal.Text, "MM/dd/yyyy") & "#"
    
    If cboCodigoCliente.Text <> "0" And cboCodigoCliente.Text <> "" Then
      sSql = sSql & " AND Cliente = " & Trim(cboCodigoCliente.Text)
    End If
    
    If frmGerente.gbSenhaGerente Then
      db.Execute sSql
    End If
    
    Call frmInformacoes.Monta_Grade_Compras
    Unload Me
  End If
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Dim rsClientes As Recordset
  Dim nCodigoCliente As Long: nCodigoCliente = frmCliFor.cboCodigo.Text
  
  Set rsClientes = db.OpenRecordset("SELECT Código, Nome FROM Cli_For WHERE Código = " & nCodigoCliente, dbOpenSnapshot)
  
  With rsClientes
    cboCodigoCliente.Text = nCodigoCliente
    lblNomeCliente.Caption = rsClientes!Nome
  End With
End Sub

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      mskDataFinal.Text = frmCalendario.gsDateCalender(mskDataFinal.Text)
  End Select
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal.Text = Ajusta_Data(mskDataFinal.Text)
End Sub

Private Sub mskDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      mskDataInicial.Text = frmCalendario.gsDateCalender(mskDataInicial.Text)
  End Select
End Sub

Private Sub mskDataInicial_LostFocus()
  mskDataInicial.Text = Ajusta_Data(mskDataInicial.Text)
End Sub
