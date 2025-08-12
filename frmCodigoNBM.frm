VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCodigoNBM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Manutenção de Códigos NCM"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCodigoNBM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   15300
   Begin VB.Frame fraX 
      Height          =   8130
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15225
      Begin VB.ComboBox cmb_FCP 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmCodigoNBM.frx":4E95A
         Left            =   13950
         List            =   "frmCodigoNBM.frx":4E964
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   7298
         Width           =   1215
      End
      Begin VB.TextBox txt_CEST 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   11640
         MaxLength       =   8
         TabIndex        =   21
         ToolTipText     =   "O Máximo de caracteres para o campo ""Código"" é de oito (8)"
         Top             =   7290
         Width           =   1845
      End
      Begin VB.ComboBox cmb_opcoes 
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
         ItemData        =   "frmCodigoNBM.frx":4E972
         Left            =   1620
         List            =   "frmCodigoNBM.frx":4E997
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   172
         Width           =   5775
      End
      Begin VB.CommandButton cmd_pesquisar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pesquisar NCM"
         Height          =   405
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   15030
      End
      Begin VB.TextBox txt_codigoNCM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   9150
         MaxLength       =   8
         TabIndex        =   15
         ToolTipText     =   "O Máximo de caracteres para o campo ""Código"" é de oito (8)"
         Top             =   165
         Width           =   1515
      End
      Begin MSFlexGridLib.MSFlexGrid gridNCM 
         Height          =   5340
         Left            =   135
         TabIndex        =   14
         Top             =   1035
         Width           =   15030
         _ExtentX        =   26511
         _ExtentY        =   9419
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16777215
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
      Begin VB.TextBox txb_tabela 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6135
         MaxLength       =   8
         TabIndex        =   5
         ToolTipText     =   "O Máximo de caracteres para o campo ""Código"" é de oito (8)"
         Top             =   7290
         Width           =   1845
      End
      Begin VB.TextBox txb_ex 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   8130
         MaxLength       =   8
         TabIndex        =   6
         ToolTipText     =   "O Máximo de caracteres para o campo ""Código"" é de oito (8)"
         Top             =   7290
         Width           =   1845
      End
      Begin VB.TextBox txb_aliqImportacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   4
         ToolTipText     =   "O Máximo de caracteres para o campo ""Código"" é de oito (8)"
         Top             =   7290
         Width           =   1845
      End
      Begin VB.TextBox txb_aliqNacional 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   135
         MaxLength       =   8
         TabIndex        =   3
         ToolTipText     =   "O Máximo de caracteres para o campo ""Código"" é de oito (8)"
         Top             =   7290
         Width           =   1845
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4080
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "O Máximo de caracteres para o campo ""Nome"" é de cem (100)"
         Top             =   6645
         Width           =   11085
      End
      Begin VB.TextBox Codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   135
         MaxLength       =   8
         TabIndex        =   1
         ToolTipText     =   "O Máximo de caracteres para o campo ""Código"" é de oito (8)"
         Top             =   6660
         Width           =   1845
      End
      Begin VB.Label Label10 
         Caption         =   "Tem FCP"
         Height          =   210
         Left            =   13950
         TabIndex        =   23
         Top             =   7050
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Código CEST"
         Height          =   240
         Left            =   11640
         TabIndex        =   22
         Top             =   7035
         Width           =   1485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Opção de pesquisa"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl_aguardar 
         BackColor       =   &H00C0C0FF&
         Caption         =   ">>>    Aguarde a pesquisa terminar...."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11760
         TabIndex        =   18
         Top             =   180
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.Label lbl_opcaoPesq 
         AutoSize        =   -1  'True
         Caption         =   "Parte do nome NCM"
         Height          =   195
         Left            =   7635
         TabIndex        =   16
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label7 
         Caption         =   "Tabela"
         Height          =   240
         Left            =   6150
         TabIndex        =   13
         Top             =   7035
         Width           =   600
      End
      Begin VB.Label Label6 
         Caption         =   "EX"
         Height          =   240
         Left            =   8175
         TabIndex        =   12
         Top             =   7035
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Aliquota de Importação"
         Height          =   195
         Left            =   4080
         TabIndex        =   11
         Top             =   7065
         Width           =   2010
      End
      Begin VB.Label Label4 
         Caption         =   "Aliquota Nacional"
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   7035
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ex: 01062000"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   6705
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   6435
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   4080
         TabIndex        =   7
         Top             =   6420
         Width           =   405
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
         Left            =   11340
         Top             =   7785
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
         Bands           =   "frmCodigoNBM.frx":4EA7C
      End
   End
End
Attribute VB_Name = "frmCodigoNBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Este Cadastro foi criado em 20/07/2005 por Daniel R. Rodrigues
'para atender a necessidade das empresas contribuintes de IPI
Dim Num_Registro As Variant

Private Sub cmb_opcoes_Click()
  txt_codigoNCM.Text = ""
  
  If cmb_opcoes.ListIndex = 0 Then
      txt_codigoNCM.Visible = True
      lbl_opcaoPesq.Visible = True
      lbl_opcaoPesq.Caption = "Número NCM"
  ElseIf cmb_opcoes.ListIndex = 1 Then
      txt_codigoNCM.Visible = True
      lbl_opcaoPesq.Visible = True
      lbl_opcaoPesq.Caption = "Parte do nome NCM"
  Else
      txt_codigoNCM.Visible = False
      lbl_opcaoPesq.Visible = False
  End If
End Sub

Private Sub cmd_pesquisar_Click()
On Error GoTo Processa_Erro

  Dim rstCodigoNBM As Recordset
  Dim sSql As String
  Dim sFCP As String

  'Número NCM
  'Parte do nome NCM
  'NCM começando com 0 ou 1
  'NCM começando com 2
  'NCM começando com 3
  'NCM começando com 4
  'NCM começando com 5
  'NCM começando com 6
  'NCM começando com 7
  'NCM começando com 8
  'NCM começando com 9
  
  If cmb_opcoes.Text = "" Then
      MsgBox "Selecione uma opção de pesquisa", vbInformation, "Atenção"
      cmb_opcoes.SetFocus
      Exit Sub
  End If
  
  If cmb_opcoes.ListIndex = 0 Then
      
      If Trim(txt_codigoNCM.Text) = "" Then
          MsgBox "Digite o Número NCM que deseja pesquisar", vbInformation, "Atenção"
          Exit Sub
      End If
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      gridNCM.Rows = 1
    
      sSql = "SELECT * FROM AliquotasNCM "
      sSql = sSql & " Where Codigo ='" & txt_codigoNCM.Text & "' "
      
      Set rstCodigoNBM = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
    
      If Not (rstCodigoNBM.EOF And rstCodigoNBM.BOF) Then
        With rstCodigoNBM
          .MoveFirst
          While Not .EOF
          
            If Not IsNull(.Fields("TemFCP").Value) Then
                If .Fields("TemFCP").Value = True Then
                    sFCP = "SIM"
                Else
                    sFCP = ""
                End If
            Else
                sFCP = ""
            End If
          
            gridNCM.AddItem vbTab & .Fields("Codigo").Value & vbTab & _
                        .Fields("Nome").Value & vbTab & _
                        .Fields("AliqNacional").Value & vbTab & _
                        .Fields("AliqImportacao").Value & vbTab & _
                        .Fields("Tabela").Value & vbTab & _
                        .Fields("Ex").Value & vbTab & _
                        .Fields("CEST").Value & vbTab & _
                        sFCP

            .MoveNext
          Wend
          rstCodigoNBM.Close
          Set rstCodigoNBM = Nothing
          
          Me.Show
          gbPodeGravar = True
          Call ActiveBarLoadToolTips(Me)
          Call ClearScreen
        End With
      End If
      
      lbl_aguardar.Visible = False
  
  ElseIf cmb_opcoes.ListIndex = 1 Then
      If Trim(txt_codigoNCM.Text) = "" Then
          MsgBox "Digite parte do nome do NCM que deseja pesquisar", vbInformation, "Atenção"
          Exit Sub
      End If
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      gridNCM.Rows = 1
    
      sSql = "SELECT * FROM AliquotasNCM "
      sSql = sSql & " Where Nome like '*" & txt_codigoNCM.Text & "*' "
      
      Set rstCodigoNBM = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
    
      If Not (rstCodigoNBM.EOF And rstCodigoNBM.BOF) Then
        With rstCodigoNBM
          .MoveFirst
          While Not .EOF
          
            If Not IsNull(.Fields("TemFCP").Value) Then
                If .Fields("TemFCP").Value = True Then
                    sFCP = "SIM"
                Else
                    sFCP = ""
                End If
            Else
                sFCP = ""
            End If
          
            gridNCM.AddItem vbTab & .Fields("Codigo").Value & vbTab & _
                        .Fields("Nome").Value & vbTab & _
                        .Fields("AliqNacional").Value & vbTab & _
                        .Fields("AliqImportacao").Value & vbTab & _
                        .Fields("Tabela").Value & vbTab & _
                        .Fields("Ex").Value & vbTab & _
                        .Fields("CEST").Value & vbTab & _
                        sFCP
                     
            .MoveNext
          Wend
          rstCodigoNBM.Close
          Set rstCodigoNBM = Nothing
          
          Me.Show
          gbPodeGravar = True
          Call ActiveBarLoadToolTips(Me)
          Call ClearScreen
        End With
      End If
      
      lbl_aguardar.Visible = False
  ElseIf cmb_opcoes.ListIndex = 2 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 1
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  ElseIf cmb_opcoes.ListIndex = 3 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 2
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  ElseIf cmb_opcoes.ListIndex = 4 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 3
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  ElseIf cmb_opcoes.ListIndex = 5 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 4
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  ElseIf cmb_opcoes.ListIndex = 6 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 5
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  ElseIf cmb_opcoes.ListIndex = 7 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 6
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  ElseIf cmb_opcoes.ListIndex = 8 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 7
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  ElseIf cmb_opcoes.ListIndex = 9 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 8
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  ElseIf cmb_opcoes.ListIndex = 10 Then
      gridNCM.Rows = 1
      
      DoEvents
      lbl_aguardar.Visible = True
      DoEvents
      cmd_pesquisar.Enabled = False
      pesquisaNCM 9
      lbl_aguardar.Visible = False
      cmd_pesquisar.Enabled = True
  End If
  
  Exit Sub
Processa_Erro:
  MsgBox "Erro na pesquisa." & Err.Number & " " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
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

Private Sub pesquisaNCM(iRange As Integer)
On Error GoTo Processa_Erro
  Dim rstCodigoNBM As Recordset
  Dim sSql As String
  Dim sFCP As String

  sSql = "SELECT * FROM AliquotasNCM "
  
  If iRange = 1 Then
      sSql = sSql & " Where Codigo > '0' and Codigo < '20000000'"
  ElseIf iRange = 2 Then
      sSql = sSql & " Where Codigo > '19999999' and Codigo < '30000000'"
  ElseIf iRange = 3 Then
      sSql = sSql & " Where Codigo > '29999999' and Codigo < '40000000'"
  ElseIf iRange = 4 Then
      sSql = sSql & " Where Codigo > '39999999' and Codigo < '50000000'"
  ElseIf iRange = 5 Then
      sSql = sSql & " Where Codigo > '49999999' and Codigo < '60000000'"
  ElseIf iRange = 6 Then
      sSql = sSql & " Where Codigo > '59999999' and Codigo < '70000000'"
  ElseIf iRange = 7 Then
      sSql = sSql & " Where Codigo > '69999999' and Codigo < '80000000'"
  ElseIf iRange = 8 Then
      sSql = sSql & " Where Codigo > '79999999' and Codigo < '90000000'"
  ElseIf iRange = 9 Then
      sSql = sSql & " Where Codigo > '89999999'"
  End If
  
  sSql = sSql & " ORDER BY Codigo"
  
  Set rstCodigoNBM = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)

  If Not (rstCodigoNBM.EOF And rstCodigoNBM.BOF) Then
    With rstCodigoNBM
      .MoveFirst
      While Not .EOF
      
        If Not IsNull(.Fields("TemFCP").Value) Then
            If .Fields("TemFCP").Value = True Then
                sFCP = "SIM"
            Else
                sFCP = ""
            End If
        Else
            sFCP = ""
        End If
            
        gridNCM.AddItem vbTab & .Fields("Codigo").Value & vbTab & _
                    .Fields("Nome").Value & vbTab & _
                    .Fields("AliqNacional").Value & vbTab & _
                    .Fields("AliqImportacao").Value & vbTab & _
                    .Fields("Tabela").Value & vbTab & _
                    .Fields("Ex").Value & vbTab & _
                    .Fields("CEST").Value & vbTab & _
                    sFCP
                 
        .MoveNext
      Wend
      rstCodigoNBM.Close
      Set rstCodigoNBM = Nothing
      
      Me.Show
      gbPodeGravar = True
      Call ActiveBarLoadToolTips(Me)
      Call ClearScreen
    End With
  End If
  
  Exit Sub
Processa_Erro:
  MsgBox "Erro na pesquisa." & Err.Number & " " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Load()
On Error GoTo Processa_Erro

  Call CenterForm(Me)
  
  gridNCM.ColWidth(0) = 0
  gridNCM.ColWidth(1) = 1100
  gridNCM.ColWidth(2) = 7320
  gridNCM.ColWidth(3) = 1000
  gridNCM.ColWidth(4) = 1000
  gridNCM.ColWidth(5) = 1100
  gridNCM.ColWidth(6) = 1000
  gridNCM.ColWidth(7) = 1200
  gridNCM.ColWidth(8) = 1000
  
  gridNCM.Row = 0
  gridNCM.TextMatrix(0, 1) = "Código NCM"
  gridNCM.TextMatrix(0, 2) = "Nome"
  gridNCM.TextMatrix(0, 3) = "Aliq.Nac."
  gridNCM.TextMatrix(0, 4) = "Aliq.Imp."
  gridNCM.TextMatrix(0, 5) = "Tabela"
  gridNCM.TextMatrix(0, 6) = "Ex"
  gridNCM.TextMatrix(0, 7) = "Código CEST"
  gridNCM.TextMatrix(0, 8) = "Tem FCP"
  
  cmb_opcoes.ListIndex = 1
  
  Exit Sub
Processa_Erro:
  MsgBox "Erro na carga da tela." & Err.Number & " " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub gridNCM_Click()
  If gridNCM.RowSel > 0 Then
      Codigo.Text = gridNCM.TextMatrix(gridNCM.RowSel, 1)
      txtNome.Text = gridNCM.TextMatrix(gridNCM.RowSel, 2)
      txb_aliqNacional.Text = gridNCM.TextMatrix(gridNCM.RowSel, 3)
      txb_aliqImportacao.Text = gridNCM.TextMatrix(gridNCM.RowSel, 4)
      txb_tabela.Text = gridNCM.TextMatrix(gridNCM.RowSel, 5)
      txb_ex.Text = gridNCM.TextMatrix(gridNCM.RowSel, 6)
      txt_CEST.Text = gridNCM.TextMatrix(gridNCM.RowSel, 7)
      
      If gridNCM.TextMatrix(gridNCM.RowSel, 8) = "SIM" Then
          cmb_FCP.ListIndex = 0
      Else
          cmb_FCP.ListIndex = 1
      End If
  End If
End Sub

Private Sub gridNCM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridNCM.Redraw = False
End Sub

Private Sub gridNCM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridNCM.RowSel = gridNCM.Row
  gridNCM.Redraw = True
End Sub

Private Sub txb_aliqNacional_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(txb_aliqNacional)
End Sub

Private Sub txb_aliqImportacao_Validate(Cancel As Boolean)
  Call FormatCurrencyValue(txb_aliqImportacao)
End Sub

Private Sub UpdateRecord()
  Dim tipoAcao As String
  Dim strSQL As String
  Dim iFCP As Integer
    
  On Error GoTo Processa_Erro
  
  'Validação do Código
  If Len(RTrim(RTrim(Codigo.Text))) <= 0 Then
    MsgBox "Código inválido, verifique.", vbExclamation, "Quick Store"
    Codigo.SetFocus
    Exit Sub
  End If
  
  Dim rstCodigoNCM As Recordset
  Set rstCodigoNCM = db.OpenRecordset("SELECT * FROM AliquotasNCM Where Codigo='" & Codigo.Text & "' ", dbOpenDynaset, dbReadOnly)
  
  If rstCodigoNCM.EOF And rstCodigoNCM.BOF Then
    'NCM novo - insert
    tipoAcao = "insert"
  Else
    'NCM existe - update
    tipoAcao = "update"
  End If
  
  rstCodigoNCM.Close
  Set rstCodigoNCM = Nothing

  'Validação do Nome
  If Len(txtNome.Text) <= 0 Then
    MsgBox "Nome inválido, verifique.", vbExclamation, "Quick Store"
    txtNome.SetFocus
    Exit Sub
  End If
  
  If Len(txb_aliqNacional.Text) <= 0 Then
    MsgBox "Aliquota Nacional inválida, verifique.", vbExclamation, "Quick Store"
    txb_aliqNacional.SetFocus
    Exit Sub
  End If

  If Len(txb_aliqImportacao.Text) <= 0 Then
    txb_aliqImportacao.Text = "0"
  End If
  
  If Len(txb_tabela.Text) <= 0 Then
    txb_tabela.Text = "0"
  End If
  
  If Len(txb_ex.Text) <= 0 Then
    txb_ex.Text = "0"
  End If
  
  If cmb_FCP.ListIndex = 0 Then
      ' SIM...este NCM tem FCP
      iFCP = 1
  Else
      iFCP = 0
  End If
  
  '-----------------------------
  ' Gravando...
  '-----------------------------
  Call StatusMsg("Gravando ...")
  DoEvents
  
  If tipoAcao = "insert" Then
    txb_aliqNacional.Text = Replace(txb_aliqNacional.Text, ",", ".")
    txb_aliqImportacao.Text = Replace(txb_aliqImportacao.Text, ",", ".")
    strSQL = "INSERT INTO AliquotasNCM (Codigo, Ex, Tabela, AliqNacional, AliqImportacao, Nome, CEST, TemFCP)  "
    strSQL = strSQL & " Values ('" & LTrim(RTrim(Codigo.Text)) & "','" & LTrim(RTrim(txb_ex.Text)) & "',"
    strSQL = strSQL & txb_tabela.Text & "," & txb_aliqNacional.Text & "," & txb_aliqImportacao.Text & ",'" & txtNome.Text & "','" & txt_CEST.Text & "'," & iFCP & ") "
  
    db.Execute strSQL
    
    gridNCM.AddItem vbTab & LTrim(RTrim(Codigo.Text)) & vbTab & _
                txtNome.Text & vbTab & _
                txb_aliqNacional.Text & vbTab & _
                txb_aliqImportacao.Text & vbTab & _
                txb_tabela.Text & vbTab & _
                txb_ex.Text & vbTab & _
                txt_CEST.Text
  Else '- update
    txb_aliqNacional.Text = Replace(txb_aliqNacional.Text, ",", ".")
    txb_aliqImportacao.Text = Replace(txb_aliqImportacao.Text, ",", ".")
    strSQL = "UPDATE AliquotasNCM SET Ex='" & LTrim(RTrim(txb_ex.Text)) & "',"
    strSQL = strSQL & " Tabela=" & txb_tabela.Text & ", AliqNacional=" & txb_aliqNacional.Text & ","
    strSQL = strSQL & " AliqImportacao=" & txb_aliqImportacao.Text & ", Nome='" & txtNome.Text & "', CEST ='" & txt_CEST.Text & "', "
    strSQL = strSQL & " TemFCP=" & iFCP
    strSQL = strSQL & " WHERE Codigo='" & LTrim(RTrim(Codigo.Text)) & "' "
  
    db.Execute strSQL
  
    gridNCM.TextMatrix(gridNCM.RowSel, 2) = txtNome.Text
    gridNCM.TextMatrix(gridNCM.RowSel, 3) = txb_aliqNacional.Text
    gridNCM.TextMatrix(gridNCM.RowSel, 4) = txb_aliqImportacao.Text
    gridNCM.TextMatrix(gridNCM.RowSel, 5) = txb_tabela.Text
    gridNCM.TextMatrix(gridNCM.RowSel, 6) = txb_ex.Text
    gridNCM.TextMatrix(gridNCM.RowSel, 7) = txt_CEST.Text
    
    If iFCP = 1 Then
        gridNCM.TextMatrix(gridNCM.RowSel, 8) = "SIM"
    Else
        gridNCM.TextMatrix(gridNCM.RowSel, 8) = ""
    End If
    
    ' Update da aliquota no cadastro de produtos que possuem este NCM
    Dim sAliqNac As String
    sAliqNac = txb_aliqNacional.Text
    sAliqNac = Replace(sAliqNac, ",", ".")
    strSQL = "Update Produtos set AliqNCM = " + sAliqNac + " Where CodigoNBM = '" + LTrim(RTrim(Codigo.Text)) + "' "
    db.Execute strSQL
  End If
  
  gridNCM.Refresh

  MsgBox "NCM gravado com sucesso!", vbInformation, "Sucesso"
  
  Call ClearScreen
  Call StatusMsg("")
  
  Exit Sub
  
Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar gravar registro."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & " - " & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  
  Codigo.Text = ""
  txtNome.Text = ""
  txb_aliqNacional.Text = "0"
  txb_aliqImportacao.Text = "0"
  txb_ex.Text = "0"
  txb_tabela.Text = "0"
  cmb_FCP.ListIndex = -1
  
  Codigo.SetFocus
  Num_Registro = Null
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
  End Select
End Sub

Private Sub Código_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Function ExisteProdutos(ByVal CodProd As String) As Boolean
  Dim rstProdutos As Recordset
  
  Set rstProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE CodigoNBM = '" & CodProd & "'", dbOpenDynaset)

  If rstProdutos.RecordCount > 0 Then ExisteProdutos = True

  rstProdutos.Close
  Set rstProdutos = Nothing

End Function
