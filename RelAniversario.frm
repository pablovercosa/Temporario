VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRelAniver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Contatos por Data de Aniversário"
   ClientHeight    =   5235
   ClientLeft      =   1890
   ClientTop       =   1755
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelAniversario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5235
   ScaleWidth      =   11430
   Begin VB.Frame frm_dataAniversarioClientes 
      Caption         =   "Data de Aniversário dos Clientes"
      Height          =   3675
      Left            =   7260
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmd_imprimirGrade 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Imprimir"
         Height          =   435
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3180
         Visible         =   0   'False
         Width           =   11160
      End
      Begin VB.CommandButton cmd_PesquisarCli 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pesquisar"
         Height          =   435
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   11160
      End
      Begin MSFlexGridLib.MSFlexGrid gridAniversarioCli 
         Height          =   2400
         Left            =   60
         TabIndex        =   12
         Top             =   720
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   4233
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   8454143
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16250871
         AllowBigSelection=   0   'False
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
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   60
      TabIndex        =   13
      Top             =   780
      Width           =   11295
      Begin VB.TextBox Dia_Ini 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1830
         MaxLength       =   2
         TabIndex        =   16
         Top             =   217
         Width           =   405
      End
      Begin VB.TextBox Dia_Fim 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   2730
         MaxLength       =   2
         TabIndex        =   15
         Top             =   217
         Width           =   405
      End
      Begin VB.ComboBox Lista 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "RelAniversario.frx":4E95A
         Left            =   5100
         List            =   "RelAniversario.frx":4E982
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   210
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Do dia"
         Height          =   255
         Left            =   1290
         TabIndex        =   19
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "até"
         Height          =   255
         Left            =   2385
         TabIndex        =   18
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label3 
         Caption         =   "Mês"
         Height          =   225
         Left            =   4710
         TabIndex        =   17
         Top             =   255
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Relatório"
      Height          =   675
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   11295
      Begin VB.OptionButton opt_cliente 
         Caption         =   "Clientes"
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton opt_listaDeContatosCliFor 
         Caption         =   "Lista de Contatos"
         Height          =   255
         Left            =   1260
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.Frame frm_listaContatos 
      Caption         =   "Todos os contatos dos Clientes/Fornecedores"
      Height          =   3675
      Left            =   60
      TabIndex        =   0
      Top             =   1500
      Width           =   11295
      Begin VB.Frame Frame1 
         Caption         =   "Saída"
         Height          =   705
         Left            =   5700
         TabIndex        =   5
         Top             =   300
         Width           =   5415
         Begin VB.OptionButton B_Vídeo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Vídeo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   450
            TabIndex        =   7
            Top             =   270
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton B_Impressora 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Impressora"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1710
            TabIndex        =   6
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.CommandButton B_Imprime 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Imprimir"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3090
         Width           =   11055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ordem"
         Height          =   705
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   5415
         Begin VB.OptionButton optEmpresa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Empresa"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1740
            TabIndex        =   3
            Top             =   300
            Width           =   975
         End
         Begin VB.OptionButton optDia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Dia"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   480
            TabIndex        =   2
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin Crystal.CrystalReport Rel1 
         Left            =   8910
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
End
Attribute VB_Name = "frmRelAniver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Imprime_Click()
  Dim Aux_Mes As String
  Dim Str1 As String
  Dim Str_Rel As String

  If IsNull(Dia_Ini.Text) Then Dia_Ini.Text = 0
  If Dia_Ini.Text = "" Then Dia_Ini.Text = 0
  If Val(Dia_Ini.Text) < 1 Or Val(Dia_Ini.Text) > 31 Then
    DisplayMsg "Digite data inicial entre 1 e 31."
    Dia_Ini.SetFocus
    Exit Sub
  End If
  
  If IsNull(Dia_Fim.Text) Then Dia_Fim.Text = 0
  If Dia_Fim.Text = "" Then Dia_Fim.Text = 0
  If Val(Dia_Fim.Text) < 1 Or Val(Dia_Fim.Text) > 31 Then
    DisplayMsg "Digite data final entre 1 e 31."
    Dia_Fim.SetFocus
    Exit Sub
  End If
  
  If Val(Dia_Fim.Text) < Val(Dia_Ini.Text) Then
    DisplayMsg "Dia final deve ser menor ou igual ao dia inicial."
    Dia_Ini.SetFocus
    Exit Sub
  End If
  
  
  If Lista.Text = "" Then
    DisplayMsg "Escolha o mês."
    Lista.SetFocus
    Exit Sub
  End If
  
  Aux_Mes = UCase(Left(Lista.Text, 3))
  
   Rem  Seta Valores e Manda Relatório

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel1.DataFiles(0) = Str1

 Rem Saída
 If B_Vídeo = True Then Rel1.Destination = 0
 If B_Impressora = True Then Rel1.Destination = 1

 Rem Nome do arquivo .rpt
  Str1 = gsReportPath & "MALA3.RPT"
 
 Rel1.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 Rem Seleção
 Str_Rel = "{Contatos.Dia Aniversário} >=" + Dia_Ini.Text
 Str_Rel = Str_Rel + " And {Contatos.Dia Aniversário} <=" + Dia_Fim.Text
 Str_Rel = Str_Rel + " And {Contatos.Mês Aniversário} = '" + Aux_Mes + "'"

 Rel1.SelectionFormula = Str_Rel
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel1.Formulas(0) = Str_Rel

 Str_Rel = "dia_ini = '"
 Str_Rel = Str_Rel + Dia_Ini.Text + "'"
 Rel1.Formulas(1) = Str_Rel
 
 Str_Rel = "dia_fim = '"
 Str_Rel = Str_Rel + Dia_Fim.Text + "'"
 Rel1.Formulas(2) = Str_Rel

 Rem mes
 Str_Rel = "mes = '"
 Str_Rel = Str_Rel + Lista.Text + "'"
 Rel1.Formulas(3) = Str_Rel

  If optDia.Value Then
    Rel1.SortFields(0) = "+{Contatos.Dia Aniversário}"
  Else
    Rel1.SortFields(0) = "+{Cli_For.Nome}"
  End If

 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  

 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

  
End Sub

Private Sub cmd_PesquisarCli_Click()
On Error GoTo Erro
  Dim rsCli As Recordset
  Dim strSQL As String
  Dim iMes As Integer
  Dim sFones As String
  Dim sEndereco As String
  Dim sEmail As String
  Dim sUF As String
  Dim sCidade As String
 
  If IsNull(Dia_Ini.Text) Then Dia_Ini.Text = 0
  If Dia_Ini.Text = "" Then Dia_Ini.Text = 0
  If Val(Dia_Ini.Text) < 1 Or Val(Dia_Ini.Text) > 31 Then
    DisplayMsg "Digite data inicial entre 1 e 31."
    Dia_Ini.SetFocus
    Exit Sub
  End If
  
  If IsNull(Dia_Fim.Text) Then Dia_Fim.Text = 0
  If Dia_Fim.Text = "" Then Dia_Fim.Text = 0
  If Val(Dia_Fim.Text) < 1 Or Val(Dia_Fim.Text) > 31 Then
    DisplayMsg "Digite data final entre 1 e 31."
    Dia_Fim.SetFocus
    Exit Sub
  End If
  
  If Val(Dia_Fim.Text) < Val(Dia_Ini.Text) Then
    DisplayMsg "Dia final deve ser menor ou igual ao dia inicial."
    Dia_Ini.SetFocus
    Exit Sub
  End If
  
  If Lista.Text = "" Then
    DisplayMsg "Escolha o mês."
    Lista.SetFocus
    Exit Sub
  End If
  
  iMes = Lista.ListIndex + 1
  
  gridAniversarioCli.Rows = 1
  
  strSQL = " Select mid(DataNascimento,1,2), * from Cli_For Order by 1,Nome "
  Set rsCli = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  While Not rsCli.EOF
  
      sEmail = ""
      sFones = ""
      sUF = ""
      sCidade = ""
      sEndereco = ""
  
      If Not IsNull(rsCli.Fields("DataNascimento").Value) Then
          If Len(rsCli.Fields("DataNascimento").Value) = 10 Then
              If CInt(Mid(rsCli.Fields("DataNascimento").Value, 1, 2)) >= CInt(Dia_Ini.Text) And _
                  CInt(Mid(rsCli.Fields("DataNascimento").Value, 1, 2)) <= CInt(Dia_Fim.Text) And _
                  CInt(Mid(rsCli.Fields("DataNascimento").Value, 4, 2)) = iMes Then
                  
                  If Not IsNull(rsCli.Fields("email").Value) Then
                      sEmail = rsCli.Fields("email").Value
                  End If
                  
                  If Not IsNull(rsCli.Fields("Fone 1").Value) Then
                      sFones = rsCli.Fields("Fone 1").Value
                  End If
                  If Not IsNull(rsCli.Fields("Fone 2").Value) Then
                      sFones = sFones & rsCli.Fields("Fone 2").Value
                  End If
                  
                  If Not IsNull(rsCli.Fields("Estado").Value) Then
                      sUF = rsCli.Fields("Estado").Value
                  End If

                  If Not IsNull(rsCli.Fields("Cidade").Value) Then
                      sCidade = rsCli.Fields("Cidade").Value
                  End If

                  If Not IsNull(rsCli.Fields("Endereço").Value) Then
                      sEndereco = rsCli.Fields("Endereço").Value
                  End If
                  If Not IsNull(rsCli.Fields("Endereço Número").Value) Then
                      sEndereco = sEndereco & " " & rsCli.Fields("Endereço Número").Value
                  End If
                  If Not IsNull(rsCli.Fields("Bairro").Value) Then
                      sEndereco = sEndereco & " " & rsCli.Fields("Bairro").Value
                  End If
                  If Not IsNull(rsCli.Fields("Cep").Value) Then
                      sEndereco = sEndereco & " " & rsCli.Fields("Cep").Value
                  End If
                  
                  gridAniversarioCli.AddItem vbTab & rsCli.Fields("Código").Value & vbTab & _
                      rsCli.Fields("Nome").Value & vbTab & _
                      rsCli.Fields("DataNascimento").Value & vbTab & _
                      sEmail & vbTab & _
                      sFones & vbTab & _
                      sUF & vbTab & _
                      sCidade & vbTab & _
                      sEndereco
              End If
          End If
      End If
    
    rsCli.MoveNext
  Wend
  rsCli.Close
  Set rsCli = Nothing

  Exit Sub
Erro:

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Dia_Ini_KeyPress(KeyAscii As Integer)
 KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub


Private Sub Dia_Fim_KeyPress(KeyAscii As Integer)
 KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Lista.ListIndex = Month(Date) - 1

  ' Grade D.R.E.
  gridAniversarioCli.ColWidth(0) = 0
  gridAniversarioCli.ColWidth(1) = 600
  gridAniversarioCli.ColWidth(2) = 2800
  gridAniversarioCli.ColWidth(3) = 1000
  gridAniversarioCli.ColWidth(4) = 3000
  gridAniversarioCli.ColWidth(5) = 2800
  gridAniversarioCli.ColWidth(6) = 330
  gridAniversarioCli.ColWidth(7) = 1500
  gridAniversarioCli.ColWidth(8) = 3500
  
  gridAniversarioCli.Row = 0
  gridAniversarioCli.TextMatrix(0, 1) = "Código"
  gridAniversarioCli.TextMatrix(0, 2) = "Nome"
  gridAniversarioCli.TextMatrix(0, 3) = "Aniversário"
  gridAniversarioCli.TextMatrix(0, 4) = "E-mail"
  gridAniversarioCli.TextMatrix(0, 5) = "Fones"
  gridAniversarioCli.TextMatrix(0, 6) = "UF"
  gridAniversarioCli.TextMatrix(0, 7) = "Cidade"
  gridAniversarioCli.TextMatrix(0, 8) = "Endereço"

End Sub

Private Sub opt_cliente_Click()
    If opt_cliente.Value = True Then
        frm_listaContatos.Visible = False
        frm_dataAniversarioClientes.Top = 1500
        frm_dataAniversarioClientes.Left = 60
        frm_dataAniversarioClientes.Visible = True
    End If
End Sub

Private Sub opt_listaDeContatosCliFor_Click()
    If opt_listaDeContatosCliFor.Value = True Then
        frm_listaContatos.Visible = True
        frm_dataAniversarioClientes.Visible = False
    End If
End Sub
