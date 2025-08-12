VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosConfiguraTab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração de Tabela de Preços"
   ClientHeight    =   3495
   ClientLeft      =   2055
   ClientTop       =   990
   ClientWidth     =   7350
   HelpContextID   =   1090
   Icon            =   "ConfiguraTab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   7350
   Begin VB.TextBox txtPercentualComissaoDesconto 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5640
      TabIndex        =   18
      Text            =   "0"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Data datPrecos 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   225
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Preços ORDER BY Tabela"
      Top             =   3885
      Width           =   1875
   End
   Begin VB.TextBox Vezes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Multiplicador 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CheckBox Aceita_Vale 
      Caption         =   "Aceita Vale / Outros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CheckBox Aceita_Cartão 
      Caption         =   "Aceita cartão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton B_Grava 
      BackColor       =   &H0000C0C0&
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parcelamento"
      Height          =   650
      Left            =   3230
      TabIndex        =   12
      Top             =   720
      Width           =   3990
      Begin VB.CheckBox Aceita_Parcela 
         Appearance      =   0  'Flat
         Caption         =   "Aceita parcelamento (contas a receber)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   3735
      End
      Begin VB.TextBox Prazo_Parcela 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   3
         Top             =   260
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Prazo máximo, em dias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   285
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cheques"
      Height          =   650
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   3135
      Begin VB.CheckBox Aceita_Pré 
         Appearance      =   0  'Flat
         Caption         =   "Aceita cheques"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   0
         Top             =   0
         Width           =   1620
      End
      Begin VB.TextBox Prazo_Pré 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Prazo máximo, em dias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   1695
      End
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "ConfiguraTab.frx":058A
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Top             =   120
      Width           =   2895
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   5106
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Tabela"
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   3230
      X2              =   3230
      Y1              =   2880
      Y2              =   1320
   End
   Begin VB.Label Label7 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   19
      Top             =   2430
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Se houver desconto, diminuir a comissão em :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   7200
      X2              =   120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   7200
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      Caption         =   "Ao imprimir a tabela, dividir o preço por x parcelas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Multiplicador da comissão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Nome da Tabela"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   150
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrecosConfiguraTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTabelas As Recordset
'Dim rsPrecos As Recordset

Sub ShowRecord()
  Aceita_Pré.Value = 0
  If rsTabelas("Aceita Pré") = True Then Aceita_Pré.Value = 1
  Prazo_Pré.Text = rsTabelas("Prazo Pré")
  
  Aceita_Parcela.Value = 0
  If rsTabelas("Aceita Parcelamento") = True Then Aceita_Parcela.Value = 1
  Prazo_Parcela.Text = rsTabelas("Prazo Parcelamento")
  
  Aceita_Cartão.Value = 0
  If rsTabelas("Aceita Cartão") = True Then Aceita_Cartão.Value = 1
  
  Aceita_Vale.Value = 0
  If rsTabelas("Aceita Vale") = True Then Aceita_Vale.Value = 1
  
  Multiplicador.Text = rsTabelas("Multiplicador Comissão") & ""
  txtPercentualComissaoDesconto.Text = rsTabelas.Fields("PercentualComissaoDesconto") & ""
  Vezes.Text = rsTabelas("Dividir")
  
End Sub

Private Sub B_Grava_Click()
  Dim Erro As Integer
  
  Call StatusMsg("")
  
  If cboLista.Text = "" Then Exit Sub
  
  If Not IsNumeric(txtPercentualComissaoDesconto.Text) Then
    MsgBox "O valor do percentual a diminuir da comissão caso haja algum desconto, não é válido !", vbCritical, "Quick Store"
    Exit Sub
  End If

  
  If IsNull(Vezes.Text) Then Vezes.Text = 0
  If Vezes.Text = "" Then Vezes.Text = 0
  If Val(Vezes.Text) < 0 Then Vezes.Text = 0
   
  Erro = False
  If IsNull(Prazo_Pré.Text) Then Erro = True
  If Erro = False Then If Prazo_Pré.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Prazo_Pré.Text) Then Erro = True
  If Erro = False Then If Val(Prazo_Pré.Text) < 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "Prazo para cheques pré-datados inválido, verifique."
    Prazo_Pré.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Prazo_Parcela.Text) Then Erro = True
  If Erro = False Then If Prazo_Parcela.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Prazo_Parcela.Text) Then Erro = True
  If Erro = False Then If Val(Prazo_Parcela.Text) < 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "Prazo para parcelamento inválido, verifique."
    Prazo_Parcela.SetFocus
    Exit Sub
  End If
   
  Erro = False
  If IsNull(Multiplicador.Text) Then Erro = True
  If Erro = False Then If Multiplicador.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Multiplicador.Text) Then Erro = True
  If Erro = False Then If CDbl(Multiplicador.Text) < 0 Then Erro = True
    If Erro = True Then
    DisplayMsg "Multiplicador da comissão inválido, verifique."
    Multiplicador.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Aguarde...")
   
  rsTabelas.Index = "Tabela"
  rsTabelas.Seek "=", cboLista.Text
   
  If rsTabelas.NoMatch Then
    rsTabelas.AddNew
    rsTabelas("Tabela") = cboLista.Text
  Else
    rsTabelas.Edit
  End If
   
  If Aceita_Pré.Value = 0 Then rsTabelas("Aceita Pré") = False
  If Aceita_Pré.Value = 1 Then rsTabelas("Aceita Pré") = True
  rsTabelas("Prazo Pré") = Val(Prazo_Pré.Text)
  
  If Aceita_Parcela.Value = 0 Then rsTabelas("Aceita Parcelamento") = False
  If Aceita_Parcela.Value = 1 Then rsTabelas("Aceita PArcelamento") = True
  rsTabelas("Prazo Parcelamento") = Val(Prazo_Parcela.Text)
  
  If Aceita_Cartão.Value = 0 Then rsTabelas("Aceita Cartão") = False
  If Aceita_Cartão.Value = 1 Then rsTabelas("Aceita Cartão") = True
  
  If Aceita_Vale.Value = 0 Then rsTabelas("Aceita Vale") = False
  If Aceita_Vale.Value = 1 Then rsTabelas("Aceita Vale") = True
  
  rsTabelas("Multiplicador Comissão") = CDbl(Multiplicador.Text)
  rsTabelas("PercentualComissaoDesconto") = CDbl(txtPercentualComissaoDesconto.Text)
  rsTabelas("Dividir") = Val(Vezes.Text)
   
  rsTabelas("Data Alteração") = Format(Date, "dd/mm/yyyy")
  rsTabelas.Update
  
  Call StatusMsg("")
  
End Sub

Private Sub cboLista_Click()
  Call FindRecord
End Sub

Private Sub cboLista_KeyPress(KeyAscii As Integer)
  KeyAscii = gnLimitKeyPress(cboLista, 15, KeyAscii)
  If KeyAscii <> 0 Then
    KeyAscii = gnTypeValidKey(KeyAscii)
  End If
End Sub

Private Sub cboLista_LostFocus()
  Call FindRecord
End Sub

Private Sub Form_Load()
 
  Call CenterForm(Me)
  
  datPrecos.DatabaseName = gsQuickDBFileName
  Set datPrecos.Recordset = db.OpenRecordset(SQL_CONS_TAB_PRECO_T1, dbOpenSnapshot)
  
  Set rsTabelas = db.OpenRecordset("Tabela de Preços")
'  Set rsPrecos = db.OpenRecordset("Preços", , dbReadOnly)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call StatusMsg("")
  rsTabelas.Close
'  rsPrecos.Close
  Set rsTabelas = Nothing
'  Set rsPrecos = Nothing
End Sub

Private Sub Multiplicador_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Prazo_Parcela_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Prazo_Pré_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Vezes_KeyPress(KeyAscii As Integer)
 KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub FindRecord()
  If cboLista.Text = "" Then
    Exit Sub
  Else
    cboLista.Text = UCase(cboLista.Text)
  End If
  
  rsTabelas.Index = "Tabela"
  rsTabelas.Seek "=", cboLista.Text
  If Not rsTabelas.NoMatch Then
    ShowRecord
  Else
    Aceita_Pré.Value = 1
    Prazo_Pré.Text = 9999
    Aceita_Parcela.Value = 1
    Prazo_Parcela.Text = 9999
    Aceita_Cartão.Value = 1
    Aceita_Vale.Value = 1
    Multiplicador.Text = 1
    txtPercentualComissaoDesconto.Text = 0
  End If
End Sub
