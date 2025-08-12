VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmMensagens 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensagens"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10095
   Icon            =   "frmMensagens.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   10095
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   10035
      TabIndex        =   6
      Top             =   840
      Width           =   10095
      Begin SSDataWidgets_B.SSDBGrid grdMensagens 
         Height          =   4515
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   9915
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   4
         BevelColorFrame =   -2147483632
         BevelColorHighlight=   -2147483633
         BevelColorShadow=   -2147483633
         AllowUpdate     =   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         ForeColorEven   =   0
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   212
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Codigo"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Ordem"
         Columns(1).Name =   "Ordem"
         Columns(1).Alignment=   1
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   10266
         Columns(2).Caption=   "Mensagem"
         Columns(2).Name =   "Mensagem"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   6165
         Columns(3).Caption=   "Regra"
         Columns(3).Name =   "Regra"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         TabNavigation   =   1
         _ExtentX        =   17489
         _ExtentY        =   7964
         _StockProps     =   79
         BackColor       =   12632256
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
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   10035
      TabIndex        =   3
      Top             =   0
      Width           =   10095
      Begin VB.Label lblTitleDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Adicione as mensagens a serem utilizadas no sistema de acordo com as regras estipuladas."
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   9255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagens"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   5640
      Width           =   1335
   End
End
Attribute VB_Name = "frmMensagens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'26/01/2006 - mpdea
'Tela para a visualização e inserção de mensagens


Private Sub cmdIncluir_Click()
  Dim objMensagemRegra As clsMensagemRegra
  
  
  On Error GoTo ErrHandler
  
  
  Set objMensagemRegra = frmMensagensIncluirRegra.GetMensagemRegra
  
  If Not objMensagemRegra Is Nothing Then
    'Atualiza a ordem de todas as mensagens para que
    'a mensagem atual seja a primeira da ordem
    db.Execute "UPDATE Mensagens SET Ordem = Ordem + 1;", dbFailOnError
    'Adiciona a nova mensagem
    objMensagemRegra.AddNew
    'Carrega as mensagens
    Call CarregaMensagens
    'Descarrega objeto
    Set objMensagemRegra = Nothing
    'Seta o foco
    grdMensagens.SetFocus
  End If
  
  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

'Carrega as mensagens
Private Sub CarregaMensagens()
  Dim rstMensagens As Recordset
  Dim strSQL As String
  Dim strRegra As String
  Dim strAux As String
  Dim strAuxArray() As String
  
  
  On Error GoTo ErrHandler
  
  
  'Grid
  grdMensagens.Redraw = False
  grdMensagens.RemoveAll
  
  'Registros
  strSQL = "SELECT * FROM Mensagens ORDER BY Ordem"
  Set rstMensagens = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstMensagens
    Do Until .EOF
      
      '------------------------------------------------------------------------
      'Monta Regra
      '
      'Filtro Produtos
      Select Case .Fields("TipoFiltroProduto").Value
        Case tfpTodos
          strAux = "Todos"
        Case tfpGrupoFiscal
          strAux = "Grupo Fiscal = " & .Fields("FiltroProduto").Value
        Case tfpClasseSubClasse
          strAuxArray = Split(.Fields("FiltroProduto").Value, "|")
          'Classe
          If strAuxArray(0) = "" Or strAuxArray(0) = "0" Then
            strAux = "Classe = Todas"
          Else
            strAux = "Classe = " & strAuxArray(0)
          End If
          'Sub Classe
          If strAuxArray(1) = "" Or strAuxArray(1) = "0" Then
            strAux = strAux & " e Sub Classe = Todas"
          Else
            strAux = strAux & " e SubClasse = " & strAuxArray(1)
          End If
        Case tfpEspecifico
          strAux = "Código = " & .Fields("FiltroProduto").Value
      End Select
      strRegra = "Produtos: " & strAux & vbCrLf
      '
      'Filtro Op. Saída
      Select Case .Fields("TipoFiltroOpSaida").Value
        Case tfoTodas
          strAux = "Todas"
        Case tfoGrupoFiscal
          strAux = "Grupo Fiscal = " & .Fields("FiltroOpSaida").Value
        Case tfoEspecifica
          strAux = "Código = " & .Fields("FiltroOpSaida").Value
      End Select
      strRegra = strRegra & "Operações de Saída: " & strAux & vbCrLf
      '
      'Filtro Estado (UF) do Cleinte
      Select Case .Fields("TipoFiltroUF").Value
        Case tfuTodos
          strAux = "Todos"
        Case tfuEspecifico
          strAux = .Fields("FiltroUF").Value
      End Select
      strRegra = strRegra & "Estado: " & strAux
      '------------------------------------------------------------------------
      
      'Adiciona registro
      grdMensagens.AddNew
      grdMensagens.Columns("Codigo").Text = .Fields("Codigo").Value
      grdMensagens.Columns("Ordem").Text = .Fields("Ordem").Value
      grdMensagens.Columns("Regra").Text = strRegra
      grdMensagens.Columns("Mensagem").Text = .Fields("Mensagem").Value
      grdMensagens.Update
      
      .MoveNext
    Loop
    .Close
  End With
  Set rstMensagens = Nothing
  
  'Grid
  grdMensagens.RowHeight = 750!
  grdMensagens.MoveFirst
  grdMensagens.Redraw = True
  
  Exit Sub
  
ErrHandler:
  'Desabilita controles em caso de erro
  grdMensagens.Enabled = False
  cmdIncluir.Enabled = False
  cmdExcluir.Enabled = False
  'Grid
  grdMensagens.Redraw = True
  'Fecha tabela
  If Not rstMensagens Is Nothing Then
    rstMensagens.Close
    Set rstMensagens = Nothing
  End If
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmdExcluir_Click()
  With grdMensagens
    If .SelBookmarks.Count = 0 Then
      DisplayMsg "Selecione a mensagem a ser excluída."
      .SetFocus
    Else
      .DeleteSelected
    End If
  End With
End Sub

Private Sub Form_Load()
  
  On Error GoTo ErrHandler

  Call StatusMsg("")
  
  Call CenterForm(Me)
  
  'Carrega as mensagens
  Call CarregaMensagens
    
  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub grdMensagens_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  Dim intX As Integer
  Dim varBookmark As Variant
  Dim lngCodigo As Long
  
  
  On Error GoTo ErrHandler
  
  
  DispPromptMsg = False
  
  If bGridBeforeDelete() Then
    With grdMensagens
      'Limpa o registro do produto na lista
      For intX = 0 To (.SelBookmarks.Count - 1)
        varBookmark = .SelBookmarks(intX)
        lngCodigo = .Columns("Codigo").CellValue(varBookmark)
        db.Execute "DELETE FROM Mensagens WHERE Codigo = " & lngCodigo, dbFailOnError
      Next intX
    End With
    Cancel = False
  Else
    Cancel = True
  End If
    
  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub
