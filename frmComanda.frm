VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmComanda 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comanda"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   ClipControls    =   0   'False
   LinkTopic       =   "Comanda"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Vendas 
      CausesValidation=   0   'False
      Height          =   3675
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   6482
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      HighLight       =   2
      AllowUserResizing=   1
      BorderStyle     =   0
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
End
Attribute VB_Name = "frmComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sComanda As String
Private iSequencia As Long
Private bComandaOK As Boolean
Private iCount As Integer

Public Property Let Comanda(ByVal pComanda As String)
    sComanda = pComanda
    Call BuscarComanda
End Property

Public Property Get Sequencia() As Long
    Sequencia = iSequencia
End Property

Public Property Get ComandaOK() As Boolean
    ComandaOK = bComandaOK
End Property

Public Property Get Total() As Integer
    Total = iCount
End Property

Private Function QUERY(ByVal SomenteNaoEfetivada As Boolean) As String
    Dim sql As String
    If SomenteNaoEfetivada Then
        sql = "SELECT DISTINCT SaidasComandas.CodSaida "
        sql = sql & "FROM (SaidasComandas INNER JOIN Saídas ON (Saídas.Sequência = SaidasComandas.CodSaida) AND (Saídas.Filial = SaidasComandas.Filial)) "
        sql = sql & "WHERE SaidasComandas.Filial = " & gnCodFilial & " AND SaidasComandas.CodComanda = '" & sComanda & "' AND Saídas.Efetivada = False"
        sql = sql & " ORDER BY SaidasComandas.CodSaida;"
    Else
        sql = "SELECT SaidasComandas.Filial, SaidasComandas.CodSaida, Saídas.Data, SaidasComandas.CodComanda, Saídas.Efetivada, [Saídas - Produtos].Linha, [Saídas - Produtos].Código, Produtos.Nome, [Saídas - Produtos].Qtde "
        sql = sql & "FROM (Saídas INNER JOIN SaidasComandas ON (Saídas.Sequência = SaidasComandas.CodSaida) AND (Saídas.Filial = SaidasComandas.Filial)) INNER JOIN (Produtos INNER JOIN [Saídas - Produtos] ON Produtos.Código = [Saídas - Produtos].Código) ON (Saídas.Sequência = [Saídas - Produtos].Sequência) AND (Saídas.Filial = [Saídas - Produtos].Filial) "
        sql = sql & "WHERE SaidasComandas.Filial = " & gnCodFilial & " AND SaidasComandas.CodComanda = '" & sComanda & "' "
        sql = sql & "ORDER BY SaidasComandas.Filial, SaidasComandas.CodSaida, [Saídas - Produtos].Linha;"
    End If
    
    QUERY = sql
End Function

Public Function Existe(ByVal pSequencia As Long) As Boolean
    Dim sql As String
    sql = "SELECT COUNT(*) AS qtde "
    sql = sql & "FROM SaidasComandas "
    sql = sql & "WHERE SaidasComandas.Filial = " & gnCodFilial & " AND SaidasComandas.CodComanda = '" & sComanda & "' AND SaidasComandas.CodSaida = " & pSequencia & ";"

    Dim rsTemp As Recordset
    Set rsTemp = db.OpenRecordset(sql, dbOpenDynaset)
    
    rsTemp.MoveFirst
    Existe = (CInt(rsTemp("qtde")) = 1)

    rsTemp.Close
End Function

Public Sub Deleta(ByVal pSequencia As Long)
    Dim sql As String
    sql = "DELETE FROM SaidasComandas "
    sql = sql & "WHERE SaidasComandas.Filial = " & gnCodFilial & " AND SaidasComandas.CodSaida = " & pSequencia & ";"
    
    db.Execute sql
End Sub

Private Sub BuscarComanda()
  iSequencia = 0
  iCount = 0
  
  Dim rsComanda As Recordset
  Set rsComanda = db.OpenRecordset(QUERY(True), dbOpenDynaset)
  
  While Not rsComanda.EOF
    iCount = iCount + 1
    rsComanda.MoveNext
  Wend
  
  If (iCount > 0) Then
    rsComanda.MoveFirst
    If (iCount = 1) Then
      bComandaOK = True
      iSequencia = CLng(rsComanda("CodSaida"))
    Else
      bComandaOK = False
      MsgBox "Erro: Mais de uma COMANDA não efetivada!", vbCritical
    End If
  Else
    bComandaOK = True
    rsComanda.Close
    iCount = 0
    Set rsComanda = db.OpenRecordset(QUERY(False), dbOpenDynaset)
    While Not rsComanda.EOF
      iCount = iCount + 1
      rsComanda.MoveNext
    Wend
    If (iCount > 0) Then
      rsComanda.MoveFirst
      With Vendas
        .Clear
        
        .TextMatrix(0, 0) = "Seqüência"
        .ColWidth(0) = 1000
        
        .TextMatrix(0, 1) = "Data"
        .ColWidth(1) = 1100
        
        .TextMatrix(0, 2) = "Código"
        .ColWidth(2) = 1700
        
        .TextMatrix(0, 3) = "Produto"
        .ColWidth(3) = 5000
        
        .TextMatrix(0, 4) = "Qtde"
        .ColWidth(4) = 510
        
        Do While Not rsComanda.EOF
          .Rows = .Rows + 1
          
          'Seqüência
          .TextMatrix(.Rows - 1, 0) = rsComanda("CodSaida")
          
          'Data
          .TextMatrix(.Rows - 1, 1) = rsComanda("Data")
          
          'Código
          .TextMatrix(.Rows - 1, 2) = rsComanda("Código")
          
          'Produto
          .TextMatrix(.Rows - 1, 3) = rsComanda("Nome")
          
          'Qtde
          .TextMatrix(.Rows - 1, 4) = rsComanda("Qtde")

          rsComanda.MoveNext
        Loop
        .FixedRows = 1
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        .Refresh
      End With
    End If
  End If
  rsComanda.Close
End Sub

Private Sub Form_Load()
    If Trim(sComanda) = "" Then
        MsgBox "Defina a comanda primeiro!"
        Unload Me
    Else
        Me.Caption = "Comanda - " & sComanda
    End If
End Sub
