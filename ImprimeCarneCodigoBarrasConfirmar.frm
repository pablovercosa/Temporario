VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmImprimeCarneCodigoBarrasConfirmar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Carnês"
   ClientHeight    =   1815
   ClientLeft      =   3405
   ClientTop       =   2295
   ClientWidth     =   3855
   Icon            =   "ImprimeCarneCodigoBarrasConfirmar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1815
   ScaleWidth      =   3855
   Begin VB.Frame Frame1 
      Caption         =   "Selecione o layout para impressão do carnê"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optEsquerda 
         Caption         =   "Alinhado à Esquerda"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optDireita 
         Caption         =   "Alinhado à Direita"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   400
      Left            =   2400
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   360
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmImprimeCarneCodigoBarrasConfirmar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intFilial As Integer
Public lngSeq As Long

Private Sub Command1_Click()

  Dim strSQL As String
  Dim strSequencia As String
  Dim lngImpressos As Long
  Dim Str_Rel As String
  Dim rsReceber As Recordset
   
  On Error GoTo ErrHandler

  Call StatusMsg("")
  
  strSQL = "SELECT * "
  strSQL = strSQL & "FROM [Contas a Receber] "
  strSQL = strSQL & "WHERE CarneCodigoBarras <>'' "
  strSQL = strSQL & "  AND Filial = " & intFilial & " "
  strSQL = strSQL & "  AND Sequência = " & lngSeq & " "
  
  strSQL = strSQL & " ORDER BY Sequência, Contador"
  
  Set rsReceber = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  strSequencia = ""
  
  With rsReceber
    Do Until .EOF
      If strSequencia <> .Fields("Sequência") Then
        strSequencia = .Fields("Sequência")
        lngImpressos = lngImpressos + 1
      End If
      .Edit
      .Fields("Carnet Impresso").Value = -1
      .Update
      .MoveNext
    Loop
  End With
  
  If lngImpressos > 0 Then
    Rem  Nome do BD
    Rel.DataFiles(0) = gsQuickDBFileName
    
    Rem Saída
    Rel.Destination = 1
    
    If optDireita.Value = True Then Rel.ReportFileName = gsReportPath & "CARNEDIREITO.RPT"
    If optEsquerda.Value = True Then Rel.ReportFileName = gsReportPath & "CARNEESQUERDO.RPT"
        
    Str_Rel = Str_Rel & "{Contas a Receber.Filial} = " & intFilial
    Str_Rel = Str_Rel & " And {Contas a Receber.Sequência} = " & lngSeq
    Str_Rel = Str_Rel & " And {Contas a Receber.CarneCodigoBarras} <>'' "
    
    Rel.SelectionFormula = Str_Rel
    
    Call StatusMsg("Aguarde, imprimindo...")
    MousePointer = vbHourglass
     
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("CARNÊ", Rel)

    Rel.Action = 1
    
    Call StatusMsg("")
    MousePointer = vbDefault
    
    DisplayMsg "Final de impressão, foram impressos " + str(lngImpressos) + " carnês."
  Else
    DisplayMsg "Não há dados para serem impressos, favor verificar os parâmetros selecionados."
  End If
  
  rsReceber.Close
  Set rsReceber = Nothing
  
  Unload Me
  Exit Sub

ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Imprimir documento."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Unload Me
  Exit Sub
  
End Sub

Private Sub Command2_Click()
  gsRetornoDoc = "NÃO"
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub


