Attribute VB_Name = "modForms"
Option Explicit

Global gbSkipKey As Boolean

Public Sub HandleKeyDown(KeyCode As Integer, Shift As Integer)
  Dim Tool As ActiveBarLibraryCtl.Tool
  
  Set Tool = New ActiveBarLibraryCtl.Tool
  gbSkipKey = False
    
  Select Case KeyCode
    Case vbKeyF3
      Tool.Name = "miComplConsultaProdutos"
      Call Screen.ActiveForm.ActiveBar1_Click(Tool)
      KeyCode = 0
    Case vbKeyF9
      Tool.Name = "miOpFirst"
      Call Screen.ActiveForm.ActiveBar1_Click(Tool)
      KeyCode = 0
    Case vbKeyF10
      Tool.Name = "miOpPrevious"
      Call Screen.ActiveForm.ActiveBar1_Click(Tool)
      KeyCode = 0
    Case vbKeyF11
      Tool.Name = "miOpNext"
      Call Screen.ActiveForm.ActiveBar1_Click(Tool)
      KeyCode = 0
    Case vbKeyF12
      Tool.Name = "miOpLast"
      Call Screen.ActiveForm.ActiveBar1_Click(Tool)
      KeyCode = 0

    '26/09/2002 - mpdea
    'Removido ações para vbKeyUp e vbKeyDown
    
'    Case vbKeyUp
'      SendKeys "+{TAB}{HOME}"
'    Case vbKeyDown
'      SendKeys "{TAB}{HOME}"
    
    Case vbKeyN
      If Shift = vbCtrlMask Then
        Tool.Name = "miOpClear"
        Call Screen.ActiveForm.ActiveBar1_Click(Tool)
        KeyCode = 0
      End If
    Case vbKeyG
      If Shift = vbCtrlMask Then
        'Verifica se a Tool está ativada para que possa ser executada a função
        'Caso o form não possua a Tool, ignora passando para a próxima execução
        On Error Resume Next
        If Screen.ActiveForm.ActiveBar1.Tools("miOpUpdate").Enabled Then
          Tool.Name = "miOpUpdate"
          Call Screen.ActiveForm.ActiveBar1_Click(Tool)
          KeyCode = 0
        End If
        On Error GoTo 0
      End If
    Case vbKeyA
      If Shift = vbCtrlMask Then
        'Verifica se a Tool está ativada para que possa ser executada a função
        'Caso o form não possua a Tool, ignora passando para a próxima execução
        On Error Resume Next
        If Screen.ActiveForm.ActiveBar1.Tools("miOpDelete").Enabled Then
          Tool.Name = "miOpDelete"
          Call Screen.ActiveForm.ActiveBar1_Click(Tool)
          KeyCode = 0
        End If
        On Error GoTo 0
      End If
    Case vbKeyP
      If Shift = vbCtrlMask Then
        Tool.Name = "miOpSearch"
        Call Screen.ActiveForm.ActiveBar1_Click(Tool)
        KeyCode = 0
      End If

    '23/10/2002 - mpdea
    'Removido tratamento para fechamento da janelas filhas
    'por ser um tratamento padrão do sistema operacional
    
'    Case vbKeyF4
'      If Shift = vbCtrlMask Then
'        Unload Screen.ActiveForm
'      End If

    Case Else
      'Exceções
      If Shift = vbCtrlMask And (KeyCode = vbKeyX Or KeyCode = vbKeyC Or KeyCode = vbKeyV) Then
        Exit Sub
      Else
      '30/01/2009 - mpdea
      'Adaptado para o novo menu
      'Key: Q7MENU
'        If frmMain.ActiveBar1.OnKeyDown(KeyCode, Shift) Then
'          KeyCode = 0
'          Shift = 0
'        End If
      End If
  End Select
  Set Tool = Nothing
End Sub

Private Function bGoodKeyNumber(ByVal nKeyCode As Integer) As Boolean
  bGoodKeyNumber = nKeyCode >= vbKey0 And nKeyCode <= vbKey9
  bGoodKeyNumber = bGoodKeyNumber Or (nKeyCode >= vbKeyNumpad0 And nKeyCode <= vbKeyNumpad9)
  bGoodKeyNumber = bGoodKeyNumber Or nKeyCode = 188  ' Ponto Decimal
  bGoodKeyNumber = bGoodKeyNumber Or nKeyCode = vbKeyBack
  bGoodKeyNumber = bGoodKeyNumber Or nKeyCode <= vbKeyHelp
End Function

Public Sub PaintTheForm(F As Form)
  Dim nX As Integer, nY As Integer
  Dim nPicWidth As Integer
  Dim nPicHeight As Integer
  Dim nFrmWidth As Integer
  Dim nFrmHeight As Integer
  
  nPicWidth = F.picBackground.Width
  nPicHeight = F.picBackground.Height
  nFrmWidth = F.Width
  nFrmHeight = F.Height
  For nX = 0 To nFrmWidth Step nPicWidth
    For nY = 0 To nFrmHeight Step nPicHeight
      F.PaintPicture F.picBackground, nX, nY
    Next nY
  Next nX
End Sub
