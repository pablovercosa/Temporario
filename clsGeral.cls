VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// Constant declarations
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_HELP_CONTEXT = &HF

'// API declaration
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Public Sub Show(NewFile As String, Optional WindowPane As String, Optional ContextID)

    Dim strfile As String
    
    strfile = NewFile
    
    If Len(WindowPane) Then
        strfile = Trim(strfile) & ">" & Trim(WindowPane)
    End If
    
    If IsMissing(ContextID) Then
        Call HtmlHelp(0, strfile, HH_DISPLAY_TOPIC, ByVal 0&)
    Else
        Call HtmlHelp(0, strfile, HH_HELP_CONTEXT, ContextID)
    End If


End Sub

