Attribute VB_Name = "modBrowserPasta"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long

Private Const BIF_RETURNONLYFSDIRS = 1

Private Type BrowseInfo
  hWndOwner As Long
  pidlRoot As Long
  sDisplayName As String
  sTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Function sFindDir(ByVal sTitle As String, ByVal hwnd As Long) As String
  Dim tFolder As BrowseInfo
  Dim nItem As Long
  Dim sDir As String
  
  With tFolder
    .hWndOwner = hwnd
    .pidlRoot = 0
    .sDisplayName = Space$(260)
    .sTitle = sTitle
    .ulFlags = BIF_RETURNONLYFSDIRS
    .lpfn = 0
    .lParam = 0
    .iImage = 0
  End With
  nItem = SHBrowseForFolder(tFolder)
  If nItem Then
    sDir = Space$(260)
    If SHGetPathFromIDList(nItem, sDir) Then
      sFindDir = Left(sDir, InStr(sDir, Chr$(0)) - 1)
    Else
      sFindDir = ""
    End If
  End If
End Function
