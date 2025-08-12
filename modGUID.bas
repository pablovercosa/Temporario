Attribute VB_Name = "modGUID"
Option Explicit

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGUID As GUID) As Long

Private Declare Function StringFromGUID2 Lib "OLE32.DLL" (pGUID As GUID, ByVal PointerToString As Long, ByVal MaxLength As Long) As Long

Private Type GUID
  Guid1 As Long
  Guid2 As Long
  Guid3 As Long
  Guid4(0 To 7) As Byte
End Type

Public Function CreateGUID() As String

  Dim udtGUID As GUID
  Dim sGUID As String
  Dim lResult As Long
  
  lResult = CoCreateGuid(udtGUID)
  
  If lResult Then
     sGUID = ""
  Else
     sGUID = String$(38, 0)
     StringFromGUID2 udtGUID, StrPtr(sGUID), 39
  End If
  
  CreateGUID = sGUID
  
End Function
