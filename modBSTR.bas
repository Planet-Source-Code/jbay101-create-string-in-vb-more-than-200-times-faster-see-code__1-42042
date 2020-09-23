Attribute VB_Name = "modFastString"
Option Explicit

Public Function AllocString(ByVal lSize As Long) As String
RtlMoveMemory ByVal VarPtr(AllocString_ADVANCED), SysAllocStringByteLen(0&, lSize + lSize), 4&
End Function

