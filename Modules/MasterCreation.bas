Attribute VB_Name = "MasterCreation"
Option Explicit

Public Function GSTNo2digt(P_code As Long) As String
Dim q As String
Dim rst As Recordset
q = "Select statecodelong from masteraddressinfo where mastercode = " + CStr(P_code) + ""
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
GSTNo2digt = g_MS.StateCode2TinDigit(rst(0))
End If
End Function

Public Sub LoadAccount()
   

End Sub


