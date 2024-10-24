Attribute VB_Name = "ScreenReps"
Option Explicit



Public Function GetAccList() As Collection
    Dim Qry As String
    Dim rst As Recordset
    Dim RetVal As Collection
    
    FlushCol RetVal
    
    'qry = "SELECT NAME FROM MASTER1 WHERE MASTERTYPE=" & ACC_MAST
    Qry = "SELECT NAMEALIAS FROM HELP1 WHERE RECTYPE = " & CStr(H1_PARTYCASHBANK) & " AND NAMEORALIAS=" & CStr(NA_NAME)
    Set rst = g_OS.GetRecordset(Qry)
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RetVal.Add rst!NameAlias.Value
            rst.MoveNext
        Loop
    End If
    
    Set GetAccList = RetVal
    
    Set rst = Nothing
    
    
End Function
Public Function GetParty() As Collection
    Dim Qry As String
    Dim rst As Recordset
    Dim RetVal As Collection
    
    FlushCol RetVal
    
    'qry = "SELECT NAME FROM MASTER1 WHERE MASTERTYPE=" & ACC_MAST
    Qry = "SELECT NAMEALIAS FROM HELP1 WHERE RECTYPE = " & CStr(H1_PARTY) & " AND NAMEORALIAS=" & CStr(NA_ALIAS)
    Set rst = g_OS.GetRecordset(Qry)
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RetVal.Add rst!NameAlias.Value
            rst.MoveNext
        Loop
    End If
    
    Set GetParty = RetVal
    
    Set rst = Nothing
    
    
End Function
Public Function GetCon() As Collection
    Dim Qry As String
    Dim rst As Recordset
    Dim RetVal As Collection
    
    FlushCol RetVal
    
    'qry = "SELECT NAME FROM MASTER1 WHERE MASTERTYPE=" & ACC_MAST
    Qry = "SELECT NAMEALIAS FROM HELP1 WHERE RECTYPE = " & CStr(H1_CASHBANK) & " AND NAMEORALIAS=" & CStr(NA_NAME)
    Set rst = g_OS.GetRecordset(Qry)
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RetVal.Add rst!NameAlias.Value
            rst.MoveNext
        Loop
    End If
    
    Set GetCon = RetVal
    
    Set rst = Nothing
    
    
End Function


Public Function GetAccgrp() As Collection
    Dim Qry As String
    Dim rst As Recordset
    Dim RetVal As Collection
    
    FlushCol RetVal
    
    'qry = "SELECT NAME FROM MASTER1 WHERE MASTERTYPE=" & ACC_MAST
'    Qry = "SELECT NAMEALIAS FROM HELP1 WHERE RECTYPE = " & CStr(H1_AG_PARTY) & "AND NAMEORALIAS=" & CStr(NA_NAME)
'    Set rst = g_OS.GetRecordset(Qry)
'
'    If rst.RecordCount > 0 Then
'        rst.MoveFirst
'        Do While Not rst.EOF
'            RetVal.Add rst!NameAlias.Value
'            rst.MoveNext
'        Loop
'    End If
    Qry = "SELECT NAMEALIAS FROM HELP1 WHERE RECTYPE = " & CStr(H1_AG_PARTYCASHBANK) & "AND NAMEORALIAS=" & CStr(NA_NAME)
    Set rst = g_OS.GetRecordset(Qry)
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RetVal.Add rst!NameAlias.Value
            rst.MoveNext
        Loop
    End If
    Set GetAccgrp = RetVal
    
    Set rst = Nothing
    
    
End Function

