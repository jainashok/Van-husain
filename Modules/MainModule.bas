Attribute VB_Name = "MainModule"
Option Explicit

Public g_CDataManager As Busy2175.CDataManager
Public g_MS As Busy2175.CMasterServices
Public g_TS As Busy2175.CTranServices
Public g_OS As Busy2175.COtherServices
Public g_CC As Busy2175.CCompany
Public g_CL As Busy2175.CCommonLibrary

Public g_BatchCol As Collection  'use in main modules
Public g_DN
Public g_FI

Public g_BusyPath As String
Public g_DataPath As String
Public g_CompCode As String
Public g_User As String
Public g_PWD As String
Public g1_n As String
Public g1_c As String
Public g1_p As String
Public g_Database As String
Public g_Server As String
Public g_UID As String
Public g_Password As String

Public g_BusyOpen As Boolean

Public g_fMain As frmSplash

Public g_Provider As Integer

'Public Sub Main()
'    On Error GoTo EH
'    Dim InvalMsg As String
'    Dim fDataPath As frmSetDataPath
'
'    Screen.MousePointer = vbHourglass
'
'    g_BusyPath = "C:\Busywin\"
'    g_DataPath = "C:\BusyWin\Data\"
'    g_CompCode = "Comp0001"
'    g_User = "q"
'    g_PWD = "q"
'
'    Set fDataPath = New frmSetDataPath
'    fDataPath.Show
'
'    Exit Sub
'
'EH:
'    If Err.Number = 1230 Then
'        MsgBox "Invalid User Name or Password"
'    Else
'        MsgBox Err.Description
'        End
'    End If
'
'End Sub

Public Sub CloseApp()

    If g_BusyOpen Then
        
        Unload g_fMain
        Set g_fMain = Nothing
        
        g_CDataManager.CloseCompDataBase
        
        Set g_CC = Nothing
        Set g_MS = Nothing
        Set g_TS = Nothing
        Set g_OS = Nothing
        Set g_CDataManager = Nothing
        
    End If
    
End Sub
