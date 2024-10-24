VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBusyWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ActivateDB(p_Data As Collection)
    Set g_CDataManager = New Busy2175.CDataManager
'MsgBox "AddOn Started"
    g_Provider = p_Data(6)
    g_CDataManager.SetProvider g_Provider
    g_CDataManager.ActivateDataBase p_Data
    g_DN = p_Data(12)
    Set g_FI = p_Data(50)
    'Set g_BatchCol = p_Data(15)        'to get data from Custom Validation use in ibusywin
End Sub

Public Sub Start()
    Dim InvalMsg As String


    Screen.MousePointer = vbHourglass
    
    g_BusyOpen = True
    
    Set g_CC = New Busy2175.CCompany
    Set g_MS = New Busy2175.CMasterServices
    Set g_TS = New Busy2175.CTranServices
    Set g_OS = New Busy2175.COtherServices
    Set g_CL = New Busy2175.CCommonLibrary
    If g_CC.Load(InvalMsg) = False Then
        MsgBox InvalMsg
    End If
'If g_DN = "19079513" Or g_DN = "21055165" Or g_DN = "19089529" Or g_DN = "1709A482" Or g_DN = "17082739" Or g_DN = "19031001" Or g_DN = "89809651" Then
 '   MsgBox "AddOn Started"
    Set g_fMain = New frmSplash
            g_fMain.Show vbModal
'Else
'MsgBox "Check Dongal"
'End If
     
    Exit Sub
    
End Sub




