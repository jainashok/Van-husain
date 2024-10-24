Attribute VB_Name = "SystemModule"
Option Explicit

Public Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Function SelectFolder(ByVal p_Hwnd As Long) As String
        
        Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
                 
         Dim tBrowseInfo As BrowseInfo
               
         szTitle = "Select Folder Name"
         
         With tBrowseInfo
            .hWndOwner = p_Hwnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS  '  + BIF_DONTGOBELOWDOMAIN
            
         End With
        
         lpIDList = SHBrowseForFolder(tBrowseInfo)
            
         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            SelectFolder = sBuffer
         End If
         
End Function
