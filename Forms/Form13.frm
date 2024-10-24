VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form13 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecting Group"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   ForeColor       =   &H00C000C0&
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Select All in Sub Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   7200
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5895
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   10398
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sub Group"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Group"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecting Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox Check2 
         Caption         =   "Seasonal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   6600
         Width           =   2415
      End
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   13500
      Left            =   0
      Picture         =   "Form13.frx":0000
      Top             =   0
      Width           =   21000
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

Dim t As Integer
If Check1.Value = 0 Then
For t = 1 To ListView2.ListItems.Count
If ListView2.ListItems(t).Checked = True Then
ListView2.ListItems(t).Checked = False
Else
End If
Next
End If
If Check1.Value = 1 Then
For t = 1 To ListView2.ListItems.Count
If ListView2.ListItems(t).Checked = False Then
ListView2.ListItems(t).Checked = True
Else
End If
Next
End If

End Sub

Private Sub Command1_Click()
Dim j As Boolean
j = False
Dim SrNo As String
SrNo = g_DN
If SrNo = "15065170" Then
If ListView1.ListItems.Count > 0 Then

    Dim q As String
    Dim q1 As String
    Dim rst As Recordset
    Dim iname As String
    Dim i As Integer
    Dim sgname As String
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
        iname = ListView1.ListItems(i).Text
       j = True
       Exit For
        Else
        j = False
        End If
    Next
    If j = True And Check2.Value = 0 Then
    q = "Select d1 from Externaldata where c1 ='5' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ""
    Set rst = g_OS.GetRecordset(q)
        If rst.RecordCount > 0 Then
        q1 = "Delete * from externaldata where c1 ='5' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ""
        g_OS.ExecuteQuerytmp (q1)
        End If
        
        q1 = "Insert into externaldata(c1,L1,D1) Values('5'," + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + "," + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ")"
        g_OS.ExecuteQuerytmp (q1)
        If ListView2.ListItems.Count > 0 Then
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Checked = True Then
        sgname = ListView2.ListItems(i).Text
        q1 = "Insert into externaldata(c1,L1,D1) Values('5'," + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + "," + CStr(g_MS.MasterName2CodeIfExist(sgname, 5)) + ")"
        g_OS.ExecuteQuerytmp (q1)
        'MsgBox "Record Updated"
        Else

        End If
    Next
    End If
    
    MsgBox "Record Updated & Please Make order of Groups"
    Else
    MsgBox "Please Select one Primery Group"
    End If
If j = True And Check2.Value = 1 Then
    q = "Select d1 from Externaldata where c1 ='6' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ""
    Set rst = g_OS.GetRecordset(q)
        If rst.RecordCount > 0 Then
        q1 = "Delete * from externaldata where c1 ='6' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ""
        g_OS.ExecuteQuerytmp (q1)
        End If
        
        q1 = "Insert into externaldata(c1,L1,D1) Values('6'," + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + "," + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ")"
        g_OS.ExecuteQuerytmp (q1)
        If ListView2.ListItems.Count > 0 Then
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Checked = True Then
        sgname = ListView2.ListItems(i).Text
        q1 = "Insert into externaldata(c1,L1,D1) Values('6'," + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + "," + CStr(g_MS.MasterName2CodeIfExist(sgname, 5)) + ")"
        g_OS.ExecuteQuerytmp (q1)
        'MsgBox "Record Updated"
        Else

        End If
    Next
    End If
    
    MsgBox "Record Updated & Please Make order of Groups"
    Else
    MsgBox "Please Select one Primery Group"
    End If

End If
Else
MsgBox "Please Check Dongal No."
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim j As Boolean
j = False
Dim SrNo As String
SrNo = g_DN
If SrNo = "15065170" Then
If ListView1.ListItems.Count > 0 Then

    Dim q As String
    Dim q1 As String
    Dim rst As Recordset
    Dim iname As String
    Dim i As Integer
    Dim sgname As String
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
        iname = ListView1.ListItems(i).Text
       j = True
       Exit For
        Else
        j = False
        End If
    Next
If j = True And Check2.Value = 0 Then
    q = "Select d1 from Externaldata where c1 ='5' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ""
    Set rst = g_OS.GetRecordset(q)
        If rst.RecordCount > 0 Then
        q1 = "Delete * from externaldata where c1 ='5' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ""
        g_OS.ExecuteQuerytmp (q1)
        End If
        
        MsgBox "Record Updated & Please Make order of Groups"
Else
    MsgBox "Please Select one Primery Group"
End If
If j = True And Check2.Value = 1 Then
    q = "Select d1 from Externaldata where c1 ='6' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ""
    Set rst = g_OS.GetRecordset(q)
        If rst.RecordCount > 0 Then
        q1 = "Delete * from externaldata where c1 ='6' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + ""
        g_OS.ExecuteQuerytmp (q1)
        End If
        
        MsgBox "Record Updated & Please Make order of Groups"
Else
    MsgBox "Please Select one Primery Group"
End If

End If
Else
MsgBox "Please Check Dongal No."
End If

End Sub

Private Sub Form_Load()
Dim q As String
Dim rst As Recordset
Dim gp As Long
Dim xt As ListItem
Dim gpn As String
gpn = "General"
gp = g_MS.MasterName2Code(gpn, 5)
ListView1.ListItems.Clear
q = "Select name from master1 where  mastertype =5 and (parentgrp =0 or parentgrp= " + CStr(gp) + ") order by name"

Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
Set xt = ListView1.ListItems.Add(, , rst(0))
rst.MoveNext
Loop
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim q As String
Dim rst As Recordset
Dim i As Integer
Dim ac As New Collection
Dim strArray() As String
Dim av As Long
Dim xt As ListItem
Dim n As Integer
Dim inccount As Integer
Dim q2 As String
Dim rst2 As Recordset
Dim r3 As String
Dim iname As String
ListView2.ListItems.Clear
For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems(i).Checked = True Then
    Dim p As Integer
    g_CL.FlushCol ac
    p = Label3.Caption
    If p = 0 Then
    Label3.Caption = ListView1.ListItems(i).Index
    iname = ListView1.ListItems(i).Text
    Else
    ListView1.ListItems(p).Checked = False
    Label3.Caption = ListView1.ListItems(i).Index
    iname = ListView1.ListItems(i).Text
    End If
End If
Next

For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems(i).Checked = True Then
    av = g_MS.MasterName2Code(ListView1.ListItems(i).Text, 5)
    q = g_OS.GenQryInStrForParentGrps(av, 5)
    q = Replace(q, "(", "")
    q = Replace(q, ")", "")
    strArray = Split(q, ",")
    For intcount = LBound(strArray) To UBound(strArray)
    ac.Add strArray(intcount)
    Next
    For n = 1 To ac.Count
        If Trim(ListView1.ListItems(i).Text) = Trim(g_MS.MasterCode2Name(ac(n))) Then
        Else
        Set xt = ListView2.ListItems.Add(, , g_MS.MasterCode2Name(ac(n)))
        End If
    Next
End If
Next
If Check1.Value = 0 Then
For i = 1 To ListView2.ListItems.Count
r3 = ListView2.ListItems(i).Text
q2 = "Select d1 from Externaldata where c1 ='5' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + " and D1 =" + CStr(g_MS.MasterName2CodeIfExist(r3, 5)) + ""
Set rst2 = g_OS.GetRecordset(q2)


If rst2.RecordCount > 0 Then

ListView2.ListItems(i).Checked = True
Else
End If
Next
End If
If Check1.Value = 1 Then
For i = 1 To ListView2.ListItems.Count
r3 = ListView2.ListItems(i).Text
q2 = "Select d1 from Externaldata where c1 ='6' and L1 = " + CStr(g_MS.MasterName2CodeIfExist(iname, 5)) + " and D1 =" + CStr(g_MS.MasterName2CodeIfExist(r3, 5)) + ""
Set rst2 = g_OS.GetRecordset(q2)


If rst2.RecordCount > 0 Then

ListView2.ListItems(i).Checked = True
Else
End If
Next
End If

End Sub

