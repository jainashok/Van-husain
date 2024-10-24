VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Update Account"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7080
   LinkTopic       =   "Form5"
   ScaleHeight     =   7845
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin Taxcat.VijayList V1 
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   3600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      InvalidCharacters=   ""
      AutoSelect      =   -1  'True
      EnterKeySupport =   -1  'True
      Text            =   ""
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   22
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   21
      Top             =   4680
      Width           =   3975
   End
   Begin Taxcat.VijayList G1 
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   4320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      InvalidCharacters=   ""
      AutoSelect      =   -1  'True
      EnterKeySupport =   -1  'True
      Text            =   ""
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Taxcat.VijayList P1 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      InvalidCharacters=   ""
      AutoSelect      =   -1  'True
      EnterKeySupport =   -1  'True
      Text            =   ""
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   19
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   16
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   15
      Top             =   3240
      Width           =   4000
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   14
      Top             =   2880
      Width           =   4000
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   13
      Top             =   2520
      Width           =   4000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   12
      Top             =   2160
      Width           =   4000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create New Account In Busy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   25
      Top             =   5160
      Width           =   870
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Station"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   24
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acc. Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   23
      Top             =   4440
      Width           =   1155
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   18
      Top             =   1230
      Width           =   1665
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GST No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   4
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Busy Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Code "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1665
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim InvalMsg As String
Dim Accdata As BusyDDC2175.udtAccMast
Dim AccMast As Busy2175.CAccMast
Dim AccAlias As BusyDDC2175.udtGeneral
Dim col2 As Collection
Dim t As String
Dim MYS As String
Dim MYS1 As String
Dim check_Alias As Boolean
Dim i As Integer
Dim udtAddInfo As BusyDDC2175.udtMasterAddressinfo




        Set AccMast = New Busy2175.CAccMast
    
    If AccMast.Load2(P1.Text) Then
        Accdata = AccMast.GetState
        
        With Accdata
        
            g_CL.FlushCol col2
            Set col2 = .MultiAliasCol
            Set .MultiAliasCol = Nothing
            udtAddInfo.C1 = Trim(Label5.Caption)    'Alias Name
            'AccAlias.S1 = Trim(Label5.Caption)
            .MultipleAliasBasis = 0
            col2.Add udtAddInfo
            Set .MultiAliasCol = col2
            
        End With
        'Set ItemMast = New Busy2175.CItemMast
         AccMast.SetState Accdata

            If AccMast.CanBeSaved(InvalMsg) Then
                If AccMast.Save(InvalMsg) Then
                    'MsgBox "Bill finished successfuly."
                Else
                    MsgBox InvalMsg
                End If
            End If

    End If

MsgBox "Saved"
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub VijayList1_Change()

End Sub

Private Sub Command3_Click()
Dim o As Integer
Dim InvalMsg As String
Dim Accdata As BusyDDC2175.udtAccMast
Dim AccMast As Busy2175.CAccMast
Dim rs1 As Recordset
Dim kt As Integer
Dim ktr As Recordset
Dim pcode As Long
Dim ih  As Integer
ih = 0

pcode = g_MS.MasterName2CodeIfExist(Trim(Label4.Caption), 2)
If pcode <> "0" Then
ih = 1
Else
ih = 2
 End If
 If ih <> 0 And Trim(Text6.Text) <> "" And V1.Text <> "" And G1.Text <> "" Then
g_OS.InitUDTAccMast Accdata
    With Accdata
    
        .Name = Trim(Text6.Text)
        
    

        .PrintName = Trim(Label4.Caption)
        .ParentGrpName = G1.Text
        .OpBal = CDbl(0#)
        .BillByBillBalancing = True
        .ChequePrintName = Trim(Label4.Caption)
        .udtAddressInfo.Address1 = Trim(Text1.Text)
        .udtAddressInfo.Address2 = Trim(Text2.Text)
        .udtAddressInfo.Address2 = Trim(Text3.Text)
        .udtAddressInfo.Address2 = Trim(Text4.Text)
        .udtAddressInfo.GSTNo = Trim(Text5.Text)
        .udtAddressInfo.ITPAN = Mid(Trim(Text5.Text), 3, 10)
        If V1.Text = "" Then
        .udtAddressInfo.StateName = "---Others---"
        .udtAddressInfo.CountryName = "---Others---"
        Else
        
        .udtAddressInfo.CountryName = "India"
         .udtAddressInfo.CountryCode = g_MS.MasterName2Code("India", COUNTRY_MAST)
       .udtAddressInfo.StateName = V1.Text
         .udtAddressInfo.StateCode = g_MS.MasterName2Code(V1.Text, 56)
        End If
       
        If .udtAddressInfo.GSTNo = "" Then
        .TypeOfDealer = 0
        Else
        .TypeOfDealer = 1
        End If
        .udtAddressInfo.Station = Trim(Text7.Text)
        .udtAddressInfo.PINCode = Trim(Text8.Text)
    End With
    
    Set AccMast = New Busy2175.CAccMast
         AccMast.SetState Accdata

If AccMast.CanBeSaved(InvalMsg) Then
    If AccMast.Save(InvalMsg) Then
        MsgBox "Account Created successfuly."
    Else
        MsgBox InvalMsg
    End If
End If

End If


End Sub

Private Sub Form_Load()
Label5.Caption = g1_c
Label4.Caption = g1_n
Text5.Text = g1_p
Dim ih  As Integer
ih = 0
G1.Text = "Sundry Debtors"
pcode = g_MS.MasterName2CodeIfExist(Trim(Label4.Caption), 2)
If pcode <> "0" Then
ih = 1
Text6.Text = ""
Text6.Enabled = True
Else
ih = 2
Text6.Text = Label4.Caption
Text6.Enabled = False
 End If

End Sub

Private Sub G1_GotFocus()
G1.List.Clear
Dim q As String
Dim rst As Recordset
q = "Select Name from Master1 where mastertype=1"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF

G1.List.AddItem rst(0)
rst.MoveNext
Loop
End If
Set rst = Nothing

End Sub

Private Sub P1_GotFocus()
P1.List.Clear
Dim q As String
Dim rst As Recordset
q = "SELECT NAMEALIAS FROM HELP1 WHERE RECTYPE = " & CStr(H1_PARTY) & " AND NAMEORALIAS=" & CStr(NA_NAME)
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
P1.List.AddItem rst!NameAlias.Value
rst.MoveNext
Loop
End If

End Sub

Private Sub P1_LostFocus()
If P1.Text <> "" Then
Dim q As String
Dim rst As Recordset
q = "Select parentgrp from master1 where code = " + CStr(g_MS.MasterName2Code(P1.Text, 2)) + ""
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
G1.Text = g_MS.MasterCode2Name(rst(0))
End If
End If
End Sub

Private Sub V1_GotFocus()
V1.List.Clear
Dim q As String
Dim rst As Recordset
q = "Select name from master1 where mastertype =56"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V1.List.AddItem rst(0)

rst.MoveNext
Loop
End If
End Sub
