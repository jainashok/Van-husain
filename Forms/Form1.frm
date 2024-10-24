VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format Master"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Delete Format"
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
      Left            =   480
      TabIndex        =   15
      Top             =   4680
      Width           =   1815
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
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   720
      Width           =   3615
   End
   Begin Taxcat.VijayList Combo1 
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.CommandButton Command4 
      Caption         =   "Exiting Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   2175
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mapping of Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   7695
      Begin VB.CommandButton Command6 
         Caption         =   "OK"
         Height          =   495
         Left            =   6720
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin Taxcat.VijayList Combo3 
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
      Begin Taxcat.VijayList Combo2 
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
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
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Remove All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   1935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Busy Field"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Excel Field"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excel Field"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Busy Field"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Format Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   13500
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   21000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If ListView1.ListItems.Count > 0 Then

ListView1.ListItems.Clear
End If
End Sub


Private Sub Combo1_GotFocus()
Combo1.List.Clear
Dim q As String
Dim rst As Recordset
q = "Select c1,d1 from externaldata where l1=4"
Set rst = g_OS.GetRecordset(q)
  If rst.RecordCount > 0 Then
  rst.MoveFirst
     Do While Not rst.EOF
     Combo1.List.AddItem rst(0)
     rst.MoveNext
Loop
End If
End Sub

Private Sub Combo1_LostFocus()
Dim q As String
Dim rst As Recordset
Dim pt As ListItem
Dim jk As String
Dim jk1() As String
Dim code2  As Long
q = "Select d1 from externaldata where l1=4 and c1 = '" + Combo1.Text + "' "
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
code2 = rst(0)
End If

q = "select c1,d1 from externaldata where i1= 5 and l1=" + CStr(code2) + ""
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
  rst.MoveFirst
     Do While Not rst.EOF
     
     jk = rst(0)
     jk1 = Split(jk, "_")
     Set pt = ListView1.ListItems.Add(, , jk1(0))
     pt.SubItems(1) = jk1(1)
     rst.MoveNext
Loop
End If

End Sub

Private Sub Combo2_GotFocus()
Combo2.List.Clear
Combo2.List.AddItem "F1"
Combo2.List.AddItem "F2"
Combo2.List.AddItem "F3"
Combo2.List.AddItem "F4"
Combo2.List.AddItem "F5"
Combo2.List.AddItem "F6"
Combo2.List.AddItem "F7"
Combo2.List.AddItem "F8"
Combo2.List.AddItem "F9"
Combo2.List.AddItem "F10"
Combo2.List.AddItem "F11"
Combo2.List.AddItem "F12"
Combo2.List.AddItem "Item Name"
Combo2.List.AddItem "Item Alias"
Combo2.List.AddItem "Item Print Name"
Combo2.List.AddItem "Item Group"
Combo2.List.AddItem "Item Unit"
Combo2.List.AddItem "Item Tax Category"
Combo2.List.AddItem "HSN Code"
Combo2.List.AddItem "Item Sale Price"
Combo2.List.AddItem "Item MRP"
Combo2.List.AddItem "Item Master OF1"
Combo2.List.AddItem "Item Master OF2"
Combo2.List.AddItem "Item Master OF3"
Combo2.List.AddItem "Item Master OF4"
Combo2.List.AddItem "Item Master OF5"
Combo2.List.AddItem "Item Master OF6"
Combo2.List.AddItem "Item Master OF7"
Combo2.List.AddItem "Item Master OF8"
Combo2.List.AddItem "Item Master OF9"
Combo2.List.AddItem "Item Master OF10"
Combo2.List.AddItem "Item Master Desc1"
Combo2.List.AddItem "Item Master Desc2"
Combo2.List.AddItem "Item Master Desc3"
Combo2.List.AddItem "Item Master Desc4"
Combo2.List.AddItem "Vch Date"
Combo2.List.AddItem "Vch No."
Combo2.List.AddItem "Vch Pur Type"
Combo2.List.AddItem "Vch Qty"
Combo2.List.AddItem "Vch List Price"
Combo2.List.AddItem "Vch Disc. %"
Combo2.List.AddItem "Vch Item Amt"
Combo2.List.AddItem "Vch BS 1 Name"
Combo2.List.AddItem "Vch BS 1 %"
Combo2.List.AddItem "Vch BS 1 Amt."
Combo2.List.AddItem "Vch BS 2 Name"
Combo2.List.AddItem "Vch BS 2 %"
Combo2.List.AddItem "Vch BS 2 Amt."
Combo2.List.AddItem "Vch BS 3 Name"
Combo2.List.AddItem "Vch BS 3 %"
Combo2.List.AddItem "Vch BS 3 Amt."
Combo2.List.AddItem "Vch BS 4 Name"
Combo2.List.AddItem "Vch BS 4 %"
Combo2.List.AddItem "Vch BS 4 Amt."
Combo2.List.AddItem "Vch OF1"
Combo2.List.AddItem "Vch OF2"
Combo2.List.AddItem "Vch OF3"
Combo2.List.AddItem "Vch OF4"
Combo2.List.AddItem "Vch OF5"
Combo2.List.AddItem "Vch OF6"
Combo2.List.AddItem "Vch OF7"
Combo2.List.AddItem "Vch OF8"
Combo2.List.AddItem "Vch OF9"
Combo2.List.AddItem "Vch OF10"


End Sub

Private Sub Combo3_GotFocus()
Combo3.List.Clear
Combo3.List.AddItem "A"
Combo3.List.AddItem "B"
Combo3.List.AddItem "C"
Combo3.List.AddItem "D"
Combo3.List.AddItem "E"
Combo3.List.AddItem "F"
Combo3.List.AddItem "G"
Combo3.List.AddItem "H"
Combo3.List.AddItem "I"
Combo3.List.AddItem "J"
Combo3.List.AddItem "K"
Combo3.List.AddItem "L"
Combo3.List.AddItem "M"
Combo3.List.AddItem "N"
Combo3.List.AddItem "O"
Combo3.List.AddItem "P"
Combo3.List.AddItem "Q"
Combo3.List.AddItem "R"
Combo3.List.AddItem "S"
Combo3.List.AddItem "T"
Combo3.List.AddItem "U"
Combo3.List.AddItem "V"
Combo3.List.AddItem "W"
Combo3.List.AddItem "X"
Combo3.List.AddItem "Y"
Combo3.List.AddItem "Z"
Combo3.List.AddItem "AA"
Combo3.List.AddItem "AB"
Combo3.List.AddItem "AC"
Combo3.List.AddItem "AD"
Combo3.List.AddItem "AE"
Combo3.List.AddItem "AF"
Combo3.List.AddItem "AG"
Combo3.List.AddItem "AH"
Combo3.List.AddItem "AI"
Combo3.List.AddItem "AJ"
Combo3.List.AddItem "AK"
Combo3.List.AddItem "AL"
Combo3.List.AddItem "AM"
Combo3.List.AddItem "AN"
Combo3.List.AddItem "AO"
Combo3.List.AddItem "AP"
Combo3.List.AddItem "AQ"
Combo3.List.AddItem "AR"
Combo3.List.AddItem "AS"
Combo3.List.AddItem "AT"
Combo3.List.AddItem "AU"
Combo3.List.AddItem "AV"
Combo3.List.AddItem "AW"
Combo3.List.AddItem "AX"
Combo3.List.AddItem "AY"
Combo3.List.AddItem "AZ"
Combo3.List.AddItem "BA"
Combo3.List.AddItem "BB"
Combo3.List.AddItem "BC"
Combo3.List.AddItem "BD"
Combo3.List.AddItem "BE"
Combo3.List.AddItem "BF"
Combo3.List.AddItem "BG"
Combo3.List.AddItem "BH"
Combo3.List.AddItem "BI"
Combo3.List.AddItem "BJ"
Combo3.List.AddItem "BK"
Combo3.List.AddItem "BL"
Combo3.List.AddItem "BM"
Combo3.List.AddItem "BN"
Combo3.List.AddItem "BO"
Combo3.List.AddItem "BP"
Combo3.List.AddItem "BQ"
Combo3.List.AddItem "BR"
Combo3.List.AddItem "BS"
Combo3.List.AddItem "BT"
Combo3.List.AddItem "BU"
Combo3.List.AddItem "BV"
Combo3.List.AddItem "BW"
Combo3.List.AddItem "BX"
Combo3.List.AddItem "BY"
Combo3.List.AddItem "BZ"
Combo3.List.AddItem "CA"
Combo3.List.AddItem "CB"
Combo3.List.AddItem "CC"
Combo3.List.AddItem "CD"
Combo3.List.AddItem "CE"
Combo3.List.AddItem "CF"
Combo3.List.AddItem "CG"
Combo3.List.AddItem "CH"
Combo3.List.AddItem "CI"
Combo3.List.AddItem "CJ"
Combo3.List.AddItem "CK"
Combo3.List.AddItem "CL"
Combo3.List.AddItem "CM"
Combo3.List.AddItem "CN"
Combo3.List.AddItem "CO"
Combo3.List.AddItem "CP"
Combo3.List.AddItem "CQ"
Combo3.List.AddItem "CR"
Combo3.List.AddItem "CS"
Combo3.List.AddItem "CT"
Combo3.List.AddItem "CU"
Combo3.List.AddItem "CV"
Combo3.List.AddItem "CW"
Combo3.List.AddItem "CX"
Combo3.List.AddItem "CY"
Combo3.List.AddItem "CZ"
Combo3.List.AddItem "DA"
Combo3.List.AddItem "DB"
Combo3.List.AddItem "DC"
Combo3.List.AddItem "DD"
Combo3.List.AddItem "DE"
Combo3.List.AddItem "DF"
Combo3.List.AddItem "DG"
Combo3.List.AddItem "DH"
Combo3.List.AddItem "DI"
Combo3.List.AddItem "DJ"
Combo3.List.AddItem "DK"
Combo3.List.AddItem "DL"
Combo3.List.AddItem "DM"
Combo3.List.AddItem "DN"
Combo3.List.AddItem "DO"
Combo3.List.AddItem "DP"
Combo3.List.AddItem "DQ"
Combo3.List.AddItem "DR"
Combo3.List.AddItem "DS"
Combo3.List.AddItem "DT"
Combo3.List.AddItem "DU"
Combo3.List.AddItem "DV"
Combo3.List.AddItem "DW"
Combo3.List.AddItem "DX"
Combo3.List.AddItem "DY"
Combo3.List.AddItem "DZ"
Combo3.List.AddItem "EA"
Combo3.List.AddItem "EB"
Combo3.List.AddItem "EC"
Combo3.List.AddItem "ED"
Combo3.List.AddItem "EE"
Combo3.List.AddItem "EF"
Combo3.List.AddItem "EG"
Combo3.List.AddItem "EH"
Combo3.List.AddItem "EI"
Combo3.List.AddItem "EJ"
Combo3.List.AddItem "EK"
Combo3.List.AddItem "EL"
Combo3.List.AddItem "EM"
Combo3.List.AddItem "EN"
Combo3.List.AddItem "EO"
Combo3.List.AddItem "EP"
Combo3.List.AddItem "EQ"
Combo3.List.AddItem "ER"
Combo3.List.AddItem "ES"
Combo3.List.AddItem "ET"
Combo3.List.AddItem "EU"
Combo3.List.AddItem "EV"
Combo3.List.AddItem "EW"
Combo3.List.AddItem "EX"
Combo3.List.AddItem "EY"
Combo3.List.AddItem "EZ"
Combo3.List.AddItem "FA"
Combo3.List.AddItem "FB"
Combo3.List.AddItem "FC"
Combo3.List.AddItem "FD"
Combo3.List.AddItem "FE"
Combo3.List.AddItem "FF"
Combo3.List.AddItem "FG"
Combo3.List.AddItem "FH"
Combo3.List.AddItem "FI"
Combo3.List.AddItem "FJ"
Combo3.List.AddItem "FK"
Combo3.List.AddItem "FL"
Combo3.List.AddItem "FM"
Combo3.List.AddItem "FN"
Combo3.List.AddItem "FO"
Combo3.List.AddItem "FP"
Combo3.List.AddItem "FQ"
Combo3.List.AddItem "FR"
Combo3.List.AddItem "FS"
Combo3.List.AddItem "FT"
Combo3.List.AddItem "FU"
Combo3.List.AddItem "FV"
Combo3.List.AddItem "FW"
Combo3.List.AddItem "FX"
Combo3.List.AddItem "FY"
Combo3.List.AddItem "FZ"
Combo3.List.AddItem "GA"
Combo3.List.AddItem "GB"
Combo3.List.AddItem "GC"
Combo3.List.AddItem "GD"
Combo3.List.AddItem "GE"
Combo3.List.AddItem "GF"
Combo3.List.AddItem "GG"
Combo3.List.AddItem "GH"
Combo3.List.AddItem "GI"
Combo3.List.AddItem "GJ"
Combo3.List.AddItem "GK"
Combo3.List.AddItem "GL"
Combo3.List.AddItem "GM"
Combo3.List.AddItem "GN"
Combo3.List.AddItem "GO"
Combo3.List.AddItem "GP"
Combo3.List.AddItem "GQ"
Combo3.List.AddItem "GR"
Combo3.List.AddItem "GS"
Combo3.List.AddItem "GT"
Combo3.List.AddItem "GU"
Combo3.List.AddItem "GV"
Combo3.List.AddItem "GW"
Combo3.List.AddItem "GX"
Combo3.List.AddItem "GY"
Combo3.List.AddItem "GZ"
Combo3.List.AddItem "HA"
Combo3.List.AddItem "HB"
Combo3.List.AddItem "HC"
Combo3.List.AddItem "HD"
Combo3.List.AddItem "HE"
Combo3.List.AddItem "HF"
Combo3.List.AddItem "HG"
Combo3.List.AddItem "HH"
Combo3.List.AddItem "HI"
Combo3.List.AddItem "HJ"
Combo3.List.AddItem "HK"
Combo3.List.AddItem "HL"
Combo3.List.AddItem "HM"
Combo3.List.AddItem "HN"
Combo3.List.AddItem "HO"
Combo3.List.AddItem "HP"
Combo3.List.AddItem "HQ"
Combo3.List.AddItem "HR"
Combo3.List.AddItem "HS"
Combo3.List.AddItem "HT"
Combo3.List.AddItem "HU"
Combo3.List.AddItem "HV"
Combo3.List.AddItem "HW"
Combo3.List.AddItem "HX"
Combo3.List.AddItem "HY"
Combo3.List.AddItem "HZ"
Combo3.List.AddItem "IA"
Combo3.List.AddItem "IB"
Combo3.List.AddItem "IC"
Combo3.List.AddItem "ID"
Combo3.List.AddItem "IE"
Combo3.List.AddItem "IF"
Combo3.List.AddItem "IG"
Combo3.List.AddItem "IH"
Combo3.List.AddItem "II"
Combo3.List.AddItem "IJ"
Combo3.List.AddItem "IK"
Combo3.List.AddItem "IL"
Combo3.List.AddItem "IM"
Combo3.List.AddItem "IN"
Combo3.List.AddItem "IO"
Combo3.List.AddItem "IP"
Combo3.List.AddItem "IQ"
Combo3.List.AddItem "IR"
Combo3.List.AddItem "IS"
Combo3.List.AddItem "IT"
Combo3.List.AddItem "IU"
Combo3.List.AddItem "IV"
Combo3.List.AddItem "IW"
Combo3.List.AddItem "IX"
Combo3.List.AddItem "IY"
Combo3.List.AddItem "IZ"









End Sub

Private Sub Command1_Click()
Dim str As String
Dim num1 As Integer
Dim Num2 As Integer
Dim num3 As Integer
Dim counter As Integer
Dim fc As String
Dim sc As String
Dim fc1 As String
num1 = 0
Num2 = 0
num3 = 0
Dim pt As ListItem
Dim i As Integer
Dim q As String
Dim rst As Recordset
Dim b As Boolean
Dim code1 As Long
code1 = 1000
Dim code2 As Long
b = False
If Label4.Caption = "1" Then
q = "Select c1,d1 from externaldata where l1=4"
Set rst = g_OS.GetRecordset(q)
  If rst.RecordCount > 0 Then
  rst.MoveFirst
     Do While Not rst.EOF
         If Trim(Text1.Text) = Trim(rst(0)) Then
         MsgBox " This Format is Already Saved"
         b = True
         Exit Do
         Else
         code1 = rst(1)
         End If
      rst.MoveNext
      Loop
   Else
   End If


    If b = False Then
    code1 = code1 + 1
    q = "Insert into externaldata(L1,D1,c1) Values(4," + CStr(code1) + ",'" + Text1.Text + "')"
    g_OS.ExecuteQuerytmp (q)
    Label5.Caption = Text1.Text
    End If
End If
If Label4.Caption = "2" Then
Label5.Caption = Combo1.Text
End If
q = "Select d1 from externaldata where l1=4 and c1 = '" + Label5.Caption + "' "
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
code2 = rst(0)
End If

If Label4.Caption = "2" Or (Label4.Caption = "1" And b = False) Then
q = "delete * from externaldata where i1=5 and l1=" + CStr(code2) + ""
g_OS.ExecuteQuerytmp (q)

For i = 1 To ListView1.ListItems.Count

str = ListView1.ListItems(i).SubItems(1)

str = Trim(str) ' Remove All Spaces
If Len(str) = 1 Then
fc = str
'MsgBox FC
End If
If Len(str) = 2 Then
fc1 = Mid(str, 1, 1)

sc = Mid(str, 2, 1)
'MsgBox FC
'MsgBox SC
End If

If fc = "A" Then
num1 = 1
End If
If fc = "B" Then
num1 = 2
End If
If fc = "C" Then
num1 = 3
End If
If fc = "D" Then
num1 = 4
End If
If fc = "E" Then
num1 = 5
End If
If fc = "F" Then
num1 = 6
End If
If fc = "G" Then
num1 = 7
End If
If fc = "H" Then
num1 = 8
End If
If fc = "I" Then
num1 = 9
End If
If fc = "J" Then
num1 = 10
End If
If fc = "K" Then
num1 = 11
End If
If fc = "L" Then
num1 = 12
End If
If fc = "M" Then
num1 = 13
End If
If fc = "N" Then
num1 = 14
End If
If fc = "O" Then
num1 = 15
End If
If fc = "P" Then
num1 = 16
End If
If fc = "Q" Then
num1 = 17
End If
If fc = "R" Then
num1 = 18
End If
If fc = "S" Then
num1 = 19
End If
If fc = "T" Then
num1 = 20
End If
If fc = "U" Then
num1 = 21
End If
If fc = "V" Then
num1 = 22
End If
If fc = "W" Then
num1 = 23
End If
If fc = "X" Then
num1 = 24
End If
If fc = "Y" Then
num1 = 25
End If
If fc = "Z" Then
num1 = 26
End If
If sc = "A" Then
Num2 = 1
End If
If sc = "B" Then
Num2 = 2
End If
If sc = "C" Then
Num2 = 3
End If
If sc = "D" Then
Num2 = 4
End If
If sc = "E" Then
Num2 = 5
End If
If sc = "F" Then
Num2 = 6
End If
If sc = "G" Then
Num2 = 7
End If
If sc = "H" Then
Num2 = 8
End If
If sc = "I" Then
Num2 = 9
End If
If sc = "J" Then
Num2 = 10
End If
If sc = "K" Then
Num2 = 11
End If
If sc = "L" Then
Num2 = 12
End If
If sc = "M" Then
Num2 = 13
End If
If sc = "N" Then
Num2 = 14
End If
If sc = "O" Then
Num2 = 15
End If
If sc = "P" Then
Num2 = 16
End If
If sc = "Q" Then
Num2 = 17
End If
If sc = "R" Then
Num2 = 18
End If
If sc = "S" Then
Num2 = 19
End If
If sc = "T" Then
Num2 = 20
End If
If sc = "U" Then
Num2 = 21
End If
If sc = "V" Then
Num2 = 22
End If
If sc = "W" Then
Num2 = 23
End If
If sc = "X" Then
Num2 = 24
End If
If sc = "Y" Then
Num2 = 25
End If
If sc = "Z" Then
Num2 = 26
End If
If fc1 = "A" Then
num1 = 26
End If
If fc1 = "B" Then
num1 = 52
End If
If fc1 = "C" Then
num1 = 78
End If
If fc1 = "D" Then
num1 = 104
End If
If fc1 = "E" Then
num1 = 130
End If
If fc1 = "F" Then
num1 = 156
End If
If fc1 = "G" Then
num1 = 182
End If
If fc1 = "H" Then
num1 = 208
End If
If fc1 = "I" Then
num1 = 234
End If

If Len(str) = 1 Then
num3 = num1
End If
If Len(str) = 2 Then
num3 = num1 + Num2
End If
Dim jk As String
jk = ListView1.ListItems(i).Text + "_" + ListView1.ListItems(i).SubItems(1)
q = "Insert into externaldata(i1,L1,D1,c1) Values(5," + CStr(code2) + "," + CStr(num3) + ",'" + jk + "')"
g_OS.ExecuteQuerytmp (q)
Next
MsgBox "Template Saved"
End If

End Sub

Private Sub Command3_Click()
Combo1.Visible = False
Label1.Visible = True
Text1.Visible = True
Text1.Left = 2520
Text1.Top = 720
Frame1.Visible = True
Label4.Caption = "1"

End Sub

Private Sub Command4_Click()
Combo1.Visible = True
Label1.Visible = True
Text1.Visible = False
Combo1.Left = 2520
Combo1.Top = 720
Frame1.Visible = True
Label4.Caption = "2"
End Sub

Private Sub Command5_Click()
Dim q As String
Dim rst As Recordset
Dim code2 As Long
If Label4.Caption = "2" Then
Label5.Caption = Combo1.Text
'End If
q = "Select d1 from externaldata where l1=4 and c1 = '" + Label5.Caption + "' "
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
code2 = rst(0)
End If

'If Label4.Caption = "2" Then
q = "delete * from externaldata where i1=5 and l1=" + CStr(code2) + ""
g_OS.ExecuteQuerytmp (q)
q = "delete * from externaldata where i1=4 and l1=" + CStr(code2) + ""
g_OS.ExecuteQuerytmp (q)

End If
End Sub

Private Sub Command6_Click()

If Combo2.Text <> "" And Combo3.Text <> "" Then
Dim i As Integer
Dim t As Boolean
t = False
Dim p As Integer
Dim pt As ListItem
If Label6.Caption = "0" Then
Set pt = ListView1.ListItems.Add(, , Combo2.Text)
pt.SubItems(1) = Combo3.Text
Else
p = Label6.Caption
For i = 1 To ListView1.ListItems.Count
If Trim(Combo2.Text) = Trim(ListView1.ListItems(i).Text) Then
t = True
MsgBox "Already in List"
Exit For
Else
t = False
End If
Next
If t = False Then
ListView1.ListItems(p).SubItems(1) = Combo1.Text
End If
End If
End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim q As String
Dim rst As Recordset
Dim i As Integer
Dim ac As New Collection
Dim xt As ListItem
Dim n As Integer
Dim inccount As Integer
Dim q2 As String
Dim rst2 As Recordset
Dim r3 As String
Dim iname As String
If ListView1.ListItems.Count > 0 Then
'ListView2.ListItems.Clear
For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems(i).Checked = True Then
    Dim p As Integer
    'g_CL.FlushCol ac
    p = Label6.Caption
    If p = 0 Then
    Label6.Caption = ListView1.ListItems(i).Index
    Combo2.Text = ListView1.ListItems(i).Text
    Combo3.Text = ListView1.ListItems(i).SubItems(1)
    'Label1.Visible = True
    'Text1.Visible = True
    Combo2.SetFocus
    Else
    ListView1.ListItems(p).Checked = False
    Label6.Caption = ListView1.ListItems(i).Index
    Combo2.Text = ListView1.ListItems(i).Text
    Combo3.Text = ListView1.ListItems(i).SubItems(1)
    'Label1.Visible = True
    'Text1.Visible = True
    Combo2.SetFocus
    End If
End If
Next

End If

End Sub

