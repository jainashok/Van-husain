VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Import"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20400
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   20400
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView3 
      Height          =   3015
      Left            =   10440
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Busy Filed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Excel coloum no."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Excel Col"
         Object.Width           =   2540
      EndProperty
   End
   Begin Taxcat.VijayList Combo1 
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Fourmla Using"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   8295
      Begin Taxcat.VijayList Combo2 
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.CommandButton Command1 
         Caption         =   "Fourmula Used"
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
         Left            =   6200
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2295
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4048
         View            =   3
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fourmula Name"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Reset"
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
      Left            =   12960
      TabIndex        =   6
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   690
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   8295
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
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   4695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7680
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Save From Excel"
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
         Left            =   5400
         TabIndex        =   4
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
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
      Left            =   14280
      TabIndex        =   2
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
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
      Left            =   12720
      TabIndex        =   1
      Top             =   8880
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   375
      Left            =   13440
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Format"
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
      Left            =   720
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   13500
      Left            =   -120
      Picture         =   "Form4.frx":0000
      Top             =   -120
      Width           =   21000
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim fc As Long
Dim pt As ListItem
ListView2.ListItems.Clear
q = "Select d1 from externaldata where l1=4 and c1='" + Trim(Combo1.Text) + "'"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
fc = rst(0)
End If
q = "Select c1 from externaldata where i1=6 and l1=" + CStr(fc) + " "
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
Set pt = ListView2.ListItems.Add(, , rst(0))
rst.MoveNext
Loop
End If
ListView3.ListItems.Clear
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
     Set pt = ListView3.ListItems.Add(, , jk1(0))
     pt.SubItems(1) = rst(1)
    pt.SubItems(2) = jk1(1)
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

Private Sub ListView4_ItemCheck(ByVal Item As MSComctlLib.ListItem)
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
If ListView4.ListItems.Count > 0 Then
'ListView2.ListItems.Clear
For i = 1 To ListView4.ListItems.Count
If ListView4.ListItems(i).Checked = True Then
    Dim p As Integer
    'g_CL.FlushCol ac
    p = Label6.Caption
    If p = 0 Then
    Label6.Caption = ListView4.ListItems(i).Index
    Combo2.Text = ListView4.ListItems(i).Text
    Combo3.Text = ListView4.ListItems(i).SubItems(1)
    'Label1.Visible = True
    'Text1.Visible = True
    Combo2.SetFocus
    Else
    ListView4.ListItems(p).Checked = False
    Label6.Caption = ListView4.ListItems(i).Index
    Combo2.Text = ListView4.ListItems(i).Text
    Combo3.Text = ListView4.ListItems(i).SubItems(1)
    'Label1.Visible = True
    'Text1.Visible = True
    Combo2.SetFocus
    End If
End If
Next

End If

End Sub

Private Sub Command1_Click()
Dim q As String
Dim rst As Recordset
Dim fc As Long
Dim fcc As Long
Dim tp As Integer
q = "Select d1 from externaldata where l1=4 and c1='" + Trim(Combo1.Text) + "'"
Set rst = g_OS.GetRecordset(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    fc = rst(0)
    End If
q = "Select d1 from externaldata where i1=6 and l1=" + CStr(fc) + " and c1='" + Trim(Label1.Caption) + "'"
Set rst = g_OS.GetRecordset(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    
    fcc = rst(0)
    End If
q = "Select d1,b1 from externaldata where i1=7 and l1=" + CStr(fcc) + " "
Set rst = g_OS.GetRecordset(q)
       
If rst.RecordCount > 0 Then
rst.MoveFirst
tp = rst(0)
End If

If tp = 1 Then
q = "Select t1,t2 from msconfig where rectype=8 and fcode= " + CStr(fcc) + " and ftype=1"
Set rst = g_OS.GetRecordsetFromDB(q)
    If rst.RecordCount >= 0 Then
        rst.MoveFirst
        Dim jk1() As String
        Dim intCount As Integer
        Dim nv
        Dim i As Integer
        jk1 = Split(rst(1), "+")
        For i = 1 To ListView1.ListItems.Count
                nv = ""
                For intCount = LBound(jk1) To UBound(jk1)
                      If jk1(intCount) = 1 Then
                        If nv = "" Then
                            nv = nv + ListView1.ListItems(i).Text
                        Else
                            nv = nv + " " + ListView1.ListItems(i).Text
                        End If
                      Else
                        If nv = "" Then
                        nv = nv + ListView1.ListItems(i).SubItems(jk1(intCount) - 1)
                        Else
                        nv = nv + " " + ListView1.ListItems(i).SubItems(jk1(intCount) - 1)
                        End If
                      End If
                Next
                If jk1(0) = 1 Then
                ListView1.ListItems(i).Text = nv
               Else
                  ListView1.ListItems(i).SubItems(jk1(intCount) - 1) = nv
                  End If
        Next
    End If
End If

If tp = 2 Then
        'q1 = "insert into msconfig (SSC,SSV,Con1,SCC,SCV,T1,SSC1,SCC1)values(8," + CStr(fcc) + ",2,'" + Trim(Combo3.Text) + "','" + Trim(Text3.Text) + "'," + CStr(mj) + ",'" + Trim(Combo5.Text) + "','" + Trim(Text4.Text) + "','" + Trim(Text5.Text) + "','" + Trim(Label16.Caption) + "','" + Trim(Label17.Caption) + "')"
        q = "Select SSC,SSV,Con1,SCC,SCV,T1,SSC1,SCC1 from msconfig where rectype=8 and fcode= " + CStr(fcc) + " and ftype=2"
        Set rst = g_OS.GetRecordsetFromDB(q)
        If rst.RecordCount >= 0 Then
            rst.MoveFirst
            If Combo4.Text = "=" Then
            mj = 1
            End If
            If Combo4.Text = ">" Then
            mj = 2
            End If
            If Combo4.Text = ">=" Then
            mj = 3
            End If
            If Combo4.Text = "<" Then
            mj = 4
            End If
            If Combo4.Text = "<=" Then
            mj = 5
            End If
            If Combo4.Text = "<>" Then
            mj = 6
            End If
            
            If rst(2) = 1 Then
                If rst(7) = "1" Then
                For i = 1 To ListView1.ListItems.Count
                If Trim(ListView1.ListItems(i).Text) = Trim(rst(4)) Then
                If Trim(ListView1.ListItems(i).SubItems(rst(6) - 1)) = rst(1) Then
                ListView1.ListItems(i).SubItems(rst(6) - 1) = rst(5)
                End If
                End If
                Next
                End If
            If rst(7) <> "1" Then
            If rst(6) = 1 Then
            
                For i = 1 To ListView1.ListItems.Count
                If Trim(ListView1.ListItems(i).SubItems(rst(7) - 1)) = Trim(rst(4)) Then
                ListView1.ListItems(i).Text = rst(5)
                Else
                ListView1.ListItems(i).SubItems(rst(6) - 1) = rst(5)
                End If
            
 Next
        
                If rst(2) = 2 Then
                
                End If
                If rst(2) = 3 Then
                
                End If
                If rst(2) = 4 Then
                
                End If
                If rst(2) = 5 Then
                
                End If
                If rst(2) = 6 Then
                
                End If
                
        
            End If
            End If
            End If
            End If
            End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Frame2.Visible = True
ListView1.Visible = True
Dim lt As ListItem
Dim jt As ColumnHeader
If Text1.Text <> "" Then
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim m As Integer
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.workbooks.Open (Text1.Text)
Dim sht As Integer
sht = 1
    Set ExcelBook = ExcelObj.workbooks(1)
    Set ExcelSheet = ExcelBook.WorkSheets(sht)

    'Dim l As ListItem
    ListView1.ListItems.Clear
    With ExcelSheet
    Dim rs As Integer
    Dim SrNo As String

  '  SrNo = g_DN
    

    'rs = Text2.Text
    'i = rs
    
    Dim fc As Integer
    Dim sc As Integer
    'fc = Text5.Text
    'sc = Text4.Text
    Dim er As Integer
   ' If SrNo = "12069007" Or SrNo = "12120261" Or SrNo = "15020229" Then

    'er = Text3.Text
    
'Else
'        If CDbl(Text3.Text) > (CDbl(Text2.Text) + 9) Then
'            er = (CDbl(Text2.Text) + 9)
'        Else
'            er = CDbl(Text3.Text)
'        End If
'    End If
    'Label6.Visible = True
    For i = 1 To ListView3.ListItems.Count
    Dim poi As String
    poi = Trim(ListView3.ListItems(i).Text) + "+" + Trim(ListView3.ListItems(i).SubItems(2))
    Set jt = ListView1.ColumnHeaders.Add(, , poi, 3000)
    
    
    Next
    Dim nc As Boolean
    'ListView1.Visible = True
    
    Do Until .cells(i, 1) & "" = ""
    'nc = IsNumeric(.cells(i, sc))
    'If nc = True Then
    'If CDbl(.cells(i, sc)) < 0 Then
    'Else
    For m = 1 To ListView3.ListItems.Count
    fc = ListView3.ListItems(m).SubItems(1)
    If m = 1 Then
        Set lt = ListView1.ListItems.Add(, , .cells(i, fc))
'.cells(i, sc) = Format(.cells(i, sc), "#0.00")
'.cells(i, sc) = g_CL.FormatNum(.cells(i, sc), 10, 2, False, True)
Else
        lt.SubItems(m - 1) = .cells(i, fc)
        End If
        Next
        'l.SubItems(2) = .cells(i, 3)
        'l.SubItems(3) = .cells(i, 4)
        'End If
        'Else
        'End If
        i = i + 1
        'If i > er Then
        'Exit Do
        'End If

    Loop

    End With
'ExcelObj.Activeworkbook.Save
    ExcelObj.Activeworkbook.Close
    'ExcelObj.Activeworkbook.Save
    ExcelObj.quit

    Set ExcelSheet = Nothing

    Set ExcelBook = Nothing

    Set ExcelObj = Nothing
    'Label6.Visible = False
    MsgBox "Data Import Successfully", vbOKOnly, "Import Data"
    'Combo2.Visible = True
    'Label19.Visible = True
    
    'Combo2.SetFocus
    
    End If
eh:



'End If
End Sub

Private Sub Command8_Click()
'If Frame1.Caption = Command4.Caption Then
'Dim sTempDir As String
'    On Error Resume Next
'    sTempDir = CurDir    'Remember the current active directory
'    CommonDialog1.DialogTitle = "Select a directory" 'titlebar
'    CommonDialog1.InitDir = App.Path 'start dir, might be "C:\" or so also
'    CommonDialog1.FileName = "Select a Directory"  'Something in filenamebox
'    CommonDialog1.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
'    CommonDialog1.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
'    CommonDialog1.CancelError = True 'allow escape key/cancel
'    CommonDialog1.ShowSave   'show the dialog screen
'
'    If Err <> 32755 Then    ' User didn't chose Cancel.
'        Text1.Text = CurDir + "\Price List.xlsx"
'    End If

    'ChDir sTempDir  'restore path to what it was at enteringEnd If
'End If
'If Frame1.Caption = Command5.Caption Then
CommonDialog1.Filter = "Apps (*.xls)|*.xlsx|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "xls"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen

'The FileName property gives you the variable you need to use
Text1.Text = CommonDialog1.FileName
'End If

End Sub


Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
For i = 1 To ListView2.ListItems.Count
If ListView2.ListItems(i).Checked = True Then
    Dim p As Integer
   ' g_CL.FlushCol ac
    p = Label3.Caption
    If p = 0 Then
    Label3.Caption = ListView2.ListItems(i).Index
    Label1.Caption = ListView2.ListItems(i).Text
    Label1.Visible = True
    Combo2.Visible = True
    Command1.Visible = True
    Else
    ListView2.ListItems(p).Checked = False
    Label3.Caption = ListView2.ListItems(i).Index
    Label1.Caption = ListView2.ListItems(i).Text
    Label1.Visible = True
    Combo2.Visible = True
    Command1.Visible = True
    End If
End If
Next

End Sub
