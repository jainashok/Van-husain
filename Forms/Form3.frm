VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00C0C000&
   Caption         =   "Import Data From Excel"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14100
   LinkTopic       =   "Form3"
   ScaleHeight     =   9150
   ScaleWidth      =   14100
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   7920
      TabIndex        =   48
      Top             =   6000
      Width           =   375
   End
   Begin Taxcat.VijayList T1 
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
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
   Begin Taxcat.VijayList V1 
      Height          =   255
      Left            =   2505
      TabIndex        =   1
      Top             =   1080
      Width           =   4395
      _ExtentX        =   7752
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
   Begin Taxcat.VijayList V2 
      Height          =   255
      Left            =   2505
      TabIndex        =   2
      Top             =   1560
      Width           =   4395
      _ExtentX        =   7752
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
   Begin Taxcat.VijayList V4 
      Height          =   255
      Left            =   2505
      TabIndex        =   4
      Top             =   2520
      Width           =   4395
      _ExtentX        =   7752
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
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   6000
      TabIndex        =   45
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   2055
      Left            =   9840
      TabIndex        =   44
      Top             =   5400
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3625
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
         Text            =   "Aname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Gst no."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C000&
      Caption         =   "GST No. Updating"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5880
      TabIndex        =   43
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Br&owse"
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
      Left            =   8880
      TabIndex        =   42
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
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
      Height          =   285
      Left            =   2520
      TabIndex        =   41
      Top             =   2880
      Visible         =   0   'False
      Width           =   6135
   End
   Begin Taxcat.VijayList G1 
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   5160
      Width           =   2295
      _ExtentX        =   4048
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
   Begin Taxcat.VijayList VL9 
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   5640
      Width           =   4575
      _ExtentX        =   8070
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
   Begin Taxcat.VijayList VS1 
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   503
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C000&
      Caption         =   "Make Bill Ref."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   38
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C000&
      Caption         =   "Vch Importing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   6360
      Width           =   2295
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
      Left            =   2640
      TabIndex        =   10
      Top             =   5160
      Width           =   495
   End
   Begin Taxcat.VijayList V3 
      Height          =   255
      Left            =   2505
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   450
      InvalidCharacters=   ""
      AutoSelect      =   -1  'True
      EnterKeySupport =   -1  'True
      Text            =   ""
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   0   'False
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
   Begin MSComctlLib.ListView ListView3 
      Height          =   2055
      Left            =   11520
      TabIndex        =   33
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3625
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
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Excel C"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2175
      Left            =   10920
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3836
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
         Text            =   "filed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Coloum"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   720
      TabIndex        =   31
      Top             =   7080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Taxcat.VijayList S1 
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   4680
      Width           =   4395
      _ExtentX        =   7752
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
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C000&
      Caption         =   "Party Spf. Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   29
      Top             =   4245
      Width           =   2000
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C000&
      Caption         =   "Party Name in Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   28
      Top             =   4245
      Width           =   2500
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C000&
      Caption         =   "Multiple Codes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   4245
      Width           =   2000
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
      Height          =   495
      Left            =   12480
      TabIndex        =   25
      Top             =   8280
      Width           =   1095
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
      Height          =   285
      Left            =   7800
      TabIndex        =   8
      Top             =   3720
      Width           =   735
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
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   3720
      Width           =   735
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
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   3720
      Width           =   735
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
      Left            =   2505
      TabIndex        =   5
      Top             =   3240
      Width           =   6135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   495
      Left            =   11160
      TabIndex        =   15
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Browse"
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
      Left            =   8880
      TabIndex        =   24
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TCS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   6960
      TabIndex        =   47
      Top             =   6000
      Width           =   465
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Round Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   4680
      TabIndex        =   46
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select A/C Mast Sheet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   40
      Top             =   2880
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   840
      TabIndex        =   39
      Top             =   6840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GST Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3480
      TabIndex        =   37
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select SalesMan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   360
      TabIndex        =   36
      Top             =   6000
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   360
      TabIndex        =   35
      Top             =   5640
      Width           =   1860
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name in Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   360
      TabIndex        =   34
      Top             =   5160
      Width           =   2100
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Party"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   30
      Top             =   4680
      Width           =   1275
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Vch Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   26
      Top             =   600
      Width           =   1725
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6240
      TabIndex        =   23
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3480
      TabIndex        =   22
      Top             =   3720
      Width           =   1320
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Sheet No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   21
      Top             =   3720
      Width           =   1755
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Tran. Sheet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   20
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   19
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select S/P Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select M.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   17
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Series"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   16
      Top             =   1080
      Width           =   1410
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public g_frt As Form5
Dim vcb As String
Dim vca As String
Dim tam As String
Dim dtn As String
Private Sub Command1_Click()
If Check2.Value = 1 Then
Dim p As Boolean
p = True
If Trim(Text1.Text) <> "" Then
Else
p = False
MsgBox "Please Select Excel File"
End If
If Trim(T1.Text) <> "" Then
Else
p = False
MsgBox "Please Select Vch Type"
End If
If Trim(V1.Text) <> "" Then
Else
p = False
MsgBox "Please Select Excel File"
End If
If Trim(V2.Text) <> "" Then
Else
p = False
MsgBox "Please Select M.C"
End If
'If Trim(V3.Text) <> "" Then
'Else
'p = False
'MsgBox "Please Select ST/PT Type"
'End If
If Trim(V4.Text) <> "" Then
Else
p = False
MsgBox "Please Select Format"
End If
If Trim(VL9.Text) <> "" Then
Else
p = False
MsgBox "Please Select Item Group"
End If
If Trim(Text2.Text) <> "" Then
Else
p = False
MsgBox "Please Select Excel Sheet No."
End If
If Trim(Text3.Text) <> "" Then
Else
p = False
MsgBox "Please Select Starting Row No."
End If
If Trim(Text3.Text) <> "" Then
Else
p = False
MsgBox "Please Select Ending Row No."
End If
If Option3.Value = False And Trim(Text5.Text) <> "" Then
Else
p = False
MsgBox "Please Select Party Name Coloum if Excel Name Spf. in Excel"
End If
'If Trim(VS1.Text) <> "" Then
'Else
'p = False
'MsgBox "Please Select SalesMan"
'End If
End If

If p = True Then

aph
ght
Chk_itemM
Chk_AccM
shortlist
If Option2.Value = True Then
chk_am
End If
If Text6.Text <> "" And Check3.Value = 1 Then
gst_L
gst_u

End If

If Check2.Value = 1 Then
vch_saves
End If

End If
If Text6.Text <> "" And Check3.Value = 1 Then
gst_L
gst_u

End If

End Sub
Public Sub gst_L()
Label15.Caption = "Account Master Importing"
Label15.Visible = True
Dim lt As ListItem
Dim jt As ColumnHeader
If Text6.Text <> "" Then
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim m As Integer
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.workbooks.Open (Text6.Text)
Dim sht As Integer
sht = CDbl(1)
    Set ExcelBook = ExcelObj.workbooks(1)
    Set ExcelSheet = ExcelBook.WorkSheets(sht)

    'Dim l As ListItem
    ListView4.ListItems.Clear
With ExcelSheet
Dim r As Integer
    Dim rs As Integer
    Dim SrNo As String
    Dim fc As Integer
    Dim sc As Integer
    fc = Text3.Text
    sc = Text4.Text
    Dim er As Integer
Dim k As String
For u = CDbl(fc) To CDbl(sc)
   'For u = 2 To 10000
   If Trim(.cells(u, 1)) <> "" Then
   
'Do Until .cells(i, 1) & "" = ""
'If .cells(u, 1) = "" Then
'MsgBox .cells(u, 1)
'Else
Set lt = ListView4.ListItems.Add(, , .cells(u, 14))

lt.SubItems(1) = .cells(u, 17)
lt.SubItems(2) = Left(.cells(u, 17), 2)
'End If
Else
Exit For
End If
'i = i + 1
'Loop
Next
End With

    ExcelObj.Activeworkbook.Close

    ExcelObj.quit

    Set ExcelSheet = Nothing

    Set ExcelBook = Nothing

    Set ExcelObj = Nothing


End If
Label15.Visible = False
End Sub
Public Sub gst_u()
Dim tk As String
Dim i As Integer
Dim o As Integer
    Dim InvalMsg As String
    Dim t2 As String
    Dim sname As String
    Dim Accdata As BusyDDC2175.udtAccMast
    Dim AccMast As Busy2175.CAccMast
Dim scode As Long

For i = 1 To ListView4.ListItems.Count

If g_MS.MasterName2CodeIfExist(Trim(Left(ListView4.ListItems(i).Text, 40)), 2) = 0 Then
Else

Label15.Caption = "GST No. Updation for Account Master : " & Trim(Left(ListView4.ListItems(i).Text, 40))
Label15.Visible = True
'g1_c = ListView1.ListItems(i).Text
tk = Trim(Left(ListView4.ListItems(i).Text, 40))
t2 = ListView4.ListItems(i).SubItems(2)
    scode = g_MS.TinDigit2StateCode(t2)
    sname = g_MS.MasterCode2NameIfExist(scode)

    
        Set AccMast = New Busy2175.CAccMast
    
    If AccMast.Load2(Trim(Left(ListView4.ListItems(i).Text, 40))) Then
        Accdata = AccMast.GetState

    With Accdata
  .udtAddressInfo.GSTNo = Trim(ListView4.ListItems(i).SubItems(1))
  .udtAddressInfo.CountryName = "India"
        .udtAddressInfo.StateName = sname
    
'                .OpBal = OpBal * (-1)         'Make OpBal negative to show debit status
    End With
    Set AccMast = New Busy2175.CAccMast
         AccMast.SetState Accdata

If AccMast.CanBeSaved(InvalMsg) Then
    If AccMast.Save(InvalMsg) Then
        'MsgBox "Bill finished successfuly."
    Else
        MsgBox InvalMsg
    End If
End If


'Form5.Show vbModal
End If
End If
Next
Label15.Visible = False
End Sub
Public Sub chk_am()
Label15.Caption = "Account Master Importing"
Label15.Visible = True
Dim tk As String
Dim i As Integer
Dim o As Integer
    Dim InvalMsg As String
    
    Dim Accdata As BusyDDC2175.udtAccMast
    Dim AccMast As Busy2175.CAccMast

For i = 1 To ListView1.ListItems.Count
If g_MS.MasterName2CodeIfExist(Trim(Left(ListView1.ListItems(i).SubItems(13), 40)), 2) = 0 Then
'g1_c = ListView1.ListItems(i).Text
tk = Left(ListView1.ListItems(i).SubItems(13), 40)
    
    

    
    g_OS.InitUDTAccMast Accdata
    With Accdata
        .Name = tk
        .PrintName = tk
    '    .TypeOfLedger = 0
    If T1.Text = "Sale" Then
        .ParentGrpName = "Sundry Debtors"
        Else
        .ParentGrpName = "Sundry Creditors"
        End If
        .OpBal = CDbl(0#)
        .udtAddressInfo.CountryName = "India"
        .udtAddressInfo.StateName = "Haryana"
        If Check1.Value = 1 Then
        .BillByBillBalancing = True
        End If
'                .OpBal = OpBal * (-1)         'Make OpBal negative to show debit status
    End With
    Set AccMast = New Busy2175.CAccMast
         AccMast.SetState Accdata

If AccMast.CanBeSaved(InvalMsg) Then
    If AccMast.Save(InvalMsg) Then
        'MsgBox "Bill finished successfuly."
    Else
        MsgBox InvalMsg
    End If
End If


'Form5.Show vbModal

End If
Next
Label15.Visible = False
End Sub

Public Sub shortlist()
ListView1.SortKey = 2
ListView1.SortOrder = lvwAscending
ListView1.Sorted = True
End Sub
Public Sub Chk_itemM()
Dim InvalMsg As String
Dim ItemData As BusyDDC2175.udtItemMast
Dim ItemMast As Busy2175.CItemMast
Dim ItemAlias As BusyDDC2175.udtGeneral
Dim col2 As Collection
Dim t As String
Dim MYS As String
Dim MYS1 As String
Dim check_Alias As Boolean
Dim i As Integer
Dim icode As Long
Dim ialias As String
For i = 1 To ListView1.ListItems.Count
icode = 0



MYS = g_CL.PadRight(ListView1.ListItems(i).SubItems(3), 40)
MYS = rchar(MYS)
icode = g_MS.MasterName2CodeIfExist(Trim(MYS), 6)

  'icode = g_MS.MasterName2Code(Trim(MYS), 6)
  If CDbl(icode) = 0 Then
  If Trim(MYS) = "8901326134382" Then
  MsgBox "FOUND"
  End If
Label15.Caption = "Item Importing : " & Trim(MYS)
Label15.Visible = True

g_OS.InitUDTItemMast ItemData
With ItemData
.Name = Trim(MYS)

.PrintName = Trim(MYS)
.ParentGrpName = VL9.Text
 .HSNCode = Trim(ListView1.ListItems(i).SubItems(4))
.MainUnitName = "Pcs."
.AltUnitName = "Pcs."
If V4.Text <> "Jockey" Then
.TaxCategory = "512"
Else
.TaxCategory = Trim(ListView1.ListItems(i).SubItems(12))
End If
'If CDbl(ListView1.ListItems(i).SubItems(8)) = 14 Or CDbl(ListView1.ListItems(i).SubItems(8)) = 28 Then
'.TaxCategory = "28%" 'RSTEMP.Fields(7).Value
'End If
'If CDbl(ListView1.ListItems(i).SubItems(8)) = 9 Or CDbl(ListView1.ListItems(i).SubItems(8)) = 18 Then
'.TaxCategory = "18%" 'RSTEMP.Fields(7).Value
'End If
'If CDbl(ListView1.ListItems(i).SubItems(8)) = 2.5 Or CDbl(ListView1.ListItems(i).SubItems(8)) = 5 Then
'.TaxCategory = "5%" 'RSTEMP.Fields(7).Value
'End If
'If CDbl(ListView1.ListItems(i).SubItems(8)) = 6 Or CDbl(ListView1.ListItems(i).SubItems(8)) = 12 Then
'.TaxCategory = "12%" 'RSTEMP.Fields(7).Value
'End If

End With

    Set ItemMast = New Busy2175.CItemMast
             ItemMast.SetState ItemData

        If ItemMast.CanBeSaved(InvalMsg) Then
            If ItemMast.Save(InvalMsg) Then
                'MsgBox "Bill finished successfuly."
            Else
                MsgBox InvalMsg
            End If
        End If
End If
Next
'MsgBox "Saved"
Label15.Caption = False
End Sub
Public Function rchar(schar As String) As String
Dim char As String
Dim x As Integer

's = "  " & vbTab & " ABC  " & vbTab & "a   " & vbTab

'// Now do it from the end
For x = Len(schar) To 1 Step -1
    char = Mid$(schar, x, 1)
    If char = vbTab Or char = " " Then
    
    Else
        rchar = Left$(schar, x)
        Exit For
    End If
Next
End Function
Public Sub Chk_AccM()
If Option1.Value = True Then
For i = 1 To ListView1.ListItems.Count
If g_OS.MasterMultipleAlias2CodeIfExist(Trim(ListView1.ListItems(i).Text), 2) = 0 Then
g1_c = ListView1.ListItems(i).Text
g1_n = ListView1.ListItems(i).SubItems(13)
g1_p = ListView1.ListItems(i).SubItems(14)

Form5.Show vbModal

End If
Next
MsgBox "All Accounts Mapped"
End If
End Sub
Public Sub vch_saves()

    
    Dim ItemData As BusyDDC2175.udtItemData
    Dim VchData As BusyDDC2175.udtVchData
    Dim BSData As BusyDDC2175.udtBSData
    Dim VchDataInv As Busy2175.CVchDataInv
    Dim ItemCol As New Collection
    Dim BSCol As New Collection
    Dim i As Integer
    Dim InvalMsg As String
    Dim udtRef As BusyDDC2175.udtReference
    Dim AORA As BusyDDC2175.udtAORA
    Dim RefCol As New Collection
        Dim PendingRefs As New Collection
    Dim Ref As udtReference

    Dim VN As Boolean
    Dim VM1 As String
    Dim z As Recordset
    Dim Accdata As BusyDDC2175.udtAccMast
    Dim AccMast As Busy2175.CAccMast
    Dim udtVchNum As BusyDDC2175.udtVchNumbering
        Dim tcstot  As Double
    Dim itmX As ListItem
    Screen.MousePointer = vbHourglass
    Dim oldnum As String
    oldnum = ""
    Dim PA As String
    PA = ""
    Dim n As Integer
    Dim jtot As Double
    For i = 1 To ListView1.ListItems.Count
    Label15.Caption = "Vch No. : " & ListView1.ListItems(i).SubItems(2) & " Row No. " & i + CDbl(Text3.Text) - 1 & "/" & CDbl(ListView1.ListItems.Count) + CDbl(Text3.Text) - 1 & " Importing"
    Label15.Visible = True
    If oldnum = Trim(ListView1.ListItems(i).SubItems(2)) And PA = Trim(ListView1.ListItems(i).Text) Then
    n = n + 1
    Else
    oldnum = Trim(ListView1.ListItems(i).SubItems(2))
    PA = Trim(ListView1.ListItems(i).Text)
    n = 1
    jtot = 0
    tcstot = 0
    VN = False
    Dim pcode As Long
If Option1.Value = True Then
pcode = g_OS.MasterMultipleAlias2CodeIfExist(Trim(ListView1.ListItems(i).Text), 2)
End If
If Option3.Value = True Then
pcode = g_MS.MasterName2CodeIfExist(S1.Text, 2)
End If
If Option2.Value = True Then
pcode = g_MS.MasterName2CodeIfExist(Mid(ListView1.ListItems(i).SubItems(13), 1, 40), 2)
End If
    g_OS.InitUdtVchData VchData
    
    Set VchDataInv = New Busy2175.CVchDataInv
If T1.Text = "Sale" Then
    VchDataInv.VchType = SALE
 ElseIf T1.Text = "Purchase" Then
    VchDataInv.VchType = PURCHASE
    ElseIf T1.Text = "Sale Return" Then
    VchDataInv.VchType = SALE_RETURN
    Else
    VchDataInv.VchType = PURCHASE_RETURN
    End If
    
    
    With VchData
    If T1.Text = "Sale" Then
    .VchType = SALE
 ElseIf T1.Text = "Purchase" Then
    .VchType = PURCHASE
    ElseIf T1.Text = "Sale Return" Then
    .VchType = SALE_RETURN
    Else
    .VchType = PURCHASE_RETURN
    End If

    
        .VchSeriesName = V1.Text
        
        .Date = g_CL.GetDateFromStr(ListView1.ListItems(i).SubItems(1))
        .tmpVchSeriesCode = g_MS.MasterName2CodeIfExist(g_OS.Series2MasterName(.VchType, .VchSeriesName), SERIES_MAST)
        .VchNo = ListView1.ListItems(i).SubItems(2)
        udtVchNum = g_OS.GetVchNumInfoData(.tmpVchSeriesCode)
       ' .AutoVchNo = g_TS.GenerateAutoVchNumber(udtVchNum, .VchNo, .Date)
       Dim tin1 As String
        Dim tincc As String
        tin1 = GSTNo2digt(pcode)
        tincc = g_MS.StateName2TinDigit("Haryana")
        
        
        'MsgBox tin1
        'MsgBox tincc
        If tin1 = tincc Then
        .STPTName = "L/GST-ItemWise"           'Sale Type
        Else
        .STPTName = "I/GST-ItemWise"           'Sale Type
        End If
        
        .MasterName1 = g_MS.MasterCode2Name(pcode)     'Party Name
        .MasterName2 = V2.Text  'Material Center name
        If VS1.Text <> "" Then
        .BrokerInvolved = True
        .BrokerName = VS1.Text
        End If
        
        If .VchType = SALE_RETURN Or .VchType = PURCHASE_RETURN Then
        If Trim(ListView1.ListItems(i).SubItems(16)) <> "" Then
                Dim fq As String
        Dim fqst As Recordset
        Dim vx As Long
        
        vx = getvchcode(Trim(ListView1.ListItems(i).SubItems(16)), 9)
        If vx <> 0 Then
fq = "SELECT Sum(VchGSTSumItemWise.TaxableAmt) AS SumOfTaxableAmt, VchGSTSumItemWise.VchCode, Sum(VchGSTSumItemWise.TaxAmt) AS SumOfTaxAmt, Sum(VchGSTSumItemWise.TaxAmt1) AS SumOfTaxAmt1 From VchGSTSumItemWise GROUP BY VchGSTSumItemWise.VchCode HAVING (((VchGSTSumItemWise.VchCode)=" + CStr(vx) + "))"

Set fqst = g_OS.GetRecordset(fq)
If fqst.RecordCount > 0 Then
fqst.MoveFirst
.OrgSalePurcDet.TaxableAmt = CDbl(fqst(0))
.OrgSalePurcDet.TaxAmt = CDbl(fqst(2))
.OrgSalePurcDet.TaxAmt1 = CDbl(fqst(3))
'.OrgSalePurcDet.tmpVchCode = vx
.OrgSalePurcDet.VchNo = Trim(ListView1.ListItems(i).SubItems(17))
        .OrgSalePurcDet.VchDate = Trim(ListView1.ListItems(i).SubItems(18))
End If
        End If
        End If
        End If
    End With
    End If
    
    g_OS.InitUdtItemData ItemData
    With ItemData
        .SrNo = n
        .ItemName = rchar(g_CL.PadRight(Trim(ListView1.ListItems(i).SubItems(3)), 40))
        .QtyMainUnit = CDbl(ListView1.ListItems(i).SubItems(6))
        
        .Qty = CDbl(ListView1.ListItems(i).SubItems(6))
        .UnitName = g_MS.GetItemMainUnitName(Trim(.ItemName), True)
        
        tcstot = tcstot + CDbl(ListView1.ListItems(i).SubItems(16))
        .Date = VchData.Date
        .VchSeriesCode = g_MS.MasterName2CodeIfExist(g_OS.Series2MasterName(VchData.VchType, VchData.VchSeriesName), SERIES_MAST)
        If G1.Text = "Local Rate" Then
        .STPercent = Round(CDbl(ListView1.ListItems(i).SubItems(8)), 2)
        .TaxBeforeSurcharge = Round(CDbl(ListView1.ListItems(i).SubItems(9)), 2)
        .TaxBeforeSurcharge1 = Round(CDbl(ListView1.ListItems(i).SubItems(9)), 2)
        .STPercent1 = Round(CDbl(ListView1.ListItems(i).SubItems(8)), 2)
        .STAmount = .TaxBeforeSurcharge + .TaxBeforeSurcharge1
         .Amt = Round(CDbl(ListView1.ListItems(i).SubItems(7)) + CDbl(.STAmount), 2)
         .Price = Round(g_CL.DivideNum(CDbl(ListView1.ListItems(i).SubItems(7)), CDbl(.Qty), 2), 2)

        .tmpTaxableAmt = Round(CDbl(ListView1.ListItems(i).SubItems(7)), 2)
        jtot = CDbl(jtot) + CDbl(.Amt)
        End If
If G1.Text = "Single Rate" And VchData.STPTName = "L/GST-ItemWise" Then          'Sale Type
        .STPercent = Round(CDbl(ListView1.ListItems(i).SubItems(8) / 2), 2)
        .TaxBeforeSurcharge = Round(CDbl(ListView1.ListItems(i).SubItems(9) / 2), 2)
        .TaxBeforeSurcharge1 = Round(CDbl(ListView1.ListItems(i).SubItems(9) / 2), 2)
        .STPercent1 = Round(CDbl(ListView1.ListItems(i).SubItems(8) / 2), 2)
        .STAmount = .TaxBeforeSurcharge + .TaxBeforeSurcharge1
         .Amt = Round(CDbl(ListView1.ListItems(i).SubItems(7)) + CDbl(.STAmount), 2)
         .Price = Round(g_CL.DivideNum(CDbl(ListView1.ListItems(i).SubItems(7)), CDbl(.Qty), 2), 2)

        .tmpTaxableAmt = Round(CDbl(ListView1.ListItems(i).SubItems(7)), 2)
        jtot = Round(CDbl(jtot) + CDbl(.Amt), 2)
        End If
If G1.Text = "Single Rate" And VchData.STPTName = "I/GST-ItemWise" Then          'Sale Type
        .STPercent = Round(CDbl(ListView1.ListItems(i).SubItems(8)), 2)
        .TaxBeforeSurcharge = Round(CDbl(ListView1.ListItems(i).SubItems(9)), 2)
        '.TaxBeforeSurcharge1 = CDbl(ListView1.ListItems(i).SubItems(9) / 2)
        '.STPercent1 = CDbl(ListView1.ListItems(i).SubItems(8) / 2)
        .STAmount = .TaxBeforeSurcharge
         .Amt = Round(CDbl(ListView1.ListItems(i).SubItems(7)) + CDbl(.STAmount), 2)
         .Price = Round(g_CL.DivideNum(CDbl(ListView1.ListItems(i).SubItems(7)), CDbl(.Qty), 2), 2)

        .tmpTaxableAmt = Round(CDbl(ListView1.ListItems(i).SubItems(7)), 2)
        jtot = Round(CDbl(jtot) + CDbl(.Amt), 2)
        End If
        
    End With
    

    ItemCol.Add ItemData
    Set VchData.ItemEntries = ItemCol
    Dim nb As Integer
    If ListView1.ListItems.Count > i Then
            If oldnum = Trim(ListView1.ListItems(i + 1).SubItems(2)) And PA = Trim(Trim(ListView1.ListItems(i + 1).Text)) Then
            Else
            VN = True
            End If
    Else
    VN = True
    End If
    If VN = True Then
    nb = 0
    Set BSCol = New Collection
        
        
    If Trim(ListView1.ListItems(i).SubItems(16)) = "" Or CDbl(ListView1.ListItems(i).SubItems(16)) = 0 Then
    Else
    g_OS.InitUdtBSData BSData
    With BSData
        .SrNo = nb + 1
       
        .BSName = "TCS"
    
       .PercentVal = "0.10"
        .Amt = CDbl(tcstot)
            jtot = Round(CDbl(jtot) + CDbl(.Amt), 2)
        .tmpBSMast = g_MS.GetBSMastData(.BSName, False, False)
        '.tmpBSMast = g_MS.GetBSMastData(Label1.Caption, False, False)
        'nb = nb + 1
    End With

    BSCol.Add BSData
nb = nb + 1
End If
    Dim rval As Double
    Dim nnv As Integer
    rval = 0
    If Round(jtot, 0) = CDbl(jtot) Then
    Else
    rval = Round(jtot, 0) - CDbl(jtot)
    End If
    
    
        If CDbl(rval) <> 0 Then
    With BSData
        .SrNo = nb + 1
        If CDbl(rval) < 0 Then
        .BSName = "Rounded Off (-)"
    nnv = 0
        Else
        .BSName = "Rounded Off (+)"
        nnv = 1
        End If
        '.PercentVal = CDbl(Text1.Text)

'        .PercentOperatedOn = .PercentVal
        .Amt = Abs(CDbl(rval))
        If T1.Text = "Sale" Or T1.Text = "Purchase Return" Then
        If nnv = 1 Then
            jtot = Round(CDbl(jtot) + CDbl(.Amt), 2)
            End If
If nnv = 0 Then
            jtot = Round(CDbl(jtot) - CDbl(.Amt), 2)
            End If
            
            End If
            If T1.Text = "Purchase" Or T1.Text = "Sale Return" Then
            jtot = Round(CDbl(jtot) - CDbl(.Amt), 2)
            End If
        .tmpBSMast = g_MS.GetBSMastData(.BSName, False, False)
        '.tmpBSMast = g_MS.GetBSMastData(Label1.Caption, False, False)
    End With

    BSCol.Add BSData
    Set VchData.BSEntries = BSCol

End If
    If Check1.Value = 1 And (T1.Text = "Sale" Or T1.Text = "Purchase") And CDbl(jtot) <> 0 Then
        With udtRef
          .SrNo = 1
        .Method = METHOD_NEWREF
        .Date = VchData.Date
        .DueDate = .Date
        .no = VchData.VchNo
        .tmpRecType = ACC_REF
        If VchData.VchType = SALE Then
        .Value1 = g_CL.Negative(CDbl(jtot))
        End If
        If VchData.VchType = PURCHASE Then
        .Value1 = CDbl(jtot)
        End If
        .VchType = VchData.VchType
        .tmpMasterCode1 = g_MS.MasterName2Code(VchData.MasterName1, 2)
        .tmpMasterCode2 = 0
    End With

    Set AORA.RefArr = New Collection
    AORA.RefArr.Add udtRef
    AORA.MasterCode1 = udtRef.tmpMasterCode1
    AORA.MasterCode2 = udtRef.tmpMasterCode2
    AORA.Value1 = udtRef.Value1

    RefCol.Add AORA

    Set VchData.PendingBills = RefCol
    
    End If
    If T1.Text = "Sale Return" Or T1.Text = "Purchase Return" Then
    If Trim(ListView1.ListItems(i).SubItems(16)) <> "" Then
    If CDbl(getvc(Trim(ListView1.ListItems(i).SubItems(16)), 9)) <> 0 Then
          If Check1.Value = 1 And (T1.Text = "Sale Return" Or T1.Text = "Purchase Return") Then
        With udtRef
          .SrNo = 1
        .Method = METHOD_ADJUSTMENT
        .Date = VchData.Date
        .DueDate = getdt(Trim(ListView1.ListItems(i).SubItems(16)), 9)
        .no = Trim(ListView1.ListItems(i).SubItems(16))
        .tmpRecType = ACC_REF
        If T1.Text = "Sale Return" Then
        .Value1 = CDbl(jtot)
        End If
        If T1.Text = "Purchase Return" Then
        .Value1 = g_CL.Negative(CDbl(jtot))
        End If
        .VchType = VchData.VchType
        .tmpMasterCode1 = g_MS.MasterName2Code(VchData.MasterName1, 2)
        .tmpMasterCode2 = 0
        .tmpRefCode = getvc(Trim(ListView1.ListItems(i).SubItems(16)), 9)
    End With

    Set AORA.RefArr = New Collection
    AORA.RefArr.Add udtRef
    AORA.MasterCode1 = udtRef.tmpMasterCode1
    AORA.MasterCode2 = udtRef.tmpMasterCode2
    AORA.Value1 = udtRef.Value1

    RefCol.Add AORA

    Set VchData.PendingBills = RefCol
    End If
    End If
    End If
    End If
                    VchDataInv.SetState VchData
            
                    If VchDataInv.CanBeSaved(InvalMsg) Then
                        If Not VchDataInv.Save(InvalMsg) Then
                            MsgBox InvalMsg
                        Else
                           ' MsgBox "KOT Created Successfully."
                        End If
                    Else
                        MsgBox InvalMsg
                    End If
    
    Set VchDataInv = Nothing
    'Set VchData = Nothing
    Set RefCol = Nothing
    Set ItemCol = Nothing
    Set PendingBills = Nothing
 Set BSCol = Nothing
    End If
    
   
   
   Next
    MsgBox "Data Imported Successfully."
    Screen.MousePointer = vbDefault
    Label15.Visible = False
End Sub
Public Function getvc(vchn As String, vcht As Integer) As Long
Dim q As String
Dim rst As Recordset
Dim vct As Long
q = "Select Vchcode from tran1 where vchno ='" + g_OS.Actual2Index(vchn, 25) + "' and vchtype =" + CStr(vcht) + ""
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
vct = rst(0)
Else
vct = 0
End If
Set rst = Nothing
If vct <> 0 Then
q = "Select refcode from tran3 where vchcode = " + CStr(vct) + ""
'MsgBox q
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
getvc = rst(0)
Else
getvc = 0
End If
End If
End Function
Public Function getdt(vchn As String, vcht As Integer) As Date
Dim q As String
Dim rst As Recordset
Dim vct As Long
Dim dt1 As Date
q = "Select Vchcode,date from tran1 where vchno ='" + g_OS.Actual2Index(vchn, 25) + "' and vchtype =" + CStr(vcht) + ""
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
vct = rst(0)
dt1 = rst(1)
Else
vct = 0
End If
Set rst = Nothing
If vct <> 0 Then
q = "Select duedate from tran3 where vchcode = " + CStr(vct) + ""
'MsgBox q
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
getdt = rst(0)
Else
getdt = dt1
End If
End If
End Function
Public Function getvchcode(vchn As String, vcht As Integer) As Long
Dim q As String
Dim rst As Recordset
Dim vct As Long
Dim dt1 As Date
q = "Select Vchcode from tran1 where vchno ='" + g_OS.Actual2Index(vchn, 25) + "' and vchtype =" + CStr(vcht) + ""
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
getvchcode = rst(0)

Else
getvchcode = 0
End If
Set rst = Nothing
End Function

Public Sub aph()
Dim q As String
Dim rst As Recordset
Dim lt As ListItem
If T1.Text = "Sale" Then
q = "Select c21,c2, i3 from config where rectype = 41 and i1 = 240 and i2 <>0 and c1 = '" + Trim(V4.Text) + "'"
End If
If T1.Text = "Purchase" Then
q = "Select c21,c2, i3 from config where rectype = 41 and i1 = 241 and i2 <>0 and c1 = '" + Trim(V4.Text) + "'"
End If
If T1.Text = "Sale Return" Then
q = "Select c21,c2, i3 from config where rectype = 41 and i1 = 242 and i2 <>0 and c1 = '" + Trim(V4.Text) + "'"
End If
If T1.Text = "Purchase Return" Then
q = "Select c21,c2, i3 from config where rectype = 41 and i1 = 243 and i2 <>0 and c1 = '" + Trim(V4.Text) + "'"
End If

Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
Set lt = ListView2.ListItems.Add(, , rst(0))
lt.SubItems(1) = rst(1)
lt.SubItems(2) = rst(2)
rst.MoveNext
Loop
End If
Dim m As Integer
Dim j As Integer

Dim yt As ListItem
    For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "PARTY_ALIAS" Then    '0
    Set yt = ListView3.ListItems.Add(, , "PARTY_ALIAS")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
    For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "VCH/BILL_DATE" Then    '1
    Set yt = ListView3.ListItems.Add(, , "VCH/BILL_DATE")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    dtn = ListView2.ListItems(j).SubItems(1)
    Else
    End If
    Next
    
    For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "VCH/BILL_NO" Then     '2
    Set yt = ListView3.ListItems.Add(, , "VCH/BILL_NO")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
'    vcb = ListView2.ListItems(j).SubItems(1)
    Else
    End If
    Next
    
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "ITEM_NAME" Then      '3
    Set yt = ListView3.ListItems.Add(, , "ITEM_NAME")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
    
    For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "ITEM_HSN_CODE" Then    '4
    Set yt = ListView3.ListItems.Add(, , "ITEM_HSN_CODE")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
    For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "ITEM_MRP" Then       '5
    Set yt = ListView3.ListItems.Add(, , "ITEM_MRP")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
    For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "QUANTITY" Then       '6
    Set yt = ListView3.ListItems.Add(, , "QUANTITY")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
    
    For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "AMOUNT" Then          '7
    Set yt = ListView3.ListItems.Add(, , "AMOUNT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    vca = ListView2.ListItems(j).SubItems(1)
    Else
    End If
    Next
If G1.Text = "Local Rate" Then
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "CGST_PERCENT" Then  '8
    Set yt = ListView3.ListItems.Add(, , "CGST_PERCENT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "CGST_AMOUNT" Then    '9
    Set yt = ListView3.ListItems.Add(, , "CGST_AMOUNT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "SGST_PERCENT" Then     '10
    Set yt = ListView3.ListItems.Add(, , "SGST_PERCENT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "SGST_AMOUNT" Then     '11
    Set yt = ListView3.ListItems.Add(, , "SGST_AMOUNT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
End If
If G1.Text = "Single Rate" Then
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "IGST_PERCENT" Then    '8
    Set yt = ListView3.ListItems.Add(, , "IGST_PERCENT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    vcb = ListView2.ListItems(j).SubItems(1)
    Else
    End If
    Next
    
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "IGST_AMOUNT" Then      '9
    Set yt = ListView3.ListItems.Add(, , "IGST_AMOUNT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    tam = ListView2.ListItems(j).SubItems(1)
    Else
    End If
    Next
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "IGST_PERCENT" Then    '10
    Set yt = ListView3.ListItems.Add(, , "IGST_PERCENT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
     For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "IGST_AMOUNT" Then      '11
    Set yt = ListView3.ListItems.Add(, , "IGST_AMOUNT")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
End If

    For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "ITEM_TAX_CATEGORY" Then     '12
    Set yt = ListView3.ListItems.Add(, , "ITEM_TAX_CATEGORY")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
'If Option1.Value = True Then
Set yt = ListView3.ListItems.Add(, , "Party Name")         '13
yt.SubItems(1) = Trim(Text5.Text)
yt.SubItems(2) = "0"
For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "BILLED_PARTY_GST_NO" Then     '14
    Set yt = ListView3.ListItems.Add(, , "BILLED_PARTY_GST_NO")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
Set yt = ListView3.ListItems.Add(, , "Round off")         '15
yt.SubItems(1) = Trim(Text7.Text)
yt.SubItems(2) = "0"
Set yt = ListView3.ListItems.Add(, , "TCS")         '16
yt.SubItems(1) = Trim(Text8.Text)
yt.SubItems(2) = "0"

If T1.Text = "Sale Return" Then
For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "ORIGINAL_SALE_PURC_VCH_NO" Then     '17
    Set yt = ListView3.ListItems.Add(, , "ORIGINAL_SALE_PURC_VCH_NO")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
    
For j = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(j).Text = "ORIGINAL_SALE_PURC_DATE" Then     '18
    Set yt = ListView3.ListItems.Add(, , "ORIGINAL_SALE_PURC_DATE")
    yt.SubItems(1) = ListView2.ListItems(j).SubItems(1)
    yt.SubItems(2) = ListView2.ListItems(j).SubItems(2)
    Else
    End If
    Next
End If

'End If
End Sub
Public Sub ght()
Dim pf As Collection
Dim lsv As ColumnHeader
g_CL.FlushCol pf
Dim u As Integer
For u = 1 To ListView3.ListItems.Count
pf.Add ListView3.ListItems(u).SubItems(1)
Next
ListView1.Refresh
For u = 1 To ListView3.ListItems.Count

Set lsv = ListView1.ColumnHeaders.Add()
lsv.Text = Trim(ListView3.ListItems(u).Text)
lsv.Width = 1500
Next

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
sht = CDbl(Text2.Text)
    Set ExcelBook = ExcelObj.workbooks(1)
    Set ExcelSheet = ExcelBook.WorkSheets(sht)

    'Dim l As ListItem
    ListView1.ListItems.Clear
With ExcelSheet
Dim r As Integer
    Dim rs As Integer
    Dim SrNo As String
    Dim fc As Integer
    Dim sc As Integer
    fc = Text3.Text
    sc = Text4.Text
    Dim er As Integer
Dim k As String
   'For u = 1 To pf.Count
'Do Until .cells(i, 1) & "" = ""
For i = CDbl(fc) To CDbl(sc)
If Mid(.cells(i, 1), 1, 11) = "*Sub Total*" Then
'MsgBox "1"
Else
For u = 1 To pf.Count

'MsgBox pf(u)
er = gnum(pf(u))
'MsgBox pf(u)
If pf(u) = Trim(vcb) Then
If .cells(i, er) = "" Or .cells(i, er) = "0" Then

k = CDbl(.cells(i, er - 2)) * 2
Else
k = .cells(i, er)
End If
ElseIf pf(u) = Trim(tam) Then
If .cells(i, er) = "" Or .cells(i, er) = "0" Then

k = CDbl(.cells(i, er - 2)) * 2
Else
k = .cells(i, er)
End If
'ElseIf pf(u) = Trim(dtn) Then
'MsgBox .cells(i, er)
'k = Mid(.cells(i, er), 4, 2) & "-" & Mid(.cells(i, er), 1, 2) & "-" & Mid(.cells(i, er), 7, 4)
'Dim er1 As Integer
'er1 = gnum(tam)
'k = CDbl(.cells(i, er)) - CDbl(.cells(i, er1))
Else
k = .cells(i, er)
End If
'MsgBox k
If u = 1 Then
r = 0

Set lt = ListView1.ListItems.Add(, , .cells(i, er))
Else
lt.SubItems(r) = k
End If
r = r + 1

Next
End If
'i = i + 1
'Loop
Next
End With

    ExcelObj.Activeworkbook.Close

    ExcelObj.quit

    Set ExcelSheet = Nothing

    Set ExcelBook = Nothing

    Set ExcelObj = Nothing


End If

End Sub
Private Sub Command2_Click()
CommonDialog1.Filter = "All files (*.*)|*.*"
CommonDialog1.DefaultExt = "xls"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen


Text1.Text = CommonDialog1.FileName

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

 Public Function gnum(Column As String) As Long

    Dim A As Integer, currentletter As String

    gnum = 0
    For A = 1 To Len(Column)
        currentletter = Mid(Column, Len(Column) - A + 1, 1)
        gnum = gnum + (Asc(currentletter) - 64) * 26 ^ (A - 1)
    Next

End Function


Private Sub Command4_Click()
CommonDialog1.Filter = "All files (*.*)|*.*"
CommonDialog1.DefaultExt = "xls"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen


Text6.Text = CommonDialog1.FileName

End Sub

Private Sub Command5_Click()
Dim c As String
c = "55502-PURPLE MACRON-S-PACK OF 1 "
MsgBox rchar(c)
End Sub

Private Sub Form_Load()
Check2.Value = 1
Text7.Text = "BS"
Text7.Enabled = False
Check1.Value = 1
End Sub

Private Sub G1_GotFocus()
G1.List.AddItem "Single Rate"
G1.List.AddItem "Local Rate"
'G1.List.AddItem "Both"
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Label10.Visible = False
S1.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Label10.Visible = False
S1.Visible = False
End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Label10.Visible = True
S1.Visible = True
End If

End Sub

Private Sub S1_GotFocus()
S1.List.Clear
Dim q As String
Dim rst As Recordset
q = "SELECT NAMEALIAS FROM HELP1 WHERE RECTYPE = " & CStr(H1_PARTY) & " AND NAMEORALIAS=" & CStr(NA_NAME)
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
S1.List.AddItem rst!NameAlias.Value
rst.MoveNext
Loop
End If

End Sub

Private Sub T1_GotFocus()
T1.List.Clear
T1.List.AddItem "Sale"
T1.List.AddItem "Purchase"
T1.List.AddItem "Sale Return"
T1.List.AddItem "Purchase Return"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Text2.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text2.Text = ""
    End If
    End If



  Select Case KeyAscii

  Case vbKey0 To vbKey9
  Case vbKeyBack
  Case vbKeyClear
  'Case vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
   KeyAscii = 0
  Beep
End Select



End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If Text3.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text3.Text = ""
    End If
    End If



  Select Case KeyAscii

  Case vbKey0 To vbKey9
  Case vbKeyBack
  Case vbKeyClear
  'Case vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
   KeyAscii = 0
  Beep
End Select



End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If Text4.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text4.Text = ""
    End If
    End If



  Select Case KeyAscii

  Case vbKey0 To vbKey9
  Case vbKeyBack
  Case vbKeyClear
  'Case vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
   KeyAscii = 0
  Beep
End Select



End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text1.Text = ""
    End If
    End If

End Sub

Private Sub Text5_LostFocus()
If Text5.Text <> "" Then
Text5.Text = UCase(Text5.Text)
End If
End Sub
Private Sub Text7_LostFocus()
If Text7.Text <> "" Then
Text7.Text = UCase(Text7.Text)
End If
End Sub
Private Sub Text8_LostFocus()
If Text8.Text <> "" Then
Text8.Text = UCase(Text8.Text)
End If
End Sub

Private Sub VijayList1_Change()

End Sub

Private Sub V1_GotFocus()
V1.List.Clear
If T1.Text = "Sale" Then
Dim q As String
Dim rst As Recordset
q = "Select name from master1 where mastertype =21 and i1 =9"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V1.List.AddItem Mid(rst(0), 3, 20)
rst.MoveNext
Loop
End If
End If
If T1.Text = "Purchase" Then
'Dim q As String
'Dim rst As Recordset
q = "Select name from master1 where mastertype =21 and i1 =2"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V1.List.AddItem Mid(rst(0), 3, 20)
rst.MoveNext
Loop
End If
End If
If T1.Text = "Sale Return" Then
'Dim q As String
'Dim rst As Recordset
q = "Select name from master1 where mastertype =21 and i1 =3"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V1.List.AddItem Mid(rst(0), 3, 20)
rst.MoveNext
Loop
End If
End If

End Sub

Private Sub V3_GotFocus()
V3.List.Clear
If T1.Text = "Sale" Then
Dim q As String
Dim rst As Recordset
q = "Select name from master1 where mastertype =13"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V3.List.AddItem rst(0)
rst.MoveNext
Loop
End If
End If
If T1.Text = "Purchase" Then
'Dim q As String
'Dim rst As Recordset
q = "Select name from master1 where mastertype =14"

Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V3.List.AddItem rst(0)
rst.MoveNext
Loop
End If
End If

End Sub
Private Sub V2_GotFocus()
V2.List.Clear
Dim q As String
Dim rst As Recordset
q = "Select name from master1 where mastertype =11"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V2.List.AddItem rst(0)
rst.MoveNext
Loop
End If

End Sub

Private Sub V4_GotFocus()
V4.List.Clear
If T1.Text = "Sale" Then
Dim q As String
Dim rst As Recordset
q = "Select c1 from config where rectype = 41 and i1 = 240 and i2 =0"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V4.List.AddItem rst(0)
rst.MoveNext
Loop
End If
End If
If T1.Text = "Purchase" Then
'Dim q As String
'Dim rst As Recordset
q = "Select c1 from config where rectype = 41 and i1 = 241 and i2 =0"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V4.List.AddItem rst(0)
rst.MoveNext
Loop
End If
End If
If T1.Text = "Sale Return" Then
'Dim q As String
'Dim rst As Recordset
q = "Select c1 from config where rectype = 41 and i1 = 242 and i2 =0"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
V4.List.AddItem rst(0)
rst.MoveNext
Loop
End If
End If

End Sub

Private Sub VL9_GotFocus()
VL9.List.Clear
Dim q As String
Dim rst As Recordset
q = "SELECT NAME FROM master1 WHERE masterTYPE = 5"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
VL9.List.AddItem rst(0)
rst.MoveNext
Loop
End If

End Sub

Private Sub VS1_GotFocus()
VS1.List.Clear
Dim q As String
Dim rst As Recordset
q = "SELECT NAME FROM master1 WHERE masterTYPE = 19"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
Do While Not rst.EOF
VS1.List.AddItem rst(0)
rst.MoveNext
Loop
End If

End Sub
