VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formula Master"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6600
      TabIndex        =   26
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Frame Frame6 
      Caption         =   "Replace Foumula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   17
      Top             =   6960
      Width           =   8655
      Begin Taxcat.VijayList Combo7 
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
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
      Begin Taxcat.VijayList Combo6 
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   6720
         TabIndex        =   36
         Top             =   840
         Width           =   1695
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
         Left            =   6120
         TabIndex        =   35
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   3960
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Enter Replace Value"
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
         Left            =   4440
         TabIndex        =   33
         Top             =   840
         Width           =   2160
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Condition"
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
         Left            =   360
         TabIndex        =   32
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Enter Value"
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
         Left            =   4440
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Select Coloum"
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
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Row Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   9360
      TabIndex        =   15
      Top             =   240
      Width           =   5535
      Begin MSComctlLib.ListView ListView1 
         Height          =   7695
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   13573
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
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
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Coloum ID"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Condition Foumula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   480
      TabIndex        =   14
      Top             =   4560
      Width           =   8655
      Begin Taxcat.VijayList Combo5 
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   1080
         Width           =   615
         _ExtentX        =   1085
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
      Begin Taxcat.VijayList Combo4 
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
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
      Begin Taxcat.VijayList Combo3 
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
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
         Left            =   2520
         TabIndex        =   29
         Top             =   1440
         Width           =   5895
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6000
         TabIndex        =   27
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   255
         Left            =   3960
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   255
         Left            =   3960
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Enter Repalce Value"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   2160
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Enter Value"
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
         Left            =   4560
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Condition"
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
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Select Conditional Coloum"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   2745
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enter Value"
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
         Left            =   4560
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Source Coloum"
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2310
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Joint Formula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   11
      Top             =   3480
      Width           =   8535
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
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "(Enter Coloum With + Symbol)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   360
         TabIndex        =   38
         Top             =   540
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Enter Coloum(s) Name"
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
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   2340
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Formula Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   8535
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
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
         Left            =   6840
         TabIndex        =   42
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Repalce"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Condition"
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
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Joint"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
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
      Height          =   495
      Left            =   13320
      TabIndex        =   3
      Top             =   8760
      Width           =   1335
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
      Left            =   11880
      TabIndex        =   2
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formula Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      Begin VB.CommandButton Command8 
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
         Left            =   6240
         TabIndex        =   44
         Top             =   720
         Width           =   2175
      End
      Begin Taxcat.VijayList combo1 
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
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
         Left            =   5280
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
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
      Begin VB.CommandButton Command7 
         Caption         =   "Disable"
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
         Left            =   6240
         TabIndex        =   43
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exiting Formula"
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
         Left            =   2400
         TabIndex        =   41
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "New Formula"
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
         Left            =   360
         TabIndex        =   40
         Top             =   240
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
         Left            =   2040
         TabIndex        =   6
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Fourmula Name"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   3735
      End
   End
   Begin VB.Image Image1 
      Height          =   13500
      Left            =   0
      Picture         =   "Form2.frx":0000
      Top             =   0
      Width           =   21000
   End
End
Attribute VB_Name = "Form2"
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
ListView1.ListItems.Clear
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
Dim q As String
Dim rst As Recordset
Dim fc As Long
Combo2.List.Clear
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
Combo2.List.AddItem rst(0)
rst.MoveNext
Loop
End If

End Sub

Private Sub Combo2_LostFocus()
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False

Dim q As String
Dim rst As Recordset
Dim fc As Long
Dim fcc As Long
q = "Select d1 from externaldata where l1=4 and c1='" + Trim(Combo1.Text) + "'"
Set rst = g_OS.GetRecordset(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    fc = rst(0)
    End If
q = "Select d1 from externaldata where i1=6 and l1=" + CStr(fc) + " and c1='" + Trim(Combo2.Text) + "'"
Set rst = g_OS.GetRecordset(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    
    fcc = rst(0)
    End If
q = "Select d1,b1 from externaldata where i1=7 and l1=" + CStr(fcc) + " "
Set rst = g_OS.GetRecordset(q)
       
If rst.RecordCount > 0 Then
rst.MoveFirst

    If rst(1).Value = False Then
    Command7.Caption = "Disable Format"
    Else
    Command7.Caption = "Enable Format"
    End If
    If rst(0) = "1" Then
    Frame3.Enabled = True
    Option1.Value = True
    Combo3.Enabled = False
    Combo4.Enabled = False
    Combo5.Enabled = False
    Combo6.Enabled = False
    Combo7.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text2.Enabled = True
    End If
    If rst(0) = "2" Then
    Frame4.Enabled = True
    Option2.Value = True
    Text2.Enabled = False
    Combo3.Enabled = True
    Combo4.Enabled = True
    Combo5.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = False
    Text7.Enabled = False
    Combo6.Enabled = False
    Combo7.Enabled = False
    
    End If

    If rst(0) = "3" Then
    Frame6.Enabled = True
    Option3.Value = True
    Text2.Enabled = False
    Combo3.Enabled = False
    Combo4.Enabled = False
    Combo5.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = True
    Text7.Enabled = True
    Combo6.Enabled = True
    Combo7.Enabled = True
    End If

End If
Combo1.Enabled = False
Combo2.Enabled = False
Command6.Enabled = False




End Sub

Private Sub Combo3_GotFocus()
Combo3.List.Clear
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
Combo3.List.AddItem Trim(ListView1.ListItems(i).SubItems(1))
Next


End Sub


Private Sub Combo3_LostFocus()
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
If Trim(ListView1.ListItems(i).SubItems(1)) = Trim(Combo3.Text) Then
Label16.Caption = i
Exit For
Else
End If
Next

End Sub

Private Sub Combo4_GotFocus()
Combo4.List.Clear
Combo4.List.AddItem "="
Combo4.List.AddItem ">"
Combo4.List.AddItem ">="
Combo4.List.AddItem "<"
Combo4.List.AddItem "<="
Combo4.List.AddItem "<>"



End Sub

Private Sub Combo5_LostFocus()
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
If Trim(ListView1.ListItems(i).SubItems(1)) = Trim(Combo5.Text) Then
Label17.Caption = i
Exit For
Else
End If
Next

End Sub

Private Sub Combo6_LostFocus()
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
If Trim(ListView1.ListItems(i).SubItems(1)) = Trim(Combo6.Text) Then
Label18.Caption = i
Exit For
Else
End If
Next

End Sub

Private Sub Combo7_GotFocus()
Combo7.List.Clear
Combo7.List.AddItem "="
Combo7.List.AddItem ">"
Combo7.List.AddItem ">="
Combo7.List.AddItem "<"
Combo7.List.AddItem "<="
Combo7.List.AddItem "<>"



End Sub

Private Sub Combo5_GotFocus()
Combo5.List.Clear
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
Combo5.List.AddItem Trim(ListView1.ListItems(i).SubItems(1))
Next


End Sub

Private Sub Combo6_GotFocus()
Combo6.List.Clear
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
Combo6.List.AddItem Trim(ListView1.ListItems(i).SubItems(1))
Next


End Sub

Private Sub Command1_Click()
Dim q As String
Dim q1 As String
Dim rst As Recordset
Dim fc As Long
Dim fcc As Long
q = "Select d1 from externaldata where l1=4 and c1='" + Trim(Combo1.Text) + "'"
Set rst = g_OS.GetRecordset(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    fc = rst(0)
    End If
q = "Select d1 from externaldata where i1=6 and l1=" + CStr(fc) + " and c1='" + Trim(Combo2.Text) + "'"
Set rst = g_OS.GetRecordset(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    
    fcc = rst(0)
    End If

If Option1.Value = True Then
If Text2.Text <> "" Then
q = "select fcode from msconfig where fcode =" + CStr(fcc) + " and ftype =1 "
Set rst = g_OS.GetRecordsetFromDB(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    q1 = "delete *from msconfig where fcode =" + CStr(fcc) + " and ftype =1 "
    g_OS.GetRecordsetFromDB (q1)
    
    Else
    'Insert Query
    End If
q1 = "insert into msconfig (rectype,fcode,ftype,t1,t2)values(8," + CStr(fcc) + ",1,'" + Trim(Text2.Text) + "','" + Trim(Label15.Caption) + "')"
g_OS.GetRecordsetFromDB (q1)
MsgBox "Record Update"
Else
MsgBox "Field Can't Blank"
End If
End If
If Option2.Value = True Then
If Combo4.Text <> "" Then
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
End If

q = "select fcode from msconfig where fcode =" + CStr(fcc) + " and ftype =2 "
Set rst = g_OS.GetRecordsetFromDB(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    q1 = "delete *from msconfig where fcode =" + CStr(fcc) + " and ftype =2 "
    g_OS.GetRecordsetFromDB (q1)
    
    Else
    'Insert Query
    End If
q1 = "insert into msconfig (rectype,fcode,ftype,SSC,SSV,Con1,SCC,SCV,T1,SSC1,SCC1)values(8," + CStr(fcc) + ",2,'" + Trim(Combo3.Text) + "','" + Trim(Text3.Text) + "'," + CStr(mj) + ",'" + Trim(Combo5.Text) + "','" + Trim(Text4.Text) + "','" + Trim(Text5.Text) + "','" + Trim(Label16.Caption) + "','" + Trim(Label17.Caption) + "')"
g_OS.GetRecordsetFromDB (q1)
MsgBox "Record Update"
End If
If Option3.Value = True Then
'Dim mj As Integer
If Combo7.Text <> "" Then
If Combo7.Text = "=" Then
mj = 1
End If
If Combo7.Text = ">" Then
mj = 2
End If
If Combo7.Text = ">=" Then
mj = 3
End If
If Combo7.Text = "<" Then
mj = 4
End If
If Combo7.Text = "<=" Then
mj = 5
End If
If Combo7.Text = "<>" Then
mj = 6
End If
End If

q = "select fcode from msconfig where fcode =" + CStr(fcc) + " and ftype =3 "
Set rst = g_OS.GetRecordsetFromDB(q)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    q1 = "delete *from msconfig where fcode =" + CStr(fcc) + " and ftype =3"
    g_OS.GetRecordsetFromDB (q1)
    
    Else
    'Insert Query
    End If
q1 = "insert into msconfig (rectype,fcode,ftype,SSC,SSV,Con1,T1,SSC1)values(8," + CStr(fcc) + ",3,'" + Trim(Combo6.Text) + "','" + Trim(Text6.Text) + "'," + CStr(mj) + ",'" + Trim(Text7.Text) + "','" + Trim(Label18.Caption) + "')"
g_OS.GetRecordsetFromDB (q1)
MsgBox "Record Update"
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Combo2.Visible = False
Label1.Visible = True
Text1.Visible = True
Text1.Left = 2040
Text1.Top = 1080
'Frame1.Visible = True
'Label4.Caption = "1"
Frame4.Enabled = False
Frame6.Enabled = False
Frame3.Enabled = False

End Sub

Private Sub Command4_Click()
Combo2.Visible = True
'Label1.Visible = True
Text1.Visible = False
Combo2.Left = 2040

Combo2.Top = 1080
'Frame1.Visible = True
'Label4.Caption = "2"
Frame4.Enabled = False
Frame6.Enabled = False
Frame3.Enabled = False

End Sub




Private Sub Command6_Click()
If Combo1.Text <> "" And Text1.Text <> "" Then
If Option1.Value = True Or Option2.Value = True Or Option3.Value = True Then
Dim q As String
Dim rst As Recordset
Dim fc As Long
Dim jb As Boolean
Dim fcc As Long
Dim optv As Integer
jb = False
q = "Select d1 from externaldata where l1=4 and c1='" + Trim(Combo1.Text) + "'"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
fc = rst(0)
End If
q = "Select d1 from externaldata where i1=6 and l1=" + CStr(fc) + " and c1='" + Trim(Text1.Text) + "'"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
jb = True

End If
q = "Select max(d1) from externaldata where i1=6"
Set rst = g_OS.GetRecordset(q)
If rst.RecordCount > 0 Then
rst.MoveFirst
If IsNull(rst(0)) Then
fcc = 5000
Else

fcc = rst(0)
End If
Else
fcc = 5000
End If
fcc = fcc + 1
If jb = False Then
q = "Insert into externaldata (i1,l1,d1,c1) values (6," + CStr(fc) + "," + CStr(fcc) + ",'" + Trim(Text1.Text) + "')"
g_OS.ExecuteQuerytmp (q)
If Option1.Value = True Then
optv = 1
End If
If Option2.Value = True Then
optv = 2
End If
If Option3.Value = True Then
optv = 3
End If

q = "Insert into externaldata (i1,l1,d1) values (7," + CStr(fcc) + "," + CStr(optv) + ")"
g_OS.ExecuteQuerytmp (q)


End If

End If
End If
End Sub

Private Sub Command8_Click()
Frame3.Enabled = False
Frame4.Enabled = False
'Frame5.Enabled = False
Frame6.Enabled = False

Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False
Combo6.Enabled = False
Combo7.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text2.Enabled = False

Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
Combo1.Enabled = True
Combo2.Enabled = True
End Sub

Private Sub Form_Load()
Dim p As String
Dim rs As Recordset
Dim sSQLDDL As String
Dim dbp As String
Dim fo As Boolean


'Dim sSQLDDL As String
 Dim db As DAO.Database
 Dim ws As DAO.Workspace
 Dim rst1 As DAO.Recordset
 Dim td As TableDef, MyField As Field

If g_Provider = 2 Then
 p = "SELECT name FROM sysobjects WHERE name = 'MSConfig'"
Set rs = g_OS.GetRecordsetFromDB(p)
If Not (rs.EOF And rs.BOF) Then

Else
sSQLDDL = "CREATE TABLE  MSConfig(RowId int NOT NULL IDENTITY  PRIMARY KEY,RecType int,FCODE int,FType int,SSC Nvarchar(100),SSV NvarCHAR(100),Con1 int,SCC Nvarchar(100),SCV NvARCHAR(100),T1 NVARchar(40),SSC1 Nvarchar(100),SSV1 NvarCHAR(100),SCC1 Nvarchar(100),SCV1 NvARCHAR(100),T2 NVARchar(40))"
  g_OS.ExecuteQueryInDB (sSQLDDL)

End If
     



End If
If g_Provider = 1 Then
    dbp = g_CDataManager.GetCompDataPath
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase _
 (dbp + "\db.bds", _
 False, False, "MS Access;PWD=ILoveMyINDIA")
For Each td In db.TableDefs
  If td.Name = "MSConfig" Then
  fo = True

  Exit For
  Else
    End If
Next td
If fo = False Then
sSQLDDL = "CREATE TABLE  MSConfig(RowId AUTOINCREMENT PRIMARY KEY,RecType int,FCODE long,FType int,SSC char(40),SSV CHAR(40),Con1 int,SCC char(40),SCV CHAR(40),T1 char(40),SSC1 char(40),SSV1 CHAR(40),SCC1 char(40),SCV1 CHAR(40),T2 char(40))"
  db.Execute sSQLDDL
End If
End If
Frame3.Enabled = False
Frame4.Enabled = False
'Frame5.Enabled = False
Frame6.Enabled = False

Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False
Combo6.Enabled = False
Combo7.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text2.Enabled = False


End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = vbBlack
Text1.ForeColor = vbWhite
Text1.FontBold = True
Text1.BorderStyle = none
Text1.Font = arial
Text1.FontSize = 10
Text1.SelStart = 0

End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = vbBlack
Text2.ForeColor = vbWhite
Text2.FontBold = True
Text2.BorderStyle = none
Text2.Font = arial
Text2.FontSize = 10
Text2.SelStart = 0

End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then
Else

Dim jk1() As String
Dim intCount As Integer

Dim lt As String
Dim b As Boolean
b = False
lt = 0
Dim i As Integer
If Text2.Text <> "" Then
    jk1 = Split(Text2.Text, "+")
    For intCount = LBound(jk1) To UBound(jk1)
        For i = 1 To ListView1.ListItems.Count
            If Trim(ListView1.ListItems(i).SubItems(1)) = Trim(jk1(intCount)) Then
            b = True
                If lt = 0 Then
                lt = i
                Else
                lt = CStr(lt) & "+" & CStr(i)
                End If
                Exit For
            Else
            b = False
            End If
        Next
    Next
  End If
If b = True Then
Label15.Caption = lt
Else
MsgBox "Wrong Value Entered in Joint Formula"
Text2.SetFocus
End If
End If
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = vbBlack
Text3.ForeColor = vbWhite
Text3.FontBold = True
Text3.BorderStyle = none
Text3.Font = arial
Text3.FontSize = 10
Text3.SelStart = 0

End Sub
Private Sub Text4_GotFocus()
Text4.BackColor = vbBlack
Text4.ForeColor = vbWhite
Text4.FontBold = True
Text4.BorderStyle = none
Text4.Font = arial
Text4.FontSize = 10
Text4.SelStart = 0

End Sub
Private Sub Text5_GotFocus()
Text5.BackColor = vbBlack
Text5.ForeColor = vbWhite
Text5.FontBold = True
Text5.BorderStyle = none
Text5.Font = arial
Text5.FontSize = 10
Text5.SelStart = 0

End Sub
Private Sub Text6_GotFocus()
Text6.BackColor = vbBlack
Text6.ForeColor = vbWhite
Text6.FontBold = True
Text6.BorderStyle = none
Text6.Font = arial
Text6.FontSize = 10
Text6.SelStart = 0

End Sub
Private Sub Text7_GotFocus()
Text7.BackColor = vbBlack
Text7.ForeColor = vbWhite
Text7.FontBold = True
Text7.BorderStyle = none
Text7.Font = arial
Text7.FontSize = 10
Text7.SelStart = 0

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text1.Text = ""
    End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If Text2.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text2.Text = ""
    End If
    End If


    
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If Text3.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text3.Text = ""
    End If
    End If


    
    

End Sub


    
    

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Text4.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text4.Text = ""
    End If
    End If


    If Combo4.Text <> "" Or Combo4.Text <> "=" Then
    Dim Where
    Where = InStr(Text4.Text, ".")
   If Where Then
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

    Else:

    Select Case KeyAscii

  Case vbKey0 To vbKey9
  Case vbKeyBack
 Case vbKeyClear
  Case vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    KeyAscii = 0
    Beep
End Select
End If
    End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Text5.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text5.Text = ""
    End If
    End If


    
    

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Text6.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text6.Text = ""
    End If
    End If

If Combo6.Text <> "" Or Combo6.Text <> "=" Then
    Dim Where
    Where = InStr(Text5.Text, ".")
   If Where Then
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

    Else:

    Select Case KeyAscii

  Case vbKey0 To vbKey9
  Case vbKeyBack
 Case vbKeyClear
  Case vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    KeyAscii = 0
    Beep
End Select
    End If
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Text7.SelStart = 0 Then
If KeyAscii <> 13 Then

        Text7.Text = ""
    End If
    End If


    

End Sub

