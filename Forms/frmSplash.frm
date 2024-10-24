VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3975
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   3795
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Import"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   11
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Make Blank All Selection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Format Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Formula Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   8
         Top             =   3240
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   6240
         TabIndex        =   7
         Top             =   120
         Width           =   690
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7080
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Image imgLogo 
         Height          =   1905
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H8000000E&
         Caption         =   "Contact Info :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H8000000E&
         Caption         =   "M-9992026622"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6525
         TabIndex        =   3
         Top             =   2400
         Width           =   330
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Basic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6015
         TabIndex        =   4
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Busy Custom Report Setting Add-on"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   1920
         TabIndex        =   6
         Top             =   1320
         Width           =   5010
      End
      Begin VB.Label A 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Arihant Softcare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         TabIndex        =   5
         Top             =   720
         Width           =   2790
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public g_Fr1 As Form2
Public g_Fr3 As Form1
Public g_Fr9 As Form3




Private Sub Form_KeyPress(KeyAscii As Integer)
'Dim sSQLDDL As String
'Dim SrNo As String
'    SrNo = g_DN
'    If SrNo = "12069007" Or SrNo = "12120261" Or SrNo = "15020238" Or SrNo = "12091542" Or SrNo = "12030090" Or SrNo = "06030240" Or SrNo = "13040017" Or SrNo = "14061654" Or SrNo = "13060504" Or SrNo = "13060718" Or SrNo = "12060326" Then
'    MsgBox "Dongal Checked"
'    Else
'    MsgBox "Please Check Dongal No."
'    End If
'        Screen.MousePointer = vbDefault
        
    
    
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbDefault
    
End Sub

Private Sub Frame1_Click()
    'Unload Me
End Sub

'Private Sub Label1_Click()
'Set g_Fr2 = New Form13
'            g_Fr2.Show vbModal
'
'End Sub


Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
Set g_Fr1 = New Form2
            g_Fr1.Show vbModal

End Sub

Private Sub Label5_Click()
Set g_Fr3 = New Form1
            g_Fr3.Show vbModal

End Sub

Private Sub Label6_Click()
Dim q As String

q = "Delete * from Externaldata where c1 ='7'"
g_OS.ExecuteQuerytmp (q)
q = "Delete * from Externaldata where c1 ='5'"
g_OS.ExecuteQuerytmp (q)
q = "Delete * from Externaldata where c1 ='6'"
g_OS.ExecuteQuerytmp (q)
q = "Delete * from Externaldata where c1 ='8'"
g_OS.ExecuteQuerytmp (q)

MsgBox "All Row Made Blank "

End Sub

Private Sub Label7_Click()
Set g_Fr9 = New Form3
            g_Fr9.Show vbModal

End Sub
