VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmListToExcel 
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdToExcel 
      Caption         =   "SaveToExcel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwToExcel 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3836
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmListToExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
   'this to prevent editing from user:
   lvwToExcel.LabelEdit = lvwManual
   'add some columns header:
   lvwToExcel.ColumnHeaders.Add , "Col01", "This is Col 1"
   'enlarfge the above column:
   lvwToExcel.ColumnHeaders("Col01").Width = 2000
   lvwToExcel.ColumnHeaders.Add , "Col02", "This is Col 2"
   'enlarge the above column:
   lvwToExcel.ColumnHeaders("Col02").Width = lvwToExcel.Width - 2100
   'show them
   lvwToExcel.HideColumnHeaders = False
   'Show a report of item, so you will see columns
   lvwToExcel.View = lvwReport
   
   'show selected item even when loosing focus
   lvwToExcel.HideSelection = False
   'add some items to the list
   lvwToExcel.ListItems.Add , "item01", "I will be saved in A1"
   lvwToExcel.ListItems("item01").SubItems(1) = "This descript of first item will go in cell B1"
   lvwToExcel.ListItems.Add , "item02", "I will be saved in A2"
   lvwToExcel.ListItems("item02").SubItems(1) = "This descript of second item will go in cell B2"
End Sub
Private Sub CmdToExcel_Click()
   Screen.MousePointer = vbHourglass
   'Now, you have different options:
   'Are you going to create a new excel sheet each time?
   'Are you going to update the same sheet?
   'Do you want user to choose?
   'As this is only an example, I will quicly create a new
   'excel sheet. If the one I am looking for exists, I will
   'delete it before recreating
   '**********************************
   'early binding: requires you to put a reference to
   'Microsoft Excel data Object
'   Dim objXcell As Excel.Application
'   Dim objWbook As Excel.Workbook
'   Dim objSheet As Excel.Worksheet
'   Set objXcell = New Excel.Application
   '**********************************
   'late binding: you can avoid reference as long as
   'excel is correctly installed in client machine:
   Dim objXcell As Object
   Dim objWbook As Object
   Dim objSheet As Object
   Set objXcell = CreateObject("Excel.Application")
   '**********************************
   Set objWbook = objXcell.Workbooks.Add
   Set objSheet = objWbook.Worksheets.Add
   With objSheet
   If Dir("c:\AAAGame.xls") <> "" Then
      Kill "c:\AAAGame.xls"
   End If
   'Now, you can read from your mdb and write to excel.
   'But as here I only have a sample listview,
   'I will read from listview and put in excel
   Dim lngCounter As Long
   For lngCounter = 1 To lvwToExcel.ListItems.Count
      'col 1 in first col, col 2 in second col
      objSheet.Cells(lngCounter, 1).Value = lvwToExcel.ListItems(lngCounter).Text
      objSheet.Cells(lngCounter, 2).Value = lvwToExcel.ListItems(lngCounter).SubItems(1)
   Next
      .SaveAs "c:\AAAGame.xls"
   End With
   'free memory
   objWbook.Close False
   Set objSheet = Nothing
   Set objWbook = Nothing
   objXcell.Quit
   Set objXcell = Nothing
   Screen.MousePointer = vbDefault
   End Sub
