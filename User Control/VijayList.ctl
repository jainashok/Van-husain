VERSION 5.00
Begin VB.UserControl VijayList 
   BackColor       =   &H80000005&
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   ScaleHeight     =   2160
   ScaleWidth      =   5295
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5295
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   5295
   End
End
Attribute VB_Name = "VijayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer)

Enum ReqOptList3
    [NotBlank]  ' Must be bracketed because Optional is a VB keyword
    Blank
End Enum

Enum InputTypeFormat
    [Any]  ' Must be bracketed because Any is a VB keyword
    TextOnly
    NumbersWithDecimal
    NumbersWithoutDecimal
    PhoneNo
    EmailId
    YesNo
End Enum



Enum TxtCase
    UserInput
    UpperCase
    LowerCase
    SentenceCase
End Enum

'Default Property Values:
Const m_def_Mandatory = 0
Const m_def_InputType = 0
Const m_def_TextCase = 0
Const m_def_MandatoryColor = &H80000005     '&H8080FF      '&HFF
Const m_def_EnterFocusColor = &H80000005
Const m_def_LeaveFocusColor = &H80000005                '&H0&

'Property Variables:
Dim m_Mandatory As ReqOptList3
Dim m_InputType As InputTypeFormat
Dim m_TextCase As TxtCase
Dim m_InvalidCharacters As String
Dim m_AutoSelect As Boolean

Dim m_EnterKeySupport As Boolean
Dim m_MandatoryColor As OLE_COLOR
Dim m_EnterFocusColor As OLE_COLOR
Dim m_LeaveFocusColor As OLE_COLOR

Dim m_Text_And_List As Boolean
Dim For_New_Entry As Boolean
Dim m_List_Sorted As Boolean


Public Stopper As Integer



Public Property Get List() As ListBox
    Set List = List1
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = txtText.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtText.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtText.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtText.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = txtText.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtText.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
    Set Font = txtText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtText.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub txtText_Change()


RaiseEvent Change
End Sub


Private Sub txtText_GotFocus()
'    Stopper = 0
'    txtText.Move 0, 0, UserControl.Width, UserControl.Height
    
    txtText.BackColor = &H3F3F3F      '&H8080FF    'EnterFocusColor
    txtText.ForeColor = &H80000005
    
    

    
    
    If Not txtText = "" And AutoSelect Then
        txtText.SelStart = 0
        For_New_Entry = False
        
'        txtText.SelLength = Len(txtText)
    Else
        txtText.SelStart = 0
    End If
'List1.Visible = True

        
    If Len(txtText) > 0 Then
        List1.Text = txtText
    End If

End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)

'
RaiseEvent KeyDown(KeyCode, Shift)

    


'
'    If KeyCode = 13 Or KeyCode = vbKeyF2 Or KeyCode = 40 Then
'
'        If EnterKeySupport Then
'            SendKeys "{TAB}"
'        End If
'        Exit Sub
'    End If
'
'    If KeyCode = 38 Then
'        SendKeys "+{TAB}"
'        Exit Sub
'    End If
'
'    If KeyCode = 8 And txtText = "" Then
'        SendKeys "+{TAB}"
'        Exit Sub
'    End If
'

'

'List1.Text = txtText


    

If txtText.SelStart = 0 And KeyCode = 8 Then
        
        
        List1.Visible = False
        UserControl.Height = txtText.Height
   
        KeyCode = 0
        For_New_Entry = False
        SendKeys "+{TAB}"
            
        Exit Sub

End If


    If txtText.SelStart = 0 And KeyCode = 8 And For_New_Entry = False Then
        
        
'        List1.Text = txtText
   
        KeyCode = 0
        For_New_Entry = False
        SendKeys "+{TAB}"
     
     
'     MsgBox txtText.SelStart, , List1.ListIndex
     
        
        Exit Sub

    End If



'    If txtText.SelStart = 0 And For_New_Entry = True Then
'        txtText.Text = ""
'    End If

     For_New_Entry = True

'
     If KeyCode = 38 And List1.Visible = True Then
        If List1.ListIndex > 0 Then
            List1.ListIndex = List1.ListIndex - 1
        End If
    ElseIf KeyCode = 38 And List1.Visible = False Then
'        RemarksKey38 = True
         SendKeys "+{Tab}"
         Exit Sub
    End If
    If KeyCode = 40 And List1.Visible = True Then
        If List1.ListIndex < List1.ListCount - 1 Then
            List1.ListIndex = List1.ListIndex + 1
        End If
    ElseIf KeyCode = 40 And List1.Visible = False Then
        SendKeys "{Tab}"
        
    End If
    
    
    If KeyCode = 37 Or KeyCode = 39 Then Exit Sub
    
    
'    If KeyCode = 13 Or KeyCode = vbKeyF2 Then
    If KeyCode = 13 Then
    
        If List1.ListIndex >= 0 Then
            txtText = List1.Text
            List1.Visible = False
            UserControl.Height = txtText.Height
            
 'xxxxxxxxxxxxxxxxxxxxx for department xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'            If frmStkTransfer.txtDep1.Text = frmStkTransfer.txtDep2.Text And Len(frmStkTransfer.txtDep2.Text) > 0 Then
'                With frmError
'                   .Capt = "Error!!!"
'                   .Msg = "Department Can Not Be Same !!!"
'                   .ImgStop.Visible = True
'
'                   .Show vbModal, Me
'                   frmStkTransfer.txtDep2.Text = ""
'                   frmStkTransfer.txtDep2.SetFocus
'                End With
'            Exit Sub
'            End If
  'xxxxxxxxxxxxxxxxxxxxx for department xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            
            
            
            
            
            
            SendKeys "{tab}"
            
        ElseIf Len(txtText.Text) = 0 Then
        
            List1.Visible = True
            
            
            txtText_KeyPress 32
        Else
            SendKeys "{tab}"
        End If
    End If
    
    If KeyCode = vbKeyDelete Then
        List1.ListIndex = -1
    Exit Sub
    End If
    
    If KeyCode = 35 Then Exit Sub
    
    KeyCode = 0

    
End Sub




Private Sub txtText_KeyPress(KeyAscii As Integer)


If KeyAscii = 8 And For_New_Entry = False Then


    
        KeyAscii = 0
        Exit Sub
        
    End If
    




    Dim eno, Nm As String
    Dim i As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 27 Then
        List1.Visible = False
        UserControl.Height = txtText.Height
        Exit Sub
    End If
    
    
    If txtText.SelStart = 0 And For_New_Entry = True Then
        txtText.Text = ""
        
    End If
    txtText.Move 0, 0, UserControl.Width, 255 '2175 - List1.Height
    
    
    
    List1.Width = txtText.Width
'    List1.Top = txtText.Top + txtText.Height
    List1.Left = txtText.Left
    List1.Visible = True
    
    
    
'    AuctusTextList.Move 0, 0, UserControl.Width, 2500
    
    UserControl.Height = txtText.Height + List1.Height + 100 '3000
    UserControl.BackColor = &H80000005
    
'jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
    If KeyAscii = 8 Then
        List1.ListIndex = 0
        
            eno = UCase(Trim(Left(txtText, Len(txtText) - 1)))
        For i = 0 To List1.ListCount - 1
        Nm = Mid(List1.List(i), 1, Len(eno))
        If UCase(eno) = UCase(Nm) Then
            List1.ListIndex = i
            List1.TopIndex = i
            Stopper = i
            Exit Sub
        Else
            List1.ListIndex = Stopper
        End If
        Next
    
        Exit Sub
    End If
'jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj

    eno = UCase(Trim(txtText & Chr(KeyAscii)))
    For i = 0 To List1.ListCount - 1
        Nm = Mid(List1.List(i), 1, Len(eno))
        If UCase(eno) = UCase(Nm) Then
            List1.ListIndex = i
            List1.TopIndex = i
            Stopper = i
            Exit Sub
        Else
            List1.ListIndex = Stopper
        End If
    Next
    
    
    
    If KeyAscii <> 8 And KeyAscii <> 38 And KeyAscii <> 40 Then
        KeyAscii = 0
'
        Exit Sub
    End If














Exit Sub


    'Case
    Select Case TextCase
        Case UpperCase
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case LowerCase
            KeyAscii = Asc(LCase(Chr(KeyAscii)))
        Case SentenceCase
            If txtText.SelStart = 0 Then
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Else
                If Mid(txtText, txtText.SelStart, 1) = " " Then
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
            End If
    End Select
    
    'Check for invalid characters
    If Not InvalidCharacters = "" Then
        Dim PressedChar As String
        PressedChar = Chr(KeyAscii)
        If InStr(1, InvalidCharacters, PressedChar) > 0 Then
'
            KeyAscii = 0
        End If
    End If
    
    'Input Type validation
    Select Case InputType
        Case TextOnly
            If (KeyAscii <= 90 And KeyAscii >= 65) Or (KeyAscii <= 122 And KeyAscii >= 97) Or KeyAscii = 32 Then
                'Ok
            Else
                KeyAscii = 0
            End If
        Case NumbersWithDecimal
            If KeyAscii <= 57 And KeyAscii >= 48 Or KeyAscii = 46 Or KeyAscii = 8 Then
                'Decimal should be pressed once
                If KeyAscii = 46 Then
                    If InStr(1, txtText, ".") > 0 Then
                        KeyAscii = 0
                    Else
                        'Ok
                    End If
                Else
                    'Ok
                End If
            Else
                KeyAscii = 0
            End If
        Case NumbersWithoutDecimal
            If KeyAscii <= 57 And KeyAscii >= 48 Then
                'Ok
            Else
                KeyAscii = 0
            End If
        Case PhoneNo
            'Allowed Characters: Numbers(0-9), Space(" "), Open Bracket"(", Close Bracket")" , Open Sqare Bracket"[" , Close Sqare Bracket"]", Dash "-"
        Case EmailId
            'Not Allowed Characters: Space " "
    End Select
    
End Sub

Private Sub txtText_LostFocus()
''ani;=================================================
'If Mandatory = NotBlank And Len(txtText.Text) = 0 Then
'        With frmError
'                .Capt = "Blank Field !!!"
'                .Msg = "This Field Can Not Be Blank !!!"
'                .ImgStop.Visible = True
'
'                .Show vbModal, Me
'            End With
'            txtText.SetFocus
'
'    Exit Sub
'End If

'anil=======================================================

txtText.Move 0, 0, UserControl.Width, 255
UserControl.Height = txtText.Height

txtText.BackColor = &H80000005     ' '&H8080FF    'EnterFocusColor
txtText.ForeColor = vbBlack
    

If List1.ListIndex = -1 Then
'    List1.ListIndex = 0
    Exit Sub
End If

'If txtText.SelStart = 0 And List1.Visible = False Then Exit Sub
'MsgBox txtText.Text, , Len(txtText.Text)
If Len(txtText) = 0 Then Exit Sub

txtText.Text = List1.Text
    If List1.Visible = True Then
        List1.Clear
        List1.Visible = False
        UserControl.Height = txtText.Height
    End If
Exit Sub



'    If Mandatory = Required And Trim(txtText.Text) = "" Then
'        txtText.BackColor = &H80000005&   ' &H8080FF 'm_MandatoryColor
'    Else
'        txtText.BackColor = &H80000005&    '&H8080FF 'm_LeaveFocusColor
'    End If
    txtText.ForeColor = vbBlack
    
    
    List1.Visible = False
    UserControl.Height = txtText.Height
End Sub

'Public Property Get EnterFocusColor() As OLE_COLOR
'    EnterFocusColor = m_EnterFocusColor
'End Property
'
'Public Property Let EnterFocusColor(ByVal New_EnterFocusColor As OLE_COLOR)
'    m_EnterFocusColor = New_EnterFocusColor
'    PropertyChanged "EnterFocusColor"
'End Property

'Public Property Get LeaveFocusColor() As OLE_COLOR
'    LeaveFocusColor = m_LeaveFocusColor
'End Property
'
'Public Property Let LeaveFocusColor(ByVal New_LeaveFocusColor As OLE_COLOR)
'    m_LeaveFocusColor = New_LeaveFocusColor
'    PropertyChanged "LeaveFocusColor"
'End Property

Public Property Get MandatoryColor() As OLE_COLOR
    MandatoryColor = m_MandatoryColor
End Property

Public Property Let MandatoryColor(ByVal New_MandatoryColor As OLE_COLOR)
    m_MandatoryColor = New_MandatoryColor
    
    
    
    If m_Mandatory Then txtText.BackColor = New_MandatoryColor
    PropertyChanged "MandatoryColor"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = txtText.MouseIcon
End Property

Private Sub UserControl_Initialize()
    UserControl.Height = 3500
    txtText.BackColor = &H80000005
'    UserControl.BackColor = vbRed
    
    EnterKeySupport = True
    
'     UserControl.m_def_EnterFocusColor = &H00C0E0FF&
' UserControl.m_def_LeaveFocusColor = &H8080FF            '&H0&

    '&H0080C0FF&
'    txtText.Move 0, 0, UserControl.Width, UserControl.Height
'    txtText.Move 0, 0, UserControl.Width, 285 'txtText.Height + List1.Height
    
'    txtText.Move 0, 0, 3000, 285
    
    txtText.Move 0, 0, UserControl.Width, 255
    m_AutoSelect = True
End Sub


Public Property Get InputType() As InputTypeFormat
    InputType = m_InputType
End Property

Public Property Let InputType(ByVal vNewValue As InputTypeFormat)
    m_InputType = vNewValue
    PropertyChanged "InputType"
End Property

Public Property Get Mandatory() As ReqOptList3
    Mandatory = m_Mandatory
End Property

Public Property Let Mandatory(ByVal vNewValue As ReqOptList3)
    m_Mandatory = vNewValue
    PropertyChanged "Mandatory"
End Property

Public Property Get TextCase() As TxtCase
    TextCase = m_TextCase
End Property

Public Property Let TextCase(ByVal vNewValue As TxtCase)
    m_TextCase = vNewValue
    PropertyChanged "TextCase"
End Property

Public Property Get InvalidCharacters() As String
    InvalidCharacters = m_InvalidCharacters
End Property

Public Property Let InvalidCharacters(ByVal vNewValue As String)
    m_InvalidCharacters = vNewValue
    PropertyChanged "InvalidCharacters"
End Property


'm_List_Sorted


Public Property Get TextAndList() As Boolean
    TextAndList = m_Text_And_List
End Property
Public Property Let TextAndList(ByVal vNewValue As Boolean)
    m_Text_And_List = vNewValue
    PropertyChanged "TextAndList"
End Property



Public Property Get List_Not_Sorted() As Boolean
    List_Not_Sorted = m_List_Sorted
End Property

Public Property Let List_Not_Sorted(ByVal vNewValue As Boolean)
    m_List_Sorted = vNewValue
    PropertyChanged "List_Not_Sorted"
End Property



Public Property Get AutoSelect() As Boolean
    AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(ByVal vNewValue As Boolean)
    m_AutoSelect = vNewValue
    PropertyChanged "AutoSelect"
End Property

Public Property Get EnterKeySupport() As Boolean
    EnterKeySupport = m_EnterKeySupport
End Property

Public Property Let EnterKeySupport(ByVal vNewValue As Boolean)
    m_EnterKeySupport = vNewValue
    PropertyChanged "EnterKeySupport"
End Property

Private Sub UserControl_Resize()
txtText.Move 0, 0, UserControl.Width, 255
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("InputType", m_InputType, m_def_InputType)
    Call PropBag.WriteProperty("TextCase", m_TextCase, m_def_TextCase)
    Call PropBag.WriteProperty("InvalidCharacters", m_InvalidCharacters)
    Call PropBag.WriteProperty("AutoSelect", m_AutoSelect)
    Call PropBag.WriteProperty("EnterKeySupport", m_EnterKeySupport)
    Call PropBag.WriteProperty("PasswordChar", txtText.PasswordChar, "")
    Call PropBag.WriteProperty("Text", txtText.Text, "txtText")
    Call PropBag.WriteProperty("BackColor", txtText.BackColor, &H8080FF)
    Call PropBag.WriteProperty("ForeColor", txtText.ForeColor, &H8080FF)
    Call PropBag.WriteProperty("Enabled", txtText.Enabled, True)
    Call PropBag.WriteProperty("Font", txtText.Font, Ambient.Font)
    
'
'    Call PropBag.WriteProperty("EnterFocusColor", m_EnterFocusColor, m_def_EnterFocusColor)
'    Call PropBag.WriteProperty("LeaveFocusColor", m_LeaveFocusColor, m_def_LeaveFocusColor)
'    Call PropBag.WriteProperty("MandatoryColor", m_MandatoryColor, m_def_MandatoryColor)

'    Call PropBag.WriteProperty("EnterFocusColor", m_EnterFocusColor, &H8080FF)
'    Call PropBag.WriteProperty("LeaveFocusColor", m_LeaveFocusColor, m_def_LeaveFocusColor)
'    Call PropBag.WriteProperty("MandatoryColor", m_MandatoryColor, &H8080FF)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '&H80000008&
'    txtText.BackColor = PropBag.ReadProperty("BackColor", &H8080FF)
'    txtText.ForeColor = PropBag.ReadProperty("ForeColor", &H8080FF)
'
    txtText.BackColor = PropBag.ReadProperty("BackColor", &H80000008)
    txtText.ForeColor = PropBag.ReadProperty("ForeColor", &H3F3F3F)
    
    
    txtText.Enabled = PropBag.ReadProperty("Enabled", True)
    
    Set txtText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    m_InputType = PropBag.ReadProperty("InputType", m_def_InputType)
    m_TextCase = PropBag.ReadProperty("TextCase", m_def_TextCase)
    m_InvalidCharacters = PropBag.ReadProperty("InvalidCharacters", "")
    m_AutoSelect = PropBag.ReadProperty("AutoSelect", True)
    m_EnterKeySupport = PropBag.ReadProperty("EnterKeySupport", True)
    txtText.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    txtText.Text = PropBag.ReadProperty("Text", "txtText")
'    m_EnterFocusColor = PropBag.ReadProperty("EnterFocusColor", m_def_EnterFocusColor)
'    m_LeaveFocusColor = PropBag.ReadProperty("LeaveFocusColor", m_def_LeaveFocusColor)
'    m_MandatoryColor = PropBag.ReadProperty("MandatoryColor", m_def_MandatoryColor)

'
'm_EnterFocusColor = PropBag.ReadProperty("EnterFocusColor", &H8080FF)
'    m_LeaveFocusColor = PropBag.ReadProperty("LeaveFocusColor", m_def_LeaveFocusColor)
'    m_MandatoryColor = PropBag.ReadProperty("MandatoryColor", m_def_MandatoryColor)

End Sub

Public Property Get PasswordChar() As String
    PasswordChar = txtText.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtText.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Text
Public Property Get Text() As String
    Text = txtText.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtText.Text() = New_Text
    PropertyChanged "Text"
End Property




