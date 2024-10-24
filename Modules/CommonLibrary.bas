Attribute VB_Name = "CommonLibrary"
Option Explicit

Public Function DivideNum(ByVal p_Num1 As Double, ByVal p_Num2 As Double, Optional p_Decimal As Integer = -1) As Double
    Dim RetVal As Double
    If Not (p_Num2 = 0) Then
        RetVal = p_Num1 / p_Num2
        If p_Decimal <> -1 Then
            DivideNum = Round(RetVal, p_Decimal)
        Else
            DivideNum = RetVal
        End If
    Else
        DivideNum = 0
    End If
End Function
Public Sub FlushCol(ByRef p_COl As Collection)
    Set p_COl = Nothing
    Set p_COl = New Collection

End Sub
Public Function DateStr(p_Date As Date) As String
        DateStr = "#" & Format(p_Date, "mm-dd-yyyy") & "#"
End Function
Public Function FormatDate(ByVal DateVal As Date, Optional p_PutLeadingZeros As Boolean = False) As String
    Dim sDay As String
    Dim sMonth As String
    Dim sYear As String
    Dim sRetVal As String
    
    If Not p_PutLeadingZeros Then
        sDay = CStr(Day(DateVal))
        sMonth = CStr(Month(DateVal))
    Else
        sDay = Format(Day(DateVal), "#0#")
        sMonth = Format(Month(DateVal), "#0#")
    End If
    
    sYear = CStr(Year(DateVal))
    
    If g_CC.DateFormat = DD_MM_YYYY Then
        sRetVal = sDay + g_CC.DateSeperator + sMonth + g_CC.DateSeperator + sYear
    Else
        sRetVal = sMonth + g_CC.DateSeperator + sDay + g_CC.DateSeperator + sYear
    End If
    
    FormatDate = sRetVal
    
End Function

Public Function Text2Date(ByVal sStr As String) As Date
    Dim sComp1 As String
    Dim sComp2 As String
    Dim sComp3 As String
    Dim sTempStr As String
    Dim nPos As Integer
    Dim nLen As Integer
    Dim nCompNo As Integer
    Dim i As Integer
    Dim nYear As Integer
    Dim DayVal As Integer, MonthVal As Integer, YearVal As Integer

    On Error GoTo ErrHandler

    nLen = Len(sStr)

    nCompNo = 1

    For i = 1 To nLen

        sTempStr = Mid$(sStr, i, 1)

        If Asc(sTempStr) >= vbKey0 And Asc(sTempStr) <= vbKey9 Then

            Select Case nCompNo

                    Case 1
                            sComp1 = sComp1 + sTempStr

                    Case 2
                            sComp2 = sComp2 + sTempStr

                    Case 3
                            sComp3 = sComp3 + sTempStr

            End Select
        Else
            nCompNo = nCompNo + 1
            If nCompNo > 3 Then GoTo ErrHandler
        End If
    Next

    If Len(sComp3) = 1 Or Len(sComp3) = 3 Then
        GoTo ErrHandler
    End If

    If Len(sComp3) = 2 Then
        nYear = CInt(Left$(sComp3, 2))

        If nYear < 80 Then
            nYear = nYear + 2000
        Else
            nYear = nYear + 1900
        End If

        sComp3 = CStr(nYear)
    End If

    If g_CC.DateFormat = DD_MM_YYYY Then
        DayVal = CInt(sComp1)
        MonthVal = CInt(sComp2)
    Else
        MonthVal = CInt(sComp1)
        DayVal = CInt(sComp2)
    End If

    YearVal = CInt(sComp3)

    sTempStr = MonthName(MonthVal) & " " & CStr(DayVal) & " " & CStr(YearVal)

    Text2Date = CDate(sTempStr)

    Exit Function

ErrHandler:

    Text2Date = 0

End Function

