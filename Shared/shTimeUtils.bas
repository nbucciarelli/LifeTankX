Attribute VB_Name = "shTimeUtils"
Option Explicit

Private Const SEC_IN_MIN = 60
Private Const SEC_IN_HOUR = 3600
Private Const SEC_IN_DAY As Long = 86400

Private Declare Sub GetLocalTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Public Enum eTimeFormat
    TF_LETTERS  '1h 30m 23s
    TF_NUMBERS  '01:30:23
End Enum

Public Function myFormatTime(ByVal dSeconds As Double, Optional ByVal iFormat As eTimeFormat = TF_LETTERS) As String
On Error GoTo ErrorHandler

    Dim lDays As Long
    Dim iHours As Integer
    Dim iMinutes As Integer
    Dim iSeconds As Integer
    Dim dRemaining As Double
    Dim dMod As Double
    Dim sRet As String
    
1    dRemaining = dSeconds
    
    If dRemaining > CDbl(31104000) Then
        sRet = "1+ year"
        GoTo Fin
    End If
    
    'Days
2    dMod = CDbl(dRemaining Mod SEC_IN_DAY)
3    lDays = CLng((dRemaining - dMod) / SEC_IN_DAY)
4    dRemaining = dMod
    
    'Hours
5    dMod = CDbl(dRemaining Mod SEC_IN_HOUR)
6    iHours = CInt((dRemaining - dMod) / SEC_IN_HOUR)
7    dRemaining = dMod
    
    'Minutes
8    dMod = CDbl(dRemaining Mod SEC_IN_MIN)
9    iMinutes = CInt((dRemaining - dMod) / SEC_IN_MIN)
10    dRemaining = dMod
    
    'Seconds
11    iSeconds = CInt(dRemaining)
    
    sRet = ""
    If iFormat = TF_LETTERS Then
        If lDays > 0 Then sRet = sRet & lDays & "d "
        
        If iHours > 0 Then
            'If iHours < 10 Then sRet = sRet & "0"
            sRet = sRet & iHours & "h "
        Else
            If lDays > 0 Then sRet = sRet & "0h "
        End If
        
        If iMinutes > 0 Then
            'If iMinutes < 10 Then sRet = sRet & "0"
            sRet = sRet & iMinutes & "m "
        Else
            If iHours > 0 Or lDays > 0 Then sRet = sRet & "0m "
        End If
        
        sRet = sRet & iSeconds & "s"
    Else
        If lDays > 0 Then sRet = sRet & lDays & ":"
        
        If iHours > 0 Then
            If iHours < 10 Then sRet = sRet & "0"
            sRet = sRet & iHours & ":"
        Else
            sRet = sRet & "00:"
        End If
        
        If iMinutes > 0 Then
            If iMinutes < 10 Then sRet = sRet & "0"
            sRet = sRet & iMinutes & ":"
        Else
            sRet = sRet & "00:"
        End If
        
        If iSeconds < 10 Then sRet = sRet & "0"
        sRet = sRet & iSeconds
    End If
    
Fin:
    myFormatTime = sRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "shTimeUtils.myFormatTime - " & Err.Description & " - Line: " & Erl
    sRet = "-Err-"
    Resume Fin
End Function

Public Function CurrentTime() As String
  Dim theTime As SYSTEMTIME
  Call GetLocalTime(theTime)
  With theTime
      CurrentTime = Format(.wMonth, "00") & "/" & Format(.wDay, "00") & "/" & Format(.wYear, "0000") & " " & Format(.wHour, "00") & ":" & Format(.wMinute, "00") & ":" & Format(.wSecond, "00") & "." & Format(.wMilliseconds, "000")
  End With
End Function
