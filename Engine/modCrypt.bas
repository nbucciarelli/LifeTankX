Attribute VB_Name = "modCrypt"
Option Explicit
Dim intTextCount As Integer
Dim intCryptKeyCount As Integer
Dim strUnlockKey As String
Const CryptoKey = "TwilightTank"

Public Function Crypt(strText As String) As String
On Error Resume Next
Dim intChrText As Integer
Dim intChrKey As Integer
Dim intCombineChr As Integer
Dim finalCrypt As String

intTextCount = 1
intCryptKeyCount = 1
strUnlockKey = CryptoKey

If strUnlockKey = "" Then
    PrintMessage "Invalid Cryptography Key passed to Crypt function."
    Exit Function
End If

If strText = "" Then
    PrintMessage "Invalid Text to Encrypt passed to Crypt function."
    Exit Function
End If
While intTextCount <= Len(strText)

If intCryptKeyCount >= Len(strUnlockKey) Then intCryptKeyCount = 1
    intChrText = g_Asc(Mid(strText, intTextCount, 1))
    intChrKey = g_Asc(Mid(strUnlockKey, intCryptKeyCount, 1))
    intCombineChr = intChrText + intChrKey
If intCombineChr > 255 Then intCombineChr = intCombineChr - 255
finalCrypt = finalCrypt + g_Chr(intCombineChr)
intTextCount = intTextCount + 1
intCryptKeyCount = intCryptKeyCount + 1
Wend
Crypt = finalCrypt
End Function

Public Function Decrypt(strText As String) As String
On Error Resume Next
Dim intChrText As Integer
Dim finalDecrypt As String
intTextCount = 1
intCryptKeyCount = 1
strUnlockKey = CryptoKey
If strUnlockKey = "" Then
    PrintMessage "Invalid Cryptography Key passed to Decrypt function."
    Exit Function
End If

If strText = "" Then
    PrintMessage "Invalid Text to Encrypt passed to Decrypt function."
    Exit Function
End If


While intTextCount <= Len(strText)
If intCryptKeyCount >= Len(strUnlockKey) Then intCryptKeyCount = 1
    intChrText = g_Asc(Mid(strText, intTextCount, 1)) - g_Asc(Mid(strUnlockKey, intCryptKeyCount, 1))
    intChrText = intChrText + 255

If intChrText > 255 Then intChrText = intChrText - 255
    finalDecrypt = finalDecrypt & g_Chr(intChrText)
    intTextCount = intTextCount + 1
    intCryptKeyCount = intCryptKeyCount + 1
Wend
Decrypt = finalDecrypt
End Function

Public Function DecASCString(bStrText As String) As String
Dim i As Integer
Dim splitStr() As String
Dim strOutput As String

splitStr() = Split(bStrText, "-")
For i = LBound(splitStr) To UBound(splitStr)
    strOutput = strOutput & g_Chr(CLng(splitStr(i)))
Next i

DecASCString = strOutput
End Function

