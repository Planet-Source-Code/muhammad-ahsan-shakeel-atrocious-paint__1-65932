Attribute VB_Name = "RegisterModule"
Public sKeys As Collection

Public Sub SaveKey(hKey As Long, strPath As String)

Dim Keyhand&

R = RegCreateKey(hKey, strPath, Keyhand&)

R = RegCloseKey(Keyhand&)
    
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)

Dim Keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer

R = RegOpenKey(hKey, strPath, Keyhand)

lResult = RegQueryValueEx(Keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(Keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
        Else
            GetString = strBuf
        End If
    End If
End If

End Function


Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)

Dim Keyhand As Long
Dim R As Long

R = RegCreateKey(hKey, strPath, Keyhand)

R = RegSetValueEx(Keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))

R = RegCloseKey(Keyhand)
    
End Sub


Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long

Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim R As Long
Dim Keyhand As Long

R = RegOpenKey(hKey, strPath, Keyhand)

lDataBufSize = 4
  
lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        GetDWord = lBuf
    End If
End If

R = RegCloseKey(Keyhand)
    
End Function

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    
Dim lResult As Long
Dim Keyhand As Long
Dim R As Long
    
R = RegCreateKey(hKey, strPath, Keyhand)

lResult = RegSetValueEx(Keyhand, strValueName, 0&, REG_DWORD, lData, 4)

R = RegCloseKey(Keyhand)
    
End Function

Public Function DeleteKey(ByVal hKey As Long, ByVal StrKey As String)

Dim R As Long

R = RegDeleteKey(hKey, StrKey)
    
End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)

Dim Keyhand As Long

R = RegOpenKey(hKey, strPath, Keyhand)

R = RegDeleteValue(Keyhand, strValue)

R = RegCloseKey(Keyhand)
    
End Function

Public Sub GetKeyNames(ByVal hKey As Long, ByVal strPath As String)

Dim Cnt As Long, StrBuff As String, StrKey As String, TKey As Long
    
RegOpenKey hKey, strPath, TKey

    Do
        StrBuff = String(255, vbNullChar)
        If RegEnumKeyEx(TKey, Cnt, StrBuff, 255, 0, vbNullString, 0, ByVal 0&) <> 0 Then Exit Do
        Cnt = Cnt + 1
        StrKey = Left(StrBuff, InStr(StrBuff, vbNullChar) - 1)
        sKeys.Add StrKey
    Loop
    
End Sub

