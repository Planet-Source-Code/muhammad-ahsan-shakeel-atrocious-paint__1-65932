Attribute VB_Name = "mdlGeneral"

Option Explicit

Public Enum enmError
  conErrWrite = 1
  conErrPrint = 2
  conErrReadImage = 3
  conErrDrawing = 4
  conErrPermission = 70
  conErrCancel = 32755
  conErrOthers = 0
End Enum

Public Function blnFileExist(strFileName As String) As Boolean
  
  Dim blnReturn As Boolean
  Dim fso As Scripting.FileSystemObject

  Set fso = New Scripting.FileSystemObject
  blnReturn = fso.FileExists(strFileName)
  Set fso = Nothing
  blnFileExist = blnReturn
  Exit Function

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Function

Public Function ForceSave(strFileName As String) As Boolean

  Dim fso As Scripting.FileSystemObject

  If MsgBox("The file is read-only/hidden. " & vbNewLine & _
            "Are you sure you want to write into the file " & _
            "and remove the read-only/hidden property?", _
            vbYesNo + vbQuestion) = vbYes Then
    Set fso = New Scripting.FileSystemObject
    fso.GetFile(strFileName).Attributes = 0
    ForceSave = True
  Else
    ForceSave = False
  End If
  Exit Function

ErrorHandler:
  ForceSave = False
  ShowErrMessage intErr:=conErrWrite
End Function

Public Sub ShowErrMessage(intErr As enmError, Optional strErrMessage As String)
  Select Case intErr
    Case conErrWrite
      MsgBox "Cannot write to the disk." & vbNewLine & vbNewLine & _
               "Make sure the disk is not full or write-protected.", _
             vbOKOnly + vbCritical
    Case conErrPrint
      MsgBox "Cannot print the file." & vbNewLine & vbNewLine & _
               "Make sure the print is ready.", vbOKOnly + vbCritical
    Case conErrReadImage
      MsgBox "Cannot open the file." & vbNewLine & vbNewLine & _
               "The file may be corrupt or not a valid picture file.", _
             vbOKOnly + vbCritical
    Case conErrDrawing
      MsgBox "Cannot drawing using the selected tool." & _
               vbNewLine & vbNewLine & _
               "The needed file may be missing.", _
             vbOKOnly + vbCritical
    Case conErrOthers
      MsgBox strErrMessage, vbOKOnly + vbCritical
  End Select
End Sub

Public Function strGetFileName(strPath As String, _
                               Optional blnNoExt As Boolean = True, _
                               Optional blnNoPath As Boolean = True) As String
  Dim intIxDot As Integer
  Dim strReturn As String
  
  If blnNoPath Then
    strReturn = Dir(strPath)
  Else
    strReturn = strPath
  End If
  If blnNoExt Then
    intIxDot = InStrRev(strReturn, ".")
    If intIxDot <> 0 Then
      strReturn = Left(strReturn, intIxDot - 1)
    End If
  End If
  strGetFileName = strReturn
  Exit Function

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Function

Public Function varIIf(blnCondition As Boolean, _
                        varTrue As Variant, varFalse As Variant) As Variant
    
  If blnCondition Then
    varIIf = varTrue
  Else
    varIIf = varFalse
  End If
  Exit Function

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Function
