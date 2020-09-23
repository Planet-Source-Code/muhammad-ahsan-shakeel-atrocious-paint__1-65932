Attribute VB_Name = "mdlAPI"

Option Explicit

Public Type typPoint
  x As Long
  y As Long
End Type

Public Declare Sub _
  ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
                            ByVal y As Long, ByVal crColor As Long, _
                            ByVal wFillType As Long)

Public Declare Function _
  GetPixel Lib "gdi32" (ByVal hDC As Long, _
                        ByVal x As Long, ByVal y As Long) As Long

Public Declare Sub _
  PolyBezier Lib "gdi32" (ByVal hDC As Long, _
                          lppt As typPoint, ByVal cPoints As Long)

Public Declare Sub _
  Polygon Lib "gdi32" (ByVal hDC As Long, _
                       lpPoint As typPoint, ByVal nCount As Long)

Public Declare Sub _
  RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, _
                         ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
                         ByVal X3 As Long, ByVal Y3 As Long)

Public Declare Sub _
  SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
                        ByVal y As Long, ByVal crColor As Long)

Public Declare Sub _
  ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
     ByVal lpFile As String, ByVal lpParameters As String, _
     ByVal lpDirectory As String, ByVal nShowCmd As Long)
