Attribute VB_Name = "mdlFilter"

Option Explicit

Public Enum enmFilter
  conFltBlacknWhite = 0
  conFltBlur = 1
  conFltBrightness = 2
  conFltCrease = 3
  conFltDarkness = 4
  conFltDiffuse = 5
  conFltEmboss = 6
  conFltGrayBlacknWhite = 7
  conFltGrayscale = 8
  conFltInvertColors = 9
  conFltReplaceColors = 10
  conFltSharpen = 11
  conFltSnow = 12
  conFltWave = 13
End Enum

Public lngReplacedColor As Long
Public lngReplaceWithColor As Long

Public Sub ApplyFilter(intFilter As enmFilter, ByRef pic As PictureBox, _
                       Optional X1 As Long = -1, Optional Y1 As Long = -1, _
                       Optional X2 As Long = -1, Optional Y2 As Long = -1)
  Dim blnSmallArea As Boolean
                                        
  Dim intDrawMode As Integer
  Dim lngColor() As Long
                               
                               
  Dim lngReadColor As Long
  Dim lngTransColor As Long
  Dim lngWriteColor As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
  Dim sngFilterFactor As Single
  Dim x As Long
  Dim y As Long
  
  On Error GoTo ErrorHandler
  
  If (X1 = -1) And (Y1 = -1) And (X2 = -1) And (Y2 = -1) Then
    X1 = 0
    Y1 = 0
    X2 = pic.ScaleWidth
    Y2 = pic.ScaleHeight
  End If
  blnSmallArea = (((X2 - X1) * (Y2 - Y1)) < (16 * 16))
  With pic
    intDrawMode = .DrawMode
    .DrawMode = vbCopyPen
    Select Case intFilter
      Case conFltBlacknWhite
        sngFilterFactor = 192
                                   
                                   
                                   
        For x = X1 To X2
          For y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hDC:=.hDC, x:=x, y:=y)
            R = lngReadColor Mod 256
            If (R >= sngFilterFactor) Then
              lngWriteColor = vbWhite
            Else
              G = (lngReadColor \ 256) Mod 256
              If (G >= sngFilterFactor) Then
                lngWriteColor = vbWhite
              Else
                B = (lngReadColor \ 256) \ 256
                If (B >= sngFilterFactor) Then
                  lngWriteColor = vbWhite
                Else
                  lngWriteColor = vbBlack
                End If
              End If
            End If
            mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=((x * 100) \ X2)
          End If
        Next
      Case conFltBlur
        sngFilterFactor = 10
                                     
                                     
                                     
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, _
                                 blnShowProgress:=(Not blnSmallArea)
        For x = X1 + 1 To X2 - 1
          For y = Y1 + 1 To Y2 - 1
            R = lngColor(0, x - 1, y - 1) + lngColor(0, x, y - 1) + _
                lngColor(0, x + 1, y - 1) + lngColor(0, x - 1, y) + _
                lngColor(0, x, y) + lngColor(0, x + 1, y) + _
                lngColor(0, x - 1, y + 1) + lngColor(0, x, y + 1) + _
                lngColor(0, x + 1, y + 1)
            G = lngColor(1, x - 1, y - 1) + lngColor(1, x, y - 1) + _
                lngColor(1, x + 1, y - 1) + lngColor(1, x - 1, y) + _
                lngColor(1, x, y) + lngColor(1, x + 1, y) + _
                lngColor(1, x - 1, y + 1) + lngColor(1, x, y + 1) + _
                lngColor(1, x + 1, y + 1)
            B = lngColor(2, x - 1, y - 1) + lngColor(2, x, y - 1) + _
                lngColor(2, x + 1, y - 1) + lngColor(2, x - 1, y) + _
                lngColor(2, x, y) + lngColor(2, x + 1, y) + _
                lngColor(2, x - 1, y + 1) + lngColor(2, x, y + 1) + _
                lngColor(2, x + 1, y + 1)
            lngWriteColor = RGB(Abs(R / sngFilterFactor), _
                                Abs(G / sngFilterFactor), _
                                Abs(B / sngFilterFactor))
            mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                    intPercentage:=(((x + 1) * 100) \ X2)
          End If
        Next
      Case conFltBrightness, conFltDarkness
        Select Case intFilter
          Case conFltBrightness
            If Not blnSmallArea Then
              sngFilterFactor = 32
                                     
                                     
                                     
            Else
              sngFilterFactor = 2
            End If
          Case conFltDarkness
            If Not blnSmallArea Then
              sngFilterFactor = -32
                                     
                                     
                                     
            Else
              sngFilterFactor = -2
            End If
        End Select
        For x = X1 To X2
          For y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hDC:=.hDC, x:=x, y:=y)
            GetRGBColor lngColor:=lngReadColor, R:=R, G:=G, B:=B
            lngWriteColor = RGB(Abs(R + sngFilterFactor), _
                                Abs(G + sngFilterFactor), _
                                Abs(B + sngFilterFactor))
            mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
            
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=((x * 100) \ X2)
          End If
        Next
      Case conFltCrease, conFltWave
        Select Case intFilter
          Case conFltCrease
            sngFilterFactor = 512
                                    
                                    
                                    
          Case conFltWave
            sngFilterFactor = 4
                                  
                                  
        End Select
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, blnAll:=True, _
                                 blnShowProgress:=(Not blnSmallArea)
        For x = X1 To X2
          For y = Y1 To Y2
            lngWriteColor = lngColor(3, x, y)
            mdlAPI.SetPixel hDC:=.hDC, x:=x, _
                            y:=(Sin(x) * sngFilterFactor) + (y), _
                            crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                    intPercentage:=((x * 100) \ X2)
          End If
        Next
      Case conFltDiffuse
        sngFilterFactor = 5
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, blnAll:=True, _
                                 blnShowProgress:=(Not blnSmallArea)
        For x = X1 + 2 To X2 - 3
          For y = Y1 + 2 To Y2 - 3
            lngReadColor = lngColor(3, x, y + Int((Rnd * sngFilterFactor) - 2))
            R = Abs(lngReadColor Mod 256)
            lngReadColor = lngColor(3, x + Int((Rnd * sngFilterFactor) - 2), y)
            G = Abs((lngReadColor \ 256) Mod 256)
            lngReadColor = lngColor(3, x + Int((Rnd * sngFilterFactor) - 2), _
                                       y + Int((Rnd * sngFilterFactor) - 2))
            B = Abs((lngReadColor \ 256) \ 256)
            lngWriteColor = RGB(Red:=R, Green:=G, Blue:=B)
            mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=(((x + 3) * 100) \ X2)
          End If
        Next
      Case conFltEmboss
        sngFilterFactor = -128
                                  
                                  
                                  
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, _
                                 blnShowProgress:=(Not blnSmallArea)
        For x = X1 To X2 - 1
          For y = Y1 To Y2 - 1
            R = Abs(lngColor(0, x, y) - lngColor(0, x + 1, y + 1) + _
                    sngFilterFactor)
            G = Abs(lngColor(1, x, y) - lngColor(1, x + 1, y + 1) + _
                    sngFilterFactor)
            B = Abs(lngColor(2, x, y) - lngColor(2, x + 1, y + 1) + _
                    sngFilterFactor)
            lngWriteColor = RGB(Red:=R, Green:=G, Blue:=B)
            mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=(((x + 1) * 100) \ X2)
          End If
        Next
      Case conFltGrayBlacknWhite
        sngFilterFactor = 3
                                    
                                    
                                    
        For x = X1 To X2
          For y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hDC:=.hDC, x:=x, y:=y)
            GetRGBColor lngColor:=lngReadColor, R:=R, G:=G, B:=B
            R = Abs(R * (G - B + G + R)) / 256
            G = Abs(R * (B - G + B + R)) / 256
            B = Abs(G * (B - G + B + R)) / 256
            lngReadColor = RGB(Red:=R, Green:=G, Blue:=B)
            GetRGBColor lngColor:=lngReadColor, R:=R, G:=G, B:=B
            lngReadColor = (R + G + B) / sngFilterFactor
            lngWriteColor = RGB(Red:=lngReadColor, _
                                Green:=lngReadColor, Blue:=lngReadColor)
            mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                    intPercentage:=((x * 100) \ X2)
          End If
        Next
      Case conFltGrayscale
        sngFilterFactor = 0.32
                               
                               
                               
        For x = X1 To X2
          For y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hDC:=.hDC, x:=x, y:=y)
            GetRGBColor lngColor:=lngReadColor, R:=R, G:=G, B:=B
            lngTransColor = Abs((R * sngFilterFactor) + _
                                (G * sngFilterFactor) + (B * sngFilterFactor))
            lngWriteColor = RGB(Red:=lngTransColor, _
                                Green:=lngTransColor, Blue:=lngTransColor)
            mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                            intPercentage:=((x * 100) \ X2)
          End If
        Next
      Case conFltReplaceColors
        For x = X1 To X2
          For y = Y1 To Y2
            lngReadColor = mdlAPI.GetPixel(hDC:=.hDC, x:=x, y:=y)
            If lngReadColor = lngReplacedColor Then
              lngWriteColor = lngReplaceWithColor
              mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
            End If
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=((x * 100) \ X2)
          End If
        Next
      Case conFltSharpen, conFltSnow
        Select Case intFilter
          Case conFltSharpen
            sngFilterFactor = 0.5
                                         
                                         
                                         
          Case conFltSnow
            sngFilterFactor = 24
                                         
                                         
                                         
        End Select
        RetrieveColorInformation pic:=pic, lngColor:=lngColor, _
                                 X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2, _
                                 blnShowProgress:=(Not blnSmallArea)
        For x = X1 + 1 To X2
          For y = Y1 + 1 To Y2
            R = lngColor(0, x, y) + _
                (sngFilterFactor * _
                 (lngColor(0, x, y) - lngColor(0, x - 1, y - 1)))
            G = lngColor(1, x, y) + _
                (sngFilterFactor * _
                 (lngColor(1, x, y) - lngColor(1, x - 1, y - 1)))
            B = lngColor(2, x, y) + _
                (sngFilterFactor * _
                 (lngColor(2, x, y) - lngColor(2, x - 1, y - 1)))
            lngWriteColor = RGB(Abs(R), Abs(G), Abs(B))
            mdlAPI.SetPixel hDC:=.hDC, x:=x, y:=y, crColor:=lngWriteColor
          Next
          If Not blnSmallArea Then
            pic.Refresh
            frmPaint.UpdateStatusBar intInfo:=conStFiltering, _
                                     intPercentage:=((x * 100) \ X2)
          End If
        Next
    End Select
    .DrawMode = intDrawMode
    .Refresh
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub GetRGBColor(lngColor As Long, ByRef R As Long, _
                        ByRef G As Long, ByRef B As Long)
  On Error GoTo ErrorHandler
  
  R = lngColor Mod 256
  G = (lngColor \ 256) Mod 256
  B = (lngColor \ 256) \ 256
  Exit Sub
  
ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub RetrieveColorInformation( _
              pic As PictureBox, ByRef lngColor() As Long, _
              Optional X1 As Long = -1, Optional Y1 As Long = -1, _
              Optional X2 As Long = -1, Optional Y2 As Long = -1, _
              Optional blnAll As Boolean = False, _
              Optional blnShowProgress = True _
            )
  Dim R As Long
  Dim G As Long
  Dim B As Long
  Dim x As Long
  Dim y As Long
  
  On Error GoTo ErrorHandler
  
  If (X1 = -1) Or (Y1 = -1) Or (X2 = -1) Or (Y2 = -1) Then
    X1 = 0
    Y1 = 0
    X2 = pic.ScaleWidth
    Y2 = pic.ScaleHeight
  End If
  If blnAll Then
    ReDim lngColor(3, X2, Y2)
  Else
    ReDim lngColor(2, X2, Y2)
  End If
  For x = X1 To X2
    For y = Y1 To Y2
      If blnAll Then
        lngColor(3, x, y) = mdlAPI.GetPixel(pic.hDC, x, y)
      Else
        GetRGBColor lngColor:=mdlAPI.GetPixel(pic.hDC, x, y), R:=R, G:=G, B:=B
        lngColor(0, x, y) = R
        lngColor(1, x, y) = G
        lngColor(2, x, y) = B
      End If
    Next
    If blnShowProgress Then
      frmPaint.UpdateStatusBar intInfo:=conStRetrieveingColor, _
                               intPercentage:=((x * 100) \ X2)
    End If
  Next
  Exit Sub
  
ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub
