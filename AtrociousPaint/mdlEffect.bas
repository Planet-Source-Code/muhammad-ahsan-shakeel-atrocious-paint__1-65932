Attribute VB_Name = "mdlEffect"

Option Explicit

Public Enum enmEffect
  conEffFlipHorizontal = 0
  conEffFlipVertical = 1
  conEffResize = 2
  conEffRotate = 3
  conEffInvertColors = 4
End Enum

Public Const conZoomFactor = 1.25
Public Const conMaxImageWidth = 50000
Public Const conMaxImageHeight = 50000

Public sngResizeWidth As Single
Public sngResizeHeight As Single

Public blnRotateClockWise As Boolean
Public sngRotateAngle As Single

Public Sub ApplyEffect(intEffect As enmEffect, _
                       ByRef pic As PictureBox, picTemp As PictureBox)
  Dim blnAutoSize As Boolean
  
  On Error GoTo ErrorHandler
  
  With pic
    blnAutoSize = picTemp.AutoSize
    picTemp.AutoSize = True
    picTemp.Width = .Width
    picTemp.Height = .Height
    picTemp.Picture = .Image
    .Picture = Nothing
    Select Case intEffect
      Case conEffFlipHorizontal
        .PaintPicture picTemp.Image, .ScaleWidth, 0, _
                      -.ScaleWidth, .ScaleHeight, , , , , vbSrcCopy
      Case conEffFlipVertical
        .PaintPicture picTemp.Image, 0, .ScaleHeight, _
                      .ScaleWidth, -.ScaleHeight, , , , , vbSrcCopy
      Case conEffInvertColors
        .PaintPicture picTemp.Image, 0, 0, _
                      .ScaleWidth, .ScaleHeight, , , , , vbSrcInvert
      Case conEffResize
        frmPaint.DrawSelectionRect
        .Visible = False
        .Width = .Width * sngResizeWidth
        .Height = .Height * sngResizeHeight
        .PaintPicture picTemp.Image, 0, 0, _
                      .ScaleWidth, .ScaleHeight, , , , , vbSrcCopy
        .Visible = True
        frmPaint.DrawSelectionRect
        frmPaint.AdjustPaintResizeBox
        frmPaint.Form_Resize
      Case conEffRotate
        If sngRotateAngle = 180 Then
          .PaintPicture picTemp.Image, .ScaleWidth, .ScaleHeight, _
                        -.ScaleWidth, -.ScaleHeight, , , , , vbSrcCopy
        Else
          ImageRotate picSource:=picTemp, picDestination:=pic, _
                      sngRotateAngle:=sngRotateAngle, _
                      blnClockWise:=blnRotateClockWise
        End If
    End Select
    picTemp.AutoSize = blnAutoSize
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub ImageRotate(picSource As PictureBox, _
                        picDestination As PictureBox, _
                        sngRotateAngle As Single, blnClockWise As Boolean)
  Const conPi = 3.14159265358979
  
  Dim A As Single
  Dim intMaxXY As Single
  Dim dXs As Long
  Dim dYs As Long
  Dim dXd As Long
  Dim dYd As Long
  Dim lngAdjustX As Long
  Dim lngAdjustY As Long
                         
                         
  Dim lngColor(3) As Long
  Dim R As Integer
  Dim Xs As Integer
  Dim Ys As Integer
  Dim Xd As Integer
  Dim Yd As Integer
                              
  If blnClockWise Then
    sngRotateAngle = 360 - sngRotateAngle
  End If
  Xs = picSource.ScaleWidth / 2
  Ys = picSource.ScaleHeight / 2
  Xd = picDestination.ScaleWidth / 2
  Yd = picDestination.ScaleHeight / 2
  intMaxXY = varIIf(picDestination.ScaleWidth > picDestination.ScaleHeight, _
                    picDestination.ScaleWidth / 2, _
                    picDestination.ScaleHeight / 2)
  If (sngRotateAngle = 90) Or (sngRotateAngle = 270) Then
    lngAdjustX = ((picDestination.ScaleHeight - _
                   picDestination.ScaleWidth) / 2) - 2
    lngAdjustY = ((picDestination.ScaleWidth - _
                   picDestination.ScaleHeight) / 2)
    frmPaint.DrawSelectionRect
    picDestination.Tag = CStr(picDestination.Width)
    picDestination.Width = picDestination.Height
    picDestination.Height = CLng(picDestination.Tag)
    With frmPaint
      .DrawSelectionRect
      .AdjustPaintResizeBox
      .Form_Resize
      .Refresh
    End With
  Else
    lngAdjustX = 0
    lngAdjustY = 0
  End If
  sngRotateAngle = sngRotateAngle * (conPi / 180)
  
  picDestination.DrawMode = vbCopyPen
  For dXd = 0 To intMaxXY
    For dYd = 0 To intMaxXY
      If dXd = 0 Then
        A = conPi / 2
      Else
        A = Atn(dYd / dXd)
      End If
      R = Sqr((dXd * dXd) + (dYd * dYd))
      dXs = R * Cos(A + sngRotateAngle)
      dYs = R * Sin(A + sngRotateAngle)

      lngColor(0) = GetPixel(picSource.hDC, Xs + dXs, Ys + dYs)
      lngColor(1) = GetPixel(picSource.hDC, Xs - dXs, Ys - dYs)
      lngColor(2) = GetPixel(picSource.hDC, Xs + dYs, Ys - dXs)
      lngColor(3) = GetPixel(picSource.hDC, Xs - dYs, Ys + dXs)

      If lngColor(0) <> -1 Then
        SetPixel picDestination.hDC, Xd + dXd + lngAdjustX, _
                 Yd + dYd + lngAdjustY, lngColor(0)
      End If
      If lngColor(1) <> -1 Then
        SetPixel picDestination.hDC, Xd - dXd + lngAdjustX, _
                 Yd - dYd + lngAdjustY, lngColor(1)
      End If
      If lngColor(2) <> -1 Then
        SetPixel picDestination.hDC, Xd + dYd + lngAdjustX, _
                 Yd - dXd + lngAdjustY, lngColor(2)
      End If
      If lngColor(3) <> -1 Then
        SetPixel picDestination.hDC, Xd - dYd + lngAdjustX, _
                 Yd + dXd + lngAdjustY, lngColor(3)
      End If
    Next
    picDestination.Refresh
  Next
  picDestination.Refresh
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub


