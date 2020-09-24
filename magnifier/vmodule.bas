Attribute VB_Name = "Module1"
Option Explicit
Option Base 0

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hobj As Long) As Integer
Public Declare Function GetObjectA Lib "gdi32" (ByVal hobj As Long, ByVal buffsize As Integer, ByRef buff As bitmap) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdcd As Long, ByVal xd As Long, ByVal yd As Long, ByVal widthd As Long, ByVal heightd As Long, ByVal hdcs As Long, ByVal xs As Long, ByVal ys As Long, ByVal widths As Long, ByVal heights As Long, ByVal opr As Long) As Integer
Public Declare Function LoadImageA Lib "user32" (ByVal hInst As Long, ByVal pfilename As String, ByVal typeimg As Long, ByVal width As Long, ByVal height As Long, ByVal flag As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hdcd As Long, ByVal xd As Long, ByVal yd As Long, ByVal widthd As Long, ByVal heightd As Long, ByVal hdcs As Long, ByVal xs As Long, ByVal ys As Long, ByVal opr As Long) As Integer
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal width As Long, ByVal height As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Integer
Public Declare Function DeleteObject Lib "gdi32" (ByVal hobj As Long) As Integer
   
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal xs As Long, ByVal ys As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal xd As Long, ByVal yd As Long, ByVal color As Long) As Integer

Type bitmap
 btype As Long
 bwidth As Long
 bheight As Long
 bwidthbytes As Long
 bplanes As Integer
 bbitpixels As Integer
 bbits As Integer
End Type

Public Const IMAGE_BITMAP = &O0
Public Const LR_LOADFROMFILE = 16

Public picinfo As bitmap

Public Const pi = 3.14159265358979
Public Const normal = 0.8
Public Const shrink = 15

Public himg As Long
Public hpic As Long
Public err As Integer

Public hframe() As Long

Public Sub Magnifier(ByVal radius As Integer, ByRef hframe As Long)
Dim w As Integer, h As Integer
Dim xc As Integer, yc As Integer
Dim xi As Integer, yi As Integer, xm As Double, ym As Double
'----------------
Dim xsb As Integer, ysb As Integer, wb As Integer, hb As Integer
'----------------
Dim hipotenusa  As Double, side1 As Double, side2 As Double
Dim angle As Double, distance As Double
Dim side1s As Double, side2s As Double

 
 
 err = BitBlt(hframe, 0, 0, picinfo.bwidth, picinfo.bheight, himg, 0, 0, vbSrcCopy)
 
 w = picinfo.bwidth
 h = picinfo.bheight
 
 xc = Int(w / 2)
 yc = Int(h / 2)
 
 xsb = xc - radius
 ysb = yc - radius
 
 wb = 2 * radius
 hb = 2 * radius
 
 For xi = xsb To wb + xsb
  For yi = ysb To hb + ysb
    
    side1 = xi - xc
    side2 = yi - yc
    hipotenusa = Sqr((side1 * side1) + (side2 * side2))
    'MsgBox CStr(hipotenusa)
    
    If hipotenusa < radius And Not hipotenusa = 0 Then
     angle = (hipotenusa / radius) * 90
     'MsgBox CStr(angle)
     distance = Abs(1 - Cos(angle * pi / 180))
     'MsgBox CStr(distance)
     
     side1s = Abs(side1 / hipotenusa)
     side2s = Abs(side2 / hipotenusa)
     
     If xi <= xc And yi <= yc Then 'area 1
      xm = radius - (radius * distance * side1s * normal) + xsb
      ym = radius - (radius * distance * side2s * normal) + ysb
     ElseIf xi <= xc And yi >= yc Then 'area 3
      xm = radius - (radius * distance * side1s * normal) + xsb
      ym = radius - (radius * distance * side2s * normal) + (2 * (yc - ysb - (radius - (radius * distance * side2s * normal)))) + ysb
     ElseIf xi > xc And yi < yc Then 'area 2
      xm = radius - (radius * distance * side1s * normal) + (2 * (xc - xsb - (radius - (radius * distance * side1s * normal)))) + xsb
      ym = radius - (radius * distance * side2s * normal) + ysb
     ElseIf xi > xc And yi > yc Then 'area 4
      xm = radius - (radius * distance * side1s * normal) + (2 * (xc - xsb - (radius - (radius * distance * side1s * normal)))) + xsb
      ym = radius - (radius * distance * side2s * normal) + (2 * (yc - ysb - (radius - (radius * distance * side2s * normal)))) + ysb
     End If
         
     err = SetPixel(hframe, xi, yi, GetPixel(himg, Round(xm), Round(ym)))
    
    Else
     
     If hipotenusa = 0 Then
      xm = xi
      ym = yi
      err = SetPixel(hframe, xi, yi, GetPixel(himg, Round(xm), Round(ym)))
     Else
      err = SetPixel(hframe, xi, yi, GetPixel(himg, xi, yi))
     End If
    End If
    
  Next yi
Next xi
 
End Sub



