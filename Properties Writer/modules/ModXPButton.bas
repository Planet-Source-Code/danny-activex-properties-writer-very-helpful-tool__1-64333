Attribute VB_Name = "ModXPButton"
Option Explicit

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum COLOR_STYLE
    [XP Blue] = 1
    [XP Silver] = 2
    [XP Olive Green] = 3
End Enum

Public Enum PICTURE_ALIGN
    [Left Justify] = 1
    [Right Justify] = 2
End Enum

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

'Invert a color; get the opposite color for another color (i.e: white = black).
Public Function InvertColor(ByVal RValue As Integer, ByVal GValue As Integer, ByVal BValue As Integer) As Long
Dim intR As Integer, intG As Integer, intB As Integer

intR = Abs(255 - RValue)
intG = Abs(255 - GValue)
intB = Abs(255 - BValue)

InvertColor = RGB(intR, intG, intB)
End Function

'Convert a long color value to an RGB value.
Public Sub LongToRGB(ByRef RValue As Integer, ByRef GValue As Integer, ByRef BValue As Integer, ByVal ColorValue As Long)
Dim intR As Integer, intG As Integer, intB As Integer

intR = ColorValue Mod 256
intG = ((ColorValue And &HFF00) / 256&) Mod 256&
intB = (ColorValue And &HFF0000) / 65536

RValue = intR
GValue = intG
BValue = intB
End Sub

'Lightens a color judging by the offset value.
Public Function LightenColor(ByVal RValue As Integer, ByVal GValue As Integer, ByVal BValue As Integer, Optional ByVal OffSet As Long = 1) As Long
Dim intR As Integer, intG As Integer, intB As Integer

intR = Abs(RValue + OffSet)
intG = Abs(GValue + OffSet)
intB = Abs(BValue + OffSet)

LightenColor = RGB(intR, intG, intB)
End Function

'Darkens a color judging by the offset value.
Public Function DarkenColor(ByVal RValue As Integer, ByVal GValue As Integer, ByVal BValue As Integer, Optional ByVal OffSet As Long = 1) As Long
Dim intR As Integer, intG As Integer, intB As Integer

intR = Abs(RValue - OffSet)
intG = Abs(GValue - OffSet)
intB = Abs(BValue - OffSet)

DarkenColor = RGB(intR, intG, intB)
End Function

'Replace one color with another color.
Public Sub ReplaceColor(PictureObject As PictureBox, ColorValue As Long, ReplaceWith As Long)
Dim lonSW As Long, lonSH As Long
Dim lonLoopW As Long, lonLoopH As Long

PictureObject.ScaleMode = vbPixels
lonSW = PictureObject.ScaleWidth
lonSH = PictureObject.ScaleHeight

For lonLoopW = 1 To lonSW
    
    For lonLoopH = 1 To lonSH
        
        If PictureObject.Point(lonLoopW, lonLoopH) = ColorValue Then
            PictureObject.PSet (lonLoopW, lonLoopH), ReplaceWith
        End If
    
    Next lonLoopH

Next lonLoopW
End Sub
