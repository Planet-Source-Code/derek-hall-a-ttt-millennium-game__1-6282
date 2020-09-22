Attribute VB_Name = "Globals"
Option Explicit
' Graphics functions and constants used in the example.
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const BBMASK = &H8800C6    ' Masks
Public Const BBPAINT = &HEE0086  ' onto masks
Public Const BBCOPY = &HCC0020   ' backgrounds

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long



