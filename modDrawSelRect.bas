Attribute VB_Name = "modDrawSelRect"


'scalemode = 3 on form and pictureboxes
'autoredraw = true for the pictureboxes

Option Explicit

'keeps cursor confined to picturebox
Private Type RECT
   Left As Integer
   Top As Integer
   Right As Integer
   Bottom As Integer
End Type

Private Type POINT
   xx As Long
   yy As Long
End Type

Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
'end cursor confined

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "GDI32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Function DrawSR(pic As Object)
 'confine cursor ...start
   Dim rt As RECT
   Dim pt As POINT
   
   'be sure to set picturebox on form, autodraw = true and scalemode = 3
   
   GetClientRect pic.hWnd, rt
   pt.xx = rt.Left
   pt.yy = rt.Top
   ClientToScreen pic.hWnd, pt
   OffsetRect rt, pt.xx, pt.yy
   ClipCursor rt
   'confine cursor ...end
End Function

Public Function ReleaseSR()
   ClipCursor ByVal 0&     'release cursor
End Function
