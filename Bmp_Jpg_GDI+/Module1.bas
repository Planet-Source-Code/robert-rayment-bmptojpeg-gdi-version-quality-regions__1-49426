Attribute VB_Name = "Module1"
' Module1 (Module1.bas)
Option Explicit

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
''--------------------------------------------------------------------------------
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' To invert picture box
Public Declare Function SetRect Lib "user32" (lpRect As RECT, _
ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function InvertRect Lib "user32" _
(ByVal hdc As Long, lpRect As RECT) As Long

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public IR As RECT
'ar = SetRect(IR, 0, 0, picWidth - 1, picHeight - 1)
'ar = InvertRect(PIC(1).hdc, IR)
'------------------------------------------------------------------------------

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1

Public PadBytes As Long
Public BytesPerScanLine As Long

Public m_hDIb As Long, m_hBmpOld As Long
Public m_hDC As Long, DIBPtr As Long

Public picWidth As Long, picHeight As Long

Public Quality As Long
Public SelectionQuality As Long
Public aSelect As Boolean
Public aSelectDone As Boolean
Public SelectType As Long

' Rectangle selection coords
Public XS1 As Single
Public YS1 As Single
Public XS2 As Single
Public YS2 As Single

Public NumLassoLines As Long

Public fraX As Single
Public fraY As Single

Public STX As Long
Public STY As Long
Public Const pi# = 3.1459265


Public Sub SETBMI()
Dim SBI As BITMAPINFOHEADER
   With SBI
      .biSize = 40
      .biWidth = picWidth
      .biHeight = picHeight
      .biPlanes = 1
      .biBitCount = 32 '24
      .biCompression = 0
   
      BytesPerScanLine = (((.biWidth * .biBitCount) + 31) \ 32) * 4
      PadBytes = BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
      .biSizeImage = BytesPerScanLine * Abs(.biHeight)
      
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   
   m_hDC = CreateCompatibleDC(0)
   m_hDIb = CreateDIBSection(m_hDC, SBI, 0, DIBPtr, 0, 0)
   m_hBmpOld = SelectObject(m_hDC, m_hDIb)
End Sub

Public Sub FixScrollbars(picC As PictureBox, picP As PictureBox, HS As HScrollBar, VS As VScrollBar)
   ' picC = Container = picFrame
   ' picP = Picture   = picDisplay
      HS.Max = picP.Width - picC.Width + 12   ' +4 to allow for border
      VS.Max = picP.Height - picC.Height + 12 ' +4 to allow for border
      HS.LargeChange = picC.Width \ 10
      HS.SmallChange = 1
      VS.LargeChange = picC.Height \ 10
      VS.SmallChange = 1
      HS.Top = picC.Top + picC.Height + 1
      HS.Left = picC.Left
      HS.Width = picC.Width
      If picP.Width < picC.Width Then
         HS.Visible = False
         'HS.Enabled = False
      Else
         HS.Visible = True
         'HS.Enabled = True
      End If
      VS.Top = picC.Top
      VS.Left = picC.Left - VS.Width - 1
      VS.Height = picC.Height
      If picP.Height < picC.Height Then
         VS.Visible = False
         'VS.Enabled = False
      Else
         VS.Visible = True
         'VS.Enabled = True
      End If
End Sub

Public Sub FixExtension(FSpec$, Ext$)
' Enter Ext$ as jpg, bmp etc  no dot
Dim p As Long

If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & "." & Ext$
   Else
      If LCase$(Mid$(FSpec$, p + 1)) <> Ext$ Then FSpec$ = Mid$(FSpec$, 1, p) & Ext$
   End If

End Sub


'### GENERAL FRAME MOVER #####################################
Public Sub fraMOVER(frm As Form, fra As Frame, Button As Integer, X As Single, Y As Single)
'Public fraX As Single
'Public fraY As Single
Dim fraLeft As Long
Dim fraTop As Long

   If Button = vbLeftButton Then
      
      fraLeft = fra.Left + (X - fraX) \ STX
      If fraLeft < 0 Then fraLeft = 0
      If fraLeft + fra.Width > frm.Width \ STX + fra.Width \ 2 Then
         fraLeft = frm.Width \ STX - fra.Width \ 2
      End If
      fra.Left = fraLeft
      
      fraTop = fra.Top + (Y - fraY) \ STY
      If fraTop < 8 Then fraTop = 8
      If fraTop + fra.Height > frm.Height \ STY + fra.Height \ 2 Then
         fraTop = frm.Height \ STY - fra.Height \ 2
      End If
      fra.Top = fraTop
      
   End If
End Sub
'### END GENERAL FRAME MOVER #####################################


