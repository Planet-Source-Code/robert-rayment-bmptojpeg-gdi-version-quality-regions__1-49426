VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   " Bmp to Jpeg using GDI+  by Robert Rayment"
   ClientHeight    =   5355
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9795
   DrawWidth       =   2
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   653
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowFra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "?"
      Height          =   240
      Left            =   9495
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   15
      Width           =   300
   End
   Begin VB.Frame fraInstructions 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Instructions"
      Height          =   2055
      Left            =   6825
      MousePointer    =   15  'Size All
      TabIndex        =   26
      Top             =   1245
      Width           =   3075
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1665
         Left            =   90
         ScaleHeight     =   107
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   188
         TabIndex        =   28
         Top             =   285
         Width           =   2880
      End
      Begin VB.CommandButton cmdCloseFra 
         Caption         =   "X"
         Height          =   195
         Left            =   2730
         TabIndex        =   27
         Top             =   120
         Width           =   225
      End
   End
   Begin VB.PictureBox picSelect 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   8445
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   25
      Top             =   5550
      Width           =   495
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   1500
      Picture         =   "Main.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Select oval "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   1230
      Picture         =   "Main.frx":0F14
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Select rounded rectangle "
      Top             =   15
      Width           =   270
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   9165
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   22
      Top             =   5550
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   660
      Picture         =   "Main.frx":0FE6
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Save jpeg "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   2040
      Picture         =   "Main.frx":1570
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Deselect "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   1755
      Picture         =   "Main.frx":1642
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Select lasso "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   960
      Picture         =   "Main.frx":1714
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Select rectangle "
      Top             =   15
      Width           =   270
   End
   Begin VB.HScrollBar HSQuality 
      Height          =   210
      Index           =   1
      LargeChange     =   10
      Left            =   3345
      Max             =   100
      Min             =   1
      TabIndex        =   15
      Top             =   270
      Value           =   10
      Width           =   900
   End
   Begin VB.HScrollBar HSQuality 
      Height          =   210
      Index           =   0
      LargeChange     =   10
      Left            =   7650
      Max             =   100
      Min             =   1
      TabIndex        =   11
      Top             =   270
      Value           =   10
      Width           =   900
   End
   Begin VB.CommandButton cmdSave_Show 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   300
      Picture         =   "Main.frx":17E6
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Show jpeg "
      Top             =   15
      Width           =   345
   End
   Begin VB.CommandButton cmdLoadPic 
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   15
      Picture         =   "Main.frx":1930
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Load picture "
      Top             =   15
      Width           =   270
   End
   Begin VB.HScrollBar HS 
      Height          =   210
      Index           =   1
      Left            =   5265
      TabIndex        =   7
      Top             =   4920
      Width           =   2490
   End
   Begin VB.HScrollBar HS 
      Height          =   210
      Index           =   0
      Left            =   525
      TabIndex        =   6
      Top             =   4980
      Width           =   2490
   End
   Begin VB.VScrollBar VS 
      Height          =   2130
      Index           =   1
      Left            =   5025
      TabIndex        =   5
      Top             =   780
      Width           =   210
   End
   Begin VB.VScrollBar VS 
      Height          =   2130
      Index           =   0
      Left            =   330
      TabIndex        =   4
      Top             =   765
      Width           =   195
   End
   Begin VB.PictureBox picC 
      Height          =   4095
      Index           =   1
      Left            =   5310
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   1
      Top             =   780
      Width           =   3810
      Begin VB.PictureBox PIC 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   1
         Left            =   0
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   3
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.PictureBox picC 
      Height          =   4155
      Index           =   0
      Left            =   540
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   0
      Top             =   765
      Width           =   4320
      Begin VB.PictureBox PIC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   795
         Index           =   0
         Left            =   15
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   2
         Top             =   15
         Width           =   930
         Begin VB.Shape SRR 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            Height          =   225
            Left            =   525
            Shape           =   4  'Rounded Rectangle
            Top             =   420
            Width           =   240
         End
         Begin VB.Shape SO 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            Height          =   195
            Left            =   150
            Shape           =   2  'Oval
            Top             =   405
            Width           =   195
         End
         Begin VB.Shape SR 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            Height          =   165
            Left            =   495
            Top             =   135
            Width           =   195
         End
         Begin VB.Line SL 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Index           =   0
            X1              =   21
            X2              =   10
            Y1              =   9
            Y2              =   16
         End
      End
   End
   Begin VB.Label LabQuality 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      Height          =   240
      Index           =   1
      Left            =   4260
      TabIndex        =   17
      Top             =   270
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selection Quality"
      Height          =   195
      Index           =   1
      Left            =   3360
      TabIndex        =   16
      Top             =   30
      Width           =   1425
   End
   Begin VB.Label LabInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "File length = 666666 WxH = 244 x 555"
      Height          =   195
      Index           =   1
      Left            =   5295
      TabIndex        =   14
      Top             =   510
      Width           =   2730
   End
   Begin VB.Label LabInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "File length = 666666 WxH = 244 x 555"
      Height          =   195
      Index           =   0
      Left            =   525
      TabIndex        =   13
      Top             =   495
      Width           =   2730
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quality"
      Height          =   195
      Index           =   0
      Left            =   7680
      TabIndex        =   12
      Top             =   60
      Width           =   615
   End
   Begin VB.Label LabQuality 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      Height          =   240
      Index           =   0
      Left            =   8565
      TabIndex        =   10
      Top             =   270
      Width           =   405
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuFileOPS 
         Caption         =   "&Load picture"
         Index           =   0
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "&Show jpeg"
         Index           =   1
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "Save &As jpeg"
         Index           =   3
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "&Exit"
         Index           =   5
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "&SELECT"
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Rectangle"
         Index           =   0
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "Ro&unded rectangle"
         Index           =   1
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Oval"
         Index           =   2
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Lasso"
         Index           =   3
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Deselect"
         Checked         =   -1  'True
         Index           =   5
      End
   End
   Begin VB.Menu mnuFileSpec 
      Caption         =   "FileSpec"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' BmpToJpeg  by Robert Rayment Oct 2003

' GDI+ version

' The GDI+ dll can be downloaded at...
' http://www.microsoft.com/downloads/release.asp?releaseid=32738

' For use of GDI+ see Carles P V, PSC CodeId=42376
' and  MrBoBo, PSC CodeId=42488

' NB jpeg does badly for Black & White (1 bpp) bmps
' and does best compression for True Color (24 bpp) bmps
' and best of all if there are large areas of same color.

' Resizing
Private Type CSizes
  xL As Single
  yT As Single
  xW As Single
  yH As Single
End Type
Dim SizeArr() As CSizes
Dim XMag As Single
Dim YMag As Single
Dim LabFontSize0

Dim A$
Dim i As Long
Dim ORGW As Long
Dim ORGH As Long


Dim PathSpec$, CurrentPath$, FileSpec$, FileSpecPath$
Dim JPGFileSpec$
Dim JPGSelectSpec$
Dim FileLength0 As Long
Dim FileLength1 As Long

Dim CommonDialog1 As OSDialog

Private Sub cmdLoadPic_Click()
Dim Title$, Filt$, InDir$

   ' LOAD STANDARD VB PICTURES
   
   MousePointer = vbDefault
   
   If CommonDialog1 Is Nothing Then Set CommonDialog1 = New OSDialog
   
   Title$ = "Load a picture file"
   Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
   InDir$ = CurrentPath$
   
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
   If Len(FileSpec$) = 0 Then
      Close
      Set CommonDialog1 = Nothing
      Exit Sub
   End If
   
   FileLength0 = FileLen(FileSpec$)
   
   FileSpecPath$ = Left$(FileSpec$, InStrRev(FileSpec$, "\"))
   
   CurrentPath$ = FileSpecPath$
   
   PIC(0).Picture = LoadPicture
   PIC(0).Picture = LoadPicture(FileSpec$)
   
   Set CommonDialog1 = Nothing
   
   FixScrollbars picC(0), PIC(0), HS(0), VS(0)
   
   picWidth = PIC(0).Width
   picHeight = PIC(0).Height
   
   LabInfo(0) = "In: File size =" & Str$(FileLength0) & " B  WxH = " & Str$(picWidth) & " x" & Str$(picHeight)
   
   cmdSave_Show.Enabled = True
   cmdSave.Enabled = True
   
   mnuFileOPS(1).Enabled = True
   mnuFileOPS(3).Enabled = True
   mnuSelect.Enabled = True
   optSelect(0).Enabled = True
   optSelect(1).Enabled = True
   optSelect(2).Enabled = True
   optSelect(3).Enabled = True
   optSelect(4).Enabled = True
   ' For masking
   picBack.Width = picWidth
   picBack.Height = picHeight
   
   mnuFileSpec.Caption = FileSpec$
   
   cmdSave_Show_Click
   
   ' Cancel any selection
   aSelect = False
   SR.Visible = False
   SRR.Visible = False
   SO.Visible = False
   If NumLassoLines > 1 Then ' Clear extra lasso lines SL(1)-SL(NumLassoLines-1)
      For i = 1 To NumLassoLines - 1
         Unload SL(i)
      Next i
      NumLassoLines = 1
   End If
   SL(0).Visible = False
   
   PIC(0).MousePointer = vbDefault
   optSelect(4).Value = True
   optSelect(4).Value = False
   DoEvents
End Sub


Private Sub cmdSave_Show_Click()
' J
Dim ar As Long

   Screen.MousePointer = vbHourglass
   
   JPGFileSpec$ = PathSpec$ & "~~Temp.jpg"
   
   SAVEJPEG JPGFileSpec$, Quality, PIC(0)
   
   PIC(1).Picture = LoadPicture
   PIC(1).Picture = LoadPicture(JPGFileSpec$)
   PIC(1).Refresh
   
   If Not aSelect And NumLassoLines > 1 Then
      For i = 1 To NumLassoLines - 1
         Unload SL(i)
      Next i
      NumLassoLines = 1
   End If
   
   If aSelect Then
      ' Save PIC(0) at SelectionQuality
      JPGSelectSpec$ = PathSpec$ & "~~Temp2.jpg"
      SAVEJPEG JPGSelectSpec$, SelectionQuality, PIC(0)
      With picSelect
         .Width = picWidth
         .Height = picHeight
      End With
      picSelect.Picture = LoadPicture
      picSelect.Picture = LoadPicture(JPGSelectSpec$)
      picSelect.Refresh
      Kill JPGSelectSpec$
      ' WE have PIC(1) @ Quality
      '      picSelect @ SelectionQuality
      ' Mask picBack done @ PIC(0)_MouseUp [B(W)]
      ' Do
      '    Invert mask [W(B)}
      '    AND PIC(1) with mask gives PIC(1) @ Quality gives
      '    picture with black hole.
      '    Invert mask [B(W)}
      '    AND picSelect with mask gives picSelect @ SelectionQuality
      '    picture surrounded by black hole.
      '    PIC(1) OR picSElect
   
      ar = SetRect(IR, 0, 0, picWidth, picHeight)
      ar = InvertRect(picBack.hdc, IR)
      picBack.Refresh
      
      BitBlt PIC(1).hdc, 0, 0, picWidth, picHeight, picBack.hdc, 0, 0, vbSrcAnd
      PIC(1).Refresh
      
      ar = SetRect(IR, 0, 0, picWidth, picHeight)
      ar = InvertRect(picBack.hdc, IR)
      picBack.Refresh
      
      BitBlt picSelect.hdc, 0, 0, picWidth, picHeight, picBack.hdc, 0, 0, vbSrcAnd
      picSelect.Refresh
      
      BitBlt PIC(1).hdc, 0, 0, picWidth, picHeight, picSelect.hdc, 0, 0, vbSrcPaint
      PIC(1).Refresh
      ' To get new file size @ SelectionQuality
      SAVEJPEG JPGFileSpec$, SelectionQuality, PIC(1)
   End If
   
   FixScrollbars picC(1), PIC(1), HS(1), VS(1)
   
   FileLength1 = FileLen(JPGFileSpec$)
   LabInfo(1) = "Jpeg: File size =" & Str$(FileLength1) & " B  WxH = " & Str$(picWidth) & " x" & Str$(picHeight)
   
   Kill JPGFileSpec$

   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
Dim Title$, Filt$, InDir$
   
   If CommonDialog1 Is Nothing Then Set CommonDialog1 = New OSDialog
   
   Title$ = "Save JPEG"
   Filt$ = "Pics jpg|*.jpg"
   InDir$ = CurrentPath$
   
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
   If Len(FileSpec$) = 0 Then
      Close
      Set CommonDialog1 = Nothing
      Exit Sub
   End If
   
   FixExtension FileSpec$, "jpg"
   FileSpecPath$ = Left$(FileSpec$, InStrRev(FileSpec$, "\"))
   CurrentPath$ = FileSpecPath$
   Set CommonDialog1 = Nothing
   
   Screen.MousePointer = vbHourglass
   
   If aSelect Or NumLassoLines > 1 Then
      SAVEJPEG FileSpec$, SelectionQuality, PIC(1)
   Else
      SAVEJPEG FileSpec$, Quality, PIC(0)
   End If
   
   Screen.MousePointer = vbDefault

End Sub

Private Sub mnuFileOPS_Click(Index As Integer)
   Select Case Index
   Case 0: cmdLoadPic_Click   ' Load picture
   Case 1: cmdSave_Show_Click ' Show jpeg
   Case 2   ' -
   Case 3: cmdSave_Click      ' Save jpeg
   Case 4   ' -
   Case 5: Unload Me          ' Exit
   End Select
End Sub

Private Sub SAVEJPEG(FSpec$, ByVal TheQuality As Long, APIC As PictureBox)
   ' Create DIB, get pointer & publics:-
   ' DIBPtr, m_hDC, m_hDIb, m_hBmpOld
   SETBMI
   ' Blit picture to DIB
   BitBlt m_hDC, 0, 0, picWidth, picHeight, APIC.hdc, 0, 0, vbSrcCopy
   
   Dim pvGDI As GDIPlusJPGConvertor
   
   Set pvGDI = New GDIPlusJPGConvertor
   
   pvGDI.SaveDIB picWidth, picHeight, DIBPtr, FSpec$, TheQuality
 
   Set pvGDI = Nothing
    
   SelectObject m_hDC, m_hBmpOld
   DeleteObject m_hDIb
   DeleteDC m_hDC
End Sub

'#### SELECTING #############################################################

Private Sub optSelect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
   Case 0 To 3   ' Select rectangle, rounded rectangle, Oval or lasso
      mnuSelectOPS_Click Index
   Case 4   ' Deselect
      optSelect(4).Value = False
      mnuSelectOPS_Click 5
   End Select
   
End Sub

Private Sub mnuSelectOPS_Click(Index As Integer)
   mnuSelectOPS(0).Checked = False
   mnuSelectOPS(1).Checked = False
   mnuSelectOPS(2).Checked = False
   mnuSelectOPS(3).Checked = False
   mnuSelectOPS(5).Checked = False
   
   SR.Visible = False
   SRR.Visible = False
   SO.Visible = False
   If NumLassoLines > 1 Then ' Clear extra lasso lines SL(1)-SL(NumLassoLines-1)
      For i = 1 To NumLassoLines - 1
         Unload SL(i)
      Next i
      NumLassoLines = 1
   End If
   SL(0).Visible = False
   
   Select Case Index
   Case 0 To 2   ' Rectangle, Rounded rectangle, Oval
      aSelect = True
      aSelectDone = False
      SelectType = Index
      PIC(0).MousePointer = 2    ' Cross
   Case 3   ' Lasso
      aSelect = True
      aSelectDone = False
      SelectType = 3
      PIC(0).MousePointer = 10   ' Up arrow
   Case 4   ' -
   Case 5   ' Deselect
      aSelect = False
      PIC(0).MousePointer = vbDefault
      ' Reset PIC(1) from PIC(0)
      Screen.MousePointer = vbHourglass
      JPGFileSpec$ = PathSpec$ & "~~Temp.jpg"
   
      SAVEJPEG JPGFileSpec$, Quality, PIC(0)
   
      PIC(1).Picture = LoadPicture
      PIC(1).Picture = LoadPicture(JPGFileSpec$)
      PIC(1).Refresh
      
      Kill JPGFileSpec$
      Screen.MousePointer = vbDefault
   End Select
   
   mnuSelectOPS(Index).Checked = True
   If Index < 5 Then optSelect(Index).Value = True
   PIC(1).SetFocus
End Sub

Private Sub PIC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aSelect Then
      aSelectDone = False
      Select Case SelectType
      Case 0: Start_Select SR, X, Y   ' Rectangle
      Case 1: Start_Select SRR, X, Y  ' Rounded rectangle
      Case 2: Start_Select SO, X, Y   ' Oval
      Case 3: Start_Lasso X, Y        ' Lasso
      End Select
   End If
End Sub

Private Sub Start_Select(SS As Shape, X As Single, Y As Single)
   With SS
      .Left = X
      .Top = Y
      .Width = 4
      .Height = 4
      .Visible = True
   End With
   XS1 = X: YS1 = Y
   XS2 = X + 4: YS2 = Y + 4
End Sub

Private Sub Start_Lasso(X As Single, Y As Single)
   If NumLassoLines > 1 Then ' Clear extra lasso lines SL(1)-SL(NumLassoLines-1)
      For i = 1 To NumLassoLines - 1
         Unload SL(i)
      Next i
      NumLassoLines = 1
   End If
   If X < 3 Then X = 3
   If Y < 3 Then Y = 3
   If X > picWidth - 3 Then X = picWidth - 3
   If Y > picHeight - 3 Then Y = picHeight - 3
   XS1 = X
   YS1 = Y
   XS2 = X
   YS2 = Y
   With SL(0)
      .X1 = XS1
      .Y1 = YS1
      .X2 = XS2
      .Y2 = YS2
   End With
   SL(0).Visible = True
End Sub

Private Sub PIC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aSelect Then
      If Not aSelectDone Then
         If Button = vbLeftButton Then
            Select Case SelectType
            Case 0: Draw_Select SR, X, Y   ' Rectangle
            Case 1: Draw_Select SRR, X, Y  ' Rounded rectangle
            Case 2: Draw_Select SO, X, Y   ' Oval
            Case 3: Draw_Lasso X, Y       ' Lasso
            End Select
         End If
      End If
   End If
End Sub

Private Sub Draw_Select(SS As Shape, X As Single, Y As Single)
Dim SW As Long
Dim SH As Long
   If X < 1 Then X = 1
   If Y < 1 Then Y = 1
   If X > picWidth - 2 Then X = picWidth - 2
   If Y > picHeight - 2 Then Y = picHeight - 2
   
   SW = X - XS1
   SH = Y - YS1
   If SW < 0 Then
      XS1 = X
      SW = -SW
      XS2 = X + SW
   Else
      XS2 = X
   End If
   If SH < 0 Then
      YS1 = Y
      SH = -SH
      YS2 = Y + SH
   Else
      YS2 = Y
   End If
   With SS
      .Left = XS1
      .Top = YS1
      .Width = SW
      .Height = SH
   End With
End Sub

Private Sub Draw_Lasso(X As Single, Y As Single)
   NumLassoLines = NumLassoLines + 1
   Load SL(NumLassoLines - 1)
   If X < 3 Then X = 3
   If Y < 3 Then Y = 3
   If X > picWidth - 3 Then X = picWidth - 3
   If Y > picHeight - 3 Then Y = picHeight - 3
   XS1 = XS2
   YS1 = YS2
   XS2 = X
   YS2 = Y
   With SL(NumLassoLines - 1)
      .X1 = XS1
      .Y1 = YS1
      .X2 = XS2
      .Y2 = YS2
      .Visible = True
   End With
End Sub

Private Sub PIC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim zrad As Single
Dim zaspect As Single
Dim ar As Long
   aSelectDone = True
   
   If aSelect Then
      
      With picBack
         .BackColor = 0
         .Cls
         .DrawWidth = 2
         .FillColor = vbWhite
         .FillStyle = vbFSSolid
      End With
         
      Select Case SelectType
      Case 0: ' Rectangle
         picBack.Line (XS1, YS1)-(XS2, YS2), vbWhite, B
      
      Case 1: ' Rounded rectangle
         ' Corner circ radius = 1/6 of shortest dimension
         If SRR.Width > SRR.Height Then
            zrad = SRR.Height / 6
         Else
            zrad = SRR.Width / 6
         End If
         picBack.Line (XS1 + zrad, YS1)-(XS2 - zrad, YS1), vbWhite ' Top
         picBack.Line (XS1 + zrad, YS2)-(XS2 - zrad, YS2), vbWhite ' Bottom
         picBack.Line (XS1, YS1 + zrad)-(XS1, YS2 - zrad), vbWhite ' Left
         picBack.Line (XS2, YS1 + zrad)-(XS2, YS2 - zrad), vbWhite ' Right
         picBack.Circle (XS1 + zrad, YS1 + zrad), zrad, vbWhite, pi# / 2, pi#     ' TL
         picBack.Circle (XS2 - zrad, YS1 + zrad), zrad, vbWhite, 0, pi# / 2       ' TR
         picBack.Circle (XS1 + zrad, YS2 - zrad), zrad, vbWhite, pi#, 3 * pi# / 2 ' BL
         picBack.Circle (XS2 - zrad, YS2 - zrad), zrad, vbWhite, 3 * pi# / 2, 0   ' BR
         
         Fill picBack, XS2 - zrad, YS2 - zrad
         
      Case 2: ' Oval
         ' zaspect ' <1 horz, >1 vert
         zaspect = SO.Height / SO.Width
         If zaspect >= 1 Then
            zrad = Abs(((XS1 + XS2) / 2 - XS1)) * zaspect
         Else
            If zaspect = 0 Then zaspect = 4
            zrad = Abs(((YS1 + YS2) / 2 - YS1)) / zaspect
         End If
         If zrad = 0 Then zrad = 1
         picBack.Circle ((XS1 + XS2) / 2, (YS1 + YS2) / 2), zrad, vbWhite, , , zaspect
      Case 3  ' Lasso
         If X < 1 Then X = 1
         If Y < 1 Then Y = 1
         If X > picWidth - 2 Then X = picWidth - 2
         If Y > picHeight - 2 Then Y = picHeight - 2
         ' Close shape
         NumLassoLines = NumLassoLines + 1
         Load SL(NumLassoLines - 1)
         With SL(NumLassoLines - 1)
            .X1 = X
            .Y1 = Y
            .X2 = SL(0).X1
            .Y2 = SL(0).Y1
            .Visible = True
         End With
         ' Transfer to picBack
         For i = 0 To NumLassoLines - 1
            With SL(i)
               picBack.Line (.X1, .Y1)-(.X2, .Y2), vbWhite
            End With
         Next i
         Fill picBack, 1, 1   ' [W(B)]
         ' Invert
         ar = SetRect(IR, 0, 0, picWidth, picHeight)
         ar = InvertRect(picBack.hdc, IR)
         picBack.Refresh
      
      End Select
      
      ' Have mask White surrounded by Black [B(W)]
      
      With picBack
         .DrawWidth = 1
         .FillStyle = vbFSTransparent
      End With
   
      picBack.Refresh
      
   End If
End Sub
'#### END SELECTING #############################################################

'### FILL #####################################################

Private Sub Fill(APIC As PictureBox, X As Single, Y As Single)
   ' Fill with FillColor = DrawColor at X,Y
   APIC.DrawStyle = vbSolid
   APIC.FillColor = vbWhite
   APIC.FillStyle = vbFSSolid
   
   ' FLOODFILLSURFACE = 1
   ' Fills with FillColor so long as point surrounded by
   ' color = APIC.Point(X, Y)
   
   ExtFloodFill APIC.hdc, X, Y, APIC.Point(X, Y), FLOODFILLSURFACE
   
   APIC.Refresh
End Sub
'### END FILL #####################################################


Private Sub Form_Load()
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrentPath$ = PathSpec$
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY

   With picC(1)
      .Width = picC(0).Width
      .Height = picC(0).Height
      .Top = picC(0).Top
   End With
   With PIC(0)
      .Left = 0
      .Top = 0
   End With
   With PIC(1)
      .Left = 0
      .Top = 0
   End With
   HS(0).TabStop = False
   HS(1).TabStop = False
   VS(0).TabStop = False
   VS(1).TabStop = False
   FixScrollbars picC(0), PIC(0), HS(0), VS(0)
   FixScrollbars picC(1), PIC(1), HS(1), VS(1)
   
   cmdSave_Show.Enabled = False
   cmdSave.Enabled = False
   mnuFileOPS(1).Enabled = False
   mnuFileOPS(3).Enabled = False
   
   mnuSelectOPS(5).Checked = True
   mnuSelect.Enabled = False
   optSelect(0).Enabled = False
   optSelect(1).Enabled = False
   optSelect(2).Enabled = False
   optSelect(3).Enabled = False
   optSelect(4).Enabled = False
   
   aSelect = False
   SL(0).Visible = False
   SR.Visible = False
   SRR.Visible = False
   SO.Visible = False
   NumLassoLines = 1 ' SL(0)
   
   picBack.Visible = False
   picSelect.Visible = False
   
   Quality = 50
   HSQuality(0).Value = Quality
   SelectionQuality = 100
   HSQuality(1).Value = SelectionQuality
   
   LabFontSize0 = LabInfo(0).FontSize
   
   LabInfo(0) = "File length =   WxH = "
   LabInfo(1) = "File length =   WxH = "
   
   ' For resizing
   XMag = 1
   YMag = 1
   ReDim SizeArr(0 To Controls.count - 1)
   For i = 0 To Controls.count - 1
      If Left$(Controls(i).Name, 3) <> "mnu" Then
      If Left$(Controls(i).Name, 2) <> "SL" Then
         SizeArr(i).xL = Controls(i).Left
         SizeArr(i).yT = Controls(i).Top
         SizeArr(i).xW = Controls(i).Width
         SizeArr(i).yH = Controls(i).Height
      End If
      End If
   Next i

   ORGW = Me.Width
   ORGH = Me.Height
   
   AddInstructions
   fraInstructions.Visible = False
End Sub

Private Sub Form_Resize()
'ResizeControls()
On Error Resume Next
   
   ' Cancel any selection
   aSelect = False
   SR.Visible = False
   SRR.Visible = False
   SO.Visible = False
   For i = 0 To NumLassoLines - 1
      'Unload NOT PERMITTED IN RESIZE EVENT
      SL(i).Visible = False
   Next i
  
   If Me.Width >= ORGW And Me.Height >= ORGH Then
      XMag = CSng(Me.Width / ORGW)
      YMag = CSng(Me.Height / ORGH)
   Else
      Me.Width = ORGW
      Me.Height = ORGH
      Me.Refresh
      Exit Sub
   End If
   
   
   For i = 0 To Controls.count - 1
      If TypeOf Controls(i) Is Label Then
         Controls(i).Move _
         SizeArr(i).xL * XMag, _
         SizeArr(i).yT * YMag
         Controls(i).FontSize = LabFontSize0 * (XMag + YMag) / 2
         Controls(i).Refresh
         If Controls(i).Name <> "LabInfo" Then
            Controls(i).Width = SizeArr(i).xW * XMag
            Controls(i).Height = SizeArr(i).yH * YMag
         End If
      ElseIf Controls(i).Name <> "PIC" Then
         If Controls(i).Name <> "picSelect" Then
         If Controls(i).Name <> "picBack" Then
         If Controls(i).Name <> "SL" Then
         If Left$(Controls(i).Name, 3) <> "mnu" Then
            Controls(i).Move _
            SizeArr(i).xL * XMag, _
            SizeArr(i).yT * YMag, _
            SizeArr(i).xW * XMag, _
            SizeArr(i).yH * YMag
            Controls(i).Refresh
         End If
         End If
         End If
         End If
      End If
      
   Next i
   
   FixScrollbars picC(0), PIC(0), HS(0), VS(0)
   FixScrollbars picC(1), PIC(1), HS(1), VS(1)
   
   PIC(0).MousePointer = vbDefault
   optSelect(4).Value = True
   optSelect(4).Value = False
   
End Sub

'#### QUALITY PIC(0/1) SCROLL BARS #########################################

Private Sub HSQuality_Change(Index As Integer)
   Select Case Index
   Case 0
      Quality = HSQuality(0).Value
      LabQuality(0) = Str$(Quality)
   Case 1
      SelectionQuality = HSQuality(1).Value
      LabQuality(1) = Str$(SelectionQuality)
   End Select
End Sub

Private Sub HS_Change(Index As Integer)
   PIC(Index).Left = -HS(Index).Value
   If Index = 0 Then
      PIC(1).Left = -HS(0).Value
      HS(1).Value = HS(0).Value
   Else
      PIC(0).Left = -HS(1).Value
      HS(0).Value = HS(1).Value
   End If
End Sub

Private Sub HS_Scroll(Index As Integer)
   PIC(Index).Left = -HS(Index).Value
   If Index = 0 Then
      PIC(1).Left = -HS(0).Value
      HS(1).Value = HS(0).Value
   Else
      PIC(0).Left = -HS(1).Value
      HS(0).Value = HS(1).Value
   End If
End Sub

Private Sub VS_Change(Index As Integer)
   PIC(Index).Top = -VS(Index).Value
   If Index = 0 Then
      PIC(1).Top = -VS(0).Value
      VS(1).Value = VS(0).Value
   Else
      PIC(0).Top = -VS(1).Value
      VS(0).Value = VS(1).Value
   End If
End Sub

Private Sub VS_Scroll(Index As Integer)
   PIC(Index).Top = -VS(Index).Value
   If Index = 0 Then
      PIC(1).Top = -VS(0).Value
      VS(1).Value = VS(0).Value
   Else
      PIC(0).Top = -VS(1).Value
      VS(0).Value = VS(1).Value
   End If
End Sub
'#### END QUALITY PIC(0/1) SCROLL BARS #########################################

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub AddInstructions()
A$ = " Load a picture. It will be shown at the" & vbCr
A$ = A$ & " Quality setting.  Change Quality then" & vbCr
A$ = A$ & " press J button to show result." & vbCr
A$ = A$ & " " & vbCr
A$ = A$ & " For selection press a selection button." & vbCr
A$ = A$ & " Draw on first picture then press J. The" & vbCr
A$ = A$ & " selected area will be shown at the" & vbCr
A$ = A$ & " Selection Quality setting." & vbCr
Picture1.Cls
Picture1.Print A$
Me.Caption = "- Bmp to Jpeg GDI+  by Robert Rayment -"
End Sub

Private Sub cmdShowFra_Click()
   fraInstructions.Visible = Not fraInstructions.Visible
End Sub
Private Sub cmdCloseFra_Click()
   fraInstructions.Visible = False
End Sub

Private Sub fraInstructions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   fraX = X
   fraY = Y
End Sub
Private Sub fraInstructions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   fraMOVER Form1, fraInstructions, Button, X, Y
End Sub

