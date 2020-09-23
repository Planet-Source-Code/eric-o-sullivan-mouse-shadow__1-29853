VERSION 5.00
Begin VB.Form frmMouseShadow 
   Caption         =   "Mouse Shadow"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMouseShadow.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   Picture         =   "frmMouseShadow.frx":0442
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timMouse 
      Interval        =   1
      Left            =   960
      Top             =   0
   End
End
Attribute VB_Name = "frmMouseShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ShadowHeightX = 6
Const ShadowHeightY = 4

Dim CursorBmp As BitmapStruc
Dim BackBmp As BitmapStruc

Private Sub Form_Load()
Dim hCursor As Long
Dim Result As Long

CursorBmp.Area.Bottom = 32
CursorBmp.Area.Right = 32
BackBmp.Area = CursorBmp.Area

'create the cursor and the background bitmaps
Call CreateNewBitmap(CursorBmp.hDcMemory, CursorBmp.hDcBitmap, CursorBmp.hDcPointer, CursorBmp.Area, frmMouseShadow, vbWhite, InPixels)
Call CreateNewBitmap(BackBmp.hDcMemory, BackBmp.hDcBitmap, BackBmp.hDcPointer, BackBmp.Area, frmMouseShadow, vbWhite, InPixels)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'- not used any more 16/12/2001
Call DrawFrame(X, Y)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'remove bitmaps from memory before exiting program
Call DeleteBitmap(CursorBmp.hDcMemory, CursorBmp.hDcBitmap, CursorBmp.hDcPointer)
Call DeleteBitmap(BackBmp.hDcMemory, BackBmp.hDcBitmap, BackBmp.hDcPointer)
End
End Sub

Private Sub timMouse_Timer()
Dim MousePos As PointAPI
Dim MouseRect As Rect
Dim TempRect As Rect
Dim FormRect As Rect
Dim SmallFormRect As Rect

'find the forms position and dimensions
FormRect.Top = (frmMouseShadow.Top / Screen.TwipsPerPixelY) + ((frmMouseShadow.Height / Screen.TwipsPerPixelY) - frmMouseShadow.ScaleHeight) - (((frmMouseShadow.Width / Screen.TwipsPerPixelX) - frmMouseShadow.ScaleWidth) / 2)
FormRect.Left = (frmMouseShadow.Left / Screen.TwipsPerPixelX) + (((frmMouseShadow.Width / Screen.TwipsPerPixelX) - frmMouseShadow.ScaleWidth) / 2)
FormRect.Bottom = (FormRect.Top + frmMouseShadow.ScaleHeight) '/ Screen.TwipsPerPixelY
FormRect.Right = (FormRect.Left + frmMouseShadow.ScaleWidth) '/ Screen.TwipsPerPixelX

'shrink the forms' dimensions by 32 pixles from outside, in.
SmallFormRect = FormRect
SmallFormRect.Top = SmallFormRect.Top + 32
SmallFormRect.Left = SmallFormRect.Left + 32
SmallFormRect.Bottom = SmallFormRect.Bottom - 32
SmallFormRect.Right = SmallFormRect.Right - 32

'find the mouses position and demensions
Call GetCursorPos(MousePos)
MouseRect.Top = MousePos.Y
MouseRect.Left = MousePos.X
MouseRect.Bottom = MousePos.Y + 32
MouseRect.Right = MousePos.X + 32

'if these two regions intersect each other then draw the frame
If RectIntersects(MouseRect, FormRect) And (Not RectIntersects(MouseRect, SmallFormRect)) Then 'IntersectRect(TempRect, MouseRect, FormRect)
    Call DrawFrame(MousePos.X - FormRect.Left, MousePos.Y - FormRect.Top)
End If
End Sub

Public Sub DrawFrame(ByVal X As Integer, ByVal Y As Integer)
'This will draw the mouse cursor onto the screen

Static LastX As Integer
Static LastY As Integer
Static Started As Boolean
Static LasthCursor As Long

Dim OffsetX As Integer
Dim OffsetY As Integer
Dim BackOffsetX As Integer
Dim BackOffsetY As Integer
Dim MouseOffsetX As Integer
Dim MouseOffsetY As Integer
Dim Result As Long
Dim TotalBackBmp As BitmapStruc

' Draw the cursor's mask to a picturebox
hCursor = GetCursor
If hCursor <> LasthCursor Then
    Result = DrawIconEx(CursorBmp.hDcMemory, 0, 0, hCursor, 0, 0, 0, 0, 1)
End If
LasthCursor = hCursor

'adjust for the shadow height
X = X + ShadowHeightX
Y = Y + ShadowHeightY

'calculate the difference in position
OffsetX = X - LastX
OffsetY = Y - LastY

If OffsetX > 0 Then
    BackOffsetX = 0
    MouseOffsetX = OffsetX
Else
    BackOffsetX = -OffsetX
    MouseOffsetX = 0
End If
If OffsetY > 0 Then
    BackOffsetY = 0
    MouseOffsetY = OffsetY
Else
    BackOffsetY = -OffsetY
    MouseOffsetY = 0
End If

'create the bitmap
'set the bitmap size
TotalBackBmp.Area.Top = 0
TotalBackBmp.Area.Left = 0
If Started Then
    TotalBackBmp.Area.Right = Abs(OffsetX) + 32
    TotalBackBmp.Area.Bottom = Abs(OffsetY) + 32
Else
    'only do this once
    TotalBackBmp.Area.Right = 32
    TotalBackBmp.Area.Bottom = 32
End If
Call CreateNewBitmap(TotalBackBmp.hDcMemory, TotalBackBmp.hDcBitmap, TotalBackBmp.hDcPointer, TotalBackBmp.Area, frmMouseShadow, 0, InPixels)

'capture the screen area
Result = BitBlt(TotalBackBmp.hDcMemory, 0, 0, TotalBackBmp.Area.Right, TotalBackBmp.Area.Bottom, frmMouseShadow.hDc, X - MouseOffsetX, Y - MouseOffsetY, SRCCOPY)

If Not Started Then
    'put the background into the background picture - only do this once
    Result = BitBlt(BackBmp.hDcMemory, 0, 0, 32, 32, frmMouseShadow.hDc, X, Y, SRCCOPY)
    Started = True
Else
    'copy the old background over where the mouse used to be
    Result = BitBlt(TotalBackBmp.hDcMemory, BackOffsetX, BackOffsetY, TotalBackBmp.Area.Right, TotalBackBmp.Area.Bottom, BackBmp.hDcMemory, 0, 0, SRCCOPY)
    
    'TotalBackBmp should now contain a clean picture with no cursor.
    'copy a section of this as the background for next time.
    Result = BitBlt(BackBmp.hDcMemory, 0, 0, 32, 32, TotalBackBmp.hDcMemory, MouseOffsetX, MouseOffsetY, SRCCOPY)
End If

'draw the mouse cursor onto TotalBackBmp
Result = BitBlt(TotalBackBmp.hDcMemory, MouseOffsetX, MouseOffsetY, TotalBackBmp.Area.Right, TotalBackBmp.Area.Bottom, CursorBmp.hDcMemory, 0, 0, SRCAND)

'copy the drawn picture onto the screen
Result = BitBlt(frmMouseShadow.hDc, X - MouseOffsetX, Y - MouseOffsetY, TotalBackBmp.Area.Right, TotalBackBmp.Area.Bottom, TotalBackBmp.hDcMemory, 0, 0, SRCCOPY)

LastX = X
LastY = Y

Call DeleteBitmap(TotalBackBmp.hDcMemory, TotalBackBmp.hDcBitmap, TotalBackBmp.hDcPointer)
End Sub

