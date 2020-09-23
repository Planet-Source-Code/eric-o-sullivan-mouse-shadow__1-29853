Attribute VB_Name = "APIGraphics"
'=================================
'12/11/2001
'----------------------------------------------
' Author : Eric O'Sullivan
'----------------------------------------------
'Contact : DiskJunky@hotmail.com
'----------------------------------------------
'Comments :
'This module was made for using api graphics functions in your
'programs. With the following api calls and function and procedures
'written by me, you have to tools to do almost anything. The only api
'function listed here that is not directly used by any piece of code
'in this module is BitBlt. You have the tools to create and manipulate
'graphics, but it is still necessary to display them manually. The
'functions themselves mostly need hDc or a handle to work. You can
'find this hDc in both a forms and pictureboxs' properties. I have
'also set up a data type called BitmapStruc. For my programs, I have
'used this structure almost exclusivly for the graphics. The structure
'holds all the information needed to reference a bitmap created using
'this module (CreateNewBitmap, DeleteBitmap).
'Please keep in mind that any object (bitmap, brush, pen etc) needs to
'be deleted after use or else it will stay in memory until the program is
'finished. Not doing so will eventually cause your program to take up
'ALL your computers recources.
'
'Thank you,
'Eric
'----------------------------------------------
'----------------------------------------------

'These functions are sorted alphabetically.
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (ByRef wef As Any, ByVal i As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LogBrush) As Long
Public Declare Function CreateColorSpace Lib "gdi32" Alias "CreateColorSpaceA" (lplogcolorspace As LogColorSpace) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32" (EllRect As Rect) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LogPen) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (Left As Integer, Top As Integer, Right As Integer, Bottom As Integer) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ColorRef As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As Rect, lprcTo As Rect) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDc As Long, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) As Boolean
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal A As Long, ByVal B As Long, wef As DEVMODE) As Boolean
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LogBrush, ByVal dwStyleCount As Long, lpStyle As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hwnd As Long, Fill As Rect, hBrush As Long) As Integer
Public Declare Function FillRgn Lib "gdi32" (ByVal hDc As Long, ByVal HRgn As Long, hBrush As Long) As Boolean
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long  'very usefull timing function ;)
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As Rect, lpSrc1Rect As Rect, lpSrc2Rect As Rect) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, XEnd As Integer, YEnd As Integer) As Boolean
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, X As Integer, Y As Integer, PointAPI) As Boolean
Public Declare Function PatBlt Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDc As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDc As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long
Public Declare Function PolylineTo Lib "gdi32" (ByVal hDc As Long, lppt As PointAPI, ByVal cCount As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hwnd As Long, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hDc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetColorAdjustment Lib "gdi32" (ByVal hDc As Long, lpca As COLORADJUSTMENT) As Long
Public Declare Function SetColorSpace Lib "gdi32" (ByVal hDc As Long, ByVal hcolorspace As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'the direction of the gradient
Public Enum GradientTo
    GradHorizontal = 0
    GradVertical = 1
End Enum

'in twips or pixels
Public Enum Scaling
    InTwips = 0
    InPixels = 1
End Enum



'Text metrics
Public Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type


Public Type COLORADJUSTMENT
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type

Public Type CIEXYZ
    ciexyzX As Long
    ciexyzY As Long
    ciexyzZ As Long
End Type

Public Type CIEXYZTRIPLE
    ciexyzRed As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyBlue As CIEXYZ
End Type

Public Type LogColorSpace
    lcsSignature As Long
    lcsVersion As Long
    lcsSize As Long
    lcsCSType As Long
    lcsIntent As Long
    lcsEndPoints As CIEXYZTRIPLE
    lcsGammaRed As Long
    lcsGammaGreen As Long
    lcsGammaBlue As Long
    lcsFileName As String * 26 'MAX_PATH
End Type

'display settings (800x600 etc)
Public Type DEVMODE
        dmDeviceName As String * 32
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * 32
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Enum AlignText
    vbLeftAlign = 1
    vbCentreAlign = 2
    vbRightAlign = 3
End Enum

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BitmapStruc
    hDcMemory As Long
    hDcBitmap As Long
    hDcPointer As Long
    Area As Rect
End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Type LogPen
        lopnStyle As Long
        lopnWidth As PointAPI
        lopnColor As Long
End Type

Public Type LogBrush
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Public Type FontStruc
    Name As String
    Alignment As AlignText
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
    StrikeThru As Boolean
    PointSize As Byte
    Colour As Long
End Type

Public Type LogFont
    'for the DrawText api call
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 32) As Byte
End Type

Public Type Point
    'you'll need this to reference a point on the
    'screen'
    X As Integer
    Y As Integer
End Type

'To hold the RGB value
Public Type RGBVal
    Red As Single
    Green As Single
    Blue As Single
End Type

'general constants
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GWL_WNDPROC = (-4)
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3
Public Const WM_USER = &H400
Public rctFrom As Rect
Public rctTo As Rect
Public lngTrayHand As Long
Public lngStartMenuHand As Long
Public lngChildHand As Long
Public strClass As String * 255
Public lngClassNameLen As Long
Public lngRetVal As Long

'PatBlt constants
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

'Display constants
Public Const CDS_FULLSCREEN = 4
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000

'DrawText constants
Public Const DT_CENTER = &H1
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

'CreateBrushIndirect constants
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERN8X8 = 8
Public Const BS_DIBPATTERNPT = 6
Public Const BS_HATCHED = 2
Public Const BS_HOLLOW = 1
Public Const BS_NULL = 1
Public Const BS_PATTERN = 3
Public Const BS_PATTERN8X8 = 7
Public Const BS_SOLID = 0
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_NOSHADE = 17
Public Const HS_SOLID = 8
Public Const HS_SOLIDBKCLR = 23
Public Const HS_SOLIDCLR = 19
Public Const HS_VERTICAL = 1                '  |||||

'BitBlt constants
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)

'LogFont constants
Public Const LF_FACESIZE = 32
Public Const FW_BOLD = 700
Public Const FW_DONTCARE = 0
Public Const FW_EXTRABOLD = 800
Public Const FW_EXTRALIGHT = 200
Public Const FW_HEAVY = 900
Public Const FW_LIGHT = 300
Public Const FW_MEDIUM = 500
Public Const FW_NORMAL = 400
Public Const FW_SEMIBOLD = 600
Public Const FW_THIN = 100
Public Const DEFAULT_CHARSET = 1
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_OUTLINE_PRECIS = 8
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_TT_PRECIS = 4
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_EMBEDDED = 128
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_MASK = &HF
Public Const CLIP_STROKE_PRECIS = 2
Public Const CLIP_TT_ALWAYS = 32
Public Const WM_SETFONT = &H30
Public Const LF_FULLFACESIZE = 64
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_QUALITY = 0
Public Const PROOF_QUALITY = 2

'GetDeviceCaps constants
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

'colourspace constants
Public Const MAX_PATH = 260

'pen constants
Public Const PS_COSMETIC = &H0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_DOT = 2                     '  .......
Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_GEOMETRIC = &H10000
Public Const PS_INSIDEFRAME = 6
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_ROUND = &H0
Public Const PS_SOLID = 0

'mouse cursor constants
Public Const IDC_APPSTARTING = 32650&
Public Const IDC_ARROW = 32512&
Public Const IDC_CROSS = 32515&
Public Const IDC_IBEAM = 32513&
Public Const IDC_ICON = 32641&
Public Const IDC_NO = 32648&
Public Const IDC_SIZE = 32640&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_UPARROW = 32516&
Public Const IDC_WAIT = 32514&

'some key values for GetASyncKeyState
Public Const KLeft = 37
Public Const KUp = 38
Public Const KRight = 39
Public Const KDown = 40

'some mathimatical constants
Public Const Pi = 3.14159265358979

'store information on where what was displayed in the selected
'area before the text was displayed. - not used any more
Private hDcMemOld As Long
Private hDcBmpOld As Long
Private hDcOldPtr As Long
Private OldRect As Rect

'This stores the various points of a polygon. Use DrawPoly to draw
'the polygon with the points specified in the array.
Public PolygonPoints() As PointAPI

Public Function DrawRect(hDc As Long, Colour As Long, Left As Integer, Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, Optional ByVal Measurement As Scaling = InPixels, Optional ByVal Style As Long = BS_SOLID, Optional ByVal Pattern As Long = HS_SOLID)
'this draws a rectangle using the co-ordinates
'and colour given.

Dim StartRect As Rect
Dim RetVal As Long
Dim Junk  As Long
Dim Brush As Long
Dim BrushStuff As LogBrush

If Measurement = InTwips Then
    'convert to pixels
    Left = Left / Screen.TwipsPerPixelX
    Top = Top / Screen.TwipsPerPixelY
    Right = Right / Screen.TwipsPerPixelX
    Bottom = Bottom / Screen.TwipsPerPixelY
End If

StartRect.Top = Top
StartRect.Left = Left
StartRect.Bottom = Bottom
StartRect.Right = Right

BrushStuff.lbColor = Colour
BrushStuff.lbHatch = Pattern
BrushStuff.lbStyle = Style
 
Brush = CreateBrushIndirect(BrushStuff)
Brush = SelectObject(hDc, Brush)
    
RetVal = PatBlt(hDc, Left, Top, (Right - Left), (Bottom - Top), PATCOPY)

'A "Brush" object was created. It must be removed from memory.
Junk = SelectObject(hDc, Brush)
Junk = DeleteObject(Junk)
End Function

Public Sub DrawRoundRect(ByVal hDc As Long, ByVal Colour As Long, ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, ByVal EdgeRadius As Integer, Optional ByVal Measurement As Scaling = InPixels, Optional ByVal Style As Long = BS_SOLID, Optional ByVal Pattern As Long = HS_SOLID)
'this draws a rectangle using the co-ordinates
'and colour given.

Const Width = 1 'the pixel width of the edge

Dim Pen As Long
Dim PenStuff As LogPen
Dim Brush As Long
Dim BrushStuff As LogBrush
Dim Junk  As Long
Dim Result As Long
Dim OffsetRect As Rect

If Measurement = InTwips Then
    'convert twip values to pixels
    Left = Left / Screen.TwipsPerPixelX
    Top = Top / Screen.TwipsPerPixelX
    Right = Right / Screen.TwipsPerPixelY
    Bottom = Bottom / Screen.TwipsPerPixelY
    EdgeRadius = EdgeRadius / Screen.TwipsPerPixelX
End If

'Find out if a specific colour is to be set. If so set it.
BrushStuff.lbColor = Colour
BrushStuff.lbStyle = Style
BrushStuff.lbHatch = Pattern  'ignored if style is solid

PenStuff.lopnColor = Colour
PenStuff.lopnWidth.X = Width
PenStuff.lopnStyle = PS_SOLID

Pen = CreatePenIndirect(PenStuff)
Pen = SelectObject(hDc, Pen)
Brush = CreateBrushIndirect(BrushStuff)
Brush = SelectObject(hDc, Brush)

'deflate the rectangle dimensions by the radius amount
OffsetRect.Left = Left
OffsetRect.Top = Top
OffsetRect.Right = Right
OffsetRect.Bottom = Bottom
'Result = InflateRect(OffsetRect, -EdgeRadius, -EdgeRadius)

'draw the rounded rectangle
Result = RoundRect(hDc, OffsetRect.Left, OffsetRect.Top, OffsetRect.Right, OffsetRect.Bottom, EdgeRadius, EdgeRadius)

'delete the objects created
Junk = SelectObject(hDc, Pen)
Junk = DeleteObject(Junk)
Junk = SelectObject(hDc, Brush)
Junk = DeleteObject(Junk)
End Sub

Public Function TitleToTray(frm As Form)
 
lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

Do
    lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
    If InStr(1, strClass, "TrayNotifyWnd") Then
        lngTrayHand = lngChildHand
        Exit Do
    End If
    lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
Loop

lngRetVal = GetWindowRect(frm.hwnd, rctFrom)
lngRetVal = GetWindowRect(lngTrayHand, rctTo)
lngRetVal = DrawAnimatedRects(frm.hwnd, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo)

'hide form
frm.Visible = False
frm.Hide
End Function

Public Function TrayToTitle(frm As Form)

lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)
Do
    lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
    If InStr(1, strClass, "TrayNotifyWnd") Then
        lngTrayHand = lngChildHand
        Exit Do
    End If
    lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
Loop

lngRetVal = GetWindowRect(frm.hwnd, rctFrom)
lngRetVal = GetWindowRect(lngTrayHand, rctTo)
lngRetVal = DrawAnimatedRects(frm.hwnd, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom)

frm.Visible = True
frm.Show
End Function

Public Sub DrawLine(hDc As Long, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, Optional ByVal Colour As Long = 0, Optional ByVal Width As Integer = 1, Optional ByVal Mesurement As Scaling = InTwips)
'This will draw a line from point1 to point2

Const NumOfPoints = 2

Dim Result As Long
Dim BlankPoint As Integer
Dim Pen As Long
Dim PenStuff As LogPen
Dim Brush As Long
Dim BrushStuff As LogBrush
Dim Junk  As Long
Dim Points(NumOfPoints) As PointAPI

BlankPoint = 0

If Mesurement = InTwips Then
    'convert twip values to pixels
    X1 = X1 / Screen.TwipsPerPixelX
    X2 = X2 / Screen.TwipsPerPixelX
    Y1 = Y1 / Screen.TwipsPerPixelY
    Y2 = Y2 / Screen.TwipsPerPixelY
End If

'Find out if a specific colour is to be set. If so set it.
BrushStuff.lbColor = Colour
BrushStuff.lbHatch = 0  'ignored if style is solid
BrushStuff.lbStyle = BS_SOLID

PenStuff.lopnColor = Colour
PenStuff.lopnWidth.X = Width
PenStuff.lopnStyle = PS_SOLID

Pen = ExtCreatePen((PS_GEOMETRIC Or PS_SOLID Or PS_ENDCAP_ROUND Or PS_JOIN_MITER), Width, BrushStuff, 0, 0)
Pen = SelectObject(hDc, Pen)

Points(1).X = X1
Points(1).Y = Y1
Points(2).X = X2 '- (Width * 3)
Points(2).Y = Y2 '- (Width * 3)
Result = Polyline(hDc, Points(1), NumOfPoints)
Result = GetLastError

'A "Pen" object was created. It must be removed from memory.
Junk = SelectObject(hDc, Pen)
Junk = DeleteObject(Junk)
End Sub

Public Sub DrawPoly(ByVal hDc As Long, Optional ByVal FillColour As Long = 0, Optional ByVal BorderColour As Long = 0, Optional ByVal BorderWidth As Integer = 1, Optional ByVal Measurement As Scaling = InPixels)
'This function will draw a polygon of size and colours specified.
'Modified to draw and fill all the points specified in the array,
'PolygonPoints(). Use;
'
'"ReDim Preserve PolygonPoints(UBound(PolygonPoints) + 1)"
'
'to add a point to the array, then specifiy the X and Y values
'of your point.
'
'To delete a point from the array, use the same as above
'except replace the + sign with a - sign.
'
'To clear the array of all data, use;
'
'"ReDim PolygonPoints(0)"
'
'----------------
'- Eric  15/12/2001
'----------------

Dim Sucessful As Long
Dim Counter As Integer
Dim Temp() As PointAPI
Dim PointNum As Integer
Dim hBrush As Long
Dim BrushStuff As LogBrush
Dim hPen As Long
Dim PenStuff As LogPen

'find the number of points stored
PointNum = UBound(PolygonPoints)

'convert all points to pixels if necessary, otherwise, don't waste time.
If Measurement = InTwips Then
    'create a temperory array to hold the points
    ReDim Temp(PointNum)
    
    'convert to pixels
    For Counter = 0 To PointNum
        Temp(Counter).X = PolygonPoints(Counter).X / Screen.TwipsPerPixelX
        Temp(Counter).Y = PolygonPoints(Counter).Y / Screen.TwipsPerPixelY
    Next Counter
End If

'apply the border and background colours
BrushStuff.lbColor = FillColour
BrushStuff.lbHatch = 0
BrushStuff.lbStyle = BS_SOLID

PenStuff.lopnColor = BorderColour
PenStuff.lopnWidth.X = BorderWidth
PenStuff.lopnStyle = PS_SOLID

'create the objects necessary to apply the colour and draw settings
hBrush = CreateBrushIndirect(BrushStuff)
hPen = CreatePenIndirect(PenStuff)
Sucessful = SelectObject(hDc, hBrush)
Sucessful = SelectObject(hDc, hPen)

'draw the polygon
If Measurement = InTwips Then
    Sucessful = Polygon(hDc, Temp(0), PointNum)
Else
    Sucessful = Polygon(hDc, PolygonPoints(0), PointNum)
End If

'remove the pen and brush objects
Sucessful = SelectObject(hDc, hBrush)
Sucessful = DeleteObject(Sucessful)
Sucessful = SelectObject(hDc, hPen)
Sucessful = DeleteObject(Sucessful)
End Sub

Public Function LockWindow(FormName As Form)
'Prevent the form from updating its display

Dim RetVal As Boolean

'RetVal = LockWindowUpdate(FormName.hwnd)
End Function

Public Function UnLockWindow()
'Let the form update its display
'RetVal = LockWindowUpdate(0)
End Function

Public Sub CreateNewBitmap(ByRef hDcMemory As Long, ByRef hDcBitmap As Long, ByRef hDcPointer As Long, ByRef BmpArea As Rect, CompatableWith As Form, Optional ByVal BackColour As Long = 0, Optional ByVal Measurement As Scaling)
'This procedure will create a new bitmap compatable with a given
'form (you will also be able to then use this bitmap in a picturebox).
'The space specified in "Area" should be in "Twips" and will be
'converted into pixels in the following code.

Dim Result As Long
Dim Area As Rect

Area = BmpArea
If Measurement = InTwips Then
    Call RectToPixels(Area)
End If

hDcMemory = CreateCompatibleDC(CompatableWith.hDc)
hDcBitmap = CreateCompatibleBitmap(CompatableWith.hDc, (Area.Right - Area.Left), (Area.Bottom - Area.Top))
hDcPointer = SelectObject(hDcMemory, hDcBitmap)

'set default colours and clear bitmap to selected colour
Result = SetBkMode(hDcMemory, OPAQUE)
Result = SetBkColor(hDcMemory, BackColour)
Call DrawRect(hDcMemory, BackColour, 0, 0, (Area.Right - Area.Left), (Area.Bottom - Area.Top))
End Sub

Public Sub DeleteBitmap(ByRef hDcMemory As Long, ByRef hDcBitmap As Long, ByRef hDcPointer As Long)
'This will remove the bitmap that stored what was displayed before
'the text was written to the screen, from memory.
Dim Junk As Long

If hDcMemory = 0 Then
    'there is nothing to delete. Exit the sub-routine
    Exit Sub
End If

Junk = SelectObject(hDcMemory, hDcPointer)
Junk = DeleteObject(hDcBitmap)
Junk = DeleteDC(hDcMemory)

hDcMemory = 0
hDcBitmap = 0
hDcPointer = 0
End Sub

Public Sub RectToTwips(ByRef TheRect As Rect)
'converts pixels to twips in a rect structure

TheRect.Left = TheRect.Left * Screen.TwipsPerPixelX
TheRect.Right = TheRect.Right * Screen.TwipsPerPixelX
TheRect.Top = TheRect.Top * Screen.TwipsPerPixelY
TheRect.Bottom = TheRect.Bottom * Screen.TwipsPerPixelY
End Sub

Public Sub RectToPixels(ByRef TheRect As Rect)
'converts twips to pixels in a rect structure

TheRect.Left = TheRect.Left \ Screen.TwipsPerPixelX
TheRect.Right = TheRect.Right \ Screen.TwipsPerPixelX
TheRect.Top = TheRect.Top \ Screen.TwipsPerPixelY
TheRect.Bottom = TheRect.Bottom \ Screen.TwipsPerPixelY
End Sub

Public Function NewColourSpace(Red As Integer, Green As Integer, Blue As Integer) As Long
'Returns the handle of the new colour space

Dim ColourSpace As LogColorSpace

'set values of the colourspace
ColourSpace.lcsGammaRed = Red
ColourSpace.lcsGammaGreen = Green
ColourSpace.lcsGammaBlue = Blue

NewColourSpace = CreateColorSpace(ColourSpace)
End Function

Public Sub SetColour(hDc As Long, Red As Integer, Green As Integer, Blue As Integer)
'sets the colour of a bitmap

Dim Adjustment As COLORADJUSTMENT
Dim Result As Long

'Adjustment

End Sub

Public Function AmIActive(TheForm As Form) As Boolean
'This function returns wether or not the window is active

If TheForm.hwnd = GetActiveWindow Then
    AmIActive = True
Else
    AmIActive = False
End If
End Function

Public Sub Gradient(ByVal DesthDc As Long, ByVal StartCol As Long, ByVal FinishCol As Long, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Direction As GradientTo, Optional ByVal Mesurement As Scaling = 1, Optional ByVal LineWidth As Byte = 1)
'draws a gradient from colour Start to colour Finish, and assums
'that all measurments passed to it are in pixels unless otherwise
'specified.

Dim Counter As Integer
Dim BigDiff As Integer
Dim Colour As RGBVal
Dim Start As RGBVal
Dim Finish As RGBVal
Dim AddRed As Single
Dim AddGreen As Single
Dim AddBlue As Single

'perform all necessary calculations before drawing gradient
'such as converting long to rgb values, and getting the correct
'scaling for the bitmap.
Start = GetRGB(StartCol)
Finish = GetRGB(FinishCol)

If Mesurement = InTwips Then
    Left = Left / Screen.TwipsPerPixelX
    Top = Top / Screen.TwipsPerPixelY
    Width = Width / Screen.TwipsPerPixelX
    Height = Height / Screen.TwipsPerPixelY
End If

'draw the colour gradient
Select Case Direction
Case GradVertical
    BigDiff = Width
Case GradHoizontal
    BigDiff = Height
End Select

'calculate how much to increment/decrement each colour per step
AddRed = (LineWidth * ((Finish.Red) - Start.Red) / BigDiff)
AddGreen = (LineWidth * ((Finish.Green) - Start.Green) / BigDiff)
AddBlue = (LineWidth * ((Finish.Blue) - Start.Blue) / BigDiff)
Colour = Start

'calculate the colour of each line before drawing it on the bitmap
For Counter = 0 To BigDiff Step LineWidth
    'find the point between colour Start and Colour Finish that
    'corresponds to the point between 0 and BigDiff
    
    'check for overflow
    If Colour.Red > 255 Then
        Colour.Red = 255
    Else
        If Colour.Red < 0 Then
            Colour.Red = 0
        End If
    End If
    If Colour.Green > 255 Then
        Colour.Green = 255
    Else
        If Colour.Green < 0 Then
            Colour.Green = 0
        End If
    End If
    If Colour.Blue > 255 Then
        Colour.Blue = 255
    Else
        If Colour.Blue < 0 Then
            Colour.Blue = 0
        End If
    End If
    
    'draw the gradient in the proper orientation in the calculated colour
    Select Case Direction
    Case GradVertical
        Call DrawLine(DesthDc, Counter + Left, Top, Counter + Left, Height + Top, RGB(Colour.Red, Colour.Green, Colour.Blue), LineWidth, InPixels)
    Case GradHorizontal
        Call DrawLine(DesthDc, Left, Counter + Top, Left + Width, Top + Counter, RGB(Colour.Red, Colour.Green, Colour.Blue), LineWidth, InPixels)
    End Select
    
    Colour.Red = Colour.Red + AddRed
    Colour.Green = Colour.Green + AddGreen
    Colour.Blue = Colour.Blue + AddBlue
Next Counter
End Sub

Public Sub FadeGradient(ByVal DesthDc As Long, ByVal DestLeft As Integer, ByVal DestTop As Integer, ByVal DestWidth As Integer, ByVal DestHeight As Integer, ByVal GradhDc As Long, StartFromA As Long, FinishToA As Long, StartFromB As Long, FinishToB As Long, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Direction As GradientTo, Optional ByVal Mesurement As Scaling = 1, Optional ByVal LineWidth As Byte = 1)
'This procedure will call the Gradient function to fade it into
'the colours specified.
'Note : all mesurements must me of the same scale, ie they must all
'be in pixels or all in twips.

Dim Colour(2) As RGBVal
Dim Start(2) As RGBVal
Dim Finish(2) As RGBVal
Dim GradCol(2) As Long
Dim Counter As Integer
Dim BigDiff As Integer
Dim Value As Integer
Dim Index As Integer
Dim Result As Long

Const A = 0
Const B = 1

'convert to RGB values
Start(A) = GetRGB(StartFromA)
Start(B) = GetRGB(StartFromB)
Finish(A) = GetRGB(FinishToA)
Finish(B) = GetRGB(FinishToB)

'convert to pixels if necessary
If Mesurement = InTwips Then
    DestLeft = DestLeft / Screen.TwipsPerPixelX
    DestTop = DestTop / Screen.TwipsPerPixelY
    DestWidth = DestWidth / Screen.TwipsPerPixelX
    DestHeight = DestHeight / Screen.TwipsPerPixelY
    Left = Left / Screen.TwipsPerPixelX
    Top = Top / Screen.TwipsPerPixelY
    Width = Width / Screen.TwipsPerPixelX
    Height = Height / Screen.TwipsPerPixelY
End If


'Find the largest difference between any two corresponding
'colours, and use that as the number of steps to take in the loop,
'(therefore guarenteing that it will cycle through all necessary
'colours without jumping)
For Index = A To B
    'test red
    Value = PositVal(Start(Index).Red - Finish(Index).Red)
    If Value > BigDiff Then
        BigDiff = Value
    End If
    
    'test green
    Value = PositVal(Start(Index).Green - Finish(Index).Green)
    If Value > BigDiff Then
        BigDiff = Value
    End If
    
    'test blue
    Value = PositVal(Start(Index).Blue - Finish(Index).Blue)
    If Value > BigDiff Then
        BigDiff = Value
    End If
Next Index

'if there is no difference, then just draw one gradient and exit
If BigDiff = 0 Then
    Call Gradient(GradhDc, StartFromA, StartFromB, Left, Top, Width, Height, Direction, InPixels, LineWidth)
    Exit Sub
End If

'fade the gradient
For Counter = 0 To BigDiff
    'find the point between colour Start and Colour Finish that
    'corresponds to the point between 0 and BigDiff
    
    For Index = A To B
        Colour(Index).Red = Start(Index).Red + (((Finish(Index).Red - Start(Index).Red) / BigDiff) * Counter)
        Colour(Index).Green = Start(Index).Green + (((Finish(Index).Green - Start(Index).Green) / BigDiff) * Counter)
        Colour(Index).Blue = Start(Index).Blue + (((Finish(Index).Blue - Start(Index).Blue) / BigDiff) * Counter)
    
        'convert to long value and store
        GradCol(Index) = RGB(Colour(Index).Red, Colour(Index).Green, Colour(Index).Blue)
    Next Index
    
    'draw the gradient onto the bitmap
    Call Gradient(GradhDc, GradCol(A), GradCol(B), Left, Top, Width, Height, Direction, InPixels, LineWidth)

    'blitt the bitmap to the screen
    Result = BitBlt(DesthDc, DestLeft, DestTop, DestWidth, DestHeight, GradhDc, Left, Top, SRCCOPY)
    DoEvents
Next Counter
End Sub

Public Function PositVal(Value As Integer) As Integer
'-obsolete. Use the Abs function - Eric, 6/12/2001
'Returns the positive value of a number

'PosVal = Sqr(Value ^ 2)
PositVal = Abs(Value)
End Function

Public Function ToRGB(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Long
'Convert RGB to Long
 
Dim MyVal As Long

MyVal = (CLng(Blue) * 65536) + (CLng(Green) * 256) + Red
ToRGB = MyVal
End Function

Public Function GetRGB(ByVal Colour As Long) As RGBVal
'Convert Long to RGB:

'if the colour value is greater than acceptable then half the value
If (Colour > RGB(255, 255, 255)) Or (Colour < (RGB(255, 255, 255) * -1)) Then
    Exit Function
End If

GetRGB.Blue = (Colour \ 65536)
GetRGB.Green = ((Colour - ((GetRGB.Blue) * 65536)) \ 256)
GetRGB.Red = (Colour - (GetRGB.Blue * (65536)) - ((GetRGB.Green) * 256))
End Function

Public Sub Pause(Ticks As Long)
'pause execution of the program for a specified number of ticks

If Ticks < 0 Then
    Ticks = 0
End If
Call Sleep(Ticks)
End Sub


Public Sub DrawEllipse(ByVal hDc As Long, ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal Height As Integer, ByVal Width As Integer, ByVal TiltAngle As Integer, Optional ByVal Colour As Long, Optional Thickness As Integer = 1, Optional Mesurement As Scaling)
'This procedure will draw an ellipse of the specified dimensions and
'colour, by drawing a line between each of the 360 points that make
'up the ellipse.

Const A = 1
Const B = 2

Dim MoveCenterX(2) As Single
Dim MoveCenterY(2) As Single
Dim CircleX(2) As Single
Dim CircleY(2) As Single
Dim Counter As Single
Dim DrawNew As Boolean
Dim TiltX As Single
Dim TiltY As Single
Dim NumOfPoints As Single

'set scaling values
If Mesurement = InTwips Then
    'convert parameters to pixels
    CenterX = (CenterX / Screen.TwipsPerPixelX) '- Thickness
    CenterY = (CenterY / Screen.TwipsPerPixelY) '- Thickness
    Height = (Height / Screen.TwipsPerPixelY) - (Thickness * 2)
    Width = (Width / Screen.TwipsPerPixelX) - (Thickness * 2)
    
    'values are now in pixels
    Mesurement = InPixels
End If

'calculate the radius for width and height
Height = Height / 2
Width = (Width / 2) - Height

'calculate the starting point of the ellipse
TiltX = Sin(TiltAngle * Pi / 180) * Width
TiltY = Cos(TiltAngle * Pi / 180) * Width

'draw the ellipse using one line for every three pixels
'This will increase drawing speed on small ellipses and produce
'detailed ones for large ellipses.
NumOfPoints = (360 / (((Width + Height) * 2) / 3))
For Counter = 0 To (360 + NumOfPoints) Step NumOfPoints
    'calculate points
    If DrawNew Then
        MoveCenterX(B) = MoveCenterX(A)
        MoveCenterY(B) = MoveCenterY(A)
    End If
    MoveCenterX(A) = CenterX + (Cos(Counter * Pi / 180) * TiltX) ' - Z
    MoveCenterY(A) = CenterY + (Cos(Counter * Pi / 180) * TiltY) ' - Z
    
    'record the previous position so you can draw a line to it
    If DrawNew Then
        CircleX(B) = CircleX(A)
        CircleY(B) = CircleY(A)
    End If
    
    'calculate the new position
    CircleX(A) = Sin((Counter + TiltAngle) * Pi / 180) * Height
    CircleY(A) = Cos((Counter + TiltAngle) * Pi / 180) * Height

    'draw lines
    If DrawNew Then
        Call DrawLine(hDc, MoveCenterX(A) + CircleX(A), MoveCenterY(A) + CircleY(A), MoveCenterX(B) + CircleX(B), MoveCenterY(B) + CircleY(B), Colour, Thickness, Mesurement)
    Else
        DrawNew = True
    End If
Next Counter
End Sub


Public Function GetAngle(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Integer
'returns the angle of point1 in relation to point2

Dim TempAngle As Integer

'if the values are not over the center point, then calculate the angle
If (Abs(Y1 - Y2) <> 0) And (Abs(X1 - X2) <> 0) Then
    
    TempAngle = (Atn(Slope(X1, Y1, X2, Y2)) * 180 / Pi) Mod 360
    If TempAngle > 0 Then
        TempAngle = 90 - TempAngle
    Else
        TempAngle = Abs(TempAngle) + 90
    End If
    If X1 < X2 Then
        TempAngle = TempAngle + 180
    End If
    
    GetAngle = TempAngle
End If
End Function

Function Slope(X1, Y1, X2, Y2)
'This function finds the slope of a line, where the slope, m =
'       X1 - Y1
'm = ------------
'       Y2 - Y2

Dim XVal As Integer
Dim YVal As Integer

XVal = X2 - X1
YVal = Y2 - Y1
If (XVal = 0) And (YVal = 0) Then
    'if both values were zero, then
    Slope = 0
    Exit Function
Else
    'if only one value was zero then
    If (XVal = 0) Or (YVal = 0) Then
        'the slope = the other value
        Select Case 0
        Case XVal
            Slope = XVal
        Case YVal
            Slope = YVal
        End Select
        Exit Function
    End If
End If

If (XVal <> 0) And (YVal <> 0) Then
    'otherwise the slope = the formula
    Slope = (Y2 - Y1) / (X2 - X1)
End If
End Function

Public Sub MakeText(ByVal hDcSurphase As Long, ByVal Text As String, ByVal Top As Integer, ByVal Left As Integer, ByVal Height As Integer, ByVal Width As Integer, Font As FontStruc, Optional ByVal Mesurement As Scaling = 0)
'This procedure will draw text onto the bitmap in the specified font,
'colour and position.

Dim APIFont As LogFont
Dim Alignment As Long
Dim TextRect As Rect
Dim Result As Long
Dim Junk As Long
Dim hDcFont As Long
Dim hDcOldFont As Long

'set mesurement values
TextRect.Top = Top
TextRect.Left = Left
TextRect.Right = Left + Width
TextRect.Bottom = Top + Height

If Mesurement = InTwips Then
    'convert to pixels
    Call RectToPixels(TextRect)
End If

'Create details about the font using the Font structure
'====================

'convert point size to pixels
APIFont.lfHeight = -((Font.PointSize * GetDeviceCaps(hDcSurphase, LOGPIXELSY)) / 72)
APIFont.lfCharSet = DEFAULT_CHARSET
APIFont.lfClipPrecision = CLIP_DEFAULT_PRECIS
APIFont.lfEscapement = 0

'move the name of the font into the array
For Counter = 1 To Len(Font.Name)
    APIFont.lfFaceName(Counter) = Asc(Mid(Font.Name, Counter, 1))
Next Counter
APIFont.lfFaceName(Counter) = 0   'this has to be a Null terminated string

APIFont.lfItalic = Font.Italic
APIFont.lfUnderline = Font.Underline
APIFont.lfStrikeOut = Font.StrikeThru
APIFont.lfOrientation = 0
APIFont.lfOutPrecision = OUT_DEFAULT_PRECIS
APIFont.lfPitchAndFamily = DEFAULT_PITCH
APIFont.lfQuality = PROOF_QUALITY

If Font.Bold Then
    APIFont.lfWeight = FW_BOLD
Else
    APIFont.lfWeight = FW_NORMAL
End If

APIFont.lfWidth = 0
hDcFont = CreateFontIndirect(APIFont)
hDcOldFont = SelectObject(hDcSurphase, hDcFont)
'====================

Select Case Font.Alignment
Case vbLeftAlign
    Alignment = DT_LEFT
Case vbCentreAlign
    Alignment = DT_CENTER
Case vbRightAlign
    Alignment = DT_RIGHT
End Select

'Draw the text into the off-screen bitmap before copying the
'new bitmap (with the text) onto the screen.
Result = SetBkMode(hDcSurphase, TRANSPARENT)
Result = SetTextColor(hDcSurphase, Font.Colour)
Result = DrawText(hDcSurphase, Text, Len(Text), TextRect, Alignment)

'clean up by deleting the off-screen bitmap and font
Junk = SelectObject(hDcSurphase, hDcOldFont)
Junk = DeleteObject(hDcFont)
End Sub

Public Function GetTextHeight(ByVal hDc As Long) As Integer
'This function will return the height of the text using the point size

Dim Metrics As TEXTMETRIC
Dim Result As Long
Dim Height As Integer

Result = GetTextMetrics(hDc, Metrics)

GetTextHeight = Metrics.tmHeight

'GetTextHeight = ((PointSize * GetDeviceCaps(TheForm.hdc, LOGPIXELSY)) / 72)
End Function

Public Sub GetScreenRes(ByRef Width As Integer, ByRef Height)
'This procedure sets the variable to the current screen dimensions.

Width = Screen.Width / Screen.TwipsPerPixelX
Height = Screen.Height / Screen.TwipsPerPixelY
End Sub

Public Sub ReturnOldDisplay()
'returns the display to what it was originally

If ResChanged Then
    Call ChangeDisplaySettings(Null, 0)
    ResChanged = False
End If
End Sub

Public Sub SetDisplay(Width As Integer, Height As Integer)
'changes the resolution of the screen to new size

Dim DevM    As DEVMODE
Dim Result As Long
Dim iAns    As Integer

lResult = EnumDisplaySettings(0, 0, DevM)

With DevM
    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        .dmPelsWidth = Width  'ScreenWidth
        .dmPelsHeight = Height 'ScreenHeight
        .dmBitsPerPel = 32 '(could be 8, 16, 32 or even 4)
End With
Result = ChangeDisplaySettings(DevM, 2)

If Result = 0 Then
        Call ChangeDisplaySettings(DevM, 4)
End If

ResChanged = True
End Sub

'test data
'-------------------
Public Function LoadBmp(FileName As String) As Long
'The function returns the handle of the bitmap loaded.
'
'NOTE: This function is untested code from another source
'12/12/2001 Eric

Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40

Dim hBitmap As Long

'we have to pass a Null terminated string
FileName = FileName & Chr(0)
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE) 'LR_LOADFROMFILE Or LR_CREATEDIBSECTION

LoadBmp = hBitmap
End Function

Public Function RectIntersects(Rect1 As Rect, Rect2 As Rect) As Boolean
'This function will return True if the two passed Rect structers
'intersect each other.

Dim TempRect As Rect
Dim Result As Long

Result = IntersectRect(TempRect, Rect1, Rect2)
RectIntersects = Result
End Function
