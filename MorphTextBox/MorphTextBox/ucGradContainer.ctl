VERSION 5.00
Begin VB.UserControl ucGradContainer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   ToolboxBitmap   =   "ucGradContainer.ctx":0000
End
Attribute VB_Name = "ucGradContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*************************************************************************
'* MorphContainer - Owner-drawn gradient container control.              *
'* Matthew R. Usner, April, 2005.                                        *
'*************************************************************************
'* You are encouraged to use this control in your projects, as long as   *
'* all credits remain intact.  Only lowlife code thieves like Ilia HD    *
'* (MoMoYa) download code, remove the comments, and claim they wrote it. *
'*************************************************************************
'* This is a replacement for VB's dull frame control.  Features include  *
'* separate gradients for header and container, icon/bitmap display      *
'* capability, and the ability to round each corner individually.        *
'* Container and header gradients can be drawn at any angle.             *
'* Loosely based on XP Container Control written by Cameron Groves and   *
'* initially modified by Jim Jose April, 2005.                           *
'* original gradient draw routine by Carles P.V. at txtCodeID=60580      *
'*************************************************************************

Option Explicit

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Const RGN_DIFF As Long = 4

' declares for Carles P.V.'s gradient paint routine.
Private Type BITMAPINFOHEADER
   biSize          As Long
   biWidth         As Long
   biHeight        As Long
   biPlanes        As Integer
   biBitCount      As Integer
   biCompression   As Long
   biSizeImage     As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed       As Long
   biClrImportant  As Long
End Type
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Const DIB_RGB_COLORS As Long = 0
Private Const PI             As Single = 3.14159265358979
Private Const TO_DEG         As Single = 180 / PI
Private Const TO_RAD         As Single = PI / 180
Private Const INT_ROT        As Long = 1000

' used to define the text drawing area.
Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

'  enum for determining size of icon/bitmap.
Public Enum IconSizeEnum
   [Display Full Size] = 0
   [Size To Header] = 1
End Enum

'  property variables and constants
Private m_CurveTopLeft     As Long                  ' the curvature of the top left corner.
Private m_CurveTopRight    As Long                  ' the curvature of the top right corner.
Private m_CurveBottomLeft  As Long                  ' the curvature of the bottom left corner.
Private m_CurveBottomRight As Long                  ' the curvature of the bottom right corner.
Private m_HeaderVisible    As Boolean               ' flag that shows/hides header.
Private m_BackMiddleOut    As Boolean               ' flag for container background middle-out gradient.
Private m_HeaderMiddleOut  As Boolean               ' flag for header middle-out gradient.
Private m_Enabled          As Boolean               ' enabled/disabled flag.
Private m_HeaderAngle      As Single                ' the angle of the header gradient.
Private m_BackAngle        As Single                ' background gradient display angle
Private m_Iconsize         As IconSizeEnum          ' icon size - full or size to header
Private m_HeaderColor1     As OLE_COLOR             ' the first gradient color of the header.
Private m_HeaderColor2     As OLE_COLOR             ' the second gradient color of the header.
Private m_BackColor1       As OLE_COLOR             ' the first gradient color of the background.
Private m_BackColor2       As OLE_COLOR             ' the second gradient color of the background.
Private m_BorderWidth      As Integer               ' width, in pixels, of border.
Private m_BorderColor      As OLE_COLOR             ' color of border.
Private m_CaptionColor     As OLE_COLOR             ' text color of caption.
Private m_Caption          As String                ' caption text.
Private m_HeaderHeight     As Long                  ' height, in pixels, of the header.
Private m_CaptionFont      As StdFont               ' font used to display header text.
Private m_Alignment        As AlignmentConstants    ' caption alignment (left, center, right).
Private m_Icon             As Picture               ' the icon or bitmap to display in the header.

Private Const m_def_CurveTopLeft = 0               ' initialize top left curvature to 0.
Private Const m_def_CurveTopRight = 0              ' initialize top right curvature to 0.
Private Const m_def_CurveBottomLeft = 0            ' initialize bottom left curvature to 0.
Private Const m_def_CurveBottomRight = 0           ' initialize bottom right curvature to 0.
Private Const m_def_HeaderVisible = True           ' initialize the header to be visible.
Private Const m_def_BackMiddleOut = True           ' initialize to a middle-out background gradient.
Private Const m_def_HeaderMiddleOut = True         ' initialize to a middle-out header gradient.
Private Const m_def_Enabled = 0                    ' initialize to disabled.
Private Const m_def_HeaderAngle = 90               ' initialize to horizontal header gradient.
Private Const m_def_BackAngle = 90                 ' initialize to horizontal background gradient.
Private Const m_def_Iconsize = 1                   ' initialize to 'size to header'
Private Const m_def_HeaderColor2 = &HF7E0D3
Private Const m_def_HeaderColor1 = &HEDC5A7
Private Const m_def_BackColor2 = &HFCF4EF
Private Const m_def_BackColor1 = &HFAE8DC
Private Const m_def_Caption = "Gradient Container" ' default caption text.
Private Const m_def_BorderWidth = 1                ' initialize border width to 1 pixel.
Private Const m_def_BorderColor = &HDCC1AD
Private Const m_def_Align = vbLeftJustify          ' initalize text to left justification.
Private Const m_def_CaptionColor = &H7B2D02
Private Const m_def_hHeight = 25                   ' initialize to 25 pixels in height.

'  events
Event Resize()

'  miscellaneous control variables
Private m_hMod         As Long

Private Sub RedrawControl()

'*************************************************************************
'* master routine for painting of container.                             *
'*************************************************************************

   SetBackGround
   If m_HeaderVisible Then
      SetHeader
   End If
   CreateBorder
   UserControl.Refresh

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* converts color long COLORREF for api coloring purposes.               *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Events >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_Initialize()
   m_hMod = LoadLibrary("shell32.dll") ' Used to prevent crashes on Windows XP
End Sub

Private Sub UserControl_Show()
   RedrawControl
End Sub

Private Sub UserControl_Resize()
   RedrawControl
End Sub

Private Sub UserControl_Terminate()
   FreeLibrary m_hMod ' Used to prevent crashes on Windows XP
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Graphics >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub SetBackGround()

'*************************************************************************
'* displays the control's background gradient.                           *
'*************************************************************************

   PaintGradient hDC, 0, 0, ScaleWidth, ScaleHeight, _
                 TranslateColor(m_BackColor1), TranslateColor(m_BackColor2), m_BackAngle, m_BackMiddleOut

End Sub

Private Sub CreateBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvatures.    *
'*************************************************************************

   Dim hRgn1   As Long
   Dim hRgn2  As Long
   Dim hBrush As Long

'  create the outer region.
   hRgn1 = pvGetRoundedRgn(0, 0, ScaleWidth, ScaleHeight, _
                           m_CurveTopLeft, m_CurveTopRight, _
                           m_CurveBottomLeft, m_CurveBottomRight)

'  create the inner region.
   hRgn2 = pvGetRoundedRgn(m_BorderWidth, m_BorderWidth, _
                           ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth, _
                           m_CurveTopLeft, m_CurveTopRight, _
                           m_CurveBottomLeft, m_CurveBottomRight)

'  combine the outer and inner regions.
   CombineRgn hRgn2, hRgn1, hRgn2, RGN_DIFF
'  create the brush used to color the combined regions.
   hBrush = CreateSolidBrush(TranslateColor(m_BorderColor))
'  color the combined regions.
   FillRgn hDC, hRgn2, hBrush

'  set the control region
   SetWindowRgn hWnd, hRgn1, True

'  delete created objects to restore memory.
   DeleteObject hBrush
   DeleteObject hRgn1
   DeleteObject hRgn2

End Sub

Private Function pvGetRoundedRgn(ByVal x1 As Long, ByVal y1 As Long, _
                                 ByVal x2 As Long, ByVal y2 As Long, _
                                 ByVal TopLeftRadius As Long, _
                                 ByVal TopRightRadius As Long, _
                                 ByVal BottomLeftRadius As Long, _
                                 ByVal BottomRightRadius As Long _
                                 ) As Long

'*************************************************************************
'* allows each corner of the container to have its own curvature.        *
'* Code by the Amazing Carles P.V.  Thanks a million (as usual) Carles.  *
'*************************************************************************

   Dim hRgnMain As Long   ' the original "starting point" region.
   Dim hRgnTmp1 As Long   ' the first region that defines a corner's radius.
   Dim hRgnTmp2 As Long   ' the second region that defines a corner's radius.

'  Bounding region
   hRgnMain = CreateRectRgn(x1, y1, x2, y2)

'  Top-left corner
   hRgnTmp1 = CreateRectRgn(x1, y1, x1 + TopLeftRadius, y1 + TopLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(x1, y1, x1 + 2 * TopLeftRadius, y1 + 2 * TopLeftRadius)
   Call CombineRgn(hRgnTmp1, hRgnTmp1, hRgnTmp2, RGN_DIFF)
   Call CombineRgn(hRgnMain, hRgnMain, hRgnTmp1, RGN_DIFF)
   Call DeleteObject(hRgnTmp1)
   Call DeleteObject(hRgnTmp2)

'  Top-right corner
   hRgnTmp1 = CreateRectRgn(x2, y1, x2 - TopRightRadius, y1 + TopRightRadius)
   hRgnTmp2 = CreateEllipticRgn(x2 + 1, y1, x2 + 1 - 2 * TopRightRadius, y1 + 2 * TopRightRadius)
   Call CombineRgn(hRgnTmp1, hRgnTmp1, hRgnTmp2, RGN_DIFF)
   Call CombineRgn(hRgnMain, hRgnMain, hRgnTmp1, RGN_DIFF)
   Call DeleteObject(hRgnTmp1)
   Call DeleteObject(hRgnTmp2)

'  Bottom-left corner
   hRgnTmp1 = CreateRectRgn(x1, y2, x1 + BottomLeftRadius, y2 - BottomLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(x1, y2 + 1, x1 + 2 * BottomLeftRadius, y2 + 1 - 2 * BottomLeftRadius)
   Call CombineRgn(hRgnTmp1, hRgnTmp1, hRgnTmp2, RGN_DIFF)
   Call CombineRgn(hRgnMain, hRgnMain, hRgnTmp1, RGN_DIFF)
   Call DeleteObject(hRgnTmp1)
   Call DeleteObject(hRgnTmp2)

'  Bottom-right corner
   hRgnTmp1 = CreateRectRgn(x2, y2, x2 - BottomRightRadius, y2 - BottomRightRadius)
   hRgnTmp2 = CreateEllipticRgn(x2 + 1, y2 + 1, x2 + 1 - 2 * BottomRightRadius, y2 + 1 - 2 * BottomRightRadius)
   Call CombineRgn(hRgnTmp1, hRgnTmp1, hRgnTmp2, RGN_DIFF)
   Call CombineRgn(hRgnMain, hRgnMain, hRgnTmp1, RGN_DIFF)
   Call DeleteObject(hRgnTmp1)
   Call DeleteObject(hRgnTmp2)

   pvGetRoundedRgn = hRgnMain

End Function

Private Sub SetHeader()

'*************************************************************************
'* displays the header gradient, header caption, and an icon if used.    *
'*************************************************************************

   If Not m_CaptionFont Is Nothing Then

'     paint the header gradient.
      PaintGradient hDC, 0, 0, ScaleWidth, m_HeaderHeight, _
                    TranslateColor(m_HeaderColor1), TranslateColor(m_HeaderColor2), m_HeaderAngle, m_HeaderMiddleOut

'     draw the caption.
      Dim TextRect As RECT  ' will define the text drawing region.

'     apply the font and text color.
      Set UserControl.Font = m_CaptionFont
      UserControl.ForeColor = TranslateColor(m_CaptionColor)

      With TextRect
'        define the text drawing area rectangle.
         If m_Alignment = vbCenter Then
            .Left = (ScaleWidth - TextWidth(m_Caption)) / 2
         ElseIf m_Alignment = vbLeftJustify Then
            If IsThere(m_Icon) Then
'              provide for a left-hand clearance of one character plus height of header.
               .Left = TextWidth("A") + m_HeaderHeight
            Else
'              provide for a left-hand clearance of on character width.
               .Left = TextWidth("A")
            End If
         Else
'           provide a right-hand clearance of one character width.
            .Left = (ScaleWidth - TextWidth(m_Caption)) - TextWidth("A")
         End If
         .Top = (m_HeaderHeight - TextHeight(m_Caption)) / 2
         .Bottom = .Top + TextHeight(m_Caption)
         .Right = .Left + TextWidth(m_Caption)
      End With

'     draw the caption.
      DrawText hDC, m_Caption, -1, TextRect, 0

'     draw the icon.
      If IsThere(m_Icon) Then
         If m_Iconsize = [Display Full Size] Then ' don't fit to header;, display full size
            PaintPicture m_Icon, m_BorderWidth + 3, 2
         Else
            PaintPicture m_Icon, m_BorderWidth + 3, 2, m_HeaderHeight - 2, m_HeaderHeight - 3 ' fit to header height
         End If
      End If

   End If

End Sub

Private Function IsThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture by checking dimensions.             *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsThere = Pic.Width <> 0
      End If
   End If

End Function

Public Sub PaintGradient(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, _
                         ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, _
                         ByVal Angle As Single, ByVal bMOut As Boolean)

'*************************************************************************
'* Carles P.V.'s routine, modified by Matthew R. Usner for middle-out    *
'* gradient capability.  Original submission at PSC, txtCodeID=60580.    *
'*************************************************************************

   Dim uBIH      As BITMAPINFOHEADER
   Dim lBits()   As Long
   Dim lGrad()   As Long, lGrad2() As Long

   Dim lClr      As Long
   Dim R1        As Long, G1 As Long, b1 As Long
   Dim R2        As Long, G2 As Long, b2 As Long
   Dim dR        As Long, dG As Long, dB As Long

   Dim Scan      As Long
   Dim i         As Long, j As Long, k As Long
   Dim jIn       As Long
   Dim iEnd      As Long, jEnd As Long
   Dim Offset    As Long

   Dim lQuad     As Long
   Dim AngleDiag As Single
   Dim AngleComp As Single

   Dim g         As Long
   Dim luSin     As Long, luCos As Long
 
   If (Width > 0 And Height > 0) Then

'     Matthew R. Usner - when angle is >= 91 and <= 270, the
'     colors invert in MiddleOut mode.  This corrects for that.
      If bMOut And Angle >= 91 And Angle <= 270 Then
         g = Color1
         Color1 = Color2
         Color2 = g
      End If

'     -- Right-hand [+] (ox=0º)
      Angle = -Angle + 90

'     -- Normalize to [0º;360º]
      Angle = Angle Mod 360
      If (Angle < 0) Then
         Angle = 360 + Angle
      End If

'     -- Get quadrant (0 - 3)
      lQuad = Angle \ 90

'     -- Normalize to [0º;90º]
        Angle = Angle Mod 90

'     -- Calc. gradient length ('distance')
      If (lQuad Mod 2 = 0) Then
         AngleDiag = Atn(Width / Height) * TO_DEG
      Else
         AngleDiag = Atn(Height / Width) * TO_DEG
      End If
      AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
      Angle = Angle * TO_RAD
      g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

'     -- Decompose colors
      If (lQuad > 1) Then
         lClr = Color1
         Color1 = Color2
         Color2 = lClr
      End If
      R1 = (Color1 And &HFF&)
      G1 = (Color1 And &HFF00&) \ 256
      b1 = (Color1 And &HFF0000) \ 65536
      R2 = (Color2 And &HFF&)
      G2 = (Color2 And &HFF00&) \ 256
      b2 = (Color2 And &HFF0000) \ 65536

'     -- Get color distances
      dR = R2 - R1
      dG = G2 - G1
      dB = b2 - b1

'     -- Size gradient-colors array
      ReDim lGrad(0 To g - 1)
      ReDim lGrad2(0 To g - 1)

'     -- Calculate gradient-colors
      iEnd = g - 1
      If (iEnd = 0) Then
'        -- Special case (1-pixel wide gradient)
         lGrad2(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
         For i = 0 To iEnd
            lGrad2(i) = b1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
         Next i
      End If

'     'if block' added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
      If bMOut Then
         k = 0
         For i = 0 To iEnd Step 2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
         For i = iEnd - 1 To 1 Step -2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
      Else
         For i = 0 To iEnd
            lGrad(i) = lGrad2(i)
         Next i
      End If

'     -- Size DIB array
      ReDim lBits(Width * Height - 1) As Long
      iEnd = Width - 1
      jEnd = Height - 1
      Scan = Width

'     -- Render gradient DIB
      Select Case lQuad

         Case 0, 2
            luSin = Sin(Angle) * INT_ROT
            luCos = Cos(Angle) * INT_ROT
            Offset = 0
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset + Scan
            Next j

         Case 1, 3
            luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
            luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
            Offset = jEnd * Scan
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset - Scan
            Next j

      End Select

'     -- Define DIB header
      With uBIH
         .biSize = 40
         .biPlanes = 1
         .biBitCount = 32
         .biWidth = Width
         .biHeight = Height
      End With

'     -- Paint it!
      Call StretchDIBits(hDC, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

    End If

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Properties >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties to the default constants.                       *
'*************************************************************************

   Set m_Icon = Nothing
   Set m_CaptionFont = Ambient.Font
   m_HeaderAngle = m_def_HeaderAngle
   m_BackAngle = m_def_BackAngle
   m_HeaderColor2 = m_def_HeaderColor2
   m_HeaderColor1 = m_def_HeaderColor1
   m_BackColor2 = m_def_BackColor2
   m_BackColor1 = m_def_BackColor1
   m_BorderColor = m_def_BorderColor
   m_CaptionColor = m_def_CaptionColor
   m_Caption = m_def_Caption
   m_Alignment = vbLeftJustify
   m_HeaderHeight = m_def_hHeight
   m_Enabled = m_def_Enabled
   m_BorderWidth = m_def_BorderWidth
   m_BackMiddleOut = m_def_BackMiddleOut
   m_HeaderMiddleOut = m_def_HeaderMiddleOut
   m_HeaderVisible = m_def_HeaderVisible
   m_CurveTopLeft = m_def_CurveTopLeft
   m_CurveTopRight = m_def_CurveTopRight
   m_CurveBottomLeft = m_def_CurveBottomLeft
   m_CurveBottomRight = m_def_CurveBottomRight

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* read properties in the property bag.                                  *
'*************************************************************************

   With PropBag
      Set m_Icon = .ReadProperty("HeaderIcon", Nothing)
      Set m_CaptionFont = .ReadProperty("CaptionFont", Ambient.Font)
      m_Iconsize = .ReadProperty("IconSize", m_def_Iconsize)
      m_HeaderAngle = .ReadProperty("HeaderAngle", m_def_HeaderAngle)
      m_BackAngle = .ReadProperty("BackAngle", m_def_BackAngle)
      m_HeaderColor2 = .ReadProperty("HeaderColor2", m_def_HeaderColor2)
      m_HeaderColor1 = .ReadProperty("HeaderColor1", m_def_HeaderColor1)
      m_BackColor2 = .ReadProperty("BackColor2", m_def_BackColor2)
      m_BackColor1 = .ReadProperty("BackColor1", m_def_BackColor1)
      m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
      m_CaptionColor = .ReadProperty("CaptionColor", m_def_CaptionColor)
      m_Caption = .ReadProperty("Caption", m_def_Caption)
      m_Alignment = .ReadProperty("CaptionAlignment", m_def_Align)
      m_HeaderHeight = .ReadProperty("HeaderHeight", m_def_hHeight)
      m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
      m_BorderWidth = .ReadProperty("BorderWidth", m_def_BorderWidth)
      m_BackMiddleOut = .ReadProperty("BackMiddleOut", m_def_BackMiddleOut)
      m_HeaderMiddleOut = .ReadProperty("HeaderMiddleOut", m_def_HeaderMiddleOut)
      m_HeaderVisible = .ReadProperty("HeaderVisible", m_def_HeaderVisible)
      m_CurveTopLeft = .ReadProperty("CurveTopLeft", m_def_CurveTopLeft)
      m_CurveTopRight = .ReadProperty("CurveTopRight", m_def_CurveTopRight)
      m_CurveBottomLeft = .ReadProperty("CurveBottomLeft", m_def_CurveBottomLeft)
      m_CurveBottomRight = .ReadProperty("CurveBottomRight", m_def_CurveBottomRight)
   End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write the properties in the property bag.                             *
'*************************************************************************

   With PropBag
      .WriteProperty "HeaderAngle", m_HeaderAngle, m_def_HeaderAngle
      .WriteProperty "BackAngle", m_BackAngle, m_def_BackAngle
      .WriteProperty "IconSize", m_Iconsize, m_def_Iconsize
      .WriteProperty "HeaderColor2", m_HeaderColor2, m_def_HeaderColor2
      .WriteProperty "HeaderColor1", m_HeaderColor1, m_def_HeaderColor1
      .WriteProperty "BackColor2", m_BackColor2, m_def_BackColor2
      .WriteProperty "BackColor1", m_BackColor1, m_def_BackColor1
      .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
      .WriteProperty "CaptionColor", m_CaptionColor, m_def_CaptionColor
      .WriteProperty "Caption", m_Caption, m_def_Caption
      .WriteProperty "CaptionAlignment", m_Alignment, vbLeftJustify
      .WriteProperty "HeaderHeight", m_HeaderHeight, m_def_hHeight
      .WriteProperty "CaptionFont", m_CaptionFont, Ambient.Font
      .WriteProperty "HeaderIcon", m_Icon, Nothing
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "BorderWidth", m_BorderWidth, m_def_BorderWidth
      .WriteProperty "BackMiddleOut", m_BackMiddleOut, m_def_BackMiddleOut
      .WriteProperty "HeaderMiddleOut", m_HeaderMiddleOut, m_def_HeaderMiddleOut
      .WriteProperty "HeaderVisible", m_HeaderVisible, m_def_HeaderVisible
      .WriteProperty "CurveTopLeft", m_CurveTopLeft, m_def_CurveTopLeft
      .WriteProperty "CurveTopRight", m_CurveTopRight, m_def_CurveTopRight
      .WriteProperty "CurveBottomLeft", m_CurveBottomLeft, m_def_CurveBottomLeft
      .WriteProperty "CurveBottomRight", m_CurveBottomRight, m_def_CurveBottomRight
   End With

End Sub

Public Property Get HeaderVisible() As Boolean
   HeaderVisible = m_HeaderVisible
End Property

Public Property Let HeaderVisible(ByVal New_HeaderVisible As Boolean)
   m_HeaderVisible = New_HeaderVisible
   PropertyChanged "HeaderVisible"
   RedrawControl
End Property

Public Property Get BackMiddleOut() As Boolean
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   PropertyChanged "BackMiddleOut"
   RedrawControl
End Property

Public Property Get HeaderMiddleOut() As Boolean
   HeaderMiddleOut = m_HeaderMiddleOut
End Property

Public Property Let HeaderMiddleOut(ByVal New_HeaderMiddleOut As Boolean)
   m_HeaderMiddleOut = New_HeaderMiddleOut
   PropertyChanged "HeaderMiddleOut"
   RedrawControl
End Property

Public Property Get HeaderAngle() As Single
   HeaderAngle = m_HeaderAngle
End Property

Public Property Let HeaderAngle(ByVal New_HeaderAngle As Single)
'  do some bounds checking.
   If New_HeaderAngle > 360 Then
      New_HeaderAngle = 360
   ElseIf New_HeaderAngle < 0 Then
      New_HeaderAngle = 0
   End If
   m_HeaderAngle = New_HeaderAngle
   PropertyChanged "HeaderAngle"
   RedrawControl
End Property

Public Property Get BackAngle() As Single
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Single)
'  do some bounds checking.
   If New_BackAngle > 360 Then
      New_BackAngle = 360
   ElseIf New_BackAngle < 0 Then
      New_BackAngle = 0
   End If
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   RedrawControl
End Property

Public Property Get BackColor1() As OLE_COLOR
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   PropertyChanged "BackColor1"
   RedrawControl
End Property

Public Property Get BackColor2() As OLE_COLOR
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   PropertyChanged "BackColor2"
   RedrawControl
End Property

Public Property Get BorderWidth() As Integer
   BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
   m_BorderWidth = New_BorderWidth
   PropertyChanged "BorderWidth"
   RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   PropertyChanged "BorderColor"
   RedrawControl
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   RedrawControl
End Property

Public Property Get HeaderColor1() As OLE_COLOR
   HeaderColor1 = m_HeaderColor1
End Property

Public Property Let HeaderColor1(ByVal New_HeaderColor1 As OLE_COLOR)
   m_HeaderColor1 = New_HeaderColor1
   PropertyChanged "HeaderColor1"
   RedrawControl
End Property

Public Property Get HeaderColor2() As OLE_COLOR
   HeaderColor2 = m_HeaderColor2
End Property

Public Property Let HeaderColor2(ByVal New_HeaderColor2 As OLE_COLOR)
   m_HeaderColor2 = New_HeaderColor2
   PropertyChanged "HeaderColor2"
   RedrawControl
End Property

Public Property Get IconSize() As IconSizeEnum
   IconSize = m_Iconsize
End Property

Public Property Let IconSize(ByVal New_IconSize As IconSizeEnum)
   m_Iconsize = New_IconSize
   PropertyChanged "IconSize"
   RedrawControl
End Property

Public Property Get CaptionColor() As OLE_COLOR
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
   m_CaptionColor = New_CaptionColor
   PropertyChanged "CaptionColor"
   RedrawControl
End Property

Public Property Get HeaderHeight() As Long
   HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal vNewHeight As Long)
   m_HeaderHeight = vNewHeight
   PropertyChanged "HeaderHeight"
   RedrawControl
End Property

Public Property Get CaptionFont() As Font
   Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal vNewCaptionFont As Font)
   Set m_CaptionFont = vNewCaptionFont
   PropertyChanged "CaptionFont"
   RedrawControl
End Property

Public Property Get CaptionAlignment() As AlignmentConstants
   CaptionAlignment = m_Alignment
End Property

Public Property Let CaptionAlignment(ByVal vNewAlignment As AlignmentConstants)
   m_Alignment = vNewAlignment
   PropertyChanged "CaptionAlignment"
   RedrawControl
End Property

Public Property Get HeaderIcon() As Picture
   Set HeaderIcon = m_Icon
End Property

Public Property Set HeaderIcon(ByVal vNewIcon As Picture)
   Set m_Icon = vNewIcon
   PropertyChanged "HeaderIcon"
   RedrawControl
End Property

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get CurveTopLeft() As Long
   CurveTopLeft = m_CurveTopLeft
End Property

Public Property Let CurveTopLeft(ByVal New_CurveTopLeft As Long)
   m_CurveTopLeft = New_CurveTopLeft
   PropertyChanged "CurveTopLeft"
   RedrawControl
End Property

Public Property Get CurveTopRight() As Long
   CurveTopRight = m_CurveTopRight
End Property

Public Property Let CurveTopRight(ByVal New_CurveTopRight As Long)
   m_CurveTopRight = New_CurveTopRight
   PropertyChanged "CurveTopRight"
   RedrawControl
End Property

Public Property Get CurveBottomLeft() As Long
   CurveBottomLeft = m_CurveBottomLeft
End Property

Public Property Let CurveBottomLeft(ByVal New_CurveBottomLeft As Long)
   m_CurveBottomLeft = New_CurveBottomLeft
   PropertyChanged "CurveBottomLeft"
   RedrawControl
End Property

Public Property Get CurveBottomRight() As Long
   CurveBottomRight = m_CurveBottomRight
End Property

Public Property Let CurveBottomRight(ByVal New_CurveBottomRight As Long)
   m_CurveBottomRight = New_CurveBottomRight
   PropertyChanged "CurveBottomRight"
   RedrawControl
End Property
