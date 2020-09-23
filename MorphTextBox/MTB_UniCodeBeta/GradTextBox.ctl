VERSION 5.00
Begin VB.UserControl MorphTextBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   122
   ToolboxBitmap   =   "GradTextBox.ctx":0000
End
Attribute VB_Name = "MorphTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'* MorphTextbox v1.20 - Owner-drawn single-line textbox usercontrol.     *
'* Written May, 2005, by Matthew R. Usner of Planet Source Code.         *
'* Copyright ©2005, Matthew R. Usner, All Rights Reserved.               *
'*************************************************************************
'* Feel free to use this control in your programs, provided ALL credits  *
'* remain intact.  Only dishonorable thieves download code that REAL pro-*
'* grammers work hard to write and freely share with their programming   *
'* peers, then remove the comments and claim that they wrote the code.   *
'*************************************************************************

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<UNICODE BETA>>>>>>>>>>>>>>>>>>>>>>>>>>>
' Note to all: Pietro Cecchi was gracious enough to add some code (and Steve McMahon's clipboard class) so that
' any unicode characters can be displayed / entered.  I would appreciate some feedback from PSC members who have
' Unicode alphabets/keyboards so I can see how this works.  Just remember this is a BETA version and not to expect
' perfect operation.  This is a 'test run' so to speak.  A MILLION THANKS to Pietro Cecchi for spending the time to
' incorporate this.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'  This control represents my initial efforts toward providing a more attractive replacement for
'  VB's boring single-line textbox control.  I've attempted to reproduce the common aspects
'  of standard textbox usage: left-right scrolling of text, selection of text via mouse drag,
'  double-click and keyboard (Shift and Ctrl), cursor movement, etc.  Background can be a
'  user-defined gradient or a bitmap.  Selected text is highlighted with a user-definable
'  (angle, colors, middle-out) gradient.  Border width and color are also user-definable.
'  The Locked and PasswordChar properties are replicated. The .SelStart and .SelLength
'  properties can be defined programmatically.  A SelectOnFocus property allows automatic
'  selection of all text when the textbox receives the focus (e.g., good for a password field).
'  There is no built-in menu that pops up on a right mouse click.  I consider the default VB
'  textbox popup menu an annoyance because I'd rather do my own custom ones.

'  I will update this control as bug reports come in (hopefully not too many...more...)
'  Comments and votes welcome.

'  Credits and Thanks to:
'  Pietro Cecchi - Unicode routines; also Steve McMahon for the clipboard class.
'  Carles P.V. - for his "any angle gradient" routine.  I modified it for middle-out gradients.
'  Jim Jose - for basis of border and text drawing code.
'  Gary ("Phantom Man") - for providing the api-caret contribution.  Thank you so much!  Nice to have
'  some of the top coders around helping me to learn.
'  Dana Seaman - for providing helpful advice and code to expand keyboard text-editing capability.
'  James Crowley - for the TrueType font detection code.
'  XeRoX - for suggestions and debugging.
'  Riccardo Cohen - for suggestions, debugging, and "triple-click select all text" code.
'  Bill Gates - for making my career simultaneously enjoyable and a living hell.

'  History:
'  06/01/2005: - v1.00 - Initial submission to Planet Source Code.
'  06/02/2005: - v1.01 - Added SelectOnFocus property.
'  06/03/2005: - v1.02 - Phantom Man's API caret addition.
'  06/04/2005: - v1.04 - Added border curvature property (thanks Paul Turcksin for idea)
'                        Expanded keyboard text editing capability (Thanks Dana Seaman for snippet!)
'  06/05/2005: - v1.06 - Tweaked DblClick to move caret to end of selected text.
'                        Prohibited copying of password-enabled textbox contents to clipboard.
'                        Added Ctrl-Insert, Shift-Insert and Shift-Delete keys. (Thanks MArio Flores G)
'                        Added a couple RaiseEvents that I missed originally.
'  06/06/2005: - v1.06 - Resubmitted with MArio Flores G's tweaks to only redraw the control in drag
'                        select mode when the selection changes.  This solves the 100% cpu usage problem
'                        noticed by Heriberto and Zin.  Thanks to all three.
'  06/07/2005: - v1.07 - Added SelGradHeight property; debugged issues found by XeRoX.
'  06/09/2005: - v1.08 - Added TrueType font detection to ensure proper caret positioning when
'                        non-TrueType, Bold fonts are used (still have a ways to go; not all fonts
'                        work properly, but more do now).
'  06/10/2005: - v1.09 - Fixed a minor bug, added CaretHeight and SelTextColor properties.
'  06/12/2005: - v1.10 - Added Changed property; resolved minor issues.
'  06/14/2005: - v1.11 - Added Ctrl+Z Undo capability.  Last 10 text changes can be undone.
'  06/20/2005: - v1.12 - Rewrote text mapping to allow more accurate caret positioning.  Thanks LaVolpe!
'  06/23/2005: - v1.13 - Corrected a minor glitch concerning selected text and LostFocus/GotFocus.
'  06/30/2005: - v1.14 - Added Riccardo Cohen's "Triple-Click" select all text capability.
'  08 Apr 2007 - v1.15 - Added .ValidateControl property so coder can check text when control loses focus.

Option Explicit

Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
'Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextA Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long

'IsWinXPPlus is now a boolean variable
'determined into UserControl_Initialize event
Private IsWinXPPlus As Boolean
Private Const VER_PLATFORM_WIN32_NT  As Long = 2
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long
   dwMajorVersion       As Long
   dwMinorVersion       As Long
   dwBuildNumber        As Long
   dwPlatformID         As Long
   szCSDVersion         As String * 128 ' Maintenance string
End Type


Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Const DT_CALCRECT As Long = &H400
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_LEFT As Long = &H0
Private Const DT_NOCLIP As Long = &H100         ' ignores right edge or rectangle when drawing
'added by Pietro Cecchi, applied into DrawText
Private Const DT_RTLREADING As Long = &H20000
'NOTE: DT_RTLREADING Layout in right to left reading order for
'bi-directional text when the font selected into the hdc is a
'Hebrew or Arabic font. The default reading order for all text
'is left to right.

'  declares for determining if font is TrueType or not.
Private Const TMPF_TRUETYPE = &H4
Private Const LF_FACESIZE = 32
Private Type LOGFONT
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
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type
Private Type TEXTMETRIC
    tmHeight As Long             ' height (ascent + descent) of characters.
    tmAscent As Long             ' ascent (units above the base line) of characters.
    tmDescent As Long            ' descent (units below the base line) of characters.
    tmInternalLeading As Long    ' amount of leading (space) inside the bounds set by tmHeight member.
    tmExternalLeading As Long    ' amount of extra leading (space) that the application adds between rows.
    tmAveCharWidth As Long       ' ave. width of characters in font; gen. defined as width of the letter x.
    tmMaxCharWidth As Long       ' width of the widest character in the font.
    tmWeight As Long             ' weight of the font.  (400=reg., 700=bold)
    tmOverhang As Long           ' extra width per string that may be added to some synthesized fonts.
    tmDigitizedAspectX As Long   ' horizontal aspect of the device for which the font was designed.
    tmDigitizedAspectY As Long   ' vertical aspect of the device for which the font was designed.
    tmFirstChar As Byte          ' value of the first character defined in the font.
    tmLastChar As Byte           ' value of the last character defined in the font.
    tmDefaultChar As Byte        ' value of the character to be substituted for characters not in the font.
    tmBreakChar As Byte          ' value of the character used to define word breaks for text justification.
    tmItalic As Byte             ' an italic font if it is nonzero.
    tmUnderlined As Byte         ' an underlined font if it is nonzero.
    tmStruckOut As Byte          ' a strikeout font if it is nonzero.
    tmPitchAndFamily As Byte     ' info about the pitch, the technology, and the family of a physical font.
    tmCharSet As Byte            ' the character set of the font.
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long

'  declares for Carles P.V.'s gradient paint routine.
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

'  rectangle structure for API drawing of text onto control.
Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

'  structure that will hold the display and selected ranges of the text in the textbox.
Private Type TextRange
   FirstCharacter As Long
   LastCharacter  As Long
End Type
Private DisplayRange  As TextRange        ' overall text display range.
Private SelectedRange As TextRange        ' selected text range.

' declarations for a fixed-size stack that allows control to undo last 10 changes to text.
Private Const MAX_UNDO_STACKSIZE As Long = 11    ' max - 1 used.
Private Type UndoStackElementType
   uText          As String
   uCursorPos     As Long
   uDisplayRange  As TextRange
   uSelectedRange As TextRange
End Type
Private UndoStack(1 To MAX_UNDO_STACKSIZE) As UndoStackElementType
Private UndoStackPointer As Long
Private Undoing As Boolean

Public Enum HeightValues
   [HeightOfText] = 0
   [HeightOfBox] = 1
End Enum

'  default property values
Private Const m_def_EnterEmulateTab = False
Private Const m_def_CaretHeight = 1
Private Const m_def_SelTextColor = 0
Private Const m_def_SelGradHeight = 1
Private Const m_DEF_Curvature = 0
Private Const m_def_SelectOnFocus = False
Private Const m_def_SelText = ""
Private Const m_def_SelStart = 0
Private Const m_def_SelLength = 0
Private Const m_def_SelColor1 = &H7000
Private Const m_def_SelColor2 = &H50FF40
Private Const m_def_PasswordChar = ""
Private Const m_def_MaxLength = 500
Private Const m_def_Locked = False
Private Const m_def_SideScroll = 5
Private Const m_def_DisabledColor1 = &H808080
Private Const m_def_DisabledColor2 = &HC0C0C0
Private Const m_def_DisabledMiddleOut = True
Private Const m_def_DisabledAngle = 90
Private Const m_def_DisabledTextColor = &H808080
Private Const m_def_DisabledBorderColor = &H808080
Private Const m_def_DisabledBorderWidth = 1
Private Const m_def_Enabled = True
Private Const m_def_Text = ""
Private Const m_def_DefaultTextColor = &H0
Private Const m_def_DefaultColor1 = &H7F90
Private Const m_def_DefaultColor2 = &H60F0FF
Private Const m_def_DefaultMiddleOut = True
Private Const m_def_DefaultAngle = 90
Private Const m_def_DefaultBorderColor = &H0
Private Const m_def_DefaultBorderWidth = 1
Private Const m_def_FocusColor1 = &H907000
Private Const m_def_FocusColor2 = &HFFEF1F
Private Const m_def_FocusMiddleOut = True
Private Const m_def_FocusAngle = 90
Private Const m_def_FocusTextColor = &H0
Private Const m_def_FocusBorderColor = &H0
Private Const m_def_FocusBorderWidth = 2

'  property variables
Private m_EnterEmulateTab     As Boolean      ' when True, Enter key acts like Tab key (for focus shift).
Private m_CaretHeight         As HeightValues ' height of caret (cursor).
Private m_SelTextColor        As OLE_COLOR    ' color of selected text.
Private m_SelGradHeight       As HeightValues ' height of text or height of box (minus border)
Private m_Curvature           As Integer      ' the amount of rounding of the textbox border.
Private m_SelectOnFocus       As Boolean      ' if true, all text selected when textbox gets the focus.
Private m_SelText             As String       ' the selected text.
Private m_SelStart            As Long         ' the start character position of the selected text.
Private m_SelLength           As Long         ' the length of the selected text.
Private m_SelColor1           As OLE_COLOR    ' selected text first gradient color.
Private m_SelColor2           As OLE_COLOR    ' selected text second gradient color.
Private m_Picture             As Picture      ' if set, supercedes regular gradient background.
Private m_PasswordChar        As String       ' when set, all typed chars appear as this character.
Private m_MaxLength           As Long         ' when 0, no limit to number of characters in textbox.
Private m_Locked              As Boolean      ' when true, text cannot be changed.
Private m_SideScroll          As Long         ' # chars to scroll when cursor passes end of textbox display.
Private m_DisabledColor1      As OLE_COLOR    ' first gradient color when textbox is disabled.
Private m_DisabledColor2      As OLE_COLOR    ' second gradient color when textbox is disabled.
Private m_DisabledMiddleOut   As Boolean      ' middle-out gradient display flag when textbox is disabled.
Private m_DisabledAngle       As Single       ' gradient angle when textbox is disabled.
Private m_DisabledTextColor   As OLE_COLOR    ' text color when textbox is disabled.
Private m_DisabledBorderColor As OLE_COLOR    ' border color when textbox is disabled.
Private m_DisabledBorderWidth As Integer      ' border width when textbox is disabled.
Private m_Enabled             As Boolean      ' control enabled flag.
Private m_Font                As Font         ' the font to write the text with.
Private m_Text                As String       ' the text contents of the control.
Private m_DefaultTextColor    As OLE_COLOR    ' text color when textbox is enabled, no focus.
Private m_DefaultColor1       As OLE_COLOR    ' first gradient color when textbox is enabled, no focus.
Private m_DefaultColor2       As OLE_COLOR    ' second gradient color when textbox is enabled, no focus.
Private m_DefaultMiddleOut    As Boolean      ' middle-out gradient flag when textbox is enabled, no focus.
Private m_DefaultAngle        As Single       ' gradient angle when textbox is enabled, no focus.
Private m_DefaultBorderColor  As OLE_COLOR    ' border color when textbox is enabled, no focus.
Private m_DefaultBorderWidth  As Integer      ' border width when textbox is enabled, no focus.
Private m_FocusColor1         As OLE_COLOR    ' first gradient color when textbox is enabled, has focus.
Private m_FocusColor2         As OLE_COLOR    ' second gradient color when textbox is enabled, has focus.
Private m_FocusMiddleOut      As Boolean      ' middle-out gradient flag when textbox is enabled, has focus.
Private m_FocusAngle          As Single       ' gradient angle when textbox is enabled, has focus.
Private m_FocusTextColor      As OLE_COLOR    ' text color when textbox is enabled, has focus.
Private m_FocusBorderColor    As OLE_COLOR    ' border color when textbox is enabled, has focus.
Private m_FocusBorderWidth    As Integer      ' border width when textbox is enabled, has focus.

'  event declarations
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Change()
Event ValidateControl()

' miscellaneous control variables
Private Const BACKSPACE   As Integer = 8       ' editing key.
Private HasFocus          As Boolean           ' reflects focus status of textbox.
Private CursorPos         As Long              ' character position of the cursor line.
Private MouseIsDown       As Boolean           ' for selecting text.
Private CharMap()         As Long              ' holds the X coordinate of left edge of each char in textbox.
Private WordMap()         As Long              ' byte position of the first character in each word of text.
Private WordCount         As Long              ' the number of separate words in the text.
Private Clearance         As Long              ' distance from left edge of textbox to start drawing text.
Private SelectModeActive  As Boolean           ' lets control know text is being selected (shift key down).
Private CharactersMapped  As Boolean           ' lets control know positions of each character are known.
Private SelectOnFocusFlag As Boolean           ' for handling the intricacies of SelectOnFocus property.
Private RightClickFlag    As Boolean           ' so I can discard right mouse button double-clicks.
Private LastSelectedChar  As Long              ' for redraw only when mouse selection changed.
Private TrueTypeFont      As Boolean           ' if true, current font is a TrueType font.
Private PreviousText      As String            ' for comparison; part of Change() event.

'  for riccardo's triple-click functionality
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private bDoubleClicked    As Boolean           ' flag indicating a double-click has taken place.
Private bTripleClicked    As Boolean           ' flag indicating a triple-click has taken place.
Private lFirstTime        As Long              ' the tick count after a double-click.
Private lSecondTime       As Long              ' the tick count after the third click.
Private lTotalTime        As Long              ' elapsed ticks between doubleclick and third click.
Private lDoubleClickTime  As Long              ' api-retrieved system double-click time, in ms.


Private Function TextWidthU(ByVal hDC As Long, sString As String) As Long

'*************************************************************************
'* a better altenative to the VB method .TextWidth.  Thanks LaVolpe!     *
'*************************************************************************

   Dim Flags    As Long
   Dim TextRect As RECT
   SetRect TextRect, 0, 0, 0, 0
   Flags = DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_LEFT
   DrawText hDC, sString, -1, TextRect, Flags
   TextWidthU = TextRect.Right + 1

End Function

Private Sub DrawText(ByVal ishDC As Long, ByVal isText As String, ByVal isLen As Long, IsRect As RECT, ByVal areFlags As Integer)
    'NOTE: DT_RTLREADING Layout in right to left reading order for
    'bi-directional text when the font selected into the hdc is a
    'Hebrew or Arabic font. The default reading order for all text
    'is left to right.
    'NOTE: since DT_RTLREADING acts automatically
    'Hebrew or Arabic font, this flag can always be set (?)
    
    'Draw text
    If IsWinXPPlus Then 'UNICODE
       DrawTextW ishDC, StrPtr(isText), isLen, IsRect, areFlags Or DT_RTLREADING
    Else 'ANSI
       DrawTextA ishDC, isText, isLen, IsRect, areFlags Or DT_RTLREADING
    End If
End Sub

'these 2 functions manage the clipboard either in ANSI either in UNICODE
'add them in the bottom of your control code
Private Function UNI_ANSI_ClipBoard_SetText(ByVal isText As String) As Boolean
   'returns true if OK
   Dim retVal As Boolean

   If IsWinXPPlus Then 'UNICODE
      Dim m_cClip As New clsClipboard
      Dim bbData() As Byte
      
      'Get access to the clipboard
      m_cClip.ClipboardOpen (UserControl.hwnd)

      'must be zero terminated
      If Right(isText, 1) <> Chr(0) Then
         bbData = isText & Chr(0)
      Else
         bbData = isText
      End If
         
      'Become the clipboard owner
      m_cClip.ClearClipboard
      'operation
      retVal = m_cClip.SetBinaryData(13, bbData())     '13 unicode
      m_cClip.ClipboardClose
   Else 'ANSI
      Clipboard.Clear
      Clipboard.SetText isText, vbCFText
      retVal = True
   End If

   UNI_ANSI_ClipBoard_SetText = retVal

End Function

Private Function UNI_ANSI_ClipBoard_GetText(ByRef isTextOut As String) As Boolean
   'returns true if OK
   Dim ssTextOut As String
   Dim retVal As Boolean

   If IsWinXPPlus Then 'UNICODE
      Dim bbData() As Byte
      Dim pos As Integer, ss As String
      Dim m_cClip As New clsClipboard

      m_cClip.ClipboardOpen (UserControl.hwnd)

      'get unicode
      retVal = m_cClip.GetBinaryData(13, bbData())   '13 unicode

      ss = bbData()
      pos = InStr(1, ss, Chr(0))
      If pos > 0 Then
         ss = Left(ss, pos - 1)
      Else
         ss = ""
      End If
   
      'pass text
      isTextOut = ss
   
      'release clipboard class
      m_cClip.ClipboardClose
   Else  'ANSI
      isTextOut = Clipboard.GetText(vbCFText)
      retVal = True
   End If
         
   UNI_ANSI_ClipBoard_GetText = retVal

End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<< Event-Handling Routines >>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_Initialize()

'*************************************************************************
'* the first event in the textbox life cycle.                            *
'*************************************************************************

   Dim osv As OSVERSIONINFO

   CursorPos = 1
   DisplayRange.FirstCharacter = 1
   SelectModeActive = False

   'OS version
   osv.dwOSVersionInfoSize = Len(osv)
   GetVersionEx osv
   
   '1) is NT 2) is Xp or above
   IsWinXPPlus = ((osv.dwPlatformID And VER_PLATFORM_WIN32_NT) = _
                 VER_PLATFORM_WIN32_NT) And _
                 ((osv.dwMajorVersion > 5) Or _
                 (osv.dwMajorVersion = 5) And (osv.dwMinorVersion >= 1))

End Sub

Private Sub UserControl_Show()

'*************************************************************************
'* sets up the text mapping and displays text for the first time.        *
'*************************************************************************

'  initialize the word map array.
   If m_MaxLength > 0 Then
      ReDim WordMap(1 To m_MaxLength)
   Else
'     arbitrary amount; for when MaxLength = 0. '+ 2' ensures at least a couple of elements.
      ReDim WordMap(1 To Len(m_Text) + 2)
   End If

'  set the Change event generation comparison variable.
   PreviousText = m_Text

'  display the textbox and map the characters for the first time.
   RedrawControl
   MapCharacters DisplayRange

'  initialize the undo stack with the beginning textbox info.
   UndoStackPointer = 0
   AddToStack

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* handles cursor movement, text selection and keyboard editing.         *
'*************************************************************************

   Dim ShiftDown As Boolean         ' shift status flag.
   Dim CtrlDown  As Boolean         ' ctrl key status flag.
   Dim i         As Long            ' loop variable.
   Dim TextLen   As Long            ' for optimizing.

'  if the mouse button is down, then ignore.  Do not return event.
   If MouseIsDown Then
      Exit Sub
   End If

'  if the EnterEmulateTab property has been set to True, the Enter key acts like the
'  Tab key, sending the focus to the next control in the Tab stop list.  Notice
'  that if you want to use this property, you must set it in ALL MorphTextBox
'  controls in the form.  This feature is good for data entry.
   If m_EnterEmulateTab Then
      If KeyCode = 13 Then
         SendKeys "{TAB}"
         Exit Sub        ' don't pass along KeyDown event.
      End If
   End If

'  determine shift key status.  This lets us know if user is selecting text.
   ShiftDown = (Shift And vbShiftMask) > 0
   If ShiftDown Then
      If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or _
         KeyCode = vbKeyDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then
         If SelectModeActive Then
'           .LastCharacter can actually be less than .FirstCharacter.  (e.g. when you select text
'           from right to left).  That's why you see swapping of the two values in certain places.
            SelectedRange.LastCharacter = CursorPos - 1
            If SelectedRange.LastCharacter < 1 Then
               SelectedRange.LastCharacter = 1
            End If
         End If
         If Not SelectModeActive Then
            SelectModeActive = True
            SelectedRange.FirstCharacter = CursorPos
         End If
      End If
   Else
'     set SelectModeActive to False, but keep any SelectedRange values.  That way when we
'     press delete, backspace, or an alphanumeric, appropriate actions will still take place.
      SelectModeActive = False
   End If

'  determine Ctrl key status.
   CtrlDown = (Shift And vbCtrlMask) > 0

'  following functions (except Ctrl+Z) kindly provided by Dana Seaman.  I made some functional
'  modifications to his snippet to fully integrate into textbox.  Thanks for the great ideas Dana.
   If (CtrlDown) And Not (m_Locked) Then
      Select Case KeyCode
         Case vbKeyZ            ' Ctrl+Z undo function
            If UndoStackPointer > 1 Then
               UndoLastAction
            End If
         Case vbKeyA            ' not a standard vb textbox function.
'           select entire text on Ctrl+A.
            SelectedRange.FirstCharacter = 1
            SelectedRange.LastCharacter = Len(m_Text) + 1
            SelectModeActive = True
            MoveCursorToEndOfText
         Case vbKeyC, vbKeyX, vbKeyInsert
'           Ctrl+Ins/Ctrl+C = Copy selection to ClipBoard, Ctrl+X = Cut selection to ClipBoard.
            If LenB(m_SelText) Then
'              only allow if password mode is not enabled.  Otherwise user could copy contents of
'              password box (into Notepad, for example), and see the password.  This emulates standard
'              vb textbox behavior.
               If m_PasswordChar = "" Then
        '          Clipboard.Clear
        '          Clipboard.SetText m_SelText
                  UNI_ANSI_ClipBoard_SetText m_SelText
'                 if Ctrl+X was pressed, remove selection.
                  If KeyCode = vbKeyX Then
                     DeleteSelection
                     RedrawControl
                  End If
               End If
            End If
         Case vbKeyV       ' Ctrl+V - paste the clipboard contents into textbox at cursor position.
            PasteClipboardContents
      End Select
   End If

   Select Case KeyCode

      Case vbKeyInsert                            ' Shift-Ins paste from clipboard. (MArio Flores G)
         If ShiftDown Then
            PasteClipboardContents
            RaiseEvent KeyDown(KeyCode, Shift)
            Exit Sub
         End If

      Case vbKeyLeft, vbKeyUp                     ' left or up arrow key.
'        this helps emulate vb textbox functionality - when cursor is at beginning
'        of text and a non-shifted arrow key is pressed, selection mode turns off.
         If CursorPos = 1 And Not ShiftDown Then
            SelectModeActive = False
            RedrawControl
            RaiseEvent KeyDown(KeyCode, Shift)
            Exit Sub
         End If

         If CursorPos > 1 Then
'           if control key is not down deal with text one character at a time.
            If Not CtrlDown Then
               CursorPos = CursorPos - 1
               If SelectModeActive Then
                  SelectedRange.LastCharacter = CursorPos
               End If
'              if the cursor goes left past the clearance, we must scroll the
'              text left by the number of characters in the SideScroll property.
               If CharMap(CursorPos) < Clearance Then
                  DisplayRange.FirstCharacter = DisplayRange.FirstCharacter - m_SideScroll
                  If DisplayRange.FirstCharacter < 1 Then
                     DisplayRange.FirstCharacter = 1
                  End If
               End If
               MapCharacters DisplayRange
               RedrawControl
            Else
'              if ctrl key is down, move cursor to beginning of previous word (or beginning of text).
'              for password character mode, just move to beginning of text.  That way, the
'              individual word segments in the text are concealed.
               If m_PasswordChar <> "" Then
                  CursorPos = 1
'                 check for shift pressed as well as ctrl key.
                  If SelectModeActive Then
                     SelectedRange.LastCharacter = 1
                  End If
                  CharactersMapped = False
                  DisplayRange.FirstCharacter = 1
                  RedrawControl
               Else
'                 otherwise go to beginning of previous word or beginning of text.
                  For i = WordCount To 1 Step -1
                     If CursorPos > WordMap(i) Then
                        CursorPos = WordMap(i)
                        Exit For
                     End If
                  Next i
'                 this is for when both Ctrl and Shift keys are down while pressing Left (or Up) arrow.
                  If SelectModeActive Then
                     SelectedRange.LastCharacter = CursorPos
                  End If
'                 if the cursor went past the left edge of the textbox,
'                 re-range the text from the cursor position.
                  If CharMap(CursorPos) < Clearance Then
                     CharactersMapped = False
                     DisplayRange.FirstCharacter = CursorPos
                  End If
                  RedrawControl
               End If
            End If
         End If

      Case vbKeyRight, vbKeyDown                  ' right or down arrow key.
'        this helps emulate vb textbox functionality - when cursor is at end
'        of text and a non-shifted arrow key is pressed, selection mode turns off.
         If CursorPos = Len(m_Text) + 1 And Not ShiftDown Then
            SelectModeActive = False
            RedrawControl
            RaiseEvent KeyDown(KeyCode, Shift)
            Exit Sub
         End If

         If CursorPos <= Len(m_Text) Then
'           if control key is not down deal with text one character at a time.
            If Not CtrlDown Then
               CursorPos = CursorPos + 1
'              if selection mode is on (shift is down), adjust the selected range.
               If SelectModeActive Then
                  SelectedRange.LastCharacter = CursorPos
               End If
'              if the cursor goes right past the last displayed character, we must scroll the
'              text right by the number of characters in the SideScroll property.
               If CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
                  CharactersMapped = False
                  DisplayRange.FirstCharacter = DisplayRange.FirstCharacter + m_SideScroll
               End If
               RedrawControl
            Else
'              if ctrl key is down, move cursor to start of next word (or end of text).
'              for password character mode, just move to end of text.  That way the
'              individual word segments in the text are concealed.
               If m_PasswordChar <> "" Then
                  CursorPos = Len(m_Text) + 1
'                 check for shift pressed as well as ctrl key.
                  If SelectModeActive Then
                     SelectedRange.LastCharacter = CursorPos
                  End If
                  CharactersMapped = False
'                 must map password chars, not the actual text, for display.
                  SetTextDisplayRangeRev String(Len(m_Text), Left(m_PasswordChar, 1)), CursorPos
                  RedrawControl
               Else
'                 otherwise, move cursor to beginning of next word (or end of text).
                  For i = 1 To WordCount
                     If CursorPos < WordMap(i) Then
                        CursorPos = WordMap(i)
                        Exit For
                     End If
                  Next i
'                 this is for when both Ctrl and Shift keys are down while pressing right (or down) arrow.
                  If SelectModeActive Then
                     SelectedRange.LastCharacter = CursorPos
                  End If
'                 if the cursor went past the edge of the textbox,
'                 re-range the text from the cursor position backwards.
                  If CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
                     SetTextDisplayRangeRev m_Text, CursorPos
                     CharactersMapped = False
                  End If
                  RedrawControl
               End If
            End If
         End If

      Case vbKeyHome                              ' home key.
'        this helps emulate vb textbox functionality - if already at first
'        character position and text is selected, remove selection.
         If CursorPos = 1 And Not ShiftDown Then
            SelectModeActive = False
            RedrawControl
            RaiseEvent KeyDown(KeyCode, Shift)
            Exit Sub
         End If
         If CursorPos <> 1 Then
'           if we are selecting text, we highlight from the beginning to the cursor position.
            If SelectModeActive Then
'              this will be swapped in the selection gradient paint routine.
'              doing it this way helps me emulate vb textbox behavior more accurately.
               SelectedRange.LastCharacter = 1
            End If
            CursorPos = 1
            DisplayRange.FirstCharacter = 1
            MapCharacters DisplayRange
            RedrawControl
         End If

      Case vbKeyEnd                               ' end key.
'        if cursor is already at end of control, redraw if necessary, raise event, and exit.
         If CursorPos = Len(m_Text) + 1 Then
'           takes care of when text is selected and cursor's at end of text.
            If Not ShiftDown Then
               SelectModeActive = False
               RedrawControl
            End If
            RaiseEvent KeyDown(KeyCode, Shift)
            Exit Sub
         End If
         MoveCursorToEndOfText

      Case vbKeyDelete                            ' delete key.
         If Not m_Locked Then
            If ShiftDown Then                     ' MArio Flores G - Shift+Delete acts like backspace
               ProcessBackspace
            Else
               TextLen = Len(m_Text)
               If TextLen > 0 And CursorPos <= TextLen + 1 Then
'                 if there is a selection made, we delete that.
                  If SelectedRange.FirstCharacter > 0 And SelectedRange.LastCharacter > 0 Then
                     DeleteSelection
                  Else
'                    otherwise we delete just the character at the cursor position.
                     If CursorPos <= TextLen Then
                        DeleteCharacterAtCursorPosition
                     End If
                  End If
                  RedrawControl
                  MapCharacters DisplayRange
               End If
            End If
         End If

      Case BACKSPACE                              ' backspace key.
         ProcessBackspace

   End Select

   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

'*************************************************************************
'* process regular alphanumeric keystrokes into the textbox.             *
'*************************************************************************

'  if the mouse button is down, then ignore.  Do not return event.
   If MouseIsDown Then
      Exit Sub
   End If

'  if invalid character, or textbox is locked, pass along the event and exit.
   If KeyAscii < 32 Or m_Locked Then
      RaiseEvent KeyPress(KeyAscii)
      Exit Sub
   End If

'  if the text is equal to the MaxLength property value, raise the event and exit.
   If Len(m_Text) = m_MaxLength And m_MaxLength > 0 Then
      RaiseEvent KeyPress(KeyAscii)
      Exit Sub
   End If

'  if any text is selected, delete selection; will replace that text with the character typed.
   If Len(m_SelText) > 0 Then
      DeleteSelection
   End If

'  sanity check.
   If m_Text = "" Then
      CursorPos = 1
   End If

'  insert the alphanumeric character at the cursor position and move the cursor right.
   m_Text = Left(m_Text, CursorPos - 1) & Chr(KeyAscii) & Right(m_Text, Len(m_Text) - CursorPos + 1)
   CursorPos = CursorPos + 1
'  map the character positions now that the text has been altered.
   MapCharacters DisplayRange

'  if the typing goes past the right edge, we must scroll the text to the left
'  by a couple of characters to help mimic standard vb textbox behavior.
   If CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
      CharactersMapped = False
      DisplayRange.FirstCharacter = DisplayRange.FirstCharacter + 3
   End If

   RedrawControl
   DoEvents

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* not used, but the event is passed on to the programmer.               *
'*************************************************************************

'  if the mouse button is down, then ignore.  Do not return event.
   If MouseIsDown Then
      Exit Sub
   End If

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_GotFocus()

'*************************************************************************
'* performs any necessary actions when the textbox receives the focus.   *
'*************************************************************************

   HasFocus = True
   If m_SelectOnFocus And (Len(m_Text) > 0) Then
      SelectModeActive = True
      SelectedRange.FirstCharacter = 1
      SelectedRange.LastCharacter = Len(m_Text) + 1
      MoveCursorToEndOfText
   End If
   RedrawControl

End Sub

Private Sub UserControl_LostFocus()

'*************************************************************************
'* sets focus flag, sets selected text to null, and draws box as default.*
'*************************************************************************

   Dim SaveSelectModeActive As Boolean
   Dim SaveSelectedRange    As TextRange
   Dim Save_m_SelStart      As Long
   Dim Save_m_SelLength     As Long
   Dim Save_m_SelText       As String

   HasFocus = False

'  save the current selection mode settings.
   SaveSelectModeActive = SelectModeActive
   SaveSelectedRange.FirstCharacter = SelectedRange.FirstCharacter
   SaveSelectedRange.LastCharacter = SelectedRange.LastCharacter
   Save_m_SelStart = m_SelStart
   Save_m_SelLength = m_SelLength
   Save_m_SelText = m_SelText

'  turn off selection mode and redraw the control.  This redraws the control
'  without the selected text highlight gradient (if one is defined).
   SelectModeActive = False
   SelectedRange.FirstCharacter = 0
   SelectedRange.LastCharacter = 0
   RedrawControl

'  restore the current selection mode settings.  This way selection gradient will be
'  redrawn if focus is re-established using tab key (or .SetFocus method call).
   SelectModeActive = SaveSelectModeActive
   SelectedRange.FirstCharacter = SaveSelectedRange.FirstCharacter
   SelectedRange.LastCharacter = SaveSelectedRange.LastCharacter
   m_SelStart = Save_m_SelStart
   m_SelLength = Save_m_SelLength
   m_SelText = Save_m_SelText

'  no need to display the caret when control has lost the focus.
   HideCaret hwnd

   RaiseEvent ValidateControl ' so programmer can have an opportunity to validate textbox contents.

End Sub

Private Sub UserControl_DblClick()

'*************************************************************************
'* when a word is double-clicked, it is selected.  When a space is       *
'* double-clicked, all the way to beginning of previous word is selected.*
'*************************************************************************

   Dim i             As Long          ' loop variable.
   Dim NonSpaceFound As Boolean       ' flag indicating a non-space has been encountered.

'  if double click event was from right mouse button, exit.  Don't pass event to project.
   If RightClickFlag Then
      RightClickFlag = False
      Exit Sub
   End If

'  for riccardo's triple-click
   bDoubleClicked = True
   bTripleClicked = False
   lDoubleClickTime = GetDoubleClickTime
   lFirstTime = GetTickCount

'  if the character x-position map needs updating (if text typed or deleted, for example), do it.
   If Not CharactersMapped Then
      MapCharacters DisplayRange
   End If

   SelectModeActive = True

'  for password character mode, just select all text.  That way,
'  the individual word segments in the text are concealed.
   If m_PasswordChar <> "" Then
      SelectedRange.FirstCharacter = 1
      SelectedRange.LastCharacter = Len(m_Text) + 1
      MoveCursorToEndOfText
      RedrawControl
      RaiseEvent DblClick
      Exit Sub
   End If

'  for when a non-space character is double-clicked.
   If Mid(m_Text, CursorPos, 1) <> " " Then

'     special case for when cursor's at end and last 1+ characters are spaces.
      If CursorPos = Len(m_Text) + 1 And Right(m_Text, 1) = " " Then
'        find end of word.
         For i = CursorPos - 1 To 1 Step -1
            If Mid(m_Text, i, 1) <> " " Then
               NonSpaceFound = True
               SelectedRange.LastCharacter = i + 1
               CursorPos = i + 1
               Exit For
            End If
         Next i
         If Not NonSpaceFound Then
'           nothing to select, just raise the event and exit.
            RaiseEvent DblClick
            Exit Sub
         End If
'        otherwise, find the beginning of the word.
         AdjustForCursorOutOfRangeLeft
         SelectedRange.FirstCharacter = BeginningOfWord(SelectedRange.LastCharacter - 1)
         RedrawControl
         RaiseEvent DblClick
         Exit Sub
      End If

'     find the beginning of the word.
      SelectedRange.FirstCharacter = BeginningOfWord(CursorPos - 1)
'     find the end of the word.
      SelectedRange.LastCharacter = EndOfWord(CursorPos + 1)
'     place the character at end of selected word to emulate vb textbox.
      CursorPos = SelectedRange.LastCharacter

   Else

'     a space was double-clicked.  loop back to beginning of previous word.
'     find the first non-space character, this is the end of the previous word.
      If CursorPos > 1 Then
         i = CursorPos - 1
      Else
         i = 1
      End If
      If Mid(m_Text, i, 1) <> " " Then
         NonSpaceFound = True
         SelectedRange.LastCharacter = i + 1
         CursorPos = i + 1
      Else
         While Mid(m_Text, i, 1) = " " And (i > 1)
            i = i - 1
            If Mid(m_Text, i, 1) <> " " Then
               NonSpaceFound = True
               SelectedRange.LastCharacter = i + 1
               CursorPos = i + 1
            End If
         Wend
      End If
'     if a non-space character was found, find the beginning of the word.
      If NonSpaceFound Then
         SelectedRange.FirstCharacter = BeginningOfWord(SelectedRange.LastCharacter - 1)
         AdjustForCursorOutOfRangeLeft
      End If

   End If

'  if the cursor went past the edge of the textbox,
'  re-range the text from the cursor position backwards.
   If CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
      SetTextDisplayRangeRev m_Text, CursorPos
      CharactersMapped = False
   End If

   RedrawControl

'  raise the DblClick event for the rest of the project.
   RaiseEvent DblClick

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* sets mousedown flag and places cursor in appropriate position.        *
'* Cursor placement mimics the intrinsic vb textbox method.              *
'*************************************************************************

   Dim i            As Long    ' loop index.
   Dim CharMidPoint As Long    ' the middle point of the clicked character.

'  allow user of control to process a right-click in their own way (pop a menu, etc.).
   If Button = vbRightButton Then
'     for some dumb reason (to me at least), VB fires a double-click event for the
'     right button too.  I use this flag to tell the DblClick routine to ignore.
      RightClickFlag = True
      RaiseEvent MouseDown(Button, Shift, X, Y)
      Exit Sub
   End If

'  set the mousedown and right-click flags.
   MouseIsDown = True
   RightClickFlag = False

'  if the "SelectOnFocus" property = True, when the textbox gets the focus because of a
'  mouse click, you'd think that the GotFocus event should handle it, and all the text
'  would be selected, right?  Wrong... because the event order is MouseDown, THEN GotFocus.
'  The text ends up only being selected to where you clicked.  The solution is to handle
'  the SelectOnFocus action HERE, rather than in the GotFocus routine.

'  if HasFocus = False, the control is getting the focus via mouse click.  The GotFocus
'  event (and EnterFocus event, for that matter) simply has not fired yet.
   If Not HasFocus Then    ' 'HasFocus' will be set in the GotFocus routine.
      If m_SelectOnFocus And (Len(m_Text) > 0) Then
         SelectOnFocusFlag = True
         SelectModeActive = True
'        highlight all the text.
         MoveCursorToEndOfText
         RaiseEvent MouseDown(Button, Shift, X, Y)
         Exit Sub
      End If
   End If

'  riccardo's triple-click
   Select Case bDoubleClicked
      Case True
         lSecondTime = GetTickCount
         lTotalTime = (lSecondTime - lFirstTime)
         If lTotalTime <= lDoubleClickTime Then
            bTripleClicked = True
'           select entire text.
            SelectedRange.FirstCharacter = 1
            SelectedRange.LastCharacter = Len(m_Text) + 1
            SelectModeActive = True
            MoveCursorToEndOfText
            bDoubleClicked = False
            RaiseEvent MouseDown(Button, Shift, X, Y)
            Exit Sub
         End If
      Case False
         bDoubleClicked = False
         bTripleClicked = False
   End Select

'  determine where to place the cursor.
'  step 1.  if the character x-position map needs updating (if text typed or deleted, for example), do it.
   If Not CharactersMapped Then
      MapCharacters DisplayRange
   End If

'  step 2.  Place the cursor.
'  2a.  special case - if text is empty, place cursor at first cursor position.
   If Len(m_Text) = 0 Then
      SetMouseDownShiftStatus Shift
      CursorPos = 1
      RedrawControl
      RaiseEvent MouseDown(Button, Shift, X, Y)
      Exit Sub
   End If

'  2b.  special case - if the mouse is clicked after the last character, place cursor at end.
   If CharMap(Len(m_Text)) < X Then
      SetMouseDownShiftStatus Shift
      CursorPos = Len(m_Text) + 1
      SelectedRange.LastCharacter = CursorPos
      RedrawControl
      RaiseEvent MouseDown(Button, Shift, X, Y)
      Exit Sub
   End If

'  2c.  special case - if the mouse is clicked before the first character, place cursor at beginning.
   If CharMap(1) > X Then
      SetMouseDownShiftStatus Shift
      CursorPos = 1
      SelectedRange.LastCharacter = CursorPos
      RedrawControl
      RaiseEvent MouseDown(Button, Shift, X, Y)
      Exit Sub
   End If

'  2d.  loop through the map and find out upon which character the mouse was clicked.
   For i = 1 To Len(m_Text) - 1        ' change so only loops in display range.
      If (X >= CharMap(i)) And (X < CharMap(i + 1)) Then
         SetMouseDownShiftStatus Shift
         CursorPos = i
         SelectedRange.LastCharacter = CursorPos
         Exit For
      End If
   Next i

'  2e.  now that we know the character, we need to know where in the character the mouse
'  was clicked.  If the mouse was clicked in the left half of the character, that
'  character gets the cursor.  If the mouse was clicked in the right half of the character,
'  the character to the right gets the cursor. (just like the vb textbox.)
   CharMidPoint = (CharMap(CursorPos) + CharMap(CursorPos + 1)) / 2
   If X > CharMidPoint Then
      SetMouseDownShiftStatus Shift
      CursorPos = CursorPos + 1
      SelectedRange.LastCharacter = CursorPos
   End If

   RedrawControl

'  raise the MouseDown event for the rest of the project.
   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* handles mouse drag selection of text.                                 *
'*************************************************************************

   Dim i     As Long          ' loop variable.

'  set the cursor to the familiar textbox i-beam.
   UserControl.MousePointer = vbIbeam

'  if all text has been selected because SelectOnFocus property is True, we
'  set SelectOnFocusFlag until a MouseUp event occurs.  MouseMove screws up
'  the total selection of text (text will only be selected to the character
'  clicked upon).  So until the MouseUp, we pass a MouseMove event and exit.
   If SelectOnFocusFlag Then
      RaiseEvent MouseMove(Button, Shift, X, Y)
      Exit Sub
   End If

'  if there's no text, just pass along the event and exit.
   If Len(m_Text) < 1 Then
      RaiseEvent MouseMove(Button, Shift, X, Y)
      Exit Sub
   End If

'  if the mouse button is down, then initiate or continue drag selection.
   If MouseIsDown And Not bTripleClicked Then ' riccardo's triple-click
      If Not SelectModeActive Then
         SelectModeActive = True
         SelectedRange.FirstCharacter = CursorPos
      Else
'        need to find the character the cursor is over.
'        cursor moved before first displayed character.
         If X < Clearance And CursorPos > 1 Then
            CharactersMapped = False
            DisplayRange.FirstCharacter = DisplayRange.FirstCharacter - m_SideScroll
'           a necessary safety net.
            If DisplayRange.FirstCharacter < 1 Then
               DisplayRange.FirstCharacter = 1
            End If
            ' fix for riccardo 06/27
            CursorPos = DisplayRange.FirstCharacter
            SelectedRange.LastCharacter = CursorPos
            RedrawControl
         ElseIf X < Clearance Then
            SelectedRange.LastCharacter = 1
         ElseIf CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
'           if the mouse goes past the right edge, we must scroll the text to the left
'           by a couple of characters to help mimic intrinsic vb textbox behavior.
            CharactersMapped = False
            DisplayRange.FirstCharacter = DisplayRange.FirstCharacter + 3
            SelectedRange.LastCharacter = CursorPos
            RedrawControl
         ElseIf CharMap(Len(m_Text)) < X Then
'           cursor moved past last character.
            SelectedRange.LastCharacter = Len(m_Text) + 1
         Else
'           cursor is somewhere in the middle of the text.
            For i = 1 To Len(m_Text) + 1
               If (X >= CharMap(i)) And (X < CharMap(i + 1)) Then
                  CursorPos = i
                  SelectedRange.LastCharacter = CursorPos
                  Exit For
               End If
            Next i
         End If
'        cursor moved after last character.
         If CharMap(Len(m_Text)) < X Then
            CursorPos = Len(m_Text) + 1
         End If
'        if text display range has changed, remap displayed text.
         If Not CharactersMapped Then
           MapCharacters DisplayRange
         End If
'        if the selection has changed, repaint the textbox.  This prevents the constant repaint when
'        mouse is down and continual firing of MouseMove event occurs.  Thanks to MArio Flores G.
         If SelectedRange.LastCharacter <> LastSelectedChar Then
            RedrawControl
         End If
'        necessary for on-the-fly MouseMove textbox redraw.
         UserControl.Refresh
'        store the position of last drag-selected character.  MArio Flores G.
         LastSelectedChar = SelectedRange.LastCharacter
      End If
   End If

'  raise the event for the rest of the project.
   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* handles actions when mouse is "unclicked".                            *
'*************************************************************************

'  if it's the right button, just pass the event along to the user and exit.  The
'  vb textbox would pop the popup menu at this point, but I ain't gonna do that.
   If Button = vbRightButton Then
      RaiseEvent MouseUp(Button, Shift, X, Y)
      Exit Sub
   End If

   MouseIsDown = False
   SelectOnFocusFlag = False

'  raise appropriate events for user. Click event order is MouseDown, MouseUp, Click.
   RaiseEvent MouseUp(Button, Shift, X, Y)
   RaiseEvent Click

End Sub

Private Sub UserControl_Resize()

'*************************************************************************
'* just for resizing purposes in design mode.  No need to raise event.   *
'*************************************************************************

   RedrawControl

End Sub

Private Sub UserControl_Terminate()

'*************************************************************************
'* destroys caret object on exit.  Thanks to Gary (Phantom Man).         *
'*************************************************************************

   On Error Resume Next
   DestroyCaret
   On Error GoTo 0

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<< String Manipulation Routines >>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub AddToStack()

'*************************************************************************
'* adds an entry to the undo stack when text is changed.                 *
'*************************************************************************

'  if the stack is full, move elements 2 -> MAX_UNDO_STACKSIZE down 1 position.
   If UndoStackPointer = MAX_UNDO_STACKSIZE Then
      MoveStackDown
   Else
      UndoStackPointer = UndoStackPointer + 1
   End If

'  place the new information on the stack.
   UndoStack(UndoStackPointer).uCursorPos = CursorPos
   UndoStack(UndoStackPointer).uDisplayRange.FirstCharacter = DisplayRange.FirstCharacter
   UndoStack(UndoStackPointer).uDisplayRange.LastCharacter = DisplayRange.LastCharacter
   UndoStack(UndoStackPointer).uSelectedRange.FirstCharacter = SelectedRange.FirstCharacter
   UndoStack(UndoStackPointer).uSelectedRange.LastCharacter = SelectedRange.LastCharacter
   UndoStack(UndoStackPointer).uText = m_Text

End Sub

Private Sub MoveStackDown()

'*************************************************************************
'* moves elements (2 through max_stacksize) down 1 element when stack    *
'* is full.  This frees up the top stack slot for a new entry.           *
'*************************************************************************

   Dim i As Long

   For i = 1 To MAX_UNDO_STACKSIZE - 1
      UndoStack(i) = UndoStack(i + 1)
   Next i

End Sub

Private Sub UndoLastAction()

'*************************************************************************
'* undoes the last text-changing event.                                  *
'*************************************************************************

   UndoStackPointer = UndoStackPointer - 1

'  restore the previous textbox state.
   m_Text = UndoStack(UndoStackPointer).uText
   DisplayRange.FirstCharacter = UndoStack(UndoStackPointer).uDisplayRange.FirstCharacter
   DisplayRange.LastCharacter = UndoStack(UndoStackPointer).uDisplayRange.LastCharacter
   SelectedRange.FirstCharacter = UndoStack(UndoStackPointer).uSelectedRange.FirstCharacter
   SelectedRange.LastCharacter = UndoStack(UndoStackPointer).uSelectedRange.LastCharacter
   If CursorPos > Len(m_Text) + 1 Then
      CursorPos = Len(m_Text) + 1
   ElseIf UndoStackPointer > 1 Then
'     the first undo stack position always holds the state of the textbox as it was
'     when the control was first initialized.  This "If" ensures that the cursor stays
'     in the appropriate place when the last "undo" is performed.  Otherwise, it
'     goes to where it was when control is initialized (first position).
      CursorPos = UndoStack(UndoStackPointer).uCursorPos
   End If

'  remap and display the text.
   MapCharacters DisplayRange
   Undoing = True     ' so the undo itself is not added to stack.
   RedrawControl
   Undoing = False

End Sub

Public Function FontIsTrueType(sFontName As String) As Boolean

'*************************************************************************
'* returns a boolean indicating if font is TrueType or not.  This is     *
'* necessary because non-TrueTypeFonts simulate a Bold characteristic,   *
'* as opposed to TrueType fonts, which actually HAVE Bold.  If the text  *
'* is in a non-TrueType bold font, I need to subtract 1 pixel from the   *
'* width of each character when I'm mapping the text for caret position- *
'* ing purposes.  This routine, and the corresponding declares at top,   *
'* by James Crowley and found on his site www.developerfusion.co.uk.     *
'* Code modified slightly for a self-contained usercontrol by Matt.      *
'*************************************************************************

   Dim lf         As LOGFONT
   Dim tm         As TEXTMETRIC
   Dim OldFont    As Long
   Dim NewFont    As Long
   Dim tmpArray() As Byte
   Dim Dummy      As Long
   Dim i          As Long

'  need to convert font name to byte array.
   tmpArray = StrConv(sFontName & vbNullString, vbFromUnicode)
   For i = 0 To UBound(tmpArray)
      lf.lfFaceName(i + 1) = tmpArray(i)
   Next i
'  create the font object
   NewFont = CreateFontIndirect(lf)
'  save the current font object and use the new font object
   OldFont = SelectObject(UserControl.hDC, NewFont)
'  get the new font object's info
   Dummy = GetTextMetrics(UserControl.hDC, tm)
'  determine whether new font object is TrueType
   FontIsTrueType = (tm.tmPitchAndFamily And TMPF_TRUETYPE)
'  restore the original font object - !!!THIS IS IMPORTANT!!!
   Dummy = SelectObject(UserControl.hDC, OldFont)
   DeleteObject NewFont ' MRU added this,  got an "out of memory" error

End Function

Private Sub PasteClipboardContents()

'*************************************************************************
'* used by Ctrl-V and Shift-Insert to paste text at caret position.      *
'*************************************************************************

   Dim sTemp As String

'   sTemp = Clipboard.GetText
   UNI_ANSI_ClipBoard_GetText sTemp

   If LenB(sTemp) Then
'     if there are any cr/lf characters, remove them.
      sTemp = Replace(sTemp, vbCrLf, "")
      If LenB(m_SelText) Then
'        replace the selected text with the contents of the clipboard.
         DeleteSelection
         m_Text = Left(m_Text, CursorPos - 1) & sTemp & Mid(m_Text, CursorPos)
      ElseIf CursorPos = 1 Then
'        insert at beginning of text.
         m_Text = sTemp & m_Text
      ElseIf CursorPos = Len(m_Text) + 1 Then
'        append to end of text.
         m_Text = m_Text & sTemp
      Else
'        insert at cursor position.
         m_Text = Left(m_Text, CursorPos - 1) & sTemp & Mid(m_Text, CursorPos)
      End If
'     make sure text adheres to MaxLength property restriction.
      If Len(m_Text) > m_MaxLength And m_MaxLength > 0 Then
         m_Text = Left(m_Text, m_MaxLength)
      End If
'     place cursor at end of newly pasted text.
      If CursorPos + Len(sTemp) <= Len(m_Text) Then
         CursorPos = CursorPos + Len(sTemp)
      Else
         CursorPos = Len(m_Text) + 1
      End If
      MapCharacters DisplayRange
'     turn off selection mode.
      SelectModeActive = False
      SelectedRange.FirstCharacter = 0
      SelectedRange.LastCharacter = 0
'     if the cursor went past the edge of the textbox,
'     re-range the text from the cursor position backwards.
      If CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
         SetTextDisplayRangeRev m_Text, CursorPos
         CharactersMapped = False
      End If
      RedrawControl
   End If

End Sub

Private Function BeginningOfWord(EndCharPos As Long) As Long

'*************************************************************************
'* determines the index of the start of a word given a starting point.   *
'*************************************************************************

'  sanity check. (06/30)
   If EndCharPos = 0 Then
      EndCharPos = 1
   End If

   BeginningOfWord = InStrRev(m_Text, " ", EndCharPos)
   If BeginningOfWord = 0 Then
'     no space found; selection goes to beginning of text.
      BeginningOfWord = 1
   Else
'     beginning of word starts at character position after the found space.
      BeginningOfWord = BeginningOfWord + 1
   End If

End Function

Private Function EndOfWord(StartCharPos As Long) As Long

'*************************************************************************
'* determines the index of the end of a word given a starting point.     *
'*************************************************************************

   Dim i As Long

   EndOfWord = InStr(StartCharPos, m_Text, " ")
   If EndOfWord = 0 Then
'     if there was no space, it's the last word.
      EndOfWord = Len(m_Text) + 1
   End If

End Function

Private Sub ProcessBackspace()

'*************************************************************************
'* used for BackSpace and Shift+Delete keys.                             *
'*************************************************************************

   If Not m_Locked Then
      If Len(m_Text) > 0 And CursorPos >= 1 Then
'        delete selection if text has been selected.
         If Len(m_SelText) > 0 Then
            DeleteSelection
         Else
'           otherwise, treat as a regular backspace.
            If CursorPos > 1 Then
               DeleteCharacterAtPreviousCursorPosition
            End If
         End If
         If Len(m_Text) = 0 Then
            CursorPos = 1
         End If
         RedrawControl
         MapCharacters DisplayRange
      End If
   End If

End Sub

Private Sub MoveCursorToEndOfText()

'*************************************************************************
'* used to handle End key, DblClick event, and SelectOnFocus property.   *
'*************************************************************************

   Dim sTemp As String    ' holds either text or len(text) worth of password char.

'  for textwidth calculating purposes.
   If m_PasswordChar = "" Then
      sTemp = m_Text
   Else
      sTemp = String(Len(m_Text), Left(m_PasswordChar, 1))
   End If
   If SelectModeActive Then
      SelectedRange.LastCharacter = Len(m_Text) + 1
   End If
'  if the whole text fits in control, just place cursor at end.
   If TextWidthU(hDC, sTemp) < ScaleWidth - Clearance - 2 Then
      CursorPos = Len(m_Text) + 1
      RedrawControl
   Else
'     otherwise, find the last 'n' characters that fill the control and place cursor at end.
      SetTextDisplayRangeRev sTemp, Len(m_Text)
      CursorPos = Len(m_Text) + 1
'     if the cursor goes right past the last displayed character, we must scroll the
'     text right by the number of characters in the SideScroll property.
      If CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
         DisplayRange.FirstCharacter = DisplayRange.FirstCharacter + m_SideScroll
      End If
      CharactersMapped = False
      RedrawControl
   End If

End Sub

Private Sub SetMouseDownShiftStatus(ByVal ShiftStatus As Integer)

'*************************************************************************
'* this is utilized by the MouseDown event handling routine. I made it   *
'* separate since this is used in several MouseDown code points.         *
'*************************************************************************

   Dim ShiftDown As Boolean

'  determine shift key status.  This lets us know if user is selecting text.
   ShiftDown = (ShiftStatus And vbShiftMask) > 0
   If ShiftDown Then
      If SelectModeActive Then
'        if shift is down when mouse clicked, existing selection is extended to new cursor
'        position, or a new selection is made from beginning of text to new cursor position.
         SelectedRange.LastCharacter = CursorPos
      End If
      If Not SelectModeActive Then
         SelectModeActive = True
         SelectedRange.FirstCharacter = CursorPos
      End If
   Else
'     set SelectMode Active to false, and also clear SelectedRange values.  This is
'     because a mouse click with no shift causes selection highlight to disappear.
      SelectModeActive = False
      SelectedRange.FirstCharacter = 0
      SelectedRange.LastCharacter = 0
   End If

End Sub

Private Sub SetTextDisplayRange(TextToRange As String)

'*************************************************************************
'* determines the first and last characters to display based on textbox  *
'* width and accumulated text widths.                                    *
'*************************************************************************

   Dim i        As Long    ' loop variable.
   Dim TextLen  As Long    ' holds length of text; used for speed optimizing purposes.
   Dim SWLessBW As Long    ' scalewidth - borderwidth; for righthand spacing purposes.

   SWLessBW = ScaleWidth - m_FocusBorderWidth   ' get the right-side clearance.

'  time saver - if the whole string can fit in the textbox, bypass the loop and exit.
   If TextWidthU(hDC, TextToRange) < (SWLessBW - Clearance) Then
      DisplayRange.LastCharacter = Len(TextToRange)
      Exit Sub
   End If

   TextLen = Len(TextToRange)     ' get total text length.

   'If Not CharactersMapped Then ' commented out 06/23
      MapCharacters DisplayRange
   'End If

'  start accumulating character widths, stopping when mapped character
'  exceeds the textbox scalewidth or we run out of characters.
   For i = DisplayRange.FirstCharacter To TextLen
      If CharMap(i) + TextWidthU(hDC, Mid$(TextToRange, i, 1)) > SWLessBW Then
'        since the character starts in textbox view, but exceeds the scalewidth
'        less the border width, use the previous character as the last to display.
         DisplayRange.LastCharacter = i - 1
         Exit Sub
      End If
   Next i

'  just a safety net - if we get here, the remainder of the string will fit in the box.
   DisplayRange.LastCharacter = Len(TextToRange)

End Sub

Private Sub SetTextDisplayRangeRev(TextToRange As String, StartPos As Long)

'*************************************************************************
'* determines the first and last characters to display based on textbox  *
'* width and accumulated text widths from the end of the text backwards. *
'*************************************************************************

   Dim i               As Long    ' loop variable.
   Dim xWidth          As Long    ' running accumulation of text character widths.
   Dim TextLen         As Long    ' holds length of text; used for speed optimizing purposes.
   Dim DisplayableArea As Long    ' the width of the displayable area.

'  get the displayable area, with a 1-letter right clearance.
   DisplayableArea = ScaleWidth - m_FocusBorderWidth - Clearance - TextWidthU(hDC, "n")
   TextLen = Len(TextToRange)     ' get total text length.

   xWidth = 0
'  start accumulating character widths, stopping when xWidth
'  exceeds the textbox scalewidth or we run out of characters.
   For i = StartPos To 1 Step -1
      xWidth = xWidth + TextWidthU(hDC, Mid(TextToRange, i, 1))
      If xWidth > DisplayableArea Then
         DisplayRange.FirstCharacter = i + 1
         DisplayRange.LastCharacter = TextLen 'Len(TextToRange)
         Exit Sub
      End If
   Next i

End Sub

Private Sub MapCharacters(DispRange As TextRange)

'*************************************************************************
'* maps the start character position of each word in the text string.    *
'* also maps the X pixel position of the leftmost edge of each character,*
'* from the first displayed character to the end of the text string.     *
'* these also are the cursor X positions.                                *
'*************************************************************************

   Dim i          As Long       ' loop variable.
   Dim X          As Long       ' holds accumulating x character positions.
   Dim TextToMap  As String     ' holds either actual text or password char string.
   Dim SpaceFound As Boolean    ' flag for use during the word count routine.
   Dim tStr As String           ' rfu

'  if there's no text, there's not much point in continuing.
   If Len(m_Text) = 0 Then
'     we do this even if the textbox has no text so that the cursor can be drawn.
      ReDim CharMap(1 To 2)
      CharMap(1) = Clearance
      Exit Sub
   End If

'  sanity check.
   If DisplayRange.FirstCharacter < 1 Then
      DisplayRange.FirstCharacter = 1
   End If

'  resize the character X-address array to the length of the text.
   ReDim CharMap(1 To Len(m_Text) + 2)
'  more elements than we'll need, but that's ok.
   ReDim WordMap(1 To Len(m_Text) + 2)

'  get the start character position of each word.  Used for ctrl-left and ctrl-right arrow cursor jumping.
'  make sure the first entry is the first position so cursor can move to beginning of text.
   WordCount = 1
   WordMap(WordCount) = 1
'  loop through the text, adding the start character position of each word in the text.
   SpaceFound = True
   For i = 1 To Len(m_Text)
      If Mid(m_Text, i, 1) = " " Then
         SpaceFound = True
      Else
         If Mid(m_Text, i, 1) <> " " Then
            If SpaceFound Then
'              start of new word encountered.
               WordCount = WordCount + 1
               WordMap(WordCount) = i
               SpaceFound = False
            End If
         End If
      End If
   Next i
'  if there are any words, add a last entry to allow cursor to move to end of text.
   If WordCount > 0 Then
      WordCount = WordCount + 1
      WordMap(WordCount) = Len(m_Text) + 1
   End If

'  determine what text to map - the actual text or password characters.
   If m_PasswordChar = "" Then
      TextToMap = m_Text
   Else
      TextToMap = String(Len(m_Text), Left(m_PasswordChar, 1))
   End If

'  the text, from the first character to display to the end of the text.
   tStr = Mid(TextToMap, DisplayRange.FirstCharacter)

'  the first character always starts at the left-side clearance.
   CharMap(DisplayRange.FirstCharacter) = Clearance
   X = 1
   For i = DisplayRange.FirstCharacter To Len(TextToMap) + 1
      CharMap(i + 1) = TextWidthU(hDC, Left(tStr, X)) + Clearance - 1
      X = X + 1
   Next i

'  set the flag so we only map when necessary (like after a character was typed or deleted).
   CharactersMapped = True

End Sub

Private Sub DeleteSelection()

'*************************************************************************
'* deletes all selected characters.                                      *
'*************************************************************************

   Dim Temp As Long       ' for selected range value swapping.

'  if the first and last selected character positions are switched due
'  to changed selection direction, swap them before deletion begins.
   If SelectedRange.FirstCharacter > SelectedRange.LastCharacter Then
      Temp = SelectedRange.FirstCharacter
      SelectedRange.FirstCharacter = SelectedRange.LastCharacter
      SelectedRange.LastCharacter = Temp
   End If

'  sanity check.
   If SelectedRange.FirstCharacter < 1 Then
      SelectedRange.FirstCharacter = 1
   End If

'  delete the selected text.
   If CursorPos = Len(m_Text) + 1 Then
      m_Text = Left(m_Text, SelectedRange.FirstCharacter - 1)
   Else
      m_Text = Left(m_Text, SelectedRange.FirstCharacter - 1) & Right(m_Text, Len(m_Text) - SelectedRange.LastCharacter + 1)
   End If

'  if the whole text was deleted, clear display range and set cursor
'  position to first position.
   If m_Text = "" Then
      DisplayRange.FirstCharacter = 0
      DisplayRange.LastCharacter = 0
      CursorPos = 1
      CharMap(CursorPos) = Clearance
   Else
'     otherwise, place the cursor at the beginning of the selected area.
      CursorPos = SelectedRange.FirstCharacter
      AdjustForCursorOutOfRangeLeft
   End If

'  since the selected area was deleted, reset the selection mode values.
   SelectModeActive = False
   SelectedRange.FirstCharacter = 0
   SelectedRange.LastCharacter = 0

End Sub

Private Sub AdjustForCursorOutOfRangeLeft()

'*************************************************************************
'* remaps the text in situations where caret ends up being positioned    *
'* before the first displayed character, such as when deleting a block   *
'* of text whose starting point is to the left of the displayed part.    *
'*************************************************************************

   If CharMap(CursorPos) < Clearance Then ' added block 6/20
      DisplayRange.FirstCharacter = 1
      DisplayRange.LastCharacter = Len(m_Text) + 1
      SetTextDisplayRangeRev m_Text, CursorPos
      MapCharacters DisplayRange
   End If

End Sub

Private Sub DeleteCharacterAtCursorPosition()

'*************************************************************************
'* deletes the character at the current cursor position.                 *
'*************************************************************************

   m_Text = Left(m_Text, CursorPos - 1) & Right(m_Text, Len(m_Text) - CursorPos)
   DoEvents

End Sub

Private Sub DeleteCharacterAtPreviousCursorPosition()

'*************************************************************************
'* deletes character at previous cursor position (for backspace key).    *
'*************************************************************************

'  delete the character.
   m_Text = Left(m_Text, CursorPos - 2) & Right(m_Text, Len(m_Text) - CursorPos + 1)
'  adjust the cursor position.
   CursorPos = CursorPos - 1
   If CursorPos < 1 Then
      CursorPos = 1
   End If
   MapCharacters DisplayRange
'  if the cursor moved left of the left edge, scroll text right by SideScroll amount.
   If CharMap(CursorPos) < Clearance Then
      CharactersMapped = False
      DisplayRange.FirstCharacter = DisplayRange.FirstCharacter - m_SideScroll
   End If
   DoEvents

End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Graphics Routines  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub RedrawControl()

'*************************************************************************
'* the master routine for displaying textbox and its contents.           *
'*************************************************************************

   SetBackGround      ' display the background gradient.
   SetText            ' write the text.
   SetBorder          ' display the textbox border.
   If HasFocus Then   ' if the textbox has the focus, display the cursor.
      SetCursor
   End If

   If m_Text <> PreviousText Then
'     since any time text is changed the control gets redrawn, this is a good place
'     to check if the Change event needs to be fired.
      PreviousText = m_Text
      RaiseEvent Change
'     since the text has been changed, update the stack if not in an undo operation.
      If Not Undoing Then
         AddToStack
      End If
   End If

End Sub

Private Sub SetBackGround()

'*************************************************************************
'* displays the control's background gradient.                           *
'*************************************************************************

   Dim Temp    As Long        ' swap variable for selected range values.
   Dim Swapped As Boolean     ' swapped flag.

   On Error GoTo ErrHandler

'  if the Picture property is set, use that.
   If IsPictureThere(m_Picture) Then
      Set UserControl.Picture = m_Picture
   Else
'     otherwise, draw the appropriate background gradient, based on textbox state.
      If HasFocus Then
         PaintGradient hDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(m_FocusColor1), _
                       TranslateColor(m_FocusColor2), m_FocusAngle, m_FocusMiddleOut
      ElseIf Not m_Enabled Then
         PaintGradient hDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(m_DisabledColor1), _
                       TranslateColor(m_DisabledColor2), m_DisabledAngle, m_DisabledMiddleOut
      Else
         PaintGradient hDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(m_DefaultColor1), _
                       TranslateColor(m_DefaultColor2), m_DefaultAngle, m_DefaultMiddleOut
      End If
   End If

'  if any text is selected, highlight it.
   If SelectModeActive And HasFocus Then

'     remap characters so that highlight works if textview has been shifted left or right.
      MapCharacters DisplayRange
'     if the first and last selected character indices are reversed, swap them.
      If SelectedRange.FirstCharacter > SelectedRange.LastCharacter Then
         Swapped = True
         Temp = SelectedRange.FirstCharacter
         SelectedRange.FirstCharacter = SelectedRange.LastCharacter
         SelectedRange.LastCharacter = Temp
      End If

'     paint the selection highlight.
      If m_SelGradHeight = [HeightOfBox] Then
'        make selection gradient the inner height of box (box height - border widths).
         PaintGradient hDC, CharMap(SelectedRange.FirstCharacter), _
                       m_FocusBorderWidth, _
                       CharMap(SelectedRange.LastCharacter) - CharMap(SelectedRange.FirstCharacter), _
                       ScaleHeight - (m_FocusBorderWidth * 2), _
                       TranslateColor(m_SelColor1), _
                       TranslateColor(m_SelColor2), _
                       m_DefaultAngle, _
                       m_DefaultMiddleOut
      Else
'        make selection gradient the height of the text.
         PaintGradient hDC, CharMap(SelectedRange.FirstCharacter), _
                       (ScaleHeight - TextHeight(m_Text)) \ 2 + 1, _
                       CharMap(SelectedRange.LastCharacter) - CharMap(SelectedRange.FirstCharacter), _
                       TextHeight(m_Text) - 1, _
                       TranslateColor(m_SelColor1), _
                       TranslateColor(m_SelColor2), _
                       m_DefaultAngle, _
                       m_DefaultMiddleOut
      End If

'     since there was a selection, set the appropriate properties.
      m_SelStart = SelectedRange.FirstCharacter
      m_SelLength = SelectedRange.LastCharacter - SelectedRange.FirstCharacter
      m_SelText = Mid(m_Text, m_SelStart, m_SelLength)

'     if first and last were swapped, swap them back.
      If Swapped Then
         Temp = SelectedRange.FirstCharacter
         SelectedRange.FirstCharacter = SelectedRange.LastCharacter
         SelectedRange.LastCharacter = Temp
      End If

   Else

'     since no selection gradient was painted we can clear the range values.
      SelectedRange.FirstCharacter = 0
      SelectedRange.LastCharacter = 0
      m_SelStart = 0
      m_SelLength = 0
      m_SelText = ""

   End If

ErrHandler:
   Exit Sub

End Sub

Private Function IsPictureThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture.  Thanks to Roger Gilchrist.        *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsPictureThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub SetBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvature.     *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim BdrCol As Long        ' border color.
   Dim hBrush As Long        ' the brush pattern used to 'paint' the border.
   Dim hRgn1  As Long        ' the outer boundary of the border region.
   Dim hRgn2  As Long        ' the inner boundary of the border region.

'  get border color and create the border region to be filled in, according to textbox state.
   If Not m_Enabled Then
      BdrCol = TranslateColor(m_DisabledBorderColor)
      hRgn1 = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, m_Curvature, m_Curvature)
      hRgn2 = CreateRoundRectRgn(m_DisabledBorderWidth, m_DisabledBorderWidth, ScaleWidth - m_DisabledBorderWidth, ScaleHeight - m_DisabledBorderWidth, m_Curvature, m_Curvature)
   ElseIf HasFocus Then
      BdrCol = TranslateColor(m_FocusBorderColor)
      hRgn1 = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, m_Curvature, m_Curvature)
      hRgn2 = CreateRoundRectRgn(m_FocusBorderWidth, m_FocusBorderWidth, ScaleWidth - m_FocusBorderWidth, ScaleHeight - m_FocusBorderWidth, m_Curvature, m_Curvature)
   Else
      BdrCol = TranslateColor(m_DefaultBorderColor)
      hRgn1 = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, m_Curvature, m_Curvature)
      hRgn2 = CreateRoundRectRgn(m_DefaultBorderWidth, m_DefaultBorderWidth, ScaleWidth - m_DefaultBorderWidth, ScaleHeight - m_DefaultBorderWidth, m_Curvature, m_Curvature)
   End If
   CombineRgn hRgn2, hRgn1, hRgn2, 3

'  create and apply the color brush.
   hBrush = CreateSolidBrush(BdrCol)
   FillRgn hDC, hRgn2, hBrush

'  set the control region.
   SetWindowRgn hwnd, hRgn1, True

'  free the memory.
   DeleteObject hRgn2
   DeleteObject hBrush
   DeleteObject hRgn1

ErrHandler:
   Exit Sub

End Sub

Private Sub SetText()

'*************************************************************************
'* displays the textbox text.  Selected text is displayed using the      *
'* SelTextColor value.                                                   *
'*************************************************************************

   If Not m_Font Is Nothing Then

      Dim r           As RECT      ' the rectangle that defines the text draw area.
      Dim tHeight     As Long      ' the height of the text.
      Dim tWidth      As Long      ' the width of the text.
      Dim DisplayText As String    ' the portion of the text to display.
      Dim PWString    As String    ' a len(text) string of .PasswordChar characters.
      Dim Temp        As Long      ' swap variable for selected text color routine.
      Dim Swapped     As Boolean   ' values swapped flag for selected text color routine.
      Dim SelTextStart As Long     ' first selected character to color using SelTextColor value.
      Dim SelTextEnd As Long       ' last selected character to color using SelTextColor value.

'     sanity check
      If DisplayRange.FirstCharacter < 1 Then
         DisplayRange.FirstCharacter = 1
      End If

'     make the left clearance one letter width.
      Clearance = TextWidthU(hDC, "n")
'     get the portion of the text to display.
      If m_PasswordChar = "" Then
         SetTextDisplayRange m_Text
         DisplayText = Mid$(m_Text, DisplayRange.FirstCharacter, Abs(DisplayRange.LastCharacter - DisplayRange.FirstCharacter + 1))
      Else
         PWString = String(Len(m_Text), Left(m_PasswordChar, 1))
         SetTextDisplayRange PWString
         DisplayText = String(Abs(DisplayRange.LastCharacter - DisplayRange.FirstCharacter + 1), Left(m_PasswordChar, 1))
      End If

'     get the height and width of the text based on the selected font.
      tHeight = TextHeight(DisplayText)
      tWidth = TextWidthU(hDC, DisplayText)

'     set the text color according to textbox status.
      If Not m_Enabled Then
         UserControl.ForeColor = TranslateColor(m_DisabledTextColor)
      ElseIf HasFocus Then
         UserControl.ForeColor = TranslateColor(m_FocusTextColor)
      Else
         UserControl.ForeColor = TranslateColor(m_DefaultTextColor)
      End If

'     define the text drawing area rectangle size.
      With r
         .Left = Clearance
         .Top = (ScaleHeight - tHeight) / 2
         .Bottom = r.Top + tHeight
         .Right = .Left + tWidth
      End With

'     draw the text using API.
      DrawText UserControl.hDC, DisplayText, -1, r, DT_LEFT Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP

'     if the textbox has the focus and text has been selected,
'     draw that text using the SelTextColor property color.
      If HasFocus Then
         If SelectedRange.FirstCharacter > 0 And SelectedRange.LastCharacter > 0 And Len(m_SelText) > 0 Then
'           if the selected range values need temporarily swapped, do it.
            If SelectedRange.FirstCharacter > SelectedRange.LastCharacter Then
               Temp = SelectedRange.FirstCharacter
               SelectedRange.FirstCharacter = SelectedRange.LastCharacter
               SelectedRange.LastCharacter = Temp
               Swapped = True
            End If
            SelTextStart = SelectedRange.FirstCharacter
'           if the last selected character to display goes past the right edge of the control, make
'           the last selected character to display the regular last character displayed.
            If SelectedRange.LastCharacter > DisplayRange.LastCharacter Then
               SelTextEnd = DisplayRange.LastCharacter
            Else
               SelTextEnd = SelectedRange.LastCharacter - 1
            End If
            With r
               If Not TrueTypeFont And SelectedRange.FirstCharacter <> DisplayRange.FirstCharacter And m_Font.Bold Then
                  .Left = CharMap(SelectedRange.FirstCharacter) - 1
                Else
                  .Left = CharMap(SelectedRange.FirstCharacter)
               End If
'              if the first selected character to display starts before the left edge of the
'              control, make the first selected character to display the first display character.
               If .Left < Clearance Then
                  .Left = Clearance
                  SelTextStart = DisplayRange.FirstCharacter
               End If
               .Right = UserControl.Width - m_FocusBorderWidth
            End With
'           make sure correct text is displayed - regular or password characters.
            If m_PasswordChar = "" Then
               DisplayText = Mid(m_Text, SelTextStart, Abs(SelTextEnd - SelTextStart + 1))
            Else
               DisplayText = String(Abs(SelTextEnd - SelTextStart + 1), Left(m_PasswordChar, 1))
            End If
'           generate the SelTextColor value and draw the text in that color.
            UserControl.ForeColor = TranslateColor(m_SelTextColor)
            DrawText UserControl.hDC, DisplayText, -1, r, DT_LEFT Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP
'           if the selected range values were swapped, swap them back.
            If Swapped Then
               Temp = SelectedRange.FirstCharacter
               SelectedRange.FirstCharacter = SelectedRange.LastCharacter
               SelectedRange.LastCharacter = Temp
            End If
         End If
      End If
   End If

End Sub

Private Sub SetCursor()

'*************************************************************************
'* draws API caret in front of appropriate character.  Thanks a million  *
'* to Gary (Phantom Man) for his great contribution to this project.     *
'*************************************************************************

'  create and display caret, at height determined by CaretHeight property.
   If m_CaretHeight = [HeightOfBox] Then
'     make caret height the box height, less the border widths, less 2 more pixels.
      CreateCaret hwnd, 0, 0, ScaleHeight - (m_FocusBorderWidth * 2) - 2
      ShowCaret hwnd
'     set the caret position.
      SetCaretPos CharMap(CursorPos), m_FocusBorderWidth + 1
   Else
'     make caret height the height of the text in the current font/size.
      CreateCaret hwnd, 0, 0, TextHeight("(")
      ShowCaret hwnd
'     set the caret position according to the custom height of the caret.
      SetCaretPos CharMap(CursorPos), ((ScaleHeight - TextHeight("(")) \ 2)
   End If

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* translates ole color into COLORREF long for drawing purposes.         *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
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

'     Matthew R. Usner - solves weird problem of when angle is
'     >= 91 and <= 270, the colors invert in MiddleOut mode.
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

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Property Routines  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties for user control.                               *
'*************************************************************************

   Set m_Font = Ambient.Font
   Set m_Picture = LoadPicture("")
   m_Enabled = m_def_Enabled
   m_Text = m_def_Text
   m_DefaultTextColor = m_def_DefaultTextColor
   m_DefaultColor1 = m_def_DefaultColor1
   m_DefaultColor2 = m_def_DefaultColor2
   m_DefaultMiddleOut = m_def_DefaultMiddleOut
   m_DefaultAngle = m_def_DefaultAngle
   m_DefaultBorderColor = m_def_DefaultBorderColor
   m_DefaultBorderWidth = m_def_DefaultBorderWidth
   m_FocusColor1 = m_def_FocusColor1
   m_FocusColor2 = m_def_FocusColor2
   m_FocusMiddleOut = m_def_FocusMiddleOut
   m_FocusAngle = m_def_FocusAngle
   m_FocusTextColor = m_def_FocusTextColor
   m_FocusBorderColor = m_def_FocusBorderColor
   m_FocusBorderWidth = m_def_FocusBorderWidth
   m_DisabledColor1 = m_def_DisabledColor1
   m_DisabledColor2 = m_def_DisabledColor2
   m_DisabledMiddleOut = m_def_DisabledMiddleOut
   m_DisabledAngle = m_def_DisabledAngle
   m_DisabledTextColor = m_def_DisabledTextColor
   m_DisabledBorderColor = m_def_DisabledBorderColor
   m_DisabledBorderWidth = m_def_DisabledBorderWidth
   m_SideScroll = m_def_SideScroll
   m_Locked = m_def_Locked
   m_MaxLength = m_def_MaxLength
   m_PasswordChar = m_def_PasswordChar
   m_SelColor1 = m_def_SelColor1
   m_SelColor2 = m_def_SelColor2
   m_SelText = m_def_SelText
   m_SelStart = m_def_SelStart
   m_SelLength = m_def_SelLength
   m_SelectOnFocus = m_def_SelectOnFocus
   m_Curvature = m_DEF_Curvature
   m_SelGradHeight = m_def_SelGradHeight
   m_CaretHeight = m_def_CaretHeight
   m_SelTextColor = m_def_SelTextColor
   m_EnterEmulateTab = m_def_EnterEmulateTab

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* load property values from storage.                                    *
'*************************************************************************

   With PropBag
      Set m_Font = .ReadProperty("Font", Ambient.Font)
      TrueTypeFont = FontIsTrueType(m_Font.Name)
      Set m_Picture = .ReadProperty("Picture", Nothing)
      Set UserControl.Font = m_Font
      m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
      UserControl.Enabled = m_Enabled
      m_Text = .ReadProperty("Text", m_def_Text)
      m_DefaultTextColor = .ReadProperty("DefaultTextColor", m_def_DefaultTextColor)
      m_DefaultColor1 = .ReadProperty("DefaultColor1", m_def_DefaultColor1)
      m_DefaultColor2 = .ReadProperty("DefaultColor2", m_def_DefaultColor2)
      m_DefaultMiddleOut = .ReadProperty("DefaultMiddleOut", m_def_DefaultMiddleOut)
      m_DefaultAngle = .ReadProperty("DefaultAngle", m_def_DefaultAngle)
      m_DefaultBorderColor = .ReadProperty("DefaultBorderColor", m_def_DefaultBorderColor)
      m_DefaultBorderWidth = .ReadProperty("DefaultBorderWidth", m_def_DefaultBorderWidth)
      m_FocusColor1 = .ReadProperty("FocusColor1", m_def_FocusColor1)
      m_FocusColor2 = .ReadProperty("FocusColor2", m_def_FocusColor2)
      m_FocusMiddleOut = .ReadProperty("FocusMiddleOut", m_def_FocusMiddleOut)
      m_FocusAngle = .ReadProperty("FocusAngle", m_def_FocusAngle)
      m_FocusTextColor = .ReadProperty("FocusTextColor", m_def_FocusTextColor)
      m_FocusBorderColor = .ReadProperty("FocusBorderColor", m_def_FocusBorderColor)
      m_FocusBorderWidth = .ReadProperty("FocusBorderWidth", m_def_FocusBorderWidth)
      m_DisabledColor1 = .ReadProperty("DisabledColor1", m_def_DisabledColor1)
      m_DisabledColor2 = .ReadProperty("DisabledColor2", m_def_DisabledColor2)
      m_DisabledMiddleOut = .ReadProperty("DisabledMiddleOut", m_def_DisabledMiddleOut)
      m_DisabledAngle = .ReadProperty("DisabledAngle", m_def_DisabledAngle)
      m_DisabledTextColor = .ReadProperty("DisabledTextColor", m_def_DisabledTextColor)
      m_DisabledBorderColor = .ReadProperty("DisabledBorderColor", m_def_DisabledBorderColor)
      m_DisabledBorderWidth = .ReadProperty("DisabledBorderWidth", m_def_DisabledBorderWidth)
      m_SideScroll = .ReadProperty("SideScroll", m_def_SideScroll)
      m_Locked = .ReadProperty("Locked", m_def_Locked)
      m_MaxLength = .ReadProperty("MaxLength", m_def_MaxLength)
      m_PasswordChar = .ReadProperty("PasswordChar", m_def_PasswordChar)
      m_SelColor1 = .ReadProperty("SelColor1", m_def_SelColor1)
      m_SelColor2 = .ReadProperty("SelColor2", m_def_SelColor2)
      m_SelText = .ReadProperty("SelText", m_def_SelText)
      m_SelStart = .ReadProperty("SelStart", m_def_SelStart)
      m_SelLength = .ReadProperty("SelLength", m_def_SelLength)
      m_SelectOnFocus = .ReadProperty("SelectOnFocus", m_def_SelectOnFocus)
      m_Curvature = .ReadProperty("Curvature", m_DEF_Curvature)
      m_SelGradHeight = .ReadProperty("SelGradHeight", m_def_SelGradHeight)
      m_CaretHeight = .ReadProperty("CaretHeight", m_def_CaretHeight)
      m_SelTextColor = .ReadProperty("SelTextColor", m_def_SelTextColor)
      m_EnterEmulateTab = .ReadProperty("EnterEmulateTab", m_def_EnterEmulateTab)
   End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write property values to storage.                                     *
'*************************************************************************

   With PropBag
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "Font", m_Font, Ambient.Font
      .WriteProperty "Text", m_Text, m_def_Text
      .WriteProperty "DefaultTextColor", m_DefaultTextColor, m_def_DefaultTextColor
      .WriteProperty "DefaultColor1", m_DefaultColor1, m_def_DefaultColor1
      .WriteProperty "DefaultColor2", m_DefaultColor2, m_def_DefaultColor2
      .WriteProperty "DefaultMiddleOut", m_DefaultMiddleOut, m_def_DefaultMiddleOut
      .WriteProperty "DefaultAngle", m_DefaultAngle, m_def_DefaultAngle
      .WriteProperty "DefaultBorderColor", m_DefaultBorderColor, m_def_DefaultBorderColor
      .WriteProperty "DefaultBorderWidth", m_DefaultBorderWidth, m_def_DefaultBorderWidth
      .WriteProperty "FocusColor1", m_FocusColor1, m_def_FocusColor1
      .WriteProperty "FocusColor2", m_FocusColor2, m_def_FocusColor2
      .WriteProperty "FocusMiddleOut", m_FocusMiddleOut, m_def_FocusMiddleOut
      .WriteProperty "FocusAngle", m_FocusAngle, m_def_FocusAngle
      .WriteProperty "FocusTextColor", m_FocusTextColor, m_def_FocusTextColor
      .WriteProperty "FocusBorderColor", m_FocusBorderColor, m_def_FocusBorderColor
      .WriteProperty "FocusBorderWidth", m_FocusBorderWidth, m_def_FocusBorderWidth
      .WriteProperty "DisabledColor1", m_DisabledColor1, m_def_DisabledColor1
      .WriteProperty "DisabledColor2", m_DisabledColor2, m_def_DisabledColor2
      .WriteProperty "DisabledMiddleOut", m_DisabledMiddleOut, m_def_DisabledMiddleOut
      .WriteProperty "DisabledAngle", m_DisabledAngle, m_def_DisabledAngle
      .WriteProperty "DisabledTextColor", m_DisabledTextColor, m_def_DisabledTextColor
      .WriteProperty "DisabledBorderColor", m_DisabledBorderColor, m_def_DisabledBorderColor
      .WriteProperty "DisabledBorderWidth", m_DisabledBorderWidth, m_def_DisabledBorderWidth
      .WriteProperty "SideScroll", m_SideScroll, m_def_SideScroll
      .WriteProperty "Locked", m_Locked, m_def_Locked
      .WriteProperty "MaxLength", m_MaxLength, m_def_MaxLength
      .WriteProperty "PasswordChar", m_PasswordChar, m_def_PasswordChar
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "SelColor1", m_SelColor1, m_def_SelColor1
      .WriteProperty "SelColor2", m_SelColor2, m_def_SelColor2
      .WriteProperty "SelText", m_SelText, m_def_SelText
      .WriteProperty "SelStart", m_SelStart, m_def_SelStart
      .WriteProperty "SelLength", m_SelLength, m_def_SelLength
      .WriteProperty "SelectOnFocus", m_SelectOnFocus, m_def_SelectOnFocus
      .WriteProperty "Curvature", m_Curvature, m_DEF_Curvature
      .WriteProperty "SelGradHeight", m_SelGradHeight, m_def_SelGradHeight
      .WriteProperty "CaretHeight", m_CaretHeight, m_def_CaretHeight
      .WriteProperty "SelTextColor", m_SelTextColor, m_def_SelTextColor
      .WriteProperty "EnterEmulateTab", m_EnterEmulateTab, m_def_EnterEmulateTab
   End With

End Sub

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   UserControl.Enabled = m_Enabled
   PropertyChanged "Enabled"
   RedrawControl
End Property

Public Property Get Font() As Font
   Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)

   Set m_Font = New_Font
   Set UserControl.Font = m_Font
   PropertyChanged "Font"
   TrueTypeFont = FontIsTrueType(m_Font.Name)
   'PropertyChanged "Text"
   RedrawControl
   MapCharacters DisplayRange

End Property

Public Property Get Text() As String
   Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)

'  make sure the text conforms to the limits of the MaxLength property.
   If m_MaxLength > 0 And Len(New_Text) > m_MaxLength Then
      m_Text = Left(New_Text, m_MaxLength)
   Else
      m_Text = New_Text
   End If
   PropertyChanged "Text"

'  reset the caret position and first displayed character.
   CursorPos = 1
   DisplayRange.FirstCharacter = 1

   RedrawControl
   MapCharacters DisplayRange

End Property

Public Property Get DefaultTextColor() As OLE_COLOR
   DefaultTextColor = m_DefaultTextColor
End Property

Public Property Let DefaultTextColor(ByVal New_DefaultTextColor As OLE_COLOR)
   m_DefaultTextColor = New_DefaultTextColor
   PropertyChanged "DefaultTextColor"
   RedrawControl
End Property

Public Property Get DefaultColor1() As OLE_COLOR
   DefaultColor1 = m_DefaultColor1
End Property

Public Property Let DefaultColor1(ByVal New_DefaultColor1 As OLE_COLOR)
   m_DefaultColor1 = New_DefaultColor1
   PropertyChanged "DefaultColor1"
   RedrawControl
End Property

Public Property Get DefaultColor2() As OLE_COLOR
   DefaultColor2 = m_DefaultColor2
End Property

Public Property Let DefaultColor2(ByVal New_DefaultColor2 As OLE_COLOR)
   m_DefaultColor2 = New_DefaultColor2
   PropertyChanged "DefaultColor2"
   RedrawControl
End Property

Public Property Get DefaultMiddleOut() As Boolean
   DefaultMiddleOut = m_DefaultMiddleOut
End Property

Public Property Let DefaultMiddleOut(ByVal New_DefaultMiddleOut As Boolean)
   m_DefaultMiddleOut = New_DefaultMiddleOut
   PropertyChanged "DefaultMiddleOut"
   RedrawControl
End Property

Public Property Get DefaultAngle() As Single
   DefaultAngle = m_DefaultAngle
End Property

Public Property Let DefaultAngle(ByVal New_DefaultAngle As Single)
   m_DefaultAngle = New_DefaultAngle
   PropertyChanged "DefaultAngle"
   RedrawControl
End Property

Public Property Get DefaultBorderColor() As OLE_COLOR
   DefaultBorderColor = m_DefaultBorderColor
End Property

Public Property Let DefaultBorderColor(ByVal New_DefaultBorderColor As OLE_COLOR)
   m_DefaultBorderColor = New_DefaultBorderColor
   PropertyChanged "DefaultBorderColor"
   RedrawControl
End Property

Public Property Get DefaultBorderWidth() As Integer
   DefaultBorderWidth = m_DefaultBorderWidth
End Property

Public Property Let DefaultBorderWidth(ByVal New_DefaultBorderWidth As Integer)
   m_DefaultBorderWidth = New_DefaultBorderWidth
   PropertyChanged "DefaultBorderWidth"
   RedrawControl
End Property

Public Property Get FocusColor1() As OLE_COLOR
   FocusColor1 = m_FocusColor1
End Property

Public Property Let FocusColor1(ByVal New_FocusColor1 As OLE_COLOR)
   m_FocusColor1 = New_FocusColor1
   PropertyChanged "FocusColor1"
   RedrawControl
End Property

Public Property Get FocusColor2() As OLE_COLOR
   FocusColor2 = m_FocusColor2
End Property

Public Property Let FocusColor2(ByVal New_FocusColor2 As OLE_COLOR)
   m_FocusColor2 = New_FocusColor2
   PropertyChanged "FocusColor2"
   RedrawControl
End Property

Public Property Get FocusMiddleOut() As Boolean
   FocusMiddleOut = m_FocusMiddleOut
End Property

Public Property Let FocusMiddleOut(ByVal New_FocusMiddleOut As Boolean)
   m_FocusMiddleOut = New_FocusMiddleOut
   PropertyChanged "FocusMiddleOut"
   RedrawControl
End Property

Public Property Get FocusAngle() As Single
   FocusAngle = m_FocusAngle
End Property

Public Property Let FocusAngle(ByVal New_FocusAngle As Single)
   m_FocusAngle = New_FocusAngle
   PropertyChanged "FocusAngle"
   RedrawControl
End Property

Public Property Get FocusTextColor() As OLE_COLOR
   FocusTextColor = m_FocusTextColor
End Property

Public Property Let FocusTextColor(ByVal New_FocusTextColor As OLE_COLOR)
   m_FocusTextColor = New_FocusTextColor
   PropertyChanged "FocusTextColor"
   RedrawControl
End Property

Public Property Get FocusBorderColor() As OLE_COLOR
   FocusBorderColor = m_FocusBorderColor
End Property

Public Property Let FocusBorderColor(ByVal New_FocusBorderColor As OLE_COLOR)
   m_FocusBorderColor = New_FocusBorderColor
   PropertyChanged "FocusBorderColor"
   RedrawControl
End Property

Public Property Get FocusBorderWidth() As Integer
   FocusBorderWidth = m_FocusBorderWidth
End Property

Public Property Let FocusBorderWidth(ByVal New_FocusBorderWidth As Integer)
   m_FocusBorderWidth = New_FocusBorderWidth
   PropertyChanged "FocusBorderWidth"
   RedrawControl
End Property

Public Property Get DisabledColor1() As OLE_COLOR
Attribute DisabledColor1.VB_Description = "The first gradient color when textbox is disabled."
   DisabledColor1 = m_DisabledColor1
End Property

Public Property Let DisabledColor1(ByVal New_DisabledColor1 As OLE_COLOR)
   m_DisabledColor1 = New_DisabledColor1
   PropertyChanged "DisabledColor1"
   RedrawControl
End Property

Public Property Get DisabledColor2() As OLE_COLOR
Attribute DisabledColor2.VB_Description = "The second gradient color when textbox is disabled."
   DisabledColor2 = m_DisabledColor2
End Property

Public Property Let DisabledColor2(ByVal New_DisabledColor2 As OLE_COLOR)
   m_DisabledColor2 = New_DisabledColor2
   PropertyChanged "DisabledColor2"
   RedrawControl
End Property

Public Property Get DisabledMiddleOut() As Boolean
Attribute DisabledMiddleOut.VB_Description = "The middle-out gradient flag when textbox is disabled."
   DisabledMiddleOut = m_DisabledMiddleOut
End Property

Public Property Let DisabledMiddleOut(ByVal New_DisabledMiddleOut As Boolean)
   m_DisabledMiddleOut = New_DisabledMiddleOut
   PropertyChanged "DisabledMiddleOut"
   RedrawControl
End Property

Public Property Get DisabledAngle() As Single
Attribute DisabledAngle.VB_Description = "The gradient angle when textbox is disabled."
   DisabledAngle = m_DisabledAngle
End Property

Public Property Let DisabledAngle(ByVal New_DisabledAngle As Single)
   m_DisabledAngle = New_DisabledAngle
   PropertyChanged "DisabledAngle"
   RedrawControl
End Property

Public Property Get DisabledTextColor() As OLE_COLOR
Attribute DisabledTextColor.VB_Description = "The Text color when textbox is disabled."
   DisabledTextColor = m_DisabledTextColor
End Property

Public Property Let DisabledTextColor(ByVal New_DisabledTextColor As OLE_COLOR)
   m_DisabledTextColor = New_DisabledTextColor
   PropertyChanged "DisabledTextColor"
   RedrawControl
End Property

Public Property Get DisabledBorderColor() As OLE_COLOR
Attribute DisabledBorderColor.VB_Description = "The border color when textbox is disabled."
   DisabledBorderColor = m_DisabledBorderColor
End Property

Public Property Let DisabledBorderColor(ByVal New_DisabledBorderColor As OLE_COLOR)
   m_DisabledBorderColor = New_DisabledBorderColor
   PropertyChanged "DisabledBorderColor"
   RedrawControl
End Property

Public Property Get DisabledBorderWidth() As Integer
Attribute DisabledBorderWidth.VB_Description = "The border width when textbox is disabled."
   DisabledBorderWidth = m_DisabledBorderWidth
End Property

Public Property Let DisabledBorderWidth(ByVal New_DisabledBorderWidth As Integer)
   m_DisabledBorderWidth = New_DisabledBorderWidth
   PropertyChanged "DisabledBorderWidth"
   RedrawControl
End Property

Public Property Get SideScroll() As Long
Attribute SideScroll.VB_Description = "The number of characters to shift the text when cursor goes past end of text window while scrolling through a long text."
   SideScroll = m_SideScroll
End Property

Public Property Let SideScroll(ByVal New_SideScroll As Long)
   m_SideScroll = New_SideScroll
   PropertyChanged "SideScroll"
   RedrawControl
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "When True, text cannot be changed."
   Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
   m_Locked = New_Locked
   PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "When 0, no limit to text length."
   MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
   m_MaxLength = New_MaxLength
   If m_MaxLength > 0 Then
      ReDim WordMap(1 To m_MaxLength + 2)
   Else
      ReDim WordMap(1 To 500)  ' arbitrary for now
   End If
   PropertyChanged "MaxLength"
'  ensures that the length of the text does not exceed the MaxLength property.
   If Len(m_Text) > m_MaxLength And m_MaxLength > 0 Then
      m_Text = Left(m_Text, m_MaxLength)
      CursorPos = 1
      CharactersMapped = False
      RedrawControl
   End If
End Property

Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "When set, all characters typed are displayed as this character."
   PasswordChar = m_PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
   m_PasswordChar = New_PasswordChar
   PropertyChanged "PasswordChar"
   RedrawControl
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
   Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set m_Picture = New_Picture
   PropertyChanged "Picture"
   RedrawControl
End Property

Public Property Get SelColor1() As OLE_COLOR
Attribute SelColor1.VB_Description = "The first gradient color for the background of selected text."
   SelColor1 = m_SelColor1
End Property

Public Property Let SelColor1(ByVal New_SelColor1 As OLE_COLOR)
   m_SelColor1 = New_SelColor1
   PropertyChanged "SelColor1"
   RedrawControl
End Property

Public Property Get SelColor2() As OLE_COLOR
Attribute SelColor2.VB_Description = "The second gradient color for the background of selected text."
   SelColor2 = m_SelColor2
End Property

Public Property Let SelColor2(ByVal New_SelColor2 As OLE_COLOR)
   m_SelColor2 = New_SelColor2
   PropertyChanged "SelColor2"
   RedrawControl
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "The text that has been selected."
   SelText = m_SelText
End Property

Private Property Let SelText(ByVal New_SelText As String)
   m_SelText = New_SelText
   PropertyChanged "SelText"
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "The starting character position of the selected text."
   SelStart = m_SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)

'*************************************************************************
'* sets the cursor position through code, and also sets the starting     *
'* point for programmatically selected text.                             *
'*************************************************************************

   If New_SelStart > Len(m_Text) + 1 Then
      m_SelStart = Len(m_Text) + 1
   Else
      m_SelStart = New_SelStart
   End If
   CursorPos = m_SelStart
   If CursorPos = 0 Then '06/21/05
      CursorPos = 1
      m_SelStart = 1
   End If
'  if the cursor went past the edge of the textbox,
'  re-range the text from the cursor position backwards.
   If CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
      SetTextDisplayRangeRev m_Text, CursorPos
      CharactersMapped = False
   End If
'  previously existing stuff
   PropertyChanged "SelStart"
   SelectModeActive = True
   SelectedRange.FirstCharacter = m_SelStart
   If HasFocus Then RedrawControl

End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "The length of the selected text."
   SelLength = m_SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)

'*************************************************************************
'* handles programmatic selection of SelLength property.                 *
'*************************************************************************

   SelectModeActive = True
'  a safety net in case user didn't set SelStart.
   If m_SelStart = 0 Then
      m_SelStart = 1
      SelectedRange.FirstCharacter = 1
   End If

   If m_SelStart + New_SelLength > Len(m_Text) + 1 Then
      m_SelLength = (Len(m_Text) + 1) - m_SelStart ' ?
   Else
      m_SelLength = New_SelLength
   End If
   PropertyChanged "SelLength"

   SelectedRange.LastCharacter = m_SelStart + m_SelLength
   m_SelText = Mid(m_Text, m_SelStart, m_SelLength)
   CursorPos = SelectedRange.LastCharacter

'  if the cursor goes right past the last displayed character, scroll the
'  text right by the appropriate number of characters.
   If CharMap(CursorPos) > (ScaleWidth - m_FocusBorderWidth - 1) Then
      CharactersMapped = False
      DisplayRange.FirstCharacter = DisplayRange.FirstCharacter + m_SideScroll
      SetTextDisplayRangeRev m_Text, CursorPos
   End If

   RedrawControl

End Property

Public Property Get SelectOnFocus() As Boolean
Attribute SelectOnFocus.VB_Description = "If True, all text is selected when control receives the focus."
   SelectOnFocus = m_SelectOnFocus
End Property

Public Property Let SelectOnFocus(ByVal New_SelectOnFocus As Boolean)
   m_SelectOnFocus = New_SelectOnFocus
   PropertyChanged "SelectOnFocus"
End Property

Public Property Get Curvature() As Integer
Attribute Curvature.VB_Description = "Allows the textbox to have rounded corners. A value of zero has no curvature."
   Curvature = m_Curvature
End Property

Public Property Let Curvature(ByVal New_Curvature As Integer)
   m_Curvature = New_Curvature
   PropertyChanged "Curvature"
   RedrawControl
End Property

Public Property Get SelGradHeight() As HeightValues
Attribute SelGradHeight.VB_Description = "Sets the text selection gradient bar to either the height of the text or the height of the textbox."
   SelGradHeight = m_SelGradHeight
End Property

Public Property Let SelGradHeight(ByVal New_SelGradHeight As HeightValues)
   m_SelGradHeight = New_SelGradHeight
   PropertyChanged "SelGradHeight"
   RedrawControl
End Property

Public Property Get CaretHeight() As HeightValues
Attribute CaretHeight.VB_Description = "Sets the height of the caret (cursor) to either the inner height of the MorphTextBox or the height of the text."
   CaretHeight = m_CaretHeight
End Property

Public Property Let CaretHeight(ByVal New_CaretHeight As HeightValues)
   m_CaretHeight = New_CaretHeight
   PropertyChanged "CaretHeight"
   RedrawControl
End Property

Public Property Get SelTextColor() As OLE_COLOR
Attribute SelTextColor.VB_Description = "Sets the color of selected text."
   SelTextColor = m_SelTextColor
End Property

Public Property Let SelTextColor(ByVal New_SelTextColor As OLE_COLOR)
   m_SelTextColor = New_SelTextColor
   PropertyChanged "SelTextColor"
   RedrawControl
End Property

Public Property Get EnterEmulateTab() As Boolean
Attribute EnterEmulateTab.VB_Description = "When True, pressing the Enter key shifts focus to next MorphTextBox as if Tab key had been pressed."
   EnterEmulateTab = m_EnterEmulateTab
End Property

Public Property Let EnterEmulateTab(ByVal New_EnterEmulateTab As Boolean)
   m_EnterEmulateTab = New_EnterEmulateTab
   PropertyChanged "EnterEmulateTab"
End Property

