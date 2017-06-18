VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'            All Thanks and Praise to My Lord Jesus Christ
'            All Thanks and Praise to My Lord Jesus Christ
'            All Thanks and Praise to My Lord Jesus Christ
'-----------------------------------------------------------------------------------------------------
'// Title:    cPrinter version 2 - API Printer Class
'// Author:   Joshy Francis
'// Version:  2.0
'// Licesnse: MIT
'-----------------------------------------------------------------------------------------------------

' This software is provided "as-is," without any express or implied warranty.
' In no event shall the author be held liable for any damages arising from the use of this software.
' If you do not agree with these terms, do not use "Grid". Use of the program implicitly means
' you have agreed to these terms.

' Permission is granted to anyone to use this software for any purpose,
' including commercial use, and to alter and redistribute it, provided that
' the following conditions are met:

' 1. All redistributions of source code files must retain all copyright
'    notices that are currently in place, and this list of conditions without
'    any modification.
' 2. All redistributions in binary form must retain all occurrences of the
'    above copyright notice and web site addresses that are currently in
'    place (for example, in the About boxes).
' 3. Modified versions in source or binary form must be plainly marked as
'    such, and must not be misrepresented as being the original software.
'-----------------------------------------------------------------------------------------------------


Option Explicit
'*********************** Begin Drawing Declarations **********************************
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'--- GetDeviceCaps constants
Private Const RASTERCAPS                As Long = 38
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const SIZEPALETTE               As Long = 104
Private Const RC_PALETTE                As Long = &H100
Private Const CAPS1                     As Long = 94 ' other caps
Private Const C1_TRANSPARENT            As Long = &H1 ' new raster cap
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
Private Type PALETTEENTRY
   peRed               As Byte
   peGreen             As Byte
   peBlue              As Byte
   peFlags             As Byte
End Type

Private Type LOGPALETTE
   palVersion          As Integer
   palNumEntries       As Integer
   palPalEntry(255)    As PALETTEENTRY
End Type
Private Type TEXTMETRIC
    tmHeight            As Long
    tmAscent            As Long
    tmDescent           As Long
    tmInternalLeading   As Long
    tmExternalLeading   As Long
    tmAveCharWidth      As Long
    tmMaxCharWidth      As Long
    tmWeight            As Long
    tmOverhang          As Long
    tmDigitizedAspectX  As Long
    tmDigitizedAspectY  As Long
    tmFirstChar         As Byte
    tmLastChar          As Byte
    tmDefaultChar       As Byte
    tmBreakChar         As Byte
    tmItalic            As Byte
    tmUnderlined        As Byte
    tmStruckOut         As Byte
    tmPitchAndFamily    As Byte
    tmCharSet           As Byte
End Type
Private Declare Function GetSystemPaletteEntries Lib "gdi32" ( _
      ByVal hDc As Long, _
      ByVal wStartIndex As Long, _
      ByVal wNumEntries As Long, _
      lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" ( _
      ByVal hDc As Long, _
      ByVal hPalette As Long, _
      ByVal bForceBackground As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkMode Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" _
      Alias "GetTextFaceA" ( _
      ByVal hDc As Long, _
      ByVal nCount As Long, _
      ByVal lpFacename As String) As Long
'--- GetTextMetrics constants
Private Const FW_NORMAL                 As Long = 400
Private Const LF_FACESIZE               As Long = 32
Private Declare Function GetTextMetrics Lib "gdi32" _
      Alias "GetTextMetricsA" ( _
      ByVal hDc As Long, _
      lpMetrics As TEXTMETRIC) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function APIFillRect Lib "user32" _
      Alias "FillRect" ( _
      ByVal hDc As Long, _
      lpRect As RECT, _
      ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function ApiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Public Enum PrintPenStyles
   PS_SOLID = 0                ' The pen is solid.
   PS_DASH = 1                 ' The pen is dashed.
   PS_DOT = 2                  ' The pen is dotted.
   PS_DASHDOT = 3              ' The pen has alternating dashes and dots.
   PS_DASHDOTDOT = 4           ' The pen has dashes and double dots.
   PS_NULL = 5                 ' The pen is invisible.
   PS_INSIDEFRAME = 6          ' The pen is solid. When this pen is used in any GDI drawing function that takes a
                               '   bounding rectangle, the dimensions of the figure are shrunk so that it fits entirely in the bounding rectangle,
                               '   taking into account the width of the pen. This applies only to geometric pens.
   PS_USERSTYLE = 7            ' <b>Windows NT/2000:</b> The pen uses a styling array supplied by the user.
   PS_ALTERNATE = 8            ' <b>Windows NT/2000:</b> The pen sets every other pixel. (This style is applicable only
                               '   for cosmetic pens.)
   PS_STYLE_MASK = &HF         ' Mask for previous PS_XXX values.
   PS_ENDCAP_ROUND = &H0       ' End caps are round.
   PS_ENDCAP_SQUARE = &H100    ' End caps are square.
   PS_ENDCAP_FLAT = &H200      ' End caps are flat.
   PS_ENDCAP_MASK = &HF00      ' Mask for previous PS_ENDCAP_XXX values.
   PS_JOIN_ROUND = &H0         ' Joins are beveled.
   PS_JOIN_BEVEL = &H1000      ' Joins are mitered when they are within the current limit set by the SetMiterLimit
                               '   function. If it exceeds this limit, the join is beveled.
   PS_JOIN_MITER = &H2000      ' Joins are round.
   PS_JOIN_MASK = &HF000       ' Mask for previous PS_JOIN_XXX values.
   PS_COSMETIC = &H0           ' The pen is cosmetic.
   PS_GEOMETRIC = &H10000      ' The pen is geometric.
   PS_TYPE_MASK = &HF0000      ' Mask for previous PS_XXX (pen type).
End Enum
Private Declare Function ApiCreatePen Lib "gdi32" _
      Alias "CreatePen" ( _
      ByVal nPenStyle As Long, _
      ByVal nWidth As Long, _
      ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" ( _
      ByVal hDc As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function GetROP2 Lib "gdi32" (ByVal hDc As Long) As Long
Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(1 To 32) As Byte
End Type

Private Type SIZEAPI
    cX                  As Long
    cY                  As Long
End Type
Private Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
      Alias "RtlMoveMemory" ( _
      Destination As Any, _
      Source As Any, _
      ByVal Length As Long)
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
'--- DrawIconEx constants
Private Const DI_NORMAL                 As Long = &H3
Private Declare Function DrawIconEx Lib "user32" ( _
      ByVal hDc As Long, _
      ByVal xLeft As Long, _
      ByVal yTop As Long, _
      ByVal hIcon As Long, _
      ByVal cxWidth As Long, _
      ByVal cyWidth As Long, _
      ByVal istepIfAniCur As Long, _
      ByVal hbrFlickerFreeDraw As Long, _
      ByVal diFlags As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
      ByVal hDc As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long) As Long
Private Declare Function APIBitBlt Lib "gdi32" _
      Alias "BitBlt" ( _
      ByVal hdcDest As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hdcSrc As Long, _
      ByVal xSrc As Long, _
      ByVal ySrc As Long, _
      ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateEnhMetaFileLong Lib "gdi32" _
      Alias "CreateEnhMetaFileA" ( _
      ByVal hdcRef As Long, _
      ByVal lpFileName As String, _
      ByVal lpRect As Long, _
      ByVal lpDescription As String) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Private Declare Function PlayMetaFile Lib "gdi32" (ByVal hDc As Long, ByVal hMF As Long) As Long
Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDc As Long, ByVal hEmf As Long, lpRect As RECT) As Long
Public Enum PrintDrawTextStyles
    DT_LEFT = &H0               ' Aligns text to the left.
    DT_TOP = &H0                ' Justifies the text to the top of the rectangle.
    DT_CENTER = &H1             ' Centers text horizontally in the rectangle.
    DT_RIGHT = &H2              ' Aligns text to the right.
    DT_VCENTER = &H4            ' Centers text vertically. This value is used only with the DT_SINGLELINE value.
    DT_BOTTOM = &H8             ' Justifies the text to the bottom of the rectangle. This value is used only with the DT_SINGLELINE value.
    DT_WORDBREAK = &H10         ' Breaks words. Lines are automatically broken between words if a word would extend past the edge of the rectangle specified by the lpRect parameter. A carriage return-line feed sequence also breaks the line.<br>If this is not specified, output is on one line.
    DT_SINGLELINE = &H20        ' Displays text on a single line only. Carriage returns and line feeds do not break the line.
    DT_EXPANDTABS = &H40        ' Expands tab characters. The default number of characters per tab is eight. The DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS, and DT_END_ELLIPSIS values cannot be used with the DT_EXPANDTABS value.
    DT_TABSTOP = &H80           ' Sets tab stops. Bits 15–8 (high-order byte of the low-order word) of the uFormat parameter specify the number of characters for each tab. The default number of characters per tab is eight. The DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, and DT_NOPREFIX values cannot be used with the DT_TABSTOP value.
    DT_NOCLIP = &H100           ' Draws without clipping. DrawText is somewhat faster when DT_NOCLIP is used.
    DT_EXTERNALLEADING = &H200  ' Includes the font external leading in line height. Normally, external leading is not included in the height of a line of text.
    DT_CALCRECT = &H400         ' Determines the width and height of the rectangle. If there are multiple lines of text, DrawText uses the width of the rectangle pointed to by the lpRect parameter and extends the base of the rectangle to bound the last line of text. If the largest word is wider than the rectangle, the width is expanded. If the text is less than the width of the rectangle, the width is reduced. If there is only one line of text, DrawText modifies the right side of the rectangle so that it bounds the last character in the line. In either case, DrawText returns the height of the formatted text but does not draw the text.
    DT_NOPREFIX = &H800         ' Turns off processing of prefix characters. Normally, DrawText interprets the mnemonic-prefix character & as a directive to underscore the character that follows, and the mnemonic-prefix characters && as a directive to print a single &. By specifying DT_NOPREFIX, this processing is turned off
    DT_INTERNAL = &H1000        ' Uses the system font to calculate text metrics.
    DT_EDITCONTROL = &H2000     ' Duplicates the text-displaying characteristics of a multiline edit control. Specifically, the average character width is calculated in the same manner as for an edit control, and the function does not display a partially visible last line.
    DT_PATH_ELLIPSIS = &H4000   ' For displayed text, replaces characters in the middle of the string with ellipses so that the result fits in the specified rectangle. If the string contains backslash (\) characters, DT_PATH_ELLIPSIS preserves as much as possible of the text after the last backslash.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_END_ELLIPSIS and DT_WORD_ELLIPSIS.
    DT_END_ELLIPSIS = &H8000    ' For displayed text, if the end of a string does not fit in the rectangle, it is truncated and ellipses are added. If a word that is not at the end of the string goes beyond the limits of the rectangle, it is truncated without ellipses.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_PATH_ELLIPSIS and DT_WORD_ELLIPSIS.
    DT_MODIFYSTRING = &H10000   ' Modifies the specified string to match the displayed text. This value has no effect unless DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is specified.
    DT_RTLREADING = &H20000     ' Layout in right-to-left reading order for bi-directional text when the font selected into the hdc is a Hebrew or Arabic font. The default reading order for all text is left-to-right.
    DT_WORD_ELLIPSIS = &H40000  ' Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.

End Enum
Private Type DRAWTEXTPARAMS
    cbSize              As Long
    iTabLength          As Long
    iLeftMargin         As Long
    iRightMargin        As Long
    uiLengthDrawn       As Long
End Type
Private Declare Function ApiDrawTextEx Lib "user32" _
      Alias "DrawTextExA" ( _
      ByVal hDc As Long, _
      ByVal lpsz As String, _
      ByVal n As Long, _
      lpRect As RECT, _
      ByVal un As Long, _
      lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function GetTextExtentPoint Lib "gdi32" _
      Alias "GetTextExtentPointA" ( _
      ByVal hDc As Long, _
      ByVal lpszString As String, _
      ByVal cbString As Long, _
      lpSize As SIZEAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
      ByVal lOleColor As Long, _
      ByVal lHPalette As Long, _
      lColorRef As Long) As Long
Private Declare Function APIRectangle Lib "gdi32" _
      Alias "Rectangle" ( _
      ByVal hDc As Long, _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDc As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" ( _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const AC_SRC_OVER = &H0
Private Type BLENDFUNCTION  'Required for Alphablend API
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal Blendfunc As Long) As Boolean
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Boolean
' Dim BF As BLENDFUNCTION, lBF As Long
'    'Set the graphics mode to persistent
'    Picture1.AutoRedraw = True
'    Picture2.AutoRedraw = True
'    'API uses pixels
'    Picture1.ScaleMode = vbPixels
'    Picture2.ScaleMode = vbPixels
'    'set the parameters
'    With BF
'        .BlendOp = AC_SRC_OVER
'        .BlendFlags = 0
'        .SourceConstantAlpha = 128
'        .AlphaFormat = 0
'    End With
'    'copy the BLENDFUNCTION-structure to a Long
'    RtlMoveMemory lBF, BF, 4
'    'AlphaBlend the picture from Picture1 over the picture of Picture2
'    GdiAlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, lBF
Private Declare Function GdiAlphaBlend Lib "gdi32.dll" (ByVal hDc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Const MM_HIMETRIC = 3
Private Const MM_HIENGLISH = 5
Private Const MM_LOENGLISH = 4
Private Const MM_LOMETRIC = 2
Private Const MM_TEXT = 1
Private Const MM_TWIPS = 6
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDc As Long) As Long
Private Type PictDesc
    Size                As Long
    Type                As Long
    hBmpOrIcon          As Long
    hPal                As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
      lpPictDesc As PictDesc, _
      riid As Any, _
      ByVal fPictureOwnsHandle As Long, _
      ipic As IPicture) As Long
Private Type ICONINFO
   fIcon               As Long
   xHotspot            As Long
   yHotspot            As Long
   hBmMask             As Long
   hbmColor            As Long
End Type
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long

'*********************** End Drawing Declarations **********************************


'*********************** Begin Printing Declarations **********************************
Private Type PRINTER_INFO_1
        flags As Long
        pDescription As String
        pName As String
        pComment As String
End Type
Private Const PRINTER_ENUM_LOCAL = &H2 'Enumerate local printer objects.
Private Const PRINTER_ENUM_CONNECTIONS = &H4     'Enumerate printer connections previously added through RpcAddPerMachineConnection.
Private Const PRINTER_ENUM_NAME = &H8     'Enumerate printers on the print server, network domain, or a specific print provider.
Private Const PRINTER_ENUM_REMOTE = &H10    'Enumerate network printers and other print servers that are in the same domain as the print server.
Private Const PRINTER_ENUM_SHARED = &H20    'Only enumerate printers with the shared attribute set. This flag MUST be combined with one or more of the other flags.
Private Const PRINTER_ENUM_NETWORK = &H40    'Enumerate network printers that are in the same domain as the print server.
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal Name As Long, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GetDefaultPrinterA Lib "winspool.drv" (ByVal str As String, pcbNeeded As Long) As Long
Private Declare Function SetDefaultPrinterA Lib "winspool.drv" (ByVal str As String) As Long
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type DEVMODE_NEW
  dmDeviceName As String * CCHDEVICENAME
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
            dmPosition As POINTAPI
            dmDisplayOrientation As Long
            dmDisplayFixedOutput As Long
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
        dmDisplayFlags As Long
        dmNup As Long
  dmDisplayFrequency As Long
        dmICMMethod As Long
        dmICMIntent As Long
        dmMediaType As Long
        dmDitherType As Long
        dmReserved1 As Long
        dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type
Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
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
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
'/* field selection bits */
Private Const DM_ORIENTATION = &H1
Private Const DM_PAPERSIZE = &H2
Private Const DM_PAPERLENGTH = &H4
Private Const DM_PAPERWIDTH = &H8
Private Const DM_SCALE = &H10
'#if(WINVER >= =&h0500)
Private Const DM_POSITION = &H20
Private Const DM_NUP = &H40
'#endif /* WINVER >= =&h0500 */
'#if(WINVER >= =&h0501)
Private Const DM_DISPLAYORIENTATION = &H80
'#endif /* WINVER >= =&h0501 */
Private Const DM_COPIES = &H100
Private Const DM_DEFAULTSOURCE = &H200
Private Const DM_PRINTQUALITY = &H400
Private Const DM_COLOR = &H800
Private Const DM_DUPLEX = &H1000
Private Const DM_YRESOLUTION = &H2000
Private Const DM_TTOPTION = &H4000
Private Const DM_COLLATE = &H8000
Private Const DM_FORMNAME = &H10000
Private Const DM_LOGPIXELS = &H20000
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const DM_DISPLAYFLAGS = &H200000
Private Const DM_DISPLAYFREQUENCY = &H400000
'#if(WINVER >= =&h0400)
Private Const DM_ICMMETHOD = &H800000
Private Const DM_ICMINTENT = &H1000000
Private Const DM_MEDIATYPE = &H2000000
Private Const DM_DITHERTYPE = &H4000000
Private Const DM_PANNINGWIDTH = &H8000000
Private Const DM_PANNINGHEIGHT = &H10000000
'#endif /* WINVER >= =&h0400 */
'#if(WINVER >= =&h0501)
Private Const DM_DISPLAYFIXEDOUTPUT = &H20000000
'#endif /* WINVER >= =&h0501 */


'/* orientation selections */
Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
'used with SetBkMode
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
'Private Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal nBkMode As Long) As Long
'/* Device Parameters for GetDeviceCaps() */
Private Const DRIVERVERSION = 0    ' /* Device driver version                    */
Private Const TECHNOLOGY = 2       ' /* Device classification                    */
Private Const HORZSIZE = 4        '  /* Horizontal size in millimeters           */
Private Const VERTSIZE = 6         ' /* Vertical size in millimeters             */
Private Const HORZRES = 8          ' /* Horizontal width in pixels               */
Private Const VERTRES = 10         ' /* Vertical height in pixels                */
Private Const BITSPIXEL = 12       ' /* Number of bits per pixel                 */
Private Const Planes = 14          ' /* Number of planes                         */
Private Const NUMBRUSHES = 16      ' /* Number of brushes the device has         */
Private Const NUMPENS = 18         ' /* Number of pens the device has            */
Private Const NUMMARKERS = 20      ' /* Number of markers the device has         */
Private Const NUMFONTS = 22       '  /* Number of fonts the device has           */
Private Const NUMCOLORS = 24      '  /* Number of colors the device supports     */
Private Const PDEVICESIZE = 26    '  /* Size required for device descriptor      */
Private Const CURVECAPS = 28      '  /* Curve capabilities                       */
Private Const LINECAPS = 30       '  /* Line capabilities                        */
Private Const POLYGONALCAPS = 32  '  /* Polygonal capabilities                   */
Private Const TEXTCAPS = 34       '  /* Text capabilities                        */
Private Const CLIPCAPS = 36       '  /* Clipping capabilities                    */
'Private Const RASTERCAPS = 38     '  /* Bitblt capabilities                      */
Private Const ASPECTX = 40        '  /* Length of the X leg                      */
Private Const ASPECTY = 42        '  /* Length of the Y leg                      */
Private Const ASPECTXY = 44       '  /* Length of the hypotenuse                 */

'Private Const LOGPIXELSX = 88      ' /* Logical pixels/inch in X                 */
'Private Const LOGPIXELSY = 90     '  /* Logical pixels/inch in Y                 */

'Private Const SIZEPALETTE = 104   '  /* Number of entries in physical palette    */
Private Const NUMRESERVED = 106   '  /* Number of reserved entries in palette    */
Private Const COLORRES = 108      '  /* Actual color resolution                  */

'// Printing related DeviceCaps. These replace the appropriate Escapes

Private Const PHYSICALWIDTH = 110  ' /* Physical Width in device units           */
Private Const PHYSICALHEIGHT = 111  '/* Physical Height in device units          */
Private Const PHYSICALOFFSETX = 112 ' /* Physical Printable Area x margin         */
Private Const PHYSICALOFFSETY = 113 '/* Physical Printable Area y margin         */
Private Const SCALINGFACTORX = 114 ' /* Scaling factor x                         */
Private Const SCALINGFACTORY = 115  '/* Scaling factor y                         */
'Private Declare Function GetDeviceCaps Lib "gdi32" _
'    (ByVal hDc As Long, ByVal nIndex As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" _
'    (ByVal hdc As Long) As Long
Private Const PRINTER_ACCESS_ADMINISTER = &H4 'Printing-specific access rights for printers to perform administrative tasks ([MS-DTYP] ACCESS_MASK Bit 29).
Private Const PRINTER_ACCESS_USE = &H8 'Printing-specific access rights for printers to perform basic printing operations ([MS-DTYP] ACCESS_MASK Bit 28).
Private Const PRINTER_ACCESS_MANAGE_LIMITED = &H40 'Printing-specific access rights for printers to perform printer data management operations ([MS-DTYP] ACCESS_MASK Bit 25).<128>
Private Const PRINTER_ALL_ACCESS = &HF000C 'Access rights for printers to perform all administrative tasks and basic printing operations except synchronization. Combines WO (Write Owner), WD (Write DACL), RC (Read Control), and DE (Delete) of ACCESS_MASK with printing-specific PRINTER_ACCESS_ADMINISTER and printing-specific PRINTER_ACCESS_USE.
Private Const PRINTER_EXECUTE = &H20008 'Access rights for printers combining RC (Read Control) of ACCESS_MASK with printing-specific PRINTER_ACCESS_USE.
Private Const PRINTER_READ = &H20008 'Access rights for printers combining RC (Read Control) of ACCESS_MASK with printing-specific PRINTER_ACCESS_USE.
Private Const PRINTER_WRITE = &H20008 'Access rights for printers combining RC (Read Control) of ACCESS_MASK with printing-specific PRINTER_ACCESS_USE.
Private Type PRINTER_DEFAULTS
    pDatatype As String
    pDevMode As DEVMODE
    DesiredAccess As Long
End Type
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, ByRef phPrinter As Long, ByRef pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Type DOCINFO
  cbSize As Long
  lpszDocName As String
  lpszOutput As String
  lpszDatatype As String
  fwType As Long 'DI_APPBANDING
End Type
Private Const DI_APPBANDING = &H1
Private Const DI_ROPS_READ_DESTINATION = &H2
Private Declare Function StartDoc Lib "gdi32" Alias _
    "StartDocA" (ByVal hDc As Long, lpdi As DOCINFO) As Long
Private Declare Function EndDoc Lib "gdi32" _
    (ByVal hDc As Long) As Long
Private Declare Function StartPage Lib "gdi32" _
    (ByVal hDc As Long) As Long
Private Declare Function EndPage Lib "gdi32" _
    (ByVal hDc As Long) As Long
'Private Declare Function Rectangle Lib "gdi32" (ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
'Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias _
    "TextOutA" (ByVal hDc As Long, ByVal x As Long, ByVal _
    y As Long, ByVal lpString As String, ByVal nCount _
    As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" _
'    (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Private Const PD_ALLPAGES = &H0 'The default flag that indicates that the All radio button is initially selected. This flag is used as a placeholder to indicate that the PD_PAGENUMS and PD_SELECTION flags are not specified.
Private Const PD_COLLATE = &H10 'If this flag is set, the Collate check box is selected.If this flag is set when the PrintDlg function returns, the application must simulate collation of multiple copies. For more information, see the description of the PD_USEDEVMODECOPIESANDCOLLATE flag.See PD_NOPAGENUMS.
Private Const PD_DISABLEPRINTTOFILE = &H80000 'Disables the Print to File check box.
Private Const PD_ENABLEPRINTHOOK = &H1000 'Enables the hook procedure specified in the lpfnPrintHook member. This enables the hook procedure for the Print dialog box.
Private Const PD_ENABLEPRINTTEMPLATE = &H4000 'Indicates that the hInstance and lpPrintTemplateName members specify a replacement for the default Print dialog box template.
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000 'Indicates that the hPrintTemplate member identifies a data block that contains a preloaded dialog box template. This template replaces the default template for the Print dialog box. The system ignores the lpPrintTemplateName member if this flag is specified.
Private Const PD_ENABLESETUPHOOK = &H2000 'Enables the hook procedure specified in the lpfnSetupHook member. This enables the hook procedure for the Print Setup dialog box.
Private Const PD_ENABLESETUPTEMPLATE = &H8000 'Indicates that the hInstance and lpSetupTemplateName members specify a replacement for the default Print Setup dialog box template.
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000 'Indicates that the hSetupTemplate member identifies a data block that contains a preloaded dialog box template. This template replaces the default template for the Print Setup dialog box. The system ignores the lpSetupTemplateName member if this flag is specified.
Private Const PD_HIDEPRINTTOFILE = &H100000 'Hides the Print to File check box.
Private Const PD_NONETWORKBUTTON = &H200000 'Hides and disables the Network button.
Private Const PD_NOPAGENUMS = &H2 'Disables the Pages radio button and the associated edit controls. Also, it causes the Collate check box to appear in the dialog.
Private Const PD_NOSELECTION = &H4 'Disables the Selection radio button.
Private Const PD_NOWARNING = &H80 'Prevents the warning message from being displayed when there is no default printer.
Private Const PD_PAGENUMS = &H2 'If this flag is set, the Pages radio button is selected. If this flag is set when the PrintDlg function returns, the nFromPage and nToPage members indicate the starting and ending pages specified by the user.
Private Const PD_PRINTSETUP = &H40 'Causes the system to display the Print Setup dialog box rather than the Print dialog box.
Private Const PD_PRINTTOFILE = &H20 'If this flag is set, the Print to File check box is selected. If this flag is set when the PrintDlg function returns, the offset indicated by the wOutputOffset member of the DEVNAMES structure contains the string "FILE:". When you call the StartDoc function to start the printing operation, specify this "FILE:" string in the lpszOutput member of the DOCINFO structure. Specifying this string causes the print subsystem to query the user for the name of the output file.
Private Const PD_RETURNDC = &H100 'Causes PrintDlg to return a device context matching the selections the user made in the dialog box. The device context is returned in hDC.
Private Const PD_RETURNDEFAULT = &H400 'If this flag is set, the PrintDlg function does not display the dialog box. Instead, it sets the hDevNames and hDevMode members to handles to DEVMODE and DEVNAMES structures that are initialized for the system default printer. Both hDevNames and hDevMode must be NULL, or PrintDlg returns an error.
Private Const PD_RETURNIC = &H200 'Similar to the PD_RETURNDC flag, except this flag returns an information context rather than a device context. If neither PD_RETURNDC nor PD_RETURNIC is specified, hDC is undefined on output.
Private Const PD_SELECTION = &H1 'If this flag is set, the Selection radio button is selected. If neither PD_PAGENUMS nor PD_SELECTION is set, the All radio button is selected.
Private Const PD_SHOWHELP = &H800 'Causes the dialog box to display the Help button. The hwndOwner member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the Help button.
Private Const PD_USEDEVMODECOPIES = &H40000 'Same as PD_USEDEVMODECOPIESANDCOLLATE.
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000 'This flag indicates whether your application supports multiple copies and collation. Set this flag on input to indicate that your application does not support multiple copies and collation. In this case, the nCopies member of the PRINTDLG structure always returns 1, and PD_COLLATE is never set in the Flags member.If this flag is not set, the application is responsible for printing and collating multiple copies. In this case, the nCopies member of the PRINTDLG structure indicates the number of copies the user wants to print, and the PD_COLLATE flag in the Flags member indicates whether the user wants collation.Regardless of whether this flag is set, an application can determine from nCopies and PD_COLLATE how many copies to render and whether to print them collated.
'If this flag is set and the printer driver does not support multiple copies, the Copies edit control is disabled. Similarly, if this flag is set and the printer driver does not support collation, the Collate check box is disabled.
'The dmCopies and dmCollate members of the DEVMODE structure contain the copies and collate information used by the printer driver. If this flag is set and the printer driver supports multiple copies, the dmCopies member indicates the number of copies requested by the user. If this flag is set and the printer driver supports collation, the dmCollate member of the DEVMODE structure indicates whether the user wants collation. If this flag is not set, the dmCopies member always returns 1, and the dmCollate member is always zero.
'Known issue on Windows 2000/XP/2003:  If this flag is not set before calling PrintDlg, PrintDlg might swap nCopies and dmCopies values when it returns. The workaround for this issue is use dmCopies if its value is larger than 1, else, use nCopies, for you to to get the actual number of copies to be printed when PrintDlg returns.



Private Const PSD_DEFAULTMINMARGINS = &H0 'Sets the minimum values that the user can specify for the page margins to be the minimum margins allowed by the printer. This is the default. This flag is ignored if the PSD_MARGINS and PSD_MINMARGINS flags are also specified.
Private Const PSD_DISABLEMARGINS = &H10 'Disables the margin controls, preventing the user from setting the margins.
Private Const PSD_DISABLEORIENTATION = &H100   'Disables the orientation controls, preventing the user from setting the page orientation.
Private Const PSD_DISABLEPAGEPAINTING = &H80000 'Prevents the dialog box from drawing the contents of the sample page. If you enable a PagePaintHook hook procedure, you can still draw the contents of the sample page.
Private Const PSD_DISABLEPAPER = &H200   'Disables the paper controls, preventing the user from setting page parameters such as the paper size and source.
Private Const PSD_DISABLEPRINTER = &H20    'Obsolete.Windows XP/2000:  Disables the Printer button, preventing the user from invoking a dialog box that contains additional printer setup information.
Private Const PSD_ENABLEPAGEPAINTHOOK = &H40000 'Enables the hook procedure specified in the lpfnPagePaintHook member.
Private Const PSD_ENABLEPAGESETUPHOOK = &H2000 'Enables the hook procedure specified in the lpfnPageSetupHook member.
Private Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000  'Indicates that the hInstance and lpPageSetupTemplateName members specify a dialog box template to use in place of the default template.
Private Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000 'Indicates that the hPageSetupTemplate member identifies a data block that contains a preloaded dialog box template. The system ignores the lpPageSetupTemplateName member if this flag is specified.
Private Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8     'Indicates that hundredths of millimeters are the unit of measurement for margins and paper size. The values in the rtMargin, rtMinMargin, and ptPaperSize members are in hundredths of millimeters. You can set this flag on input to override the default unit of measurement for the user's locale. When the function returns, the dialog box sets this flag to indicate the units used.
Private Const PSD_INTHOUSANDTHSOFINCHES = &H4     'Indicates that thousandths of inches are the unit of measurement for margins and paper size. The values in the rtMargin, rtMinMargin, and ptPaperSize members are in thousandths of inches. You can set this flag on input to override the default unit of measurement for the user's locale. When the function returns, the dialog box sets this flag to indicate the units used.
Private Const PSD_INWININIINTLMEASURE = &H0     'Reserved.
Private Const PSD_MARGINS = &H2 'Causes the system to use the values specified in the rtMargin member as the initial widths for the left, top, right, and bottom margins. If PSD_MARGINS is not set, the system sets the initial widths to one inch for all margins.
Private Const PSD_MINMARGINS = &H1     'Causes the system to use the values specified in the rtMinMargin member as the minimum allowable widths for the left, top, right, and bottom margins. The system prevents the user from entering a width that is less than the specified minimum. If PSD_MINMARGINS is not specified, the system sets the minimum allowable widths to those allowed by the printer.
Private Const PSD_NONETWORKBUTTON = &H200000 'Hides and disables the Network button.
Private Const PSD_NOWARNING = &H80    'Prevents the system from displaying a warning message when there is no default printer.
Private Const PSD_RETURNDEFAULT = &H400   'PageSetupDlg does not display the dialog box. Instead, it sets the hDevNames and hDevMode members to handles to DEVMODE and DEVNAMES structures that are initialized for the system default printer. PageSetupDlg returns an error if either hDevNames or hDevMode is not NULL.
Private Const PSD_SHOWHELP = &H800   'Causes the dialog box to display the Help button. The hwndOwner member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the Help button.

Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

Private Const RC_BITBLT = 1
Private Const RC_BANDING = 2
Private Const RC_SCALING = 4
Private Const RC_BITMAP64 = 8
Private Const RC_GDI20_OUTPUT = &H10
Private Const RC_DI_BITMAP = &H80
'Private Const RC_PALETTE = &H100
Private Const RC_DIBTODEV = &H200
Private Const RC_BIGFONT = &H400
Private Const RC_STRETCHBLT = &H800
Private Const RC_FLOODFILL = &H1000
Private Const RC_STRETCHDIB = &H2000
Private Const GMEM_FIXED = &H0
'Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
'Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_MODIFY = &H80
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Const DIB_RGB_COLORS = 0
Private Const DIB_PAL_COLORS = 1
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
'Private Const PRINTER_ACCESS_ADMINISTER = &H4
'Private Const PRINTER_ACCESS_USE = &H8
'Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Private Const DM_UPDATE = 1
Private Const DM_COPY = 2
Private Const DM_PROMPT = 4
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = 8
Private Const DM_IN_PROMPT = 4
Private Const DM_OUT_BUFFER = 2
Private Const DM_OUT_DEFAULT = 1

Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
'Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function OpenPrinterBynum Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
'Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, ByVal fMode As Long) As Long
Private Declare Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, ByVal pDevModeInput As Long) As Long
Private Declare Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hWnd As Long, ByVal flags As Long) As Long
Private Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hWnd As Long, ByVal pPortName As String) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, ByVal lpDevMode As Long) As Long
Private Declare Function CreateDCBynum Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
'Private Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndDocAPI Lib "gdi32" Alias "EndDoc" (ByVal hDc As Long) As Long
Private Declare Function AbortDoc Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetAbortProc Lib "gdi32" (ByVal hDc As Long, ByVal lpAbortProc As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const BI_RGB = 0&
Private Type BITMAPINFOHEADER '40 bytes
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
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

' BITMAPINFO for this example is for 16 color bitmap
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(256) As RGBQUAD
End Type
Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Const TA_BASELINE = 24
Private Const TA_BOTTOM = 8
Private Const TA_CENTER = 6
Private Const TA_LEFT = 0
Private Const TA_NOUPDATECP = 0
Private Const TA_RIGHT = 2
Private Const TA_TOP = 0
Private Const TA_UPDATECP = 1
Private Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDc As Long, ByVal wFlags As Long) As Long
'Private AbortPrinting As Long
Private Const SP_NOTREPORTED = &H4000
Private Const SP_ERROR = (-1)
Private Const SP_APPABORT = (-2)
Private Const SP_USERABORT = (-3)
Private Const SP_OUTOFDISK = (-4)
Private Const SP_OUTOFMEMORY = (-5)
Private AbortPrinting As Long
'Private Type PRINTER_DEFAULTS
'  pDatatype As String
'  pDevMode As Long  'DEVMODE
'  DesiredAccess As Long
'End Type

'Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Long, ByVal cbBuf As Long, pcbNeeded As Long) As Long
'Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Long, ByVal Command As Long) As Long
'Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
'Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)
'Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrcpy1 Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long

'Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const ERROR_INSUFFICIENT_BUFFER = 122
'Private Const PRINTER_ACCESS_USE = &H8
'Private Const PRINTER_ACCESS_ADMINISTER = &H4
'Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDc As Long, lpInitData As Any) As Long
Private Type SIZEL
    cX As Long
    cY As Long
End Type
' The two definitions for FORM_INFO_1 make the coding easier.
Private Type FORM_INFO_1
    flags As Long
    pName As Long            ' String
    Size As SIZEL
    ImageableArea As RECT
End Type
Private Type sFORM_INFO_1
    flags As Long
    pName As String
    Size As SIZEL
    ImageableArea As RECT
End Type
Private Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" _
        (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, _
        ByVal cbBuf As Long, ByRef pcbNeeded As Long, _
        ByRef pcReturned As Long) As Long
Private Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" _
        (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Private Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" _
        (ByVal hPrinter As Long, ByVal pFormName As String) As Long
' Optional functions not used in this sample, but may be useful.
Private Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" _
        (ByVal hPrinter As Long, ByVal pFormName As String, _
        ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, _
        pcbNeeded As Long) As Long
Private Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" _
        (ByVal hPrinter As Long, ByVal pFormName As String, _
        ByVal Level As Long, pForm As Byte) As Long
'Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
'Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long

' Custom constants for this sample's SelectForm function
Private Const FORM_NOT_SELECTED = 0
Private Const FORM_SELECTED = 1
Private Const FORM_ADDED = 2


Private Const mMMPerInch As Single = 25.4

'*********************** End Printing Declarations **********************************

'*************** Begin Class Variables ************************************************
Private hPrinter As Long
Private hdcPrinter As Long, temp_hdcPrinter As Long
Private sPrinter As String, bPreview As Boolean
Private hBmp As Long, hBmpOld As Long, temp_hBmp As Long, temp_hBmpOld As Long
Private hPen As Long, hPenOld As Long
Private hBrush As Long, hBrushOld As Long
Private hFont As Long, hFontOld As Long
Private hPalette As Long, hPaletteOld As Long
Private m_PSD As PAGESETUPDLG
Private hdcParent As Long, hwndParent As Long
Private nJobID As Long
Private pDevMode As DEVMODE
Private DPIX As Long, DPIY As Long, BitsPerPixel As Long
Private XPrintableArea As Long, YPrintableArea As Long
Private DPIX_Display As Long, DPIY_Display As Long
Private pScaleX As Double, pScaleY As Double
Private OffsetX As Long, OffsetY As Long
Public Enum Printer_Orientatons
    Portrait = DMORIENT_PORTRAIT
    Landscape = DMORIENT_LANDSCAPE
End Enum
Private sServerName As String
Private sShareName As String, sPortName As String, sDriverName As String
Private sComment As String, sLocation As String, sPrintProcessor As String
Private sDefaultDataType As String

'*************** End Class Variables ************************************************
'*************** Begin Class Functions ************************************************
Public Property Get ServerName() As String
      ServerName = sServerName
End Property
Public Property Get ShareName() As String
       ShareName = sShareName
End Property
Public Property Get DriverName() As String
        DriverName = sDriverName
End Property
Public Property Get Comment() As String
         Comment = sComment
End Property
Public Property Get Location() As String
          Location = sLocation
End Property
Public Property Get PrintProcessor() As String
           PrintProcessor = sPrintProcessor
End Property
Public Property Get DefaultDataType() As String
            DefaultDataType = sDefaultDataType
End Property
Public Property Get Preview() As Boolean
     Preview = bPreview
End Property
Public Property Let Preview(ByVal NewVal As Boolean)
    bPreview = NewVal
If bPreview = True And temp_hdcPrinter = 0 Then
    temp_hdcPrinter = hdcPrinter
    temp_hBmp = hBmp
    temp_hBmpOld = hBmpOld
    hdcPrinter = CreateCompatibleDC(hdcParent)
'    hBmp = CreateCompatibleBitmap(hdcParent, Width, Height)
    hBmp = CreateCompatibleBitmap(hdcParent, XPrintableArea, YPrintableArea)
    hBmpOld = SelectObject(hdcPrinter, hBmp)
            SetBkColor hdcPrinter, -1
            SetBkMode hdcPrinter, TRANSPARENT
            SetROP2 hdcPrinter, GetROP2(temp_hdcPrinter)
Else
    If temp_hdcPrinter <> 0 Then
        DeleteObject SelectObject(hdcPrinter, hBmpOld)
        DeleteDC hdcPrinter
        hdcPrinter = temp_hdcPrinter
        hBmp = temp_hBmp
        hBmpOld = temp_hBmpOld
        temp_hdcPrinter = 0
    End If
End If
End Property
Public Property Get Printer() As String
     Printer = sPrinter
End Property
Public Property Let Printer(ByVal NewVal As String)
    sPrinter = NewVal
        PrinterCreate
End Property

Public Property Get Duplex() As Long
     Duplex = pDevMode.dmDuplex
End Property
Public Property Let Duplex(ByVal NewVal As Long)
    pDevMode.dmDuplex = NewVal
        PrinterReset
End Property
Public Property Get PaperSize() As Long
    PaperSize = pDevMode.dmPaperSize
End Property
Public Property Let PaperSize(ByVal NewVal As Long)
    pDevMode.dmPaperSize = NewVal
        PrinterReset
End Property
Public Property Get Orientation() As Printer_Orientatons
    Orientation = pDevMode.dmOrientation
End Property
Public Property Let Orientation(ByVal NewVal As Printer_Orientatons)
    pDevMode.dmOrientation = NewVal
        PrinterReset
End Property
Public Property Get PaperLength() As Long
    PaperLength = pDevMode.dmPaperLength
End Property
Public Property Let PaperLength(ByVal NewVal As Long)
    pDevMode.dmPaperLength = NewVal
        PrinterReset
End Property
Public Property Get PaperWidth() As Long
    PaperWidth = pDevMode.dmPaperWidth
End Property
Public Property Let PaperWidth(ByVal NewVal As Long)
    pDevMode.dmPaperWidth = NewVal
        PrinterReset
End Property
Public Property Get PrinterScale() As Long
    PrinterScale = pDevMode.dmScale
End Property
Public Property Let PrinterScale(ByVal NewVal As Long)
    pDevMode.dmScale = NewVal
        PrinterReset
End Property
Public Property Get Copies() As Long
    Copies = pDevMode.dmCopies
End Property
Public Property Let Copies(ByVal NewVal As Long)
    pDevMode.dmCopies = NewVal
        PrinterReset
End Property
Public Property Get DefaultSource() As Long
    DefaultSource = pDevMode.dmDefaultSource
End Property
Public Property Let DefaultSource(ByVal NewVal As Long)
    pDevMode.dmDefaultSource = NewVal
        PrinterReset
End Property
Public Property Get PrintQuality() As Long
    PrintQuality = pDevMode.dmPrintQuality
End Property
Public Property Let PrintQuality(ByVal NewVal As Long)
    pDevMode.dmPrintQuality = NewVal
        PrinterReset
End Property
Public Property Get Color() As Long
    Color = pDevMode.dmColor
End Property
Public Property Let Color(ByVal NewVal As Long)
    pDevMode.dmColor = NewVal
        PrinterReset
End Property
Public Property Get Collate() As Long
    Collate = pDevMode.dmCollate
End Property
Public Property Let Collate(ByVal NewVal As Long)
    pDevMode.dmCollate = NewVal
        PrinterReset
End Property

Public Property Get DisplayFrequency() As Long
    DisplayFrequency = pDevMode.dmDisplayFrequency
End Property
Public Property Let DisplayFrequency(ByVal NewVal As Long)
    pDevMode.dmDisplayFrequency = NewVal
        PrinterReset
End Property
Public Property Get Port() As String
  Dim bBuffer() As Long
  Dim lngNeeded As Long
  Dim pdDefaults As PRINTER_DEFAULTS
  ReDim bBuffer(1)
  pdDefaults.DesiredAccess = PRINTER_ACCESS_USE
  pdDefaults.pDatatype = "RAW"
'  pdDefaults.pDevMode = 0

  GetPrinter hPrinter, 2, bBuffer(0), 0, lngNeeded
  If Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
    Port = ""
    Exit Property
  End If

  ReDim bBuffer((lngNeeded \ 4) + 1)
  If GetPrinter(hPrinter, 2, bBuffer(0), lngNeeded, lngNeeded) = 0 Then
    Port = ""
    Exit Property
  End If
  Port = PointerToString(bBuffer(3))
End Property
Public Property Let Port(ByVal strPort As String)
  Dim bBuffer() As Long  'This is a PRINTER_INFO_2 + the string area
  Dim lngNeeded As Long
  Dim pdDefaults As PRINTER_DEFAULTS

  ReDim bBuffer(1)

  pdDefaults.DesiredAccess = PRINTER_ALL_ACCESS
  pdDefaults.pDatatype = "RAW"
'  pdDefaults.pDevMode = 0

  GetPrinter hPrinter, 2, bBuffer(0), 0, lngNeeded
  If Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
    Exit Property
  End If


  ReDim bBuffer((lngNeeded \ 4) + (Len(strPort) \ 2) + 2)
  If GetPrinter(hPrinter, 2, bBuffer(0), lngNeeded, lngNeeded) = 0 Then
    Exit Property
  End If

  lstrcpy1 VarPtr(bBuffer((lngNeeded \ 4) + 1)), strPort
  bBuffer(3) = VarPtr(bBuffer((lngNeeded \ 4) + 1))

  If SetPrinter(hPrinter, 2, bBuffer(0), 0) = 0 Then
    Exit Property
  End If
End Property
Sub Init(ParentForm As Form)
    hdcParent = ParentForm.hDc
    hwndParent = ParentForm.hWnd
End Sub
Private Sub PrinterDestroy()
    If hdcPrinter <> 0 Then
        If hPenOld <> 0 Then
            DeleteObject SelectObject(hdcPrinter, hPenOld)
                hPenOld = 0
        End If
        If hBrushOld <> 0 Then
            DeleteObject SelectObject(hdcPrinter, hBrushOld)
                hBrushOld = 0
        End If
        If hFontOld <> 0 Then
            DeleteObject SelectObject(hdcPrinter, hFontOld)
                hFontOld = 0
        End If
        If hPaletteOld <> 0 Then
            DeleteObject SelectObject(hdcPrinter, hPaletteOld)
                hPaletteOld = 0
        End If
        DeleteDC hdcPrinter
            hdcPrinter = 0
    End If
    If hPrinter <> 0 Then
        ClosePrinter hPrinter
        hPrinter = 0
    End If
End Sub
Public Function StringFromPointer(lpString As Long, lMaxLength As Long) As String
    Dim sRet As String
    Dim lRet As Long
    If lpString = 0 Then
        StringFromPointer = ""
        Exit Function
    End If
    If IsBadStringPtrByLong(lpString, lMaxLength) Then
        ' An error has occured - do not attempt to use this pointer
        StringFromPointer = ""
        Exit Function
    End If
    ' Pre-initialise the return string...
    sRet = Space$(lMaxLength)
    CopyMemory ByVal sRet, ByVal lpString, ByVal Len(sRet)
    If Err.LastDllError = 0 Then
        If InStr(sRet, Chr$(0)) > 0 Then
            sRet = Left$(sRet, InStr(sRet, Chr$(0)) - 1)
        End If
    End If
    StringFromPointer = sRet
End Function

Private Sub PrinterCreate()
        If hPrinter <> 0 Then
            PrinterDestroy
        End If
    pDevMode.dmSize = Len(pDevMode)
'    pDevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX Or DM_PAPERSIZE
    pDevMode.dmDeviceName = sPrinter & Chr(0)
'    pDEVMODE.dmPaperWidth = Printer.Width
'    pDevMode.dmOrientation = nOrientation ' Printer.Orientation
'    pDevMode.dmPaperSize = nPaperSize 'A5'Printer.PaperSize
'    pDevMode.dmDuplex = nDuplex ' Printer.Duplex
Dim lResult As Long
Dim PDF As PRINTER_DEFAULTS
        PDF.DesiredAccess = PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE
        PDF.pDatatype = vbNullString
    lResult = OpenPrinter(sPrinter, hPrinter, PDF)
If lResult = 0 Then
     Exit Sub
End If
'Private sServerName As String
'Private sShareName As String, sPortName As String, sDriverName As String
'Private sComment As String, sLocation As String, sPrintProcessor As String
'Private sDefaultDataType As String
    Dim SizeNeeded As Long, Buffer() As Long
            'Initialize the buffer
        ReDim Preserve Buffer(0 To 0) As Long
        'Retrieve the required size (in bytes)
        lResult = GetPrinter(hPrinter, 2, Buffer(0), UBound(Buffer), SizeNeeded)
        'Resize the buffer... Note that a Long is four bytes
        ReDim Preserve Buffer(0 To (SizeNeeded / 4) + 3) As Long
        'Retrieve the Printer information
        lResult = GetPrinter(hPrinter, 2, Buffer(0), UBound(Buffer) * 4, SizeNeeded)
        'The data stored in 'buffer' corresponds with the data of a PRINTER_INFO_2 structure
        sServerName = StringFromPointer(Buffer(0), 255)
    '    sPrinterName =  StringFromPointer(Buffer(1), 255)
        sShareName = StringFromPointer(Buffer(2), 255)
        sPortName = StringFromPointer(Buffer(3), 255)
        sDriverName = StringFromPointer(Buffer(4), 255)
        sComment = StringFromPointer(Buffer(5), 255)
        sLocation = StringFromPointer(Buffer(6), 255)
        sPrintProcessor = StringFromPointer(Buffer(9), 255)
        sDefaultDataType = StringFromPointer(Buffer(10), 255)
        
    Dim bufsize As Long
    Dim dmInBuf() As Byte
    Dim dmOutBuf() As Byte
    ' Get a copy of the DEVMODE structure for this printer
    ' First find out how big the DEVMODE structure is
    bufsize = DocumentProperties(hwndParent, hPrinter, sPrinter, 0, 0, 0)

    ' Allocate buffers of that size
'    ReDim dmInBuf(bufsize)
    ReDim dmOutBuf(bufsize)

    ' Get the output DEVMODE structure
'    lResult = DocumentProperties(hwndParent, hPrinter, sPrinter, ByVal VarPtr(dmOutBuf(0)), ByVal VarPtr(dmInBuf(0)), DM_OUT_BUFFER)
    lResult = DocumentProperties(hwndParent, hPrinter, sPrinter, ByVal VarPtr(dmOutBuf(0)), ByVal 0&, DM_OUT_BUFFER)

    ' Copy the data buffer into the DEVMODE structure
     CopyMemory pDevMode, dmOutBuf(0), Len(pDevMode)
    ' Set the orientation, and set the dmField flag so that
    ' the function will know that it is valid.

    'pDevMode.dmOrientation = DMORIENT_LANDSCAPE
    'pDevMode.dmFields = dm.dmFields Or DM_ORIENTATION
'    If pDevMode.dmFields = 0 Then
        pDevMode.dmFields = pDevMode.dmFields Or DM_PAPERSIZE Or DM_DUPLEX Or _
            DM_ORIENTATION Or DM_COLLATE Or DM_BITSPERPEL Or DM_COLOR Or _
            DM_COPIES Or DM_DEFAULTSOURCE Or DM_DISPLAYFREQUENCY Or _
            DM_PRINTQUALITY Or DM_PAPERWIDTH Or DM_SCALE
'    End If
'    pDevMode.dmPaperSize = nPaperSize
'    pDevMode.dmOrientation = nOrientation
'    pDevMode.dmDuplex = nDuplex
'    pDevMode.dmCollate
'    pDevMode.dmColor
'    pDevMode.dmCopies
'    pDevMode.dmPrintQuality
    ' Now copy the data back to the buffer
    CopyMemory dmOutBuf(0), pDevMode, Len(pDevMode)
    ' We now have need DC to the default printer
    ' This DC is also initialized to landscape mode

'    hdcPrinter = CreateDC("WINSPOOL", strPrinterName & Chr(0), "", pDevMode)'ok

'    hdcPrinter = CreateDC("WINSPOOL", strPrinterName, "", pDevMode) 'ok
    hdcPrinter = CreateDC("WINSPOOL", sPrinter, "", ByVal VarPtr(dmOutBuf(0)))  'ok

        PrinterInit

'    Caption = "API Print"
'    ' We set up the abort procdure here
'    AbortPrinting = 0
''    lResult = SetAbortProc(hdcPrinter, zAddressOf(Me, 3))
'    lResult = SetAbortProc(hdcPrinter, AddressOf ModHelpers.AbortProc)


'    Dim PrintJobID As Long
'     Dim iret As Long, n As Long
'    Dim xpos As Double, ypos As Double, W As Double, H As Double
'    Dim docinf As DOCINFO
'    ' set up an initial font
'    Dim log_font As LOGFONT, New_Font As Long, old_font As Long
'    Dim str As String
'
''    With log_font
''      .lfEscapement = 0  ' desired rotation in tenths of a degree
''      .lfHeight = 12 * (-DPIY / 72)   ' 12 points
''      .lfFaceName = "Verdana" & vbNullChar
''      .lfWeight = 400   ' standard (bold = 700)
''      .lfItalic = False
''      .lfUnderline = False
''    End With
''    New_Font = CreateFontIndirect(log_font)
''    old_font = SelectObject(hdcPrinter, New_Font)
'            Dim hPen As Long, hPenOld As Long
'        hPen = ApiCreatePen(PS_STYLE_MASK Or PS_SOLID, 1, vbBlack)
'If hPen <> 0 Then
'    hPenOld = SelectObject(hdcPrinter, hPen)
'End If
'
'    ' start a document
'    docinf.cbSize = Len(docinf)  ' Size of DOCINFO structure
'            docinf.fwType = 2 '3
'            docinf.lpszDatatype = "NT EMF 1.008"
'            docinf.lpszDocName = "Doc 1"
'            docinf.lpszOutput = ""
'    PrintJobID = StartDoc(hdcPrinter, docinf)   'Start new document
'    '///////////////////Page-1 /////////////////////////////////
'    iret = StartPage(hdcPrinter)  'Start a new page
'
'         SetTextAlign hdcPrinter, TA_LEFT Or TA_TOP Or TA_NOUPDATECP
'
'    Rectangle hdcPrinter, 10, 10, 100, 100
'
''        Rectangle hdcPrinter, 1 * (DPIX - OffsetX), 1 * (DPIY - OffsetY), 10 * (DPIX - OffsetX), 10 * (DPIY - OffsetY)
'        xpos = (10 / DPIX_Display) * DPIX
'        ypos = (10 / DPIY_Display) * DPIY
'        W = (200 / DPIX_Display) * DPIX
'        H = (200 / DPIY_Display) * DPIY
'    Rectangle hdcPrinter, xpos, ypos, W, H
'        xpos = (205 / DPIX_Display) * DPIX
'        ypos = (10 / DPIY_Display) * DPIY
'        W = xpos + ((340 / DPIX_Display) * DPIX)
'        H = ((200 / DPIY_Display) * DPIY)
'    Rectangle hdcPrinter, xpos, ypos, W, H
'    Rectangle hdcPrinter, 100, 100, 1000, 1000
'        Dim p As POINTAPI
'            MoveToEx hdcPrinter, xpos, ypos, ByVal 0&
'            LineTo hdcPrinter, W, H
'    ' print a simple line of text at position (1, 1) (inches)
'    For n = 1 To 10
''      PrinterText "This is Line " & Format(n), 1, 1 * 0.16 * n
'        str = "This is Line " & Format(n)
''        xpos = 1 * (DPIX - OffsetX)
''        ypos = 1 * 0.16 * n * (DPIY - OffsetY)
''        xpos = 1 * (DPIX_Display)
''        ypos = 1 * n * (DPIY_Display)
'        xpos = (12 / DPIX_Display) * DPIX
'        ypos = (((12 * n) / DPIY_Display) * DPIY)
'
'        TextOut hdcPrinter, xpos, ypos, str, Len(str)
'    Next n
'    ' end page
'    iret = EndPage(hdcPrinter)  'End the page
'    DeleteObject New_Font  ' clear up the new font
'    '///////////////////Page-2 /////////////////////////////////
'    iret = StartPage(hdcPrinter)  'Start a new page
'
''    '//Change font again
''    log_font.lfFaceName = "Courier New" & vbNullChar
''    New_Font = CreateFontIndirect(log_font)
''    iret = SelectObject(hdcPrinter, New_Font)
'
'    ' print a simple line of text at position (1, 1) (inches)
'    For n = 1 To 10
''      PrinterText "This is Line " & Format(n), 1, 1 * 0.16 * n
'        str = "This is Line " & Format(n)
''        xpos = 1 * (DPIX - OffsetX)
''        ypos = 1 * 0.16 * n * (DPIY - OffsetY)
'        xpos = 1 * (DPIX_Display)
'        ypos = 1 * n * (DPIY_Display)
'
'        TextOut hdcPrinter, xpos, ypos, str, Len(str)
'    Next n
'    iret = EndPage(hdcPrinter)  'End the page
'    DeleteObject New_Font  ' clear up the new font
'    '////////////////////////////////////////////////////////////
'    ' end the document
'    iret = EndDoc(hdcPrinter)  'End the print job
'    iret = SelectObject(hdcPrinter, old_font)
'        If hPenOld <> 0 Then
'            DeleteObject SelectObject(hdcPrinter, hPenOld)
'        End If
'
'            DeleteDC hdcPrinter
'
'ClosePrinter hPrinter
End Sub
Public Property Get XScale(ByVal x As Long) As Long
    XScale = (x / DPIX_Display) * (DPIX - OffsetX)
End Property
Public Property Get YScale(ByVal y As Long) As Long
    YScale = (y / DPIY_Display) * (DPIY - OffsetY)
End Property

Public Property Get Width() As Long
    Width = (XPrintableArea * DPIX_Display) / (DPIX - OffsetX)
End Property
Public Property Get Height() As Long
    Height = (YPrintableArea * DPIY_Display) / (DPIY - OffsetY)
End Property
Private Sub PrinterInit()
Dim hdcDisplay As Long
    hdcDisplay = CreateDC("DISPLAY", "", "", ByVal 0&)
    DPIX = GetDeviceCaps(hdcPrinter, LOGPIXELSX) '      // x dpi
    DPIY = GetDeviceCaps(hdcPrinter, LOGPIXELSY) '      // y dpi
    BitsPerPixel = GetDeviceCaps(hdcPrinter, BITSPIXEL) '       // bit per pixel
    XPrintableArea = GetDeviceCaps(hdcPrinter, HORZRES) '         // x printable area in pixels. 0 maps to 0 here
    YPrintableArea = GetDeviceCaps(hdcPrinter, VERTRES) '         // y printable area in pixels. 0 maps to 0 here
    OffsetX = GetDeviceCaps(hdcPrinter, PHYSICALOFFSETX)
    OffsetY = GetDeviceCaps(hdcPrinter, PHYSICALOFFSETY)
        DPIX_Display = GetDeviceCaps(hdcDisplay, LOGPIXELSX)
        DPIY_Display = GetDeviceCaps(hdcDisplay, LOGPIXELSY)

If DPIX = 0 Then DPIX = DPIX_Display
If DPIY = 0 Then DPIY = DPIY_Display

        pScaleX = Max(DPIX_Display, DPIX) / Min(DPIX_Display, DPIX)
        pScaleY = Max(DPIY_Display, DPIY) / Min(DPIY_Display, DPIY)

            DeleteDC hdcDisplay
    'A5 SIze 5.83in W x 8.27in H
    'XPrintableArea/DPIX = 5.82666666666667     3496/600
    'yPrintableArea/DPIy = 8.26833333333333     4961/600
    'Convert pixels to inches (Output to Monitors/Printers)  100 Pixels
    '           Formula: Pixels ÷ DPI = Inches
    '                   eg: 100 Pixels/96 DPI  = 1.04 Inches
    '                  and: 100 Pixels/600 DPI = 0.17 Inches
    '   Exact Conversion = (Pixels / DPIX_Display) * DPIX   = 10/96*600 =62.5
    'Convert inches to Pixels (Input from Scanners)
    '           Formula: Inches X DPI = Pixels
    '                   eg:5.82666666666667*96=559.36   or 3496/600*96=559.36
  Dim nHasPalette         As Long
  Dim nPaletteSize        As Long
  Dim LogPal              As LOGPALETTE
  Dim tm                  As TEXTMETRIC
  Dim sFaceName           As String * 80
  Dim oFont               As StdFont
  Dim m_MemoryPal          As Long
   If hdcParent = 0 Then
      hdcParent = GetDC(0)
   End If
   ' Get screen properties
   nHasPalette = GetDeviceCaps(hdcParent, RASTERCAPS) And RC_PALETTE   ' Palette support
   nPaletteSize = GetDeviceCaps(hdcParent, SIZEPALETTE)                ' Size of palette
   ' If the screen has a palette make a copy and realize it

   If nHasPalette And (nPaletteSize = 256) Then
      ' Create a copy of the system palette
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      Call GetSystemPaletteEntries(hdcParent, 0&, 256, LogPal.palPalEntry(0))
      hPalette = CreatePalette(LogPal)
      hPaletteOld = SelectPalette(hdcPrinter, m_MemoryPal, 0&)
      Call RealizePalette(hdcPrinter)
   End If

      '--- Set attributes. Take from parent.
      Call SetBkColor(hdcPrinter, GetBkColor(hdcParent))
      Call SetTextColor(hdcPrinter, GetTextColor(hdcParent))
      Call SetBkMode(hdcPrinter, GetBkMode(hdcParent))
        SetTextAlign hdcPrinter, TA_LEFT Or TA_TOP Or TA_NOUPDATECP
        SetBkMode hdcPrinter, TRANSPARENT

      Call GetTextMetrics(hdcParent, tm)
      Call GetTextFace(hdcParent, 79, sFaceName)
      Set oFont = New StdFont
'
      With oFont
         .Bold = (tm.tmWeight > FW_NORMAL)
         .Charset = tm.tmCharSet
         .Italic = (tm.tmItalic <> 0)
         .Name = sFaceName
         .Strikethrough = (tm.tmStruckOut <> 0)
         .Underline = (tm.tmUnderlined <> 0)
         .Weight = tm.tmWeight
         .Size = (tm.tmHeight - tm.tmInternalLeading) * 72 / tm.tmDigitizedAspectY
      End With

           Set Font = oFont
      Set oFont = Nothing
    Dim log_font As LOGFONT, New_Font As Long, old_font As Long
    Dim str As String

'    With log_font
'      .lfEscapement = 0  ' desired rotation in tenths of a degree
'      .lfHeight = 12 * (-DPIY / 72)   ' 12 points
'      .lfFaceName = "Verdana" & vbNullChar
'      .lfWeight = 400   ' standard (bold = 700)
'      .lfItalic = False
'      .lfUnderline = False
'    End With
'    New_Font = CreateFontIndirect(log_font)
'    old_font = SelectObject(hdcPrinter, New_Font)
        hPen = ApiCreatePen(PS_STYLE_MASK Or PS_SOLID, 1, vbBlack)
If hPen <> 0 Then
    hPenOld = SelectObject(hdcPrinter, hPen)
End If
        
End Sub
Private Sub PrinterReset()
If hPrinter = 0 Then
    PrinterCreate
End If
If hPrinter = 0 Then
    Exit Sub
End If
Dim lResult As Long
Dim pDevMod2 As DEVMODE
Dim bufsize As Long
Dim dmInBuf() As Byte
Dim dmOutBuf() As Byte
    bufsize = DocumentProperties(hwndParent, hPrinter, sPrinter, 0, 0, 0)
ReDim dmInBuf(bufsize)
ReDim dmOutBuf(bufsize)
'lResult = DocumentProperties(hwndParent, hPrinter, sPrinter, ByVal VarPtr(dmOutBuf(0)), ByVal VarPtr(dmInBuf(0)), DM_OUT_BUFFER Or DM_IN_BUFFER)
'lResult = DocumentProperties(hwndParent, hPrinter, sPrinter, ByVal VarPtr(dmOutBuf(0)), ByVal VarPtr(dmInBuf(0)), DM_OUT_BUFFER)
lResult = DocumentProperties(hwndParent, hPrinter, sPrinter, ByVal VarPtr(dmOutBuf(0)), ByVal 0&, DM_OUT_BUFFER)

    CopyMemory pDevMod2, pDevMode, Len(pDevMode)
     
CopyMemory pDevMode, dmOutBuf(0), Len(pDevMode)
    pDevMode.dmFields = pDevMod2.dmFields Or DM_PAPERSIZE Or DM_DUPLEX Or _
        DM_ORIENTATION Or DM_COLLATE Or DM_BITSPERPEL Or DM_COLOR Or _
        DM_COPIES Or DM_DEFAULTSOURCE Or DM_DISPLAYFREQUENCY Or _
        DM_PRINTQUALITY Or DM_PAPERWIDTH Or DM_SCALE
    With pDevMode
        .dmBitsPerPel = pDevMod2.dmBitsPerPel
        .dmCollate = pDevMod2.dmCollate
        .dmColor = pDevMod2.dmColor
        .dmCopies = pDevMod2.dmCopies
        .dmDefaultSource = pDevMod2.dmDefaultSource
        .dmDisplayFrequency = pDevMod2.dmDisplayFrequency
        .dmDuplex = pDevMod2.dmDuplex
        .dmOrientation = pDevMod2.dmOrientation
        .dmPaperLength = pDevMod2.dmPaperLength
        .dmPaperSize = pDevMod2.dmPaperSize
        .dmPaperWidth = pDevMod2.dmPaperWidth
        .dmPrintQuality = pDevMod2.dmPrintQuality
        .dmScale = pDevMod2.dmScale
    End With
    CopyMemory dmOutBuf(0), pDevMode, Len(pDevMode)
'hdcPrinter = ResetDC(hdcPrinter, pDevMode)
    
lResult = DocumentProperties(hwndParent, hPrinter, sPrinter, ByVal VarPtr(dmOutBuf(0)), ByVal VarPtr(dmOutBuf(0)), DM_OUT_BUFFER Or DM_IN_BUFFER)

'    lResult = ResetDC(hdcPrinter, dmOutBuf(0))
    lResult = ResetDC(hdcPrinter, ByVal VarPtr(dmOutBuf(0)))
If lResult <> 0 Then
    Call CopyMemory(pDevMode, dmOutBuf(0), Len(pDevMode))
    'check h mhDCPrinter has changed
    If lResult <> hdcPrinter Then
      Call DeleteDC(hdcPrinter)
      hdcPrinter = lResult
    End If
End If
        PrinterInit
End Sub
'*************** End Class Functions ************************************************

'*********************** Begin Drawing Function **********************************
Public Property Get hDc() As Long
    hDc = hdcPrinter
End Property
Sub Cls(x As Long, y As Long, W As Long, H As Long, clrFill As Long)
'   FillRect 0, 0, Width, Height, clrFill, hbrFill
  Dim rc          As RECT
    Dim hbrFill As Long
   If clrFill <> -1 Then
      hbrFill = CreateSolidBrush(TranslateColor(clrFill))
   End If

   With rc
      .Left = ((x / DPIX_Display) * (DPIX - OffsetX))
      .Top = (y / DPIY_Display) * (DPIY - OffsetY)
      .Right = .Left + ((W / DPIX_Display) * (DPIX - OffsetX))
      .Bottom = .Top + ((H / DPIY_Display) * (DPIY - OffsetY))
   End With

   Call APIFillRect(hdcPrinter, rc, hbrFill)
   '--- cleanup the brush (if neccessary)

   If clrFill <> -1 Then
      Call ApiDeleteObject(hbrFill)
   End If

End Sub

Sub DrawLine(ByVal X1 As Long, _
                    ByVal Y1 As Long, _
                    ByVal X2 As Long, _
                    ByVal Y2 As Long, _
                    Optional ByVal clrLine As Long = -1, _
                    Optional ByVal lStyle As PrintPenStyles = PS_SOLID, _
                    Optional ByVal lWidth As Long = 1)
Dim hPen As Long, hPenOld As Long
    If hPen <> 0 Then
        DeleteObject SelectObject(hdcPrinter, hPenOld)
        hPen = 0
    End If
   If clrLine <> -1 Then
'      hPen = ApiCreatePen(PS_SOLID, 1, TranslateColor(clrLine))
      hPen = ApiCreatePen(lStyle, lWidth, TranslateColor(clrLine))
      hPenOld = SelectObject(hdcPrinter, hPen)
   End If
   '--- draw line
   If X1 >= 0 Then
      MoveToEx hdcPrinter, ((X1 / DPIX_Display) * (DPIX - OffsetX)), ((Y1 / DPIY_Display) * (DPIY - OffsetY)), 0
   End If
   LineTo hdcPrinter, ((X2 / DPIX_Display) * (DPIX - OffsetX)), ((Y2 / DPIY_Display) * (DPIY - OffsetY))

End Sub
Public Property Let DrawMode(ByVal dwValue As DrawModeConstants)
    Call SetROP2(hdcPrinter, dwValue)
End Property
Public Property Get Font() As StdFont
'Purpose: Returns or sets the font currently selected in the <b>cMemDC</b> object.
  Dim tm              As TEXTMETRIC
  Dim sFaceName       As String * 80
   GetTextMetrics hdcPrinter, tm
   GetTextFace hdcPrinter, 79, sFaceName
 Dim fnt As New StdFont
   With fnt
      .Name = sFaceName 'StrConv(sFaceName, vbUnicode)
      .Bold = (tm.tmWeight >= FW_NORMAL)
      .Charset = tm.tmCharSet
      .Italic = (tm.tmItalic <> 0)
      .Strikethrough = (tm.tmStruckOut <> 0)
      .Underline = (tm.tmUnderlined <> 0)
      .Weight = tm.tmWeight
      .Size = (tm.tmHeight - tm.tmInternalLeading) * 72 / tm.tmDigitizedAspectY
   End With
        Set Font = fnt
End Property
Public Property Set Font(ByVal oValue As StdFont)
  Dim tFont           As LOGFONT
   With tFont
      CopyMemory .lfFaceName(1), ByVal oValue.Name, Len(oValue.Name) + 1
      .lfCharSet = oValue.Charset
      .lfItalic = (-oValue.Italic)
      .lfStrikeOut = (-oValue.Strikethrough)
      .lfUnderline = (-oValue.Underline)
      .lfWeight = oValue.Weight
      .lfHeight = -(oValue.Size * GetDeviceCaps(hdcPrinter, LOGPIXELSY) / 72)
   End With

   If hFont <> 0 Then
        DeleteObject SelectObject(hdcPrinter, hFontOld)
   End If
    hFont = CreateFontIndirect(tFont)
   hFontOld = SelectObject(hdcPrinter, hFont)
End Property
Public Property Get FontHeight() As Long
'Purpose: Returns the height of the currently selected fron in a <b>cMemDC</b> object in pixels.
  Dim tm As TEXTMETRIC
   GetTextMetrics hdcPrinter, tm
     FontHeight = tm.tmHeight
End Property
Public Property Get ForeColor() As Long
    ForeColor = GetTextColor(hdcPrinter)
End Property
Public Property Let ForeColor(ByVal NewValue As Long)
      SetTextColor hdcPrinter, TranslateColor(NewValue)
End Property

Sub PaintPicture(ByVal oPic As StdPicture, _
                        Optional ByVal xDest As Long, _
                        Optional ByVal yDest As Long, _
                        Optional ByVal nWidth As Long, _
                        Optional ByVal nHeight As Long, _
                        Optional ByVal xSrc As Long, _
                        Optional ByVal ySrc As Long, _
                        Optional ByVal dwRop As RasterOpConstants = vbSrcCopy)
'Purpose: Paints a <b>StdPicture</b> objects in a rectangle in the device context using specified raster operation or skipping mask color pixels.

  Dim rc              As RECT
  Dim hdcPaint        As Long
  Dim hbmOrig         As Long
  Dim hEmf            As Long
  Dim srcW As Long, srch As Long
      
      srcW = HM2Pix_X(oPic.Width)
   If nWidth = 0 Then
        nWidth = srcW
   End If
      
      srch = HM2Pix_Y(oPic.Height)
   If nHeight = 0 Then
        nHeight = srch
   End If
xDest = ((xDest / DPIX_Display) * (DPIX - OffsetX))
yDest = ((yDest / DPIY_Display) * (DPIY - OffsetY))
nWidth = ((nWidth / DPIX_Display) * (DPIX - OffsetX))
nHeight = ((nHeight / DPIY_Display) * (DPIY - OffsetY))
'        If Not oPic Is Nothing Then
'        '        .PaintPicture pic, .ScaleX(X, vbTwips, .ScaleMode), .ScaleY(Y, vbTwips, .ScaleMode), .ScaleX(W, vbTwips, .ScaleMode), .ScaleY(H, vbTwips, .ScaleMode)
'            If oPic.handle <> 0 Then
'                Dim l As Long, t As Long ', hDc As Long
'                Dim w As Long, h As Long
'        '            l = .ScaleX(x, vbTwips, vbPixels)
'        '            t = .ScaleY(y, vbTwips, vbPixels)
'        '            w = .ScaleX(W, vbTwips, vbPixels)
'        '            h = .ScaleY(H, vbTwips, vbPixels)
'                    l = xDest
'                    t = yDest
'                    w = nWidth
'                    h = nHeight
'
'        '            If pic.Width > pic.Height Then
'        '                T = T + 20
'        '                h = h - 40
'        '            Else
'        '                L = L + 20
'        '                w = w - 40
'        '            End If
'                    oPic.Render --hdcPrinter, --l, --t, --w, --h, 0, oPic.Height, oPic.Width, -oPic.Height, ByVal 0&
'            End If
'        End If
Dim hdcTemp As Long, hBmpTemp As Long, hBmpTempOld As Long
Dim hBrushTemp As Long, hBrushTempOld As Long, LB As LOGBRUSH
Dim tmpRect As RECT
'        hdcTemp = CreateCompatibleDC(hdcParent)
'        hBmpTemp = CreateCompatibleBitmap(hdcParent, srcW, srcH)
        hdcTemp = CreateCompatibleDC(hdcPrinter)
        hBmpTemp = CreateCompatibleBitmap(hdcPrinter, srcW, srch)
        
        hBmpTempOld = SelectObject(hdcTemp, hBmpTemp)
'            SetBkColor hdcTemp, GetBkColor(hdcPrinter)
            SetBkColor hdcTemp, -1
            SetBkMode hdcTemp, TRANSPARENT
            SetROP2 hdcTemp, GetROP2(hdcPrinter)
'''            SetStretchBltMode hdcTemp, 4  'half tone
'        hBrushTemp = CreateSolidBrush(GetBkColor(hdcPrinter))
'        hBrushTemp = CreateSolidBrush(-1)
        hBrushTemp = CreateSolidBrush(vbMagenta)
''            LB.lbColor = GetBkColor(hdcPrinter)
''            LB.lbHatch = 0
''            LB.lbStyle = 1 ' Private Const BS_NULL = 1'Private Const BS_SOLID = 0
''        hBrushTemp = CreateBrushIndirect(LB)
''        hBrushTemp = CreateHatchBrush(1, vbRed)
                tmpRect.Left = 0: tmpRect.Top = 0
                tmpRect.Right = srcW: tmpRect.Bottom = srch
            FillRect hdcTemp, tmpRect, hBrushTemp
        If hBrushTemp <> 0 Then DeleteObject hBrushTemp
    Select Case oPic.Type
       Case vbPicTypeIcon
'    '      DrawIconEx hdcTemp, xDest, yDest, oPic.handle, nWidth, nHeight, 0, 0, DI_NORMAL
          DrawIconEx hdcTemp, 0, 0, oPic.Handle, srcW, srch, 0, 0, DI_NORMAL
            
       Case vbPicTypeBitmap
'          hdcPaint = CreateCompatibleDC(hdcPrinter)
          hdcPaint = CreateCompatibleDC(hdcTemp)
          hbmOrig = SelectObject(hdcPaint, oPic.Handle)
    
    '         Call APIBitBlt(hdcPrinter, xDest, yDest, nWidth, nHeight, hdcPaint, xSrc, ySrc, dwRop)
            BitBlt hdcTemp, 0, 0, srcW, srch, hdcPaint, 0, 0, dwRop
    
          Call SelectObject(hdcPaint, hbmOrig)
          Call DeleteDC(hdcPaint)
       Case vbPicTypeEMetafile, vbPicTypeMetafile
    '      rc.Left = xDest
    '      rc.Top = yDest
    '      rc.Right = xDest + nWidth
    '      rc.Bottom = yDest + nHeight
          rc.Left = 0
          rc.Top = 0
          rc.Right = srcW
          rc.Bottom = srch
    
          If oPic.Type = vbPicTypeMetafile Then
    '         hdcPaint = CreateEnhMetaFileLong(hdcPrinter, vbNullString, 0, vbNullString)
             hdcPaint = CreateEnhMetaFileLong(hdcTemp, vbNullString, 0, vbNullString)
             Call PlayMetaFile(hdcPaint, oPic.Handle)
             hEmf = CloseEnhMetaFile(hdcPaint)
          Else
             hEmf = oPic.Handle
          End If
    
    '      Call PlayEnhMetaFile(hdcPrinter, hEmf, rc)
          Call PlayEnhMetaFile(hdcTemp, hEmf, rc)
    
          If oPic.Type = vbPicTypeMetafile Then
             Call DeleteEnhMetaFile(hEmf)
          End If
   End Select
'Dim bi As BITMAPINFO, bm As BITMAP
'Dim di As Long, bufsize As Long, ghnd As Long, gptr As Long
''     di = GetObjectAPI(oPic.handle, Len(bm), bm)
'
'     di = GetObjectAPI(hBmpTemp, Len(bm), bm)
'
'    ' Can this printer handle the DIB?
'    If (GetDeviceCaps(hdcPrinter, RASTERCAPS)) And RC_DIBTODEV = 0 Then
'        MsgBox "This device does not support DIB's" + vbCrLf + "See source code for further info", 0, "No DIB support"
'    End If
'
'    ' Fill the BITMAPINFO for the desired DIB
'    bi.bmiHeader.biSize = Len(bi.bmiHeader)
'    bi.bmiHeader.biWidth = bm.bmWidth
'    bi.bmiHeader.biHeight = bm.bmHeight
'    bi.bmiHeader.biPlanes = 1
'    ' Set to 24 here to create a 24 bit DIB
'    ' Set to 8 here to create an 8 bit DIB
'    bi.bmiHeader.biBitCount = 4
'    bi.bmiHeader.biCompression = BI_RGB
'    ' Now calculate the data buffer size needed
'    bufsize = bi.bmiHeader.biWidth
'
'    ' Figure out the number of bytes based on the
'    ' number of pixels in each byte. In this case we
'    ' really don't need all this code because this example
'    ' always uses a 16 color DIB, but the code is shown
'    ' here for your future reference
'    Select Case bi.bmiHeader.biBitCount
'        Case 1
'            bufsize = (bufsize + 7) / 8
'        Case 4
'            bufsize = (bufsize + 1) / 2
'        Case 24
'            bufsize = bufsize * 3
'    End Select
'    ' And make sure it aligns on a long boundary
'    bufsize = ((bufsize + 3) / 4) * 4
'    ' And multiply by the # of scan lines
'    bufsize = bufsize * bi.bmiHeader.biHeight
'
'    ' Now allocate a buffer to hold the data
'    ' We use the global memory pool because this buffer
'    ' could easily be above 64k bytes.
'    ghnd = GlobalAlloc(GMEM_MOVEABLE, bufsize)
'    gptr = GlobalLock(ghnd)
'
''    di = GetDIBits(dctemp, pic.handle, 0, bm.bmHeight, ByVal gptr&, bi, DIB_RGB_COLORS)
''    di = GetDIBits(hdcTemp, oPic.handle, 0, bm.bmHeight, ByVal gptr, bi, DIB_RGB_COLORS)
'    di = GetDIBits(hdcTemp, hBmpTemp, 0, bm.bmHeight, ByVal gptr, bi, DIB_RGB_COLORS)
'
'        ' print over the 1:1 scaled bitmap
''    'di = SetDIBitsToDevice(hdc, PointX, PointY, bm.bmWidth, bm.bmHeight, 0, 0, 0, bm.bmHeight, ByVal gptr&, bi, DIB_RGB_COLORS)
'    di = StretchDIBits(hdcPrinter, xDest, yDest, nWidth, nHeight, 0, 0, bm.bmWidth, bm.bmHeight, ByVal gptr, bi, DIB_RGB_COLORS, dwRop)
'    ' Dump the global memory block
'    di = GlobalUnlock(ghnd)
'    di = GlobalFree(ghnd)

'    StretchBlt hdcPrinter, xDest, yDest, nWidth, nHeight, hdcTemp, 0, 0, srcW, srcH, dwRop
    TransparentBlt hdcPrinter, xDest, yDest, nWidth, nHeight, hdcTemp, 0, 0, srcW, srch, vbMagenta '-1 'vbBlack
         
         
         DeleteObject SelectObject(hdcTemp, hBmpTempOld)
        DeleteDC hdcTemp
End Sub

Function HM2Pix_X(ByVal Value As Double) As Double
   HM2Pix_X = Value * 1440 / 2540 / Screen.TwipsPerPixelX
End Function
Function HM2Pix_Y(ByVal Value As Double) As Double
   HM2Pix_Y = Value * 1440 / 2540 / Screen.TwipsPerPixelY
End Function
'Purpose: Draws a text using specified drawing flags in a rectangle in the device context.
Sub DrawText(ByVal Text As String, _
                    ByRef x As Long, _
                    ByRef y As Long, _
                    ByRef W As Long, _
                    ByRef H As Long, _
                    Optional ByVal dtFlags As PrintDrawTextStyles)

  Dim wTextParams     As DRAWTEXTPARAMS
  Dim rc              As RECT
            With rc
               .Left = ((x / DPIX_Display) * (DPIX - OffsetX))
               .Top = (y / DPIY_Display) * (DPIY - OffsetY)
               .Right = .Left + ((W / DPIX_Display) * (DPIX - OffsetX))
               .Bottom = .Top + ((H / DPIY_Display) * (DPIY - OffsetY))
            End With
        wTextParams.cbSize = Len(wTextParams)
    ApiDrawTextEx hdcPrinter, Text, -1, rc, dtFlags, wTextParams
End Sub
Sub PrintText(ByVal Text As String, _
                    ByVal x As Long, _
                    ByVal y As Long)
    x = ((x / DPIX_Display) * (DPIX - OffsetX))
    y = (y / DPIY_Display) * (DPIY - OffsetY)
    TextOut hdcPrinter, x, y, Text, Len(Text)
End Sub
'Purpose: Calculates the height of a text string using currently selected font settings.
Public Property Get TextHeight(ByVal sText As String) As Long
  Dim sz              As SIZEAPI
        GetTextExtentPoint hdcPrinter, sText, Len(sText), sz
        '--- success
    sz.cY = (sz.cY * DPIY_Display) / (DPIY - OffsetY)

         TextHeight = sz.cY
End Property

'Purpose: Calculates the width of a text string using currently selected font settings.
Public Property Get TextWidth(ByVal sText As String) As Long
  Dim sz              As SIZEAPI
        GetTextExtentPoint hdcPrinter, sText, Len(sText), sz
        '--- success
    sz.cX = (sz.cX * DPIX_Display) / (DPIX - OffsetX)
         TextWidth = sz.cX
End Property

'Purpose: Fills a rectangle in the device context optionally using a specified color for pen and brush.
Sub Rectangle(ByVal x As Long, _
                       ByVal y As Long, _
                       ByVal W As Long, _
                       ByVal H As Long, _
                     Optional ByVal clrFill As OLE_COLOR = -1, _
                     Optional ByVal clrOutline As OLE_COLOR = -1)
Dim hbrFill As Long, hbrFillOld As Long, hpnOutline As Long, hPenOutlineOld As Long
Dim rc As RECT
            With rc
               .Left = ((x / DPIX_Display) * (DPIX - OffsetX))
               .Top = (y / DPIY_Display) * (DPIY - OffsetY)
               .Right = .Left + ((W / DPIX_Display) * (DPIX - OffsetX))
               .Bottom = .Top + ((H / DPIY_Display) * (DPIY - OffsetY))
            End With

        '--- create brush if neccessary
        If clrFill <> -1 Then
            hbrFill = CreateSolidBrush(TranslateColor(clrFill))
        End If
        If hbrFillOld <> 0 Then
            hbrFillOld = SelectObject(hdcPrinter, hbrFillOld)
        End If
        '--- create pen if neccessary
        If clrOutline <> -1 Then
            hpnOutline = ApiCreatePen(PS_SOLID, 1, TranslateColor(clrOutline))
        End If
        If hpnOutline <> 0 Then
            hPenOutlineOld = SelectObject(hdcPrinter, hpnOutline)
        End If
        Call APIRectangle(hdcPrinter, rc.Left, rc.Top, rc.Right, rc.Bottom)
    If hbrFillOld <> 0 Then
        ApiDeleteObject SelectObject(hdcPrinter, hbrFillOld)
    End If
    If hPenOutlineOld <> 0 Then
        ApiDeleteObject SelectObject(hdcPrinter, hPenOutlineOld)
    End If
End Sub

'Purpose: Sets a clipping rectangle for the device context. Pixels outside clipping rectangle are not modified by drawing operations.
Sub SetClipRect(ByVal x As Long, _
                       Optional ByVal y As Long, _
                       Optional ByVal W As Long, _
                       Optional ByVal H As Long)

Dim hRgn        As Long
Dim rc As RECT
            With rc
               .Left = ((x / DPIX_Display) * (DPIX - OffsetX))
               .Top = (y / DPIY_Display) * (DPIY - OffsetY)
               .Right = .Left + ((W / DPIX_Display) * (DPIX - OffsetX))
               .Bottom = .Top + ((H / DPIY_Display) * (DPIY - OffsetY))
            End With
        hRgn = CreateRectRgn(rc.Left, rc.Top, rc.Right, rc.Bottom)
    SelectClipRgn hdcPrinter, hRgn
    If hRgn <> 0 Then
        ApiDeleteObject hRgn
    End If

End Sub
'Purpose: Translates an OLE_COLOR value to a Long value which represents an RGB tripple.
Public Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional m_MemoryPal As Long) As Long

  '--- handle invalid (none) color

    If clrColor = -1 Then
        TranslateColor = -1
      Else
        Call OleTranslateColor(clrColor, m_MemoryPal, TranslateColor)
        TranslateColor = TranslateColor And &HFFFFFF
    End If

End Function

Sub PaintTo(DestHdc As Long, W As Long, H As Long)
'    StretchBlt DestHdc, 0, 0, W, H, hdcPrinter, 0, 0, Width, Height, vbSrcCopy
    StretchBlt DestHdc, 0, 0, W, H, hdcPrinter, 0, 0, XPrintableArea, YPrintableArea, vbSrcCopy
End Sub

'*********************** End Drawing Function **********************************

Private Function StripNulls(OriginalStr As String) As String
On Error Resume Next
  If (InStr(OriginalStr, Chr$(0)) > 0) Then
    OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr$(0)) - 1)
  End If
  StripNulls = Trim$(OriginalStr)
End Function

Private Function ByteToString(ByteArray() As Byte) As String
On Error Resume Next
  Dim TempStr As String
  Dim i As Integer

  For i = 1 To CCHDEVICENAME
    TempStr = TempStr & Chr$(ByteArray(i))
  Next i
  ByteToString = StripNulls(TempStr)
End Function

Private Function PtrCtoVbString(ByVal Add As Long) As String
Dim sTemp As String * 512, x As Long
On Error Resume Next
  x = lstrcpy(sTemp, ByVal Add)
  If (InStr(1, sTemp, Chr$(0)) = 0) Then
    PtrCtoVbString = ""
  Else
    PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr$(0)) - 1)
  End If
End Function
Function SetPrinterDefault(ByVal PrinterName As String) As Boolean
Dim ret As Long
ret = SetDefaultPrinterA(PrinterName)
    SetPrinterDefault = (ret = 1)
End Function

Function PrinterDefault() As String
Dim str As String, cb As Long, ret As Long
    ret = GetDefaultPrinterA("", cb)
        str = Space$(cb - 1) '-1 for Null Char
    ret = GetDefaultPrinterA(str, cb)
If GetLastError = ERROR_FILE_NOT_FOUND Then 'There is no default printer.
    str = ""
End If
    PrinterDefault = str
End Function
Function GetPrinteres(numprinters As Long, D() As String) As String()
    Dim longbuffer() As Long  ' resizable array receives information from the function
'    Dim printinfo  As PRINTER_INFO_1  ' values inside longbuffer() will be put into here
    Dim numbytes As Long  ' size in bytes of longbuffer()
    Dim numneeded As Long  ' receives number of bytes necessary if longbuffer() is too small
'    Dim numprinters As Long  ' receives number of printers found
    Dim c As Long, RetVal As Long  ' counter variable & return value
    Dim A() As String
 
    ' Get information about the local printers
    numbytes = 0 ' 3076  ' should be sufficiently big, but it may not be
    ReDim longbuffer(0 To numbytes / 4) As Long  ' resize array -- note how 1 Long = 4 bytes
    RetVal = EnumPrinters(PRINTER_ENUM_LOCAL Or PRINTER_ENUM_CONNECTIONS, 0, 1, ByVal 0&, 0, numneeded, numprinters)
    If RetVal = 0 Then  ' try enlarging longbuffer() to receive all necessary information
        numbytes = numneeded
        ReDim longbuffer(0 To numbytes / 4) As Long  ' make it large enough
        RetVal = EnumPrinters(PRINTER_ENUM_LOCAL Or PRINTER_ENUM_CONNECTIONS, 0, 1, longbuffer(0), numbytes, numneeded, numprinters)
        If RetVal = 0 Then ' failed again!
'            Debug.Print "Could not successfully enumerate the printes."
        End  ' abort program
    End If
    End If
    ' Convert longbuffer() data into printinfo()
    If numprinters <> 0 Then
'        ReDim printinfo(0 To numprinters - 1) As PRINTER_INFO_1 ' room for each printer
        ReDim A(numprinters - 1)
        ReDim D(numprinters - 1)
    End If
    For c = 0 To numprinters - 1  ' loop, putting each set of information into each element
        ' longbuffer(4 * c) = .flags, longbuffer(4 * c + 1) = .pDescription, etc.
        ' For each string, the string is first buffered to provide enough room, and then the string is copied.
'        printinfo(c).flags = longbuffer(4 * c)
'        printinfo(c).pDescription = Space(lstrlen(longbuffer(4 * c + 1)))
'        RetVal = lstrcpy(printinfo(c).pDescription, longbuffer(4 * c + 1))
'        printinfo(c).pName = Space(lstrlen(longbuffer(4 * c + 2)))
'        RetVal = lstrcpy(printinfo(c).pName, longbuffer(4 * c + 2))
'        printinfo(c).pComment = Space(lstrlen(longbuffer(4 * c + 3)))
'        RetVal = lstrcpy(printinfo(c).pComment, longbuffer(4 * c + 3))
            A(c) = Space(lstrlen(longbuffer(4 * c + 2)))
        RetVal = lstrcpy(A(c), longbuffer(4 * c + 2))
            D(c) = Space(lstrlen(longbuffer(4 * c + 1)))
        RetVal = lstrcpy(D(c), longbuffer(4 * c + 1))
        
    Next c
        Erase longbuffer
GetPrinteres = A
    Erase A
End Function
Function PageSetup() As Long
    m_PSD.lStructSize = Len(m_PSD)
'    m_PSD.hwndOwner = Me.hWnd
    m_PSD.hInstance = App.hInstance
    m_PSD.flags = PSD_MARGINS 'Or PSD_SHOWHELP
'    m_PSD.ptPaperSize.x = Printer.Width
'    m_PSD.ptPaperSize.y = Printer.Height
'    m_PSD.rtMargin.Left = 500 '. inch
'    m_PSD.rtMinMargin.Left = 500 '. inch
'    Dim DEVMODE As DEVMODE
    Dim DevName As DEVNAMES_TYPE
    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Long
'     'Set the current orientation and duplex setting
'    DEVMODE.dmDeviceName = sPrinter ' Printer.DeviceName
'    DEVMODE.dmSize = Len(DEVMODE)
'    DEVMODE.dmFields = DM_ORIENTATION Or DM_DUPLEX
'    DEVMODE.dmPaperWidth = pDevMode.dmPaperWidth ' Printer.Width
'    DEVMODE.dmOrientation = pDevMode.dmOrientation ' Printer.Orientation
'    DEVMODE.dmPaperSize = pDevMode.dmPaperSize ' Printer.PaperSize
'    DEVMODE.dmDuplex = pDevMode.dmDuplex ' Printer.Duplex
'    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    m_PSD.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(pDevMode))
    lpDevMode = GlobalLock(m_PSD.hDevMode)
    If lpDevMode > 0 Then
'        CopyMemory ByVal lpDevMode, DEVMODE, Len(DEVMODE)
        CopyMemory ByVal lpDevMode, pDevMode, Len(pDevMode)
        bReturn = GlobalUnlock(m_PSD.hDevMode)
    End If
    
        'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(sPrinter) 'Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Port) 'Len(Printer.Port)
        .wDefault = 0
    End With

'    With Printer
'        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
        DevName.extra = "winspool" & Chr(0) & sPrinter & Chr(0) & Port & Chr(0)
'    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    m_PSD.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(m_PSD.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If
    
    'Show the pagesetup dialog
    If PAGESETUPDLG(m_PSD) Then
         
                 'First get the DevName structure.
        lpDevName = GlobalLock(m_PSD.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree m_PSD.hDevNames

                'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(m_PSD.hDevMode)
'        CopyMemory DEVMODE, ByVal lpDevMode, Len(DEVMODE)
        CopyMemory pDevMode, ByVal lpDevMode, Len(pDevMode)
        bReturn = GlobalUnlock(m_PSD.hDevMode)
        GlobalFree m_PSD.hDevMode
'        sPrinter = UCase$(Left(DEVMODE.dmDeviceName, InStr(DEVMODE.dmDeviceName, Chr$(0)) - 1))
        sPrinter = UCase$(Left(pDevMode.dmDeviceName, InStr(pDevMode.dmDeviceName, Chr$(0)) - 1))
'                CopyMemory pDevMode, DEVMODE, Len(DEVMODE)
            PrinterReset
'        If Printer.DeviceName <> NewPrinterName Then
'            For Each objPrinter In Printers
'                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
'                    Set Printer = objPrinter
'                    'set printer toolbar name at this point
'                End If
'            Next
'        End If
'
'        On Error Resume Next
'        'Set printer object properties according to selections made
'        'by user
'        Printer.Copies = DEVMODE.dmCopies
'        Printer.Duplex = DEVMODE.dmDuplex
'        Printer.Orientation = DEVMODE.dmOrientation
'        Printer.PaperSize = DEVMODE.dmPaperSize
'        Printer.PrintQuality = DEVMODE.dmPrintQuality
'        Printer.ColorMode = DEVMODE.dmColor
'        Printer.PaperBin = DEVMODE.dmDefaultSource
'
'        On Error GoTo 0
    End If
End Function
Public Sub PrinterSetup(Optional PrintFlags As Long)
    '-> Code by Donald Grover
    Dim PrintDlg As PRINTDLG_TYPE
'    Dim DEVMODE As DEVMODE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = hwndParent

    PrintDlg.flags = PrintFlags
'    On Error Resume Next
'    'Set the current orientation and duplex setting
'    DEVMODE.dmDeviceName = Printer.DeviceName
'    DEVMODE.dmSize = Len(DEVMODE)
'    DEVMODE.dmFields = DM_ORIENTATION Or DM_DUPLEX
'    DEVMODE.dmPaperWidth = Printer.Width
'    DEVMODE.dmOrientation = Printer.Orientation
'    DEVMODE.dmPaperSize = Printer.PaperSize
'    DEVMODE.dmDuplex = Printer.Duplex
'    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(pDevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, pDevMode, Len(pDevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(sPrinter) 'Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Port) 'Len(Printer.Port)
        .wDefault = 0
    End With

'    With Printer
'        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
'    End With
        DevName.extra = "winspool" & Chr(0) & sPrinter & Chr(0) & Port & Chr(0)

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory pDevMode, ByVal lpDevMode, Len(pDevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        sPrinter = UCase$(Left(pDevMode.dmDeviceName, InStr(pDevMode.dmDeviceName, Chr$(0)) - 1))
                PrinterCreate
'        If Printer.DeviceName <> NewPrinterName Then
'            For Each objPrinter In Printers
'                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
'                    Set Printer = objPrinter
'                    'set printer toolbar name at this point
'                End If
'            Next
'        End If
'
'        On Error Resume Next
'        'Set printer object properties according to selections made
'        'by user
'        Printer.Copies = DEVMODE.dmCopies
'        Printer.Duplex = DEVMODE.dmDuplex
'        Printer.Orientation = DEVMODE.dmOrientation
'        Printer.PaperSize = DEVMODE.dmPaperSize
'        Printer.PrintQuality = DEVMODE.dmPrintQuality
'        Printer.ColorMode = DEVMODE.dmColor
'        Printer.PaperBin = DEVMODE.dmDefaultSource
'        On Error GoTo 0
    End If
End Sub

Function PointerToString(pString As Long) As String
  Dim sTemp As String * 512

  lstrcpy sTemp, pString
  If (InStr(1, sTemp, Chr(0)) = 0) Then
    PointerToString = ""
  Else
    PointerToString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
  End If
  Exit Function
End Function

Private Sub Class_Initialize()
    sPrinter = PrinterDefault
        PrinterCreate
End Sub

Private Sub Class_Terminate()
If bPreview = True Then
    Preview = False
End If
    PrinterDestroy
End Sub
Function Max(ByVal x As Double, ByVal y As Double) As Double
    If x > y Then
        Max = x
    Else
        Max = y
    End If
End Function
Function Min(ByVal x As Double, ByVal y As Double) As Double
    If x < y Then
        Min = x
    Else
        Min = y
    End If
End Function
Function PrinterStartDoc(Optional ByVal DocName As String, _
    Optional ByVal sOutPut As String) As Long
If bPreview = True Then
        PrinterStartDoc = 1
    Exit Function
End If
    Dim PrintJobID As Long
    Dim docinf As DOCINFO
    ' set up an initial font
            If DocName = "" Then
                DocName = App.Title
            End If
'            If sOutPut = "" Then
'                sOutPut = App.Title
'            End If
    ' start a document
    docinf.cbSize = Len(docinf)  ' Size of DOCINFO structure
'            docinf.fwType = 2 '3
            docinf.lpszDatatype = vbNullString ' "NT EMF 1.008"
            docinf.lpszDocName = DocName ' "Doc 1"
            docinf.lpszOutput = sOutPut '""
    PrintJobID = StartDoc(hdcPrinter, docinf)   'Start new document
PrinterStartDoc = PrintJobID
End Function
Function PrinterStartPage() As Long
If bPreview = True Then
    Cls 0, 0, XPrintableArea, YPrintableArea, TranslateColor(vbWhite)
    Exit Function
End If
    PrinterStartPage = StartPage(hdcPrinter)
        SetTextAlign hdcPrinter, TA_LEFT Or TA_TOP Or TA_NOUPDATECP
        SetBkMode hdcPrinter, TRANSPARENT
End Function
Function PrinterEndPage() As Long
If bPreview = True Then Exit Function
    PrinterEndPage = EndPage(hdcPrinter)
End Function
Function PrinterEndDoc() As Long
If bPreview = True Then Exit Function
    PrinterEndDoc = EndDoc(hdcPrinter)
End Function
Function PrinterAbortDoc() As Long
If bPreview = True Then Exit Function
    PrinterAbortDoc = AbortDoc(hdcPrinter)
End Function
Public Function PrinterGetForms(Names() As String, SizeX() As Long, SizeY() As Long) As Long
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1 ' Working FI1 array
    Dim temp() As Byte       ' Temp FI1 array
    Dim FormIndex As Integer
    Dim BytesNeeded As Long
    Dim RetVal As Long

    ReDim aFI1(1)
    ' First call retrieves the BytesNeeded.
    RetVal = EnumForms(hPrinter, 1, aFI1(0), 0&, BytesNeeded, NumForms)
    ReDim temp(BytesNeeded)
    ReDim aFI1(BytesNeeded / Len(FI1))
    ' Second call actually enumerates the supported forms.
    RetVal = EnumForms(hPrinter, 1, temp(0), BytesNeeded, BytesNeeded, _
            NumForms)
    Call CopyMemory(aFI1(0), temp(0), BytesNeeded)
        ReDim Names(NumForms - 1)
        ReDim SizeX(NumForms - 1)
        ReDim SizeY(NumForms - 1)
        
    For i = 0 To NumForms - 1
        With aFI1(i)
                Names(i) = PtrCtoVbString(.pName)
                SizeX(i) = .Size.cX
                SizeY(i) = .Size.cY
'                FormIndex = i + 1
        End With
    Next i
    PrinterGetForms = NumForms
End Function
Public Function PrinterFindFormName(ByVal FormIndex As Long) As String
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1 ' Working FI1 array
    Dim temp() As Byte       ' Temp FI1 array
    Dim BytesNeeded As Long
    Dim RetVal As Long
Dim str As String
    ' Get a handle to the printer.
'    If OpenPrinter(Printer.DeviceName, PrinterHandle, 0&) Then
    

    ReDim aFI1(1)
    ' First call retrieves the BytesNeeded.
    RetVal = EnumForms(hPrinter, 1, aFI1(0), 0&, BytesNeeded, NumForms)
    ReDim temp(BytesNeeded)
    ReDim aFI1(BytesNeeded / Len(FI1))
    ' Second call actually enumerates the supported forms.
    RetVal = EnumForms(hPrinter, 1, temp(0), BytesNeeded, BytesNeeded, _
            NumForms)
    Call CopyMemory(aFI1(0), temp(0), BytesNeeded)
    For i = 0 To NumForms - 1
        With aFI1(i)
            If i + 1 = FormIndex Then
                str = PtrCtoVbString(.pName)
                Exit For
            End If
        End With
    Next i
'            ClosePrinter PrinterHandle
'    End If
    PrinterFindFormName = str  ' Returns non-zero when form is found.
End Function
Function PrinterDeleteForm(ByVal FormName As String) As Boolean
'    Dim PrinterHandle As Long ' Handle to printer
'    If OpenPrinter(Printer.DeviceName, PrinterHandle, 0&) Then
        If DeleteForm(hPrinter, FormName) = 1 Then
            PrinterDeleteForm = True
        End If
'            ClosePrinter PrinterHandle
'    End If
End Function
Public Function PrinterAddNewForm(ByVal cX As Long, ByVal cY As Long, _
                           ByVal FormName As String) As String
    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim RetVal As Long
'    Dim PrinterHandle As Long ' Handle to printer
    Dim cbBuf As Long, BytesNeeded As Long
    Dim FormSize As SIZEL
        FormSize.cX = cX
        FormSize.cY = cY
    ' Get a handle to the printer.
'If Printers.count = 0 Then Exit Function
'    If OpenPrinter(Printer.DeviceName, PrinterHandle, 0&) Then
            ReDim aFI1(Len(FI1))
                RetVal = GetForm(hPrinter, FormName, 1, aFI1(0), UBound(aFI1), BytesNeeded)
            If BytesNeeded <> 0 Then
                ReDim aFI1(BytesNeeded)
            End If
                RetVal = GetForm(hPrinter, FormName, 1, aFI1(0), UBound(aFI1), BytesNeeded)
        If RetVal Then
            Dim Fi2 As FORM_INFO_1
            
                    Call CopyMemory(Fi2, aFI1(0), Len(Fi2))
                With Fi2
                    .flags = 0
'                    .pName = FormName
                    With .Size
                        .cX = FormSize.cX
                        .cY = FormSize.cY
                    End With
                    With .ImageableArea
'                        .Left = 0
'                        .Top = 0
                        .Right = Fi2.Size.cX
                        .Bottom = Fi2.Size.cY
                    End With
                End With
                            ReDim aFI1(Len(Fi2))
                    Call CopyMemory(aFI1(0), Fi2, Len(Fi2))
            RetVal = SetForm(hPrinter, FormName, 1, aFI1(0))
            If RetVal = 1 Then
                PrinterAddNewForm = FormName
            Else
                MsgBox "Error: " & Err.LastDllError, , "Error Updating Form"
            End If
        End If
        If RetVal = 0 Then
                With FI1
                    .flags = 0
                    .pName = FormName
                    With .Size
                        .cX = FormSize.cX
                        .cY = FormSize.cY
                    End With
                    With .ImageableArea
                        .Left = 0
                        .Top = 0
                        .Right = FI1.Size.cX
                        .Bottom = FI1.Size.cY
                    End With
                End With
                Call CopyMemory(aFI1(0), FI1, Len(FI1))
            RetVal = AddForm(hPrinter, 1, aFI1(0))
            If RetVal = 0 Then
                If Err.LastDllError = 5 Then
                    MsgBox "You do not have permissions to add a form to " & _
                            sPrinter, vbExclamation, "Access Denied!"
                Else
                    MsgBox "Error: " & Err.LastDllError, , "Error Adding Form"
                End If
                    PrinterAddNewForm = "none"
            Else
                PrinterAddNewForm = FI1.pName
            End If
        End If
'            ClosePrinter PrinterHandle
'    Else
'        PrinterAddNewForm = ""
'    End If
End Function
Public Function PrinterUpdateForm(ByVal cX As Long, ByVal cY As Long, _
                           ByVal FormName As String) As String
    PrinterUpdateForm = PrinterAddNewForm(cX, cY, FormName)
End Function
Public Function PrinterFindFormIndex(FormName As String, _
    Optional ByRef cX As Long, Optional ByRef cY As Long) As Long
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1 ' Working FI1 array
    Dim temp() As Byte       ' Temp FI1 array
    Dim FormIndex As Integer
    Dim BytesNeeded As Long
    Dim RetVal As Long
      
    FormIndex = 0
    ReDim aFI1(1)
    ' First call retrieves the BytesNeeded.
    RetVal = EnumForms(hPrinter, 1, aFI1(0), 0&, BytesNeeded, NumForms)
    ReDim temp(BytesNeeded)
    ReDim aFI1(BytesNeeded / Len(FI1))
    ' Second call actually enumerates the supported forms.
    RetVal = EnumForms(hPrinter, 1, temp(0), BytesNeeded, BytesNeeded, _
            NumForms)
    Call CopyMemory(aFI1(0), temp(0), BytesNeeded)
    For i = 0 To NumForms - 1
        With aFI1(i)
'            If .Size.cx = FormSize.cx And .Size.cy = FormSize.cy Then
            If FormName = PtrCtoVbString(.pName) Then
                cX = .Size.cX
                cY = .Size.cY
                ' Found the desired form
                FormIndex = i + 1
                Exit For
            End If
        End With
    Next i
    PrinterFindFormIndex = FormIndex  ' Returns non-zero when form is found.
End Function
Function PrinterGetFormSize(FormName As String, cX As Long, cY As Long) As Long
    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim RetVal As Long
'    Dim PrinterHandle As Long ' Handle to printer
    Dim cbBuf As Long, BytesNeeded As Long
            ReDim aFI1(Len(FI1))
                RetVal = GetForm(hPrinter, FormName, 1, aFI1(0), UBound(aFI1), BytesNeeded)
            If BytesNeeded <> 0 Then
                ReDim aFI1(BytesNeeded)
            End If
                RetVal = GetForm(hPrinter, FormName, 1, aFI1(0), UBound(aFI1), BytesNeeded)
        If RetVal Then
            Dim Fi2 As FORM_INFO_1
            
                    Call CopyMemory(Fi2, aFI1(0), Len(Fi2))
                With Fi2
                    With .Size
                         cX = .cX
                         cY = .cY
                    End With
                End With
        End If
    PrinterGetFormSize = RetVal
End Function
Public Function PrinterSelectForm(FormName As String, Optional ByVal WinMM As Double = 214000, _
        Optional ByVal HinMM As Double = 216000) As Integer
    Dim nSize As Long        ' Size of DEVMODE
'    Dim pDevMode As DEVMODE
'    Dim PrinterHandle As Long ' Handle to printer
    Dim hPrtDC As Long       ' Handle to Printer DC
    Dim aDevMode() As Byte   ' Working DEVMODE
    Dim FormSize As SIZEL

'    If PrinterName = "" Then
'        PrinterName = Printer.DeviceName    ' Current printer
'    Else
'        MakeDefaultPrinter PrinterName
'    End If

    hPrtDC = hdcPrinter      ' hDC for current Printer
    PrinterSelectForm = FORM_NOT_SELECTED    ' Set for failure unless reset in code.

'    ' Get a handle to the printer.
'    If OpenPrinter(sPrinter, PrinterHandle, 0&) Then
        ' Retrieve the size of the DEVMODE.
        nSize = DocumentProperties(hwndParent, hPrinter, sPrinter, 0&, _
                0&, 0&)
        ' Reserve memory for the actual size of the DEVMODE.
        ReDim aDevMode(1 To nSize)

        ' Fill the DEVMODE from the printer.
        nSize = DocumentProperties(hwndParent, hPrinter, sPrinter, _
                aDevMode(1), 0&, DM_OUT_BUFFER)
        ' Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))

        ' If FormName is "MyCustomForm", we must make sure it exists
        ' before using it. Otherwise, it came from our EnumForms list,
        ' and we do not need to check first. Note that we could have
        ' passed in a Flag instead of checking for a literal name.
'        If FormName = "MyCustomForm" Then
            ' Use form "MyCustomForm", adding it if necessary.
            ' Set the desired size of the form needed.
            With FormSize    ' Given in thousandths of millimeters
                .cX = WinMM '214000 ' width
                .cY = HinMM '216000 ' height
            End With
            If PrinterFindFormIndex(FormName) = 0 Then
                ' Form not found - Either of the next 2 lines will work.
                'FormName = AddNewForm(PrinterHandle, FormSize, "MyCustomForm")
                PrinterAddNewForm FormSize.cX, FormSize.cY, FormName   '"MyCustomForm"
                If PrinterFindFormIndex(FormName) = 0 Then
'                    ClosePrinter (PrinterHandle)
                    PrinterSelectForm = FORM_NOT_SELECTED    ' Selection Failed!
                    Exit Function
                Else
                    PrinterSelectForm = FORM_ADDED    ' Form Added, Selection succeeded!
                End If
            End If
'        End If

        ' Change the appropriate member in the DevMode.
        ' In this case, you want to change the form name.
        pDevMode.dmFormName = FormName & Chr(0)    ' Must be NULL terminated!
        ' Set the dmFields bit flag to indicate what you are changing.
        pDevMode.dmFields = DM_FORMNAME

        ' Copy your changes back, then update DEVMODE.
        Call CopyMemory(aDevMode(1), pDevMode, Len(pDevMode))
        nSize = DocumentProperties(hwndParent, hPrinter, sPrinter, _
                aDevMode(1), aDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)

        nSize = ResetDC(hPrtDC, pDevMode)        ' Reset the DEVMODE for the DC.

        ' Close the handle when you are finished with it.
'        ClosePrinter (PrinterHandle)
        ' Selection Succeeded! But was Form Added?
        If PrinterSelectForm <> FORM_ADDED Then PrinterSelectForm = FORM_SELECTED
'    Else
'        SelectForm = FORM_NOT_SELECTED    ' Selection Failed!
'    End If
End Function
Public Function BitmapToPicture() As IPicture
'Purpose: Creates a <b>StdPicture</b> object from HBITMAP handle.
   '--- Returns a VB picture object containing the specified bitmap.
  Dim oNewPic         As Picture
  Dim lpPictDesc      As PictDesc
  Dim aGuid(0 To 3)   As Long
   '--- fill struct
   With lpPictDesc
      .Size = Len(lpPictDesc)
      .Type = vbPicTypeBitmap
      .hBmpOrIcon = hBmp
      .hPal = 0 ' hPalette
   End With
   '--- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
   aGuid(0) = &H7BF80980
   aGuid(1) = &H101ABF32
   aGuid(2) = &HAA00BB8B
   aGuid(3) = &HAB0C3000
   '--- Create picture from bitmap handle
   OleCreatePictureIndirect lpPictDesc, aGuid(0), True, oNewPic
   '--- success
   Set BitmapToPicture = oNewPic
End Function
Public Function ExtractIcon(ByVal clrMask As OLE_COLOR) As StdPicture
'Purpose: Creates a <b>StdPicture</b> object that represents an icon containing current image with pixels of specified mask color being transparent.
  Dim ii              As ICONINFO
  Dim hMaskDC         As Long
  Dim hImgDC          As Long
  Dim hImgBmp         As Long
  Dim hMaskBmp        As Long
  Dim hPrevMaskBmp    As Long
  Dim hPrevImgBmp     As Long
  Dim hIcon           As Long
   On Error Resume Next
   '--- state check
   '--- alloc resources
   hMaskDC = CreateCompatibleDC(0)
   hImgDC = CreateCompatibleDC(hdcPrinter)
   hMaskBmp = CreateCompatibleBitmap(hMaskDC, Width, Height)
   hImgBmp = CreateCompatibleBitmap(hdcPrinter, Width, Height)
   '--- select bitmaps
   hPrevMaskBmp = SelectObject(hMaskDC, hMaskBmp)
   hPrevImgBmp = SelectObject(hImgDC, hImgBmp)
   '--- copy image
   APIBitBlt hImgDC, 0, 0, Width, Height, hdcPrinter, 0, 0, vbSrcCopy
   '--- create mask
   SetBkColor hImgDC, clrMask
   SetTextColor hImgDC, vbWhite
   APIBitBlt hMaskDC, 0, 0, Width, Height, hImgDC, 0, 0, vbSrcCopy
   '--- do mask image
   SetBkColor hImgDC, vbBlack
   APIBitBlt hImgDC, 0, 0, Width, Height, hMaskDC, 0, 0, vbSrcAnd
   '--- deselect bitmaps
   Call SelectObject(hMaskDC, hPrevMaskBmp)
   Call SelectObject(hImgDC, hPrevImgBmp)
   '--- create icon

   With ii
      .fIcon = 1
      .hbmColor = hImgBmp
      .hBmMask = hMaskBmp
   End With

   hIcon = CreateIconIndirect(ii)
   '--- cleanup
   Call ApiDeleteObject(hMaskBmp)
   Call ApiDeleteObject(hImgBmp)
   Call DeleteDC(hMaskDC)
   Call DeleteDC(hImgDC)
   '--- return StdPicture
   Set ExtractIcon = IconToPicture(hIcon)
End Function
Public Function IconToPicture(ByVal hIcon As Long) As IPicture

'Purpose: Creates a <b>StdPicture</b> object from HICON handle.

   '--- Returns a VB picture object containing the specified icon.
  Dim oNewPic         As Picture
  Dim lpPictDesc      As PictDesc
  Dim aGuid(0 To 3)   As Long

   On Error Resume Next
   '--- check argument

   If hIcon = 0 Then
      Exit Function
   End If

   '--- fill struct

   With lpPictDesc
      .Size = Len(lpPictDesc)
      .Type = vbPicTypeIcon
      .hBmpOrIcon = hIcon
   End With

   '--- fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
   aGuid(0) = &H7BF80980
   aGuid(1) = &H101ABF32
   aGuid(2) = &HAA00BB8B
   aGuid(3) = &HAB0C3000
   '--- do convert
   OleCreatePictureIndirect lpPictDesc, aGuid(0), True, oNewPic
   '--- success
   Set IconToPicture = oNewPic

End Function
Sub ViewDosPrint(ByVal sFile As String)
''    pic.Font = "Lucida Console"
'Dim c As Long
'Dim sText As String, r As RECT, sChar As String, nChar As Long
''    R.Left = 10
''    R.Top = 10
'''    R.Right = 200
'''    R.Bottom = 200
''    R.Right = R.Left + 400
''    R.Bottom = R.Top + 300
''Rectangle pic.hdc, R.Left - 1, R.Top - 1, R.Right + 1, R.Bottom + 1
''    For c = 9480 To 9835
''        sText = sText & ChrW(c)
''            If c > 9480 And ((c - 9480) Mod 27 = 0) Then
''                sText = sText & vbCrLf
''            End If
''    Next
''    DrawTextExW pic.hdc, StrPtr(sText), Len(sText), R, DT_WORDBREAK, ByVal 0&
'
'Dim Obj As PictureBox, cX As Long, cY As Long, TW As Long, th As Long, MaxW As Long
'Dim DoubleWidth As Boolean
'            frmPrintPreview.cmdPrint.Visible = True
'            frmPrintPreview.cmdClose.Visible = True
'        frmPrintPreview.Show
'        frmPrintPreview.Caption = sFile
'
'        Set Obj = frmPrintPreview.PicP(0)
'        Obj.ScaleMode = 3
'        Obj.AutoRedraw = True
'        Obj.Cls
'        Obj.Font = "Lucida Console"
''        obj.Font = "Terminal"
'            TW = Obj.TextWidth("A")
'            th = Obj.TextHeight("A")
'               Obj.Width = Screen.Width 'obj.ScaleX((TW + 1) * 80, obj.ScaleMode, 1)
'               Obj.Height = Obj.ScaleY(th, Obj.ScaleMode, 1)
'Open sFile For Input As 1
'    Do Until EOF(1) = True
'                If MaxW < cX Then
'                    MaxW = cX
'                End If
'            cX = 0
'        Line Input #1, sText
'            For c = 1 To Len(sText)
'                If Obj.Height > Screen.Width \ 2 Then
'                            Obj.Width = Obj.ScaleX(MaxW + TW, Obj.ScaleMode, 1)
'                    Load frmPrintPreview.PicP(frmPrintPreview.PicP.UBound + 1)
'                    Set frmPrintPreview.PicP(frmPrintPreview.PicP.UBound).Picture = frmPrintPreview.PicP(0).Image
'                        frmPrintPreview.PicP(frmPrintPreview.PicP.UBound).Visible = False
'                        Obj.Cls
'                               Obj.Width = Screen.Width 'obj.ScaleX((TW + 1) * 80, obj.ScaleMode, 1)
'                               Obj.Height = Obj.ScaleY(th, Obj.ScaleMode, 1)
'                        cY = 0
'                End If
'
'                        r.Left = cX: r.Top = cY
'                        r.Right = r.Left + TW: r.Bottom = r.Top + th
'                    sChar = Mid$(sText, c, 1)
'                    nChar = Asc(sChar)
''                        If c > 3 And nChar > 0 Then
''                            If Mid$(sText, c - 3, 1) = Chr(27) Then
''                                nChar = 0
''                            End If
''                        End If
''                        If c > 1 And nChar > 0 Then
''                            If Mid$(sText, c - 1, 1) = Chr(27) Then
''                                nChar = 0
''                            End If
''                        End If
''                        If c > 2 And nChar > 0 Then
''                            If Mid$(sText, c - 2, 1) = Chr(27) Then
''                                nChar = 0
''                            End If
''                        End If
'                    If Mid$(sText, c, 1) = Chr(27) Then
'                        Select Case Mid$(sText, c + 1, 1)
'                            Case "j", "x", "k", "A", "P", "p", "t", "J"
'                                nChar = 0
'                                c = c + 2
'                            Case "@", "g"
'                                nChar = 0
'                                c = c + 1
'                            Case "W"
'                                nChar = 0
'                                If Mid$(sText, c + 2, 1) = Chr(1) Then
'                                    Obj.FontBold = True
'                                    DoubleWidth = True
'                                Else
'                                    Obj.FontBold = False
'                                    DoubleWidth = False
'                                End If
'                                c = c + 2
'                            Case "E"
'                                    Obj.FontBold = True
'                                c = c + 1
'                            Case "F"
'                                    Obj.FontBold = False
'                                c = c + 1
'                        End Select
'                    End If
'                        If DoubleWidth = True And nChar <> 0 Then
'                            sChar = sChar & " "
'                        End If
'                Select Case nChar
'                    Case 32 To 127
'                        DrawTextExW Obj.hDc, StrPtr(sChar), Len(sChar), r, 0, ByVal 0&
'                            cX = cX + (TW * Len(sChar))
'                    Case 176 To 223
'                            sChar = ChrW(nChar + 9355)
'                        DrawTextExW Obj.hDc, StrPtr(sChar), Len(sChar), r, 0, ByVal 0&
'                            cX = cX + (TW * Len(sChar))
'                End Select
'            Next
'                cY = cY + th
'                    Obj.Height = Obj.Height + Obj.ScaleY(th, Obj.ScaleMode, 1)
'    Loop
'Close 1
'               Obj.Width = Obj.ScaleX(MaxW + TW, Obj.ScaleMode, 1)
'
'        Load frmPrintPreview.PicP(frmPrintPreview.PicP.UBound + 1)
'        Set frmPrintPreview.PicP(frmPrintPreview.PicP.UBound).Picture = frmPrintPreview.PicP(0).Image
'            frmPrintPreview.PicP(frmPrintPreview.PicP.UBound).Visible = False
'        frmPrintPreview.SetPage
'        frmPrintPreview.Show
'        frmPrintPreview.Form_Resize
End Sub
Private Function Callback1_AbortProc(ByVal hPr As Long, ByVal code As Long) As Long
    ' We must allow events to take place, otherwise the
    ' user button press on the abortform form will never
    ' be detected!
    DoEvents
    If code = SP_OUTOFDISK Or AbortPrinting Then
        Callback1_AbortProc = 0
        Exit Function
    End If
    Callback1_AbortProc = -1
End Function
