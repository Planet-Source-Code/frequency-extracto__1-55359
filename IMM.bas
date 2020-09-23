Attribute VB_Name = "Module1"
'Icon Modification Module
'Compiled by Bad Frequency
'2003

Option Explicit
Public Dirty As Boolean
Type tIconInfo
    iWidth As Long
    iHeight As Long
    iBitCnt As Long
    iFileName As String
    iDC As Long
    iBitmap As Long
End Type
Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Type BITMAPINFOHEADER
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
Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Type BITMAPINFO1Bit
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 1) As RGBQUAD
End Type
Type BITMAPINFO4Bit
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 15) As RGBQUAD
End Type
Type BITMAPINFO8Bit
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type
Type BITMAPINFO24Bit
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Type ICONDIRENTRY
   bWidth As Byte               ' Width of the image
   bHeight As Byte              ' Height of the image (times 2)
   bColorCount As Byte          ' Number of colors in image (0 if >=8bpp)
   bReserved As Byte            ' Reserved
   wPlanes As Integer           ' Color Planes
   wBitCount As Integer         ' Bits per pixel
   dwBytesInRes As Long         ' how many bytes in this resource?
   dwImageOffset As Long        ' where in the file is this image
End Type
Type ICONDIR
   idReserved As Integer   ' Reserved
   idType As Integer       ' resource type (1 for icons)
   idCount As Integer      ' how many images?
   idEntries As ICONDIRENTRY 'array follows.
End Type
Type ICONDIRENTRY2
   bWidth As Byte               ' Width of the image
   bHeight As Byte              ' Height of the image (times 2)
   bColorCount As Byte          ' Number of colors in image (0 if >=8bpp)
   bReserved As Byte            ' Reserved
   wPlanes As Integer           ' Color Planes
   wBitCount As Integer         ' Bits per pixel
   dwBytesInRes As Long         ' how many bytes in this resource?
   dwImageOffset As Long        ' where in the file is this image
End Type
Type ICONDIR2
   idReserved As Integer   ' Reserved
   idType As Integer       ' resource type (1 for icons)
   idCount As Integer      ' how many images?
End Type
'A Cursor .cur file consists of CURSORDIR,CURSORDIRENTRY, and Image bytes

Type CURSORDIR
idReserved As Integer  '0
idType As Integer '2 for Cursor
idCount As Integer 'number of icons in file
End Type

Type CURSORDIRENTRY
bWidth As Byte
bHeight As Byte
bColorCount As Byte
bReserved As Byte
wXHotspot As Integer
wYHotspot As Integer
dwBytesInRes As Long
dwImageOffset As Long
End Type

Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As Any, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0
Public Const DIB_PAL_COLORS = 1
Public Const TransCol = 12961221

Global BitCnt As Long, IconInfo As tIconInfo, Ubnd
Global bi24BitInfo As BITMAPINFO24Bit
Global MaskBits(0 To 127) As Byte, bBits(0 To 3071) As Byte
Global CancelIt As Boolean



Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type

Public BIH As BITMAPINFOHEADER
Public ID As ICONDIR
Public IDE As ICONDIRENTRY

Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, ByVal SecAtts&, phkResult&, lpdwDisp&)
Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey&, ByVal lpValueName$)
Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, lpReserved&, lpType&, ByVal lpData$, lpcbData&)
Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, ByVal lpData$, ByVal cbData&)


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_ALL_ACCESS = (&H1F0000 Or &H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H20) And (Not &H100000)

Public Const Ttl = "Extracto"
Public curX As Single, curY As Single
Public iDone As Boolean
Public AreaSel As Boolean
Public IsCB As Boolean
Public Const MAX_PATH = 260
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const ILD_TRANSPARENT = &H1
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Public Declare Function SHGetFileInfo Lib _
   "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl As Long, ByVal i As Long, _
    ByVal hDCDest As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal flags As Long) As Long

Public shinfo As SHFILEINFO

Public Sub CheckSettings()
Dim Ret1
    #If Win32 Then
        If Not ScreenDimensionsOk Then End
        If Not ColorModeOk Then End
        Form1.Show
    #Else
        End
    #End If
End Sub
Private Function ScreenDimensionsOk() As Boolean
    Dim Msg As String
    Dim ScreenHeight As Integer
    Dim ScreenWidth As Integer

    ScreenDimensionsOk = True

    ScreenWidth = GetSystemMetrics(0)
    ScreenHeight = GetSystemMetrics(1)

    If ScreenHeight < 480 Or ScreenWidth < 640 Then
       Msg = Ttl & " requires a screen resolution of at least 640 x 480 pixels."
       Msg = Msg & vbCrLf & vbCrLf
       Msg = Msg & "Please adjust your screen resolution and try again."
       MsgBox Msg, vbCritical, Ttl & " - Error"
       ScreenDimensionsOk = False
    End If

End Function
Private Function ColorModeOk() As Boolean

    Dim Msg As String
    Dim hScreenDc As Long

    hScreenDc = GetDC(0)
    
    If GetDeviceCaps(hScreenDc, 12) > 8 Then
       ColorModeOk = True
    Else
       Msg = Ttl & " requires at least high color mode."
       Msg = Msg & vbCrLf & vbCrLf
       Msg = Msg & "Please adjust your color mode and try again."
       MsgBox Msg, vbCritical, Ttl & " - Error"
    End If

End Function
Public Sub SetUpIconDblClick()

    Dim RegData$, hKey&, Rv&

    Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, ".ico", 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, 0)
    If Rv = 0 Then
       RegSetValueEx hKey, vbNullString, 0, 1, "icofile", 7
       RegCloseKey hKey
    End If

    Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "icofile", 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, 0)
    If Rv = 0 Then
       RegSetValueEx hKey, vbNullString, 0, 1, "Windows Icon", 12
       RegCloseKey hKey
    End If

    Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "icofile\DefaultIcon", 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, 0)
    If Rv = 0 Then
       RegSetValueEx hKey, vbNullString, 0, 1, "%1", 2
       RegCloseKey hKey
    End If

    Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "icofile\Shell\Open\Command", 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, 0)
    If Rv = 0 Then
       If Right(App.Path, 1) = "\" Then
          RegData = App.Path & "vbIconMaker.exe"
       Else
          RegData = App.Path & "\vbIconMaker.exe"
       End If
       RegData = RegData & " /open %1"
       RegSetValueEx hKey, vbNullString, 0, 1, RegData, Len(RegData)
       RegCloseKey hKey
    End If

End Sub

Public Sub PrepIconHeader()

    ID.idReserved = 0
    ID.idType = 1
    ID.idCount = 1

    IDE.bWidth = 32
    IDE.bHeight = 32
    IDE.bColorCount = 0
    IDE.bReserved = 0
    IDE.wPlanes = 1
    IDE.wBitCount = 24
    IDE.dwBytesInRes = 3240
    IDE.dwImageOffset = 22

    BIH.biSize = 40
    BIH.biWidth = 32
    BIH.biHeight = 64
    BIH.biPlanes = 1
    BIH.biBitCount = 24
    BIH.biCompression = 0
    BIH.biSizeImage = 3200
    BIH.biXPelsPerMeter = 0
    BIH.biYPelsPerMeter = 0
    BIH.biClrUsed = 0
    BIH.biClrImportant = 0

End Sub

Public Sub SaveIcon(sFileName As String, nDC As Long, nBitmap As Long, BpP As Long)
    Dim CopyDC As Long, CopyBitmap As Long
    CopyDC = CreateCompatibleDC(nDC)
    bi24BitInfo.bmiHeader.biWidth = 32
    bi24BitInfo.bmiHeader.biHeight = 32
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = 32
        .biHeight = 32
    End With
    CopyBitmap = CreateDIBSection(nDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject CopyDC, CopyBitmap
    ChangePixels IconInfo.iDC, 0, 0, 32, 32, TransCol, vbBlack, CopyDC
    Select Case BpP
        Case 1
            SaveIcon1Bit sFileName, nDC, nBitmap, CopyDC, CopyBitmap
        Case 4
            SaveIcon4Bit sFileName, nDC, nBitmap, CopyDC, CopyBitmap
        Case 8
            SaveIcon8Bit sFileName, nDC, nBitmap, CopyDC, CopyBitmap
        Case 24
            SaveIcon24Bit sFileName, nDC, nBitmap, CopyDC, CopyBitmap
    End Select
    DeleteDC CopyDC
    DeleteObject CopyBitmap
End Sub
Function ChangePixels(hSrcDC As Long, X As Long, Y As Long, lWidth As Long, lHeight As Long, OldColor As Long, NewColor As Long, hDestDC As Long) As Boolean
    Dim r As RECT, mBrush As Long, CopyDC As Long, CopyBitmap As Long
    SetRect r, 0, 0, lWidth, lHeight
    mBrush = CreateSolidBrush(NewColor)
    CopyDC = CreateCompatibleDC(hSrcDC)
    bi24BitInfo.bmiHeader.biWidth = lWidth
    bi24BitInfo.bmiHeader.biHeight = lHeight
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
    End With
    CopyBitmap = CreateDIBSection(hSrcDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    If SelectObject(CopyDC, CopyBitmap) = 0 Then
        MsgBox "In ChangePixels - SelectObject(CopyDC, CopyBitmap) = 0"
        Exit Function
    End If
    If FillRect(CopyDC, r, mBrush) = 0 Then Exit Function
    If TransBlt(hSrcDC, X, Y, lWidth, lHeight, OldColor, CopyDC, hDestDC) = False Then Exit Function
    DeleteDC CopyDC
    DeleteObject CopyBitmap
    DeleteObject mBrush
    ChangePixels = True
End Function

Function TransBlt(hSrcDC As Long, X As Long, Y As Long, lWidth As Long, lHeight As Long, MaskColor As Long, hBackDC As Long, hDestDC As Long) As Boolean
    Dim MonoDC As Long, MonoBitmap As Long, CopyDC As Long, CopyBitmap As Long
    Dim AndDC As Long, AndBitmap As Long, r As RECT
    MonoDC = CreateCompatibleDC(hSrcDC)
    MonoBitmap = CreateBitmap(lWidth, lHeight, 1, 1, ByVal 0&)
    If SelectObject(MonoDC, MonoBitmap) = 0 Then Exit Function
    If CreateMask(hSrcDC, X, Y, lWidth, lHeight, MonoDC, MaskColor) = 0 Then Exit Function
    CopyDC = CreateCompatibleDC(hSrcDC)
    bi24BitInfo.bmiHeader.biWidth = lWidth
    bi24BitInfo.bmiHeader.biHeight = lHeight
    CopyBitmap = CreateDIBSection(hSrcDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    If SelectObject(CopyDC, CopyBitmap) = 0 Then Exit Function
    AndDC = CreateCompatibleDC(hSrcDC)
    bi24BitInfo.bmiHeader.biWidth = lWidth
    bi24BitInfo.bmiHeader.biHeight = lHeight
    AndBitmap = CreateDIBSection(hSrcDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    If SelectObject(AndDC, AndBitmap) = 0 Then Exit Function
    BitBlt AndDC, 0, 0, lWidth, lHeight, hSrcDC, X, Y, vbSrcCopy
    BitBlt CopyDC, 0, 0, lWidth, lHeight, hBackDC, 0, 0, vbSrcCopy
    BitBlt CopyDC, 0, 0, lWidth, lHeight, MonoDC, 0, 0, vbSrcAnd
    SetRect r, 0, 0, lWidth, lHeight
    InvertRect MonoDC, r
    BitBlt AndDC, 0, 0, lWidth, lHeight, MonoDC, 0, 0, vbSrcAnd
    BitBlt CopyDC, 0, 0, lWidth, lHeight, AndDC, 0, 0, vbSrcPaint
    If BitBlt(hDestDC, X, Y, lWidth, lHeight, CopyDC, 0, 0, vbSrcCopy) = 0 Then Exit Function
    DeleteDC MonoDC
    DeleteDC CopyDC
    DeleteDC AndDC
    DeleteObject MonoBitmap
    DeleteObject CopyBitmap
    DeleteObject AndBitmap
    TransBlt = True
End Function

Private Sub SetSaveData(BpP As Long, nDC As Long, MaskInfo As BITMAPINFO1Bit, fID As ICONDIR, nMaskDC As Long, nMaskBitmap As Long)
    Select Case BpP
        Case 1
            Ubnd = 1
            fID.idEntries.bColorCount = 0
        Case 4
            Ubnd = 15
            fID.idEntries.bColorCount = 16
        Case 8
            Ubnd = 255
           fID.idEntries.bColorCount = 0
       Case 24
            Ubnd = 0
            fID.idEntries.bColorCount = 0
    End Select
    fID.idCount = 1
    fID.idType = 1
    fID.idEntries.bHeight = 32
    fID.idEntries.bWidth = 32
    fID.idEntries.dwImageOffset = Len(fID)
    fID.idEntries.wBitCount = 0
    fID.idEntries.wPlanes = 0
    fID.idEntries.dwBytesInRes = 744
    MaskInfo.bmiHeader.biSize = Len(MaskInfo.bmiHeader)
    MaskInfo.bmiHeader.biBitCount = 1
    MaskInfo.bmiHeader.biClrUsed = 2
    MaskInfo.bmiHeader.biHeight = 32
    MaskInfo.bmiHeader.biPlanes = 1
    MaskInfo.bmiHeader.biWidth = 32
    nMaskDC = CreateCompatibleDC(GetDC(0))
    nMaskBitmap = CreateBitmap(32, 32, 1, 1, ByVal 0&)
    SelectObject nMaskDC, nMaskBitmap
    SetBkColor nDC, TransCol
    SetBkColor nMaskDC, vbBlack
    BitBlt nMaskDC, 0, 0, 32, 32, nDC, 0, 0, vbSrcCopy
End Sub
Public Sub GetIconBits(bBits() As Byte, BpP As Long, nDC As Long, nBitmap As Long, CopyArr() As RGBQUAD)
    Dim BI As BITMAPINFO8Bit
    BI.bmiHeader.biBitCount = BpP
    BI.bmiHeader.biCompression = BI_RGB
    BI.bmiHeader.biPlanes = 1
    BI.bmiHeader.biHeight = 32
    BI.bmiHeader.biWidth = 32
    BI.bmiHeader.biSize = Len(BI.bmiHeader)
    GetDIBits nDC, nBitmap, 0, 32, bBits(0), BI, DIB_RGB_COLORS
    If BpP = 1 Then
        CopyMemory CopyArr(0), BI.bmiColors(0), Len(CopyArr(0)) * 2
    ElseIf BpP = 4 Then
        CopyMemory CopyArr(0), BI.bmiColors(0), Len(CopyArr(0)) * 16
    ElseIf BpP = 8 Then
        CopyMemory CopyArr(0), BI.bmiColors(0), Len(CopyArr(0)) * 256
    End If
End Sub

Function CreateMask(hSrcDC As Long, X As Long, Y As Long, nWidth As Long, nHeight As Long, hDestDC As Long, MaskColor As Long) As Boolean
    Dim MonoDC As Long, MonoBitmap As Long, OldBkColor As Long
    MonoDC = CreateCompatibleDC(hSrcDC)
    MonoBitmap = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    If SelectObject(MonoDC, MonoBitmap) = 0 Then Exit Function
    OldBkColor = SetBkColor(hSrcDC, MaskColor)
    BitBlt MonoDC, 0, 0, nWidth, nHeight, hSrcDC, X, Y, vbSrcCopy
    If BitBlt(hDestDC, 0, 0, nWidth, nHeight, MonoDC, 0, 0, vbSrcCopy) = 0 Then Exit Function
    SetBkColor hSrcDC, OldBkColor
    DeleteObject MonoBitmap
    DeleteDC MonoDC
    CreateMask = True
End Function

Private Sub SaveIcon1Bit(sFileName As String, nDC As Long, nBitmap As Long, CopyDC As Long, CopyBitmap As Long)
    Dim fID As ICONDIR, MaskInfo As BITMAPINFO1Bit, IconInfo As BITMAPINFO1Bit
    Dim nMaskDC As Long, nMaskBitmap As Long, bBits(0 To 127) As Byte
    If Dir(sFileName) <> "" Then Kill sFileName
    SetSaveData 1, nDC, MaskInfo, fID, nMaskDC, nMaskBitmap
    fID.idEntries.wPlanes = 5
    fID.idEntries.wBitCount = 7
    fID.idEntries.dwBytesInRes = 304
    fID.idEntries.dwImageOffset = 22
    IconInfo.bmiHeader.biSize = Len(IconInfo.bmiHeader)
    IconInfo.bmiHeader.biBitCount = 1
    IconInfo.bmiHeader.biSizeImage = 128
    IconInfo.bmiHeader.biCompression = BI_RGB
    IconInfo.bmiHeader.biHeight = 64
    IconInfo.bmiHeader.biPlanes = 1
    IconInfo.bmiHeader.biWidth = 32
    Open sFileName For Binary As #1
        Put #1, , fID
        Put #1, , IconInfo.bmiHeader
        GetIconBits bBits(), 1, CopyDC, CopyBitmap, IconInfo.bmiColors()
        Put #1, , IconInfo.bmiColors
        Put #1, , bBits()
        GetDIBits nMaskDC, nMaskBitmap, 0, 32, MaskBits(0), MaskInfo, DIB_RGB_COLORS
        Put #1, , MaskBits()
    Close
    DeleteDC nMaskDC
    DeleteObject nMaskBitmap
End Sub
Private Sub SaveIcon4Bit(sFileName As String, nDC As Long, nBitmap As Long, CopyDC As Long, CopyBitmap As Long)
    Dim fID As ICONDIR, MaskInfo As BITMAPINFO1Bit, IconInfo As BITMAPINFO4Bit
    Dim nMaskDC As Long, nMaskBitmap As Long, bBits(0 To 511) As Byte
    If Dir(sFileName) <> "" Then Kill sFileName
    SetSaveData 4, nDC, MaskInfo, fID, nMaskDC, nMaskBitmap
    IconInfo.bmiHeader.biSize = Len(IconInfo.bmiHeader)
    IconInfo.bmiHeader.biBitCount = 4
    IconInfo.bmiHeader.biSizeImage = 2
    IconInfo.bmiHeader.biCompression = BI_RGB
    IconInfo.bmiHeader.biHeight = 64
    IconInfo.bmiHeader.biPlanes = 1
    IconInfo.bmiHeader.biWidth = 32
    Open sFileName For Binary As #1
        Put #1, , fID
        Put #1, , IconInfo.bmiHeader
        GetIconBits bBits(), 4, CopyDC, CopyBitmap, IconInfo.bmiColors()
        Put #1, , IconInfo.bmiColors
        Put #1, , bBits()
        GetDIBits nMaskDC, nMaskBitmap, 0, 32, MaskBits(0), MaskInfo, DIB_RGB_COLORS
        Put #1, , MaskBits()
    Close
    DeleteDC nMaskDC
    DeleteObject nMaskBitmap
End Sub
Private Sub SaveIcon8Bit(sFileName As String, nDC As Long, nBitmap As Long, CopyDC As Long, CopyBitmap As Long)
    Dim fID As ICONDIR, MaskInfo As BITMAPINFO1Bit, IconInfo As BITMAPINFO8Bit
    Dim nMaskDC As Long, nMaskBitmap As Long, bBits(0 To 1023) As Byte
    Dim IconPal(0 To 255) As RGBQUAD
    If Dir(sFileName) <> "" Then Kill sFileName
    SetSaveData 8, nDC, MaskInfo, fID, nMaskDC, nMaskBitmap
    fID.idEntries.dwBytesInRes = 2216
    IconInfo.bmiHeader.biSize = Len(IconInfo.bmiHeader)
    IconInfo.bmiHeader.biBitCount = 8
    IconInfo.bmiHeader.biSizeImage = 1152
    IconInfo.bmiHeader.biClrUsed = 256
    IconInfo.bmiHeader.biCompression = BI_RGB
    IconInfo.bmiHeader.biHeight = 64
    IconInfo.bmiHeader.biPlanes = 1
    IconInfo.bmiHeader.biWidth = 32
    Open sFileName For Binary As #1
        Put #1, , fID
        Put #1, , IconInfo.bmiHeader
            GetIconBits bBits(), 8, CopyDC, CopyBitmap, IconInfo.bmiColors()
        Put #1, , IconInfo.bmiColors
        Put #1, , bBits()
        GetDIBits nMaskDC, nMaskBitmap, 0, 32, MaskBits(0), MaskInfo, DIB_RGB_COLORS
        Put #1, , MaskBits()
    Close
    DeleteDC nMaskDC
    DeleteObject nMaskBitmap
End Sub
Private Sub SaveIcon24Bit(sFileName As String, nDC As Long, nBitmap As Long, CopyDC As Long, CopyBitmap As Long)
    Dim fID As ICONDIR, MaskInfo As BITMAPINFO1Bit, IconInfo As BITMAPINFO4Bit
    Dim nMaskDC As Long, nMaskBitmap As Long
    Dim IconPal(0 To 255) As RGBQUAD
    If Dir(App.Path & "\written.ico") <> "" Then Kill App.Path & "\written.ico"
    SetSaveData 24, nDC, MaskInfo, fID, nMaskDC, nMaskBitmap
    fID.idEntries.dwBytesInRes = 3244
    IconInfo.bmiHeader.biSize = Len(IconInfo.bmiHeader)
    IconInfo.bmiHeader.biBitCount = 24
    IconInfo.bmiHeader.biSizeImage = 1153
    IconInfo.bmiHeader.biClrUsed = 1
    IconInfo.bmiHeader.biCompression = BI_RGB
    IconInfo.bmiHeader.biHeight = 64
    IconInfo.bmiHeader.biPlanes = 1
    IconInfo.bmiHeader.biWidth = 32
    Open sFileName For Binary As #1
        Put #1, , fID
        Put #1, , IconInfo.bmiHeader
        GetIconBits bBits(), 24, CopyDC, CopyBitmap, IconInfo.bmiColors()
        Put #1, , "    "
        Put #1, , bBits()
        GetDIBits nMaskDC, nMaskBitmap, 0, 32, MaskBits(0), MaskInfo, DIB_RGB_COLORS
        Put #1, , MaskBits()
    Close
    DeleteDC nMaskDC
    DeleteObject nMaskBitmap
End Sub


