VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extracto"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Icon Extraction"
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.PictureBox picReal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         ToolTipText     =   "Icon Extracted from File"
         Top             =   240
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   120
         X2              =   2280
         Y1              =   840
         Y2              =   840
      End
   End
   Begin MSComctlLib.StatusBar sB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status: Idle"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      DragIcon        =   "Form1.frx":2CFA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      LinkTimeout     =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   570
      LinkTimeout     =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1095
      LinkTimeout     =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picTest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1680
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2400
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Dim okToMove As Boolean, prevX%, prevY%
Dim xSav, ySav, xStart, yStart
Dim pX%, pY%, pXOff%, pYOff%
Dim cvtY, cvtX
Dim filePath
Dim pixelDraw As Boolean, canDraw As Boolean
Dim ColorChg, bkClr, clr As Long, r As Integer, g As Integer, b As Integer
Dim j, p, x1, y1, colorSave, eraseIt As Boolean, pickColor As Boolean
Dim chkPix As Boolean, chgColor
Dim lineDraw As Boolean, lineOKDraw As Boolean
Dim rectDraw As Boolean, fillBoxDraw As Boolean, rectOKDraw As Boolean
Dim circleDraw As Boolean, fillCircleDraw As Boolean, circleOKDraw As Boolean
Dim textDraw As Boolean
Dim selRegion As Boolean
Dim lineX1, lineY1
Dim pasteIt As Boolean
Dim XHi, XLo, YHi, YLo, xDelLo, xdelHi, yDelLo, ydelHi
Dim canSelect As Boolean
Dim moveIt As Boolean
Dim selectIt As Boolean
Dim xOff, yOff, setDiff As Boolean
Dim xMove, yMove

Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private CurrentFile$
Private CurrentName$

Private Sub cmdExit_Click()
Dim X
X = MsgBox("Are you sure you want to exit?", vbYesNo, App.Title)
If X = vbYes Then
Unload Me
Else
Exit Sub
End If
End Sub

Private Sub Command1_Click()
Dim Answ, comD, cdIndex, Pos
Dim hImgLarge As Long
Dim hImgSmall As Long
Dim fName As String
Dim FizFil As String
Dim r As Long

sB1.SimpleText = "Status: Browsing for File..."
On Local Error GoTo cmdLoadErrorHandler

FizFil$ = "All Files (*.*)|*.*"

cd1.FileName = ""



cd1.FilterIndex = cdIndex
cd1.InitDir = comD



cd1.CancelError = True
cd1.Filter = FizFil$
cd1.ShowOpen

Pos = InStrRev(cd1.FileName, "\")
picReal.Picture = LoadPicture()
picTest.Picture = LoadPicture()
   fName$ = cd1.FileName
   

hImgSmall& = SHGetFileInfo(fName$, 0&, _
shinfo, Len(shinfo), _
BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

hImgLarge& = SHGetFileInfo(fName$, 0&, _
shinfo, Len(shinfo), _
BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
   
picTest.Picture = LoadPicture()
picTest.AutoRedraw = True
sB1.SimpleText = "Status: Icon Extracted!"
picTest.BackColor = RGB(197, 197, 197)

Call ImageList_Draw(hImgLarge&, shinfo.iIcon, picReal.hdc, 0, 0, ILD_TRANSPARENT)
SloP
Form1.Refresh
MousePointer = 0
Exit Sub

cmdLoadErrorHandler:
sB1.SimpleText = "Status: Idle"
Exit Sub
End Sub

Public Sub SloP()
Dim F
Static Pont
If Pont = &HFFC0FF Then
Line (picReal.Left - 1, picReal.Top - 1)-(picReal.Left + picReal.Width, picReal.Top + picReal.Height), QBColor(15), B
Else
Line (picReal.Left - 1, picReal.Top - 1)-(picReal.Left + picReal.Width, picReal.Top + picReal.Height), QBColor(0), B
End If
End Sub

Private Sub WriteDataToFile(Fn$)

    Dim MaskString$
    Dim Msg$
    Dim F%, h%, w%
    Dim c1&, c2&, r&, g&, b&, k%, n%

    On Error GoTo WriteError

    F = FreeFile
    Open Fn For Binary Access Write As #F


         For k = Len(Fn) To 1 Step -1
             If Mid(Fn, k, 1) = "\" Then Exit For
         Next

         Put #F, 1, ID
         Put #F, 7, IDE
         Put #F, 23, BIH
         k = 63
         For h = 31 To 0 Step -1
             For w = 0 To 31
                 c1 = GetPixel(PicImage.hdc, w, h)
                 c2 = GetPixel(picMask.hdc, w, h)
                 If c2 = &HFFFFFF Then
                    Put #F, k, 0
                    Put #F, k + 1, 0
                    Put #F, k + 2, 0
                 Else
                    b = c1 \ 65536
                    g = (c1 - b * 65536) \ 256
                    r = c1 - b * 65536 - g * 256
                    Put #F, k, b
                    Put #F, k + 1, g
                    Put #F, k + 2, r
                 End If
                 k = k + 3
             Next
         Next
         k = 0
         n = 0
         For h = 31 To 0 Step -1
             For w = 0 To 31
                 If GetPixel(picMask.hdc, w, h) = &HFFFFFF Then
                    MaskString = MaskString & "1"
                 Else
                    MaskString = MaskString & "0"
                 End If
                 k = k + 1
                 If k = 8 Then
                    k = 0
                    Put #F, n + 3135, BinaryStringToByte(MaskString)
                    MaskString = ""
                    n = n + 1
                 End If
             Next
         Next
    Close #F

    CurrentFile = Fn
    On Error GoTo 0
    Exit Sub

WriteError:

    Screen.MousePointer = 0

    If Err.Number <> cdlCancel Then
       Msg = Err.Description & "."
       Msg = Msg & vbCrLf & vbCrLf
       If CurrentFile = "Untitled" Then
          Msg = Msg & "Unable to save Untitled."
       Else
          Msg = Msg & "Unable to save " & CurrentName
       End If
       MsgBox Msg, vbExclamation, Ttl & " - Error"
    End If
    'bFileSaved = False
    Err.Clear
    Exit Sub

End Sub

Private Function BinaryStringToByte(MS$) As Byte

    Dim k%, Rv As Byte

    For k = 1 To 8
        If Mid(MS, k, 1) = "1" Then Rv = Rv + 2 ^ (8 - k)
    Next

    BinaryStringToByte = Rv

End Function

Private Sub cmdSave_Click()
Dim Ret, bmpPicInfo As BITMAPINFO, Answ, comD, cdIndex, Pos
Dim sPos, ePos

sB1.SimpleText = "Status: Saving Icon..."
On Error GoTo ExitIt

cd1.FileName = "Icon"
cd1.flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
cd1.Filter = "Icons (*.ico)|*.ico|Bitmaps (*.bmp)|*.bmp|Cursors (*.cur)|*.cur"
cd1.ShowSave
Pos = InStrRev(cd1.FileName, "\")
comD = Mid(cd1.FileName, 1, Pos)
cdIndex = cd1.FilterIndex
MousePointer = 11
BitCnt = 8

With bmpPicInfo.bmiHeader
    .biBitCount = 16
    .biCompression = BI_RGB
    .biPlanes = 1
    .biSize = Len(bmpPicInfo.bmiHeader)
    .biWidth = 32
    .biHeight = 32
End With
    
IconInfo.iDC = CreateCompatibleDC(0)
IconInfo.iWidth = 32
IconInfo.iHeight = 32
bi24BitInfo.bmiHeader.biWidth = 32
bi24BitInfo.bmiHeader.biHeight = 32
IconInfo.iBitmap = CreateDIBSection(IconInfo.iDC, bmpPicInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
SelectObject IconInfo.iDC, IconInfo.iBitmap
Ret = BitBlt(IconInfo.iDC, 0, 0, 32, 32, picReal.hdc, 0, 0, vbSrcCopy)
DoEvents
SaveIcon cd1.FileName, IconInfo.iDC, IconInfo.iBitmap, BitCnt
IconInfo.iFileName = cd1.FileName
DeleteDC IconInfo.iDC
DeleteObject IconInfo.iBitmap
picReal.BackColor = RGB(197, 197, 197)
picTest.Picture = LoadPicture(cd1.FileName)
DoEvents
SloP
MousePointer = 0
Dirty = False
Form1.Refresh
picReal.Picture = LoadPicture(cd1.FileName)
sB1.SimpleText = "Status: Icon Saved!"
Exit Sub
ExitIt:
sB1.SimpleText = "Status: Idle"
Exit Sub
End Sub

Private Sub Form_Load()
picReal.BackColor = RGB(197, 197, 197)
End Sub

Private Sub form_unload(cancel As Integer)
MsgBox ("Thank you for using Extracto from Frequency Software!")
End
End Sub
