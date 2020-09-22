VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Picture"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   1770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1635
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   555
         Left            =   120
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   32
         ScaleMode       =   0  'User
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   240
         Width           =   555
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Center"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tile"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   420
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Stretch"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   1
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Drag the arrow over a window to set its picture."
         Height          =   585
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   1455
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DC% = 1
Private Const WINDOWDC% = 2

Private Const CENTER% = 1
Private Const TILE% = 2
Private Const STRETCH% = 3

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PICTUREPROPERTIES
    pType As Long
    pWidth As Long
    pHeight As Long
    pWidthBytes As Long
    pPlanes As Integer
    pBitsPixel As Integer
    pBits As Long
End Type

Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Private Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hdc&)
Private Declare Function DeleteDC& Lib "gdi32" (ByVal hdc&)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC& Lib "user32" (ByVal hWnd&)
Private Declare Function GetObject& Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any)
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SelectObject& Lib "gdi32" (ByVal hdc&, ByVal hObject&)
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Dim thWnd&, pt As POINTAPI, sPos%

Private Sub Form_Load()

    sPos = 1

End Sub

Private Sub Option1_Click(Index As Integer)

    If Option1(Index).Value = True Then sPos = Index + 1

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Button
        Case 1
            MousePointer = vbUpArrow
    End Select

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Button
        Case 1
            Call GetCursorPos(pt)
            thWnd = WindowFromPoint(pt.X, pt.Y)
            Call SetPicture(thWnd, App.Path + "\picture.jpg", sPos, 2)
            MousePointer = vbDefault
    End Select

End Sub

Private Sub SetPicture(hWnd&, lpsz$, Optional PicPos% = CENTER, Optional TargetDC% = DC)

    Dim hDestDC&, hSrcDC&, hPic&, PicProp As PICTUREPROPERTIES, wR As RECT, Cx&, Cy&

    If Dir(lpsz) = vbNullString Then Exit Sub
    Select Case TargetDC
        Case DC
            'Sets the picture to inside of window.
            hDestDC = GetDC(hWnd)
        Case WINDOWDC
            'Sets the picture to all window.
            'Also to caption.
            hDestDC = GetWindowDC(hWnd)
    End Select
    Call GetObject(LoadPicture(lpsz).Handle, Len(PicProp), PicProp)
    hSrcDC = CreateCompatibleDC(hDestDC)
    Call SelectObject(hSrcDC, LoadPicture(lpsz).Handle)
    Call GetWindowRect(hWnd, wR)
    Select Case PicPos
        Case CENTER
            'Sets the picture to center.
            Call BitBlt(hDestDC, ((wR.Right - wR.Left) - PicProp.pWidth) / 2, ((wR.Bottom - wR.Top) - PicProp.pHeight) / 2, PicProp.pWidth, PicProp.pHeight, hSrcDC, 0, 0, vbSrcCopy)
        Case TILE
            'Tiles the picture to window.
            For Cx = 0 To wR.Right - wR.Left Step PicProp.pWidth
                For Cy = 0 To wR.Bottom - wR.Top Step PicProp.pHeight
                    Call BitBlt(hDestDC, Cx, Cy, PicProp.pWidth, PicProp.pHeight, hSrcDC, 0, 0, vbSrcCopy)
                Next Cy
            Next Cx
        Case STRETCH
            'Stretchs the picture to window.
            Call StretchBlt(hDestDC, 0, 0, wR.Right - wR.Left, wR.Bottom - wR.Top, hSrcDC, 0, 0, PicProp.pWidth, PicProp.pHeight, vbSrcCopy)
    End Select
    Call DeleteDC(hDestDC)
    Call DeleteDC(hSrcDC)

End Sub
