VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Toolbar Bitmap Editor"
   ClientHeight    =   4635
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pBar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   323
      TabIndex        =   26
      Top             =   4335
      Width           =   4845
      Begin VB.PictureBox pFore 
         BackColor       =   &H00000000&
         Height          =   180
         Left            =   3570
         ScaleHeight     =   120
         ScaleWidth      =   495
         TabIndex        =   29
         ToolTipText     =   "ForeColor"
         Top             =   75
         Width           =   555
      End
      Begin VB.PictureBox pBack 
         BackColor       =   &H00FF00FF&
         Height          =   180
         Left            =   4170
         ScaleHeight     =   120
         ScaleWidth      =   495
         TabIndex        =   28
         ToolTipText     =   "BackColor"
         Top             =   75
         Width           =   555
      End
      Begin VB.Label lblXY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         Height          =   195
         Left            =   105
         TabIndex        =   27
         Top             =   45
         Width           =   270
      End
   End
   Begin VB.Frame FraMove 
      Caption         =   "Move Image"
      Height          =   1200
      Left            =   105
      TabIndex        =   21
      Top             =   2925
      Width           =   1485
      Begin VB.CommandButton ImgMove 
         Height          =   345
         Index           =   3
         Left            =   960
         Picture         =   "frmmain.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move Right"
         Top             =   540
         Width           =   345
      End
      Begin VB.CommandButton ImgMove 
         Height          =   345
         Index           =   2
         Left            =   225
         Picture         =   "frmmain.frx":0055
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Move Left"
         Top             =   555
         Width           =   345
      End
      Begin VB.CommandButton ImgMove 
         Height          =   345
         Index           =   1
         Left            =   585
         Picture         =   "frmmain.frx":00AC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Move Down"
         Top             =   660
         Width           =   345
      End
      Begin VB.CommandButton ImgMove 
         Height          =   345
         Index           =   0
         Left            =   585
         Picture         =   "frmmain.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Move Up"
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Button Preview"
      Height          =   750
      Left            =   2685
      TabIndex        =   19
      Top             =   2295
      Width           =   1920
      Begin VB.PictureBox pView 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   150
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   20
         Top             =   270
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      Height          =   1995
      Left            =   2685
      TabIndex        =   2
      Top             =   150
      Width           =   1920
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   300
         Index           =   15
         Left            =   1380
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   18
         Tag             =   "0,&H00000000&"
         Top             =   1425
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000080&
         Height          =   300
         Index           =   14
         Left            =   1005
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   17
         Tag             =   "0,&H00000080&"
         Top             =   1425
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00800080&
         Height          =   300
         Index           =   13
         Left            =   645
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   16
         Tag             =   "0,&H00800080&"
         Top             =   1425
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00800000&
         Height          =   300
         Index           =   12
         Left            =   285
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   15
         Tag             =   "0,&H00800000&"
         Top             =   1425
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         Height          =   300
         Index           =   11
         Left            =   1380
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   14
         Tag             =   "0,&H00808080&"
         Top             =   1065
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         Height          =   300
         Index           =   10
         Left            =   1380
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   13
         Tag             =   "0,&H00808000&"
         Top             =   690
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFF00&
         Height          =   300
         Index           =   9
         Left            =   1380
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   12
         Tag             =   "0,&H00FFFF00&"
         Top             =   330
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   300
         Index           =   8
         Left            =   1005
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   11
         Tag             =   "0,&H000000FF&"
         Top             =   1065
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         Height          =   300
         Index           =   7
         Left            =   645
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   10
         Tag             =   "0,&H00FF00FF&"
         Top             =   1065
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         Height          =   300
         Index           =   6
         Left            =   285
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   9
         Tag             =   "0,&H00FF0000&"
         Top             =   1065
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
         Height          =   300
         Index           =   5
         Left            =   1005
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   8
         Tag             =   "0,&H00008000&"
         Top             =   690
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008080&
         Height          =   300
         Index           =   4
         Left            =   645
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   7
         Tag             =   "0,&H00008080&"
         Top             =   690
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   3
         Left            =   285
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   6
         Tag             =   "0,&H00C0C0C0&"
         Top             =   690
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         Height          =   300
         Index           =   2
         Left            =   1005
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   5
         Tag             =   "0,&H0000FF00&"
         Top             =   330
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         Height          =   300
         Index           =   1
         Left            =   645
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   4
         Tag             =   "0,&H0000FFFF&"
         Top             =   330
         Width           =   300
      End
      Begin VB.PictureBox pColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   285
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   3
         Tag             =   "0,&H00FFFFFF&"
         Top             =   330
         Width           =   300
      End
   End
   Begin VB.PictureBox pSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2235
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2130
      Top             =   345
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox DrawArea 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   105
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   0
      Top             =   225
      Width           =   1905
      Begin VB.Shape Shape1 
         Height          =   150
         Left            =   945
         Top             =   930
         Width           =   150
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   37
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   37
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnua 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
      End
      Begin VB.Menu mnuC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuInvert 
         Caption         =   "&Invert"
      End
      Begin VB.Menu mnuB 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFlipH 
         Caption         =   "FlipH"
      End
      Begin VB.Menu MnuFlipV 
         Caption         =   "&FlipV"
      End
      Begin VB.Menu mnuRotate 
         Caption         =   "&Rotate"
      End
      Begin VB.Menu mnuD 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFill 
         Caption         =   "&Fill"
         Begin VB.Menu mnuFore 
            Caption         =   "ForeColor"
         End
         Begin VB.Menu mnuBack 
            Caption         =   "BacColor"
         End
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_cell As Integer = 10
Private Const ImgSize As Integer = 15

Private xpos As Integer
Private ypos As Integer
Private GridArea As Integer

Private ImageData(0 To ImgSize, 0 To ImgSize) As Long

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Function ShowColorDLG() As Long
On Error GoTo CanErr:
    
    With CD1
        .CancelError = True
        .ShowColor
        ShowColorDLG = .Color
    End With
    
    Exit Function
    
CanErr:
    ShowColorDLG = -1
End Function

Private Sub CleanUp()
    xpos = 0
    ypos = 0
    GridArea = 0
    Erase ImageData
    lblXY.Caption = ""
    Set pSrc.Picture = Nothing
    DrawArea.Cls
    pView.Cls
End Sub

Private Sub MoveImage(MoveOp As Integer)
Dim X As Integer
Dim Y As Integer

    'Move Image Up
    If (MoveOp = 0) Then
        For X = 0 To 15
            For Y = 0 To (ImgSize - 1)
                pSrc.PSet (X, Y), pSrc.Point(X, Y + 1)
            Next
            pSrc.PSet (X, ImgSize), vbMagenta
        Next
    End If
   
    'Move Image Down
    If (MoveOp = 1) Then
        For X = 0 To ImgSize
            For Y = ImgSize To 0 Step -1
                pSrc.PSet (X, Y), pSrc.Point(X, Y - 1)
            Next
            pSrc.PSet (X, 0), vbMagenta
        Next
    End If
   
   'Move Image Left
    If (MoveOp = 2) Then
        For Y = 0 To ImgSize
            For X = 0 To ImgSize - 1
                pSrc.PSet (X, Y), pSrc.Point(X + 1, Y)
            Next
            pSrc.PSet (ImgSize, Y), vbMagenta
        Next
    End If
    
    'Move Image Right
    If (MoveOp = 3) Then
        For Y = 0 To ImgSize
            For X = ImgSize To 1 Step -1
                pSrc.PSet (X, Y), pSrc.Point(X - 1, Y)
            Next
            pSrc.PSet (0, Y), vbMagenta
        Next
    End If
   
    Call StoreImageData
    Call DrawImage
    
    X = 0
    Y = 0
    
End Sub

Private Sub Flip(FlipOp As Integer)
Dim X As Integer
Dim Y As Integer
    
    Set pSrc.Picture = Nothing
    
    'Set the data
    For X = 0 To ImgSize
        For Y = 0 To ImgSize
            pSrc.PSet (X, Y), ImageData(X, Y)
        Next
    Next
    
    'FlipV
    If (FlipOp = 1) Then
        For X = ImgSize To 0 Step -1
            For Y = 0 To ImgSize
                ImageData(X, Y) = pSrc.Point(X, ImgSize - Y)
            Next
        Next
    End If

    'Flip H
    If (FlipOp = 2) Then
        For X = 0 To ImgSize
            For Y = ImgSize To 0 Step -1
                ImageData(X, Y) = pSrc.Point(ImgSize - X, Y)
            Next
        Next
    End If
    
    'Rotate
    If (FlipOp = 3) Then
        For X = 0 To ImgSize
            For Y = 0 To ImgSize
                ImageData(X, Y) = pSrc.Point(Y, X)
            Next
        Next
    End If
    
    Call DrawImage
    X = 0
    Y = 0
End Sub

Private Sub LongToRgb(LngColor As Long, r As Integer, g As Integer, b As Integer)
Dim tRgb(2) As Byte
    CopyMemory tRgb(0), LngColor, 4
    
    r = tRgb(0)
    g = tRgb(1)
    b = tRgb(2)
    
    Erase tRgb
End Sub


Private Sub NewImage(Optional Color As Long = vbMagenta)
Dim X As Integer, Y As Integer

    For X = 0 To ImgSize
        For Y = 0 To ImgSize
            ImageData(X, Y) = Color
        Next
    Next
    
    X = 0: Y = 0
    
End Sub

Function CheckBitmapSize(BmpFile As String) As Boolean
Dim fp As Long
Dim h As Long
Dim w As Long

    fp = FreeFile
    
    Open BmpFile For Binary As #fp
        Seek #fp, 19
        Get #fp, , h
        Get #fp, , w
    Close #fp
    
    CheckBitmapSize = (h = 16) And (w = 16)
    
    h = 0
    w = 0
    
End Function

Private Sub PutPixel(MouseButton As Integer)
Dim cUse As Long
    
    If (MouseButton = vbLeftButton) Then
        cUse = pFore.BackColor
    End If
    
    If (MouseButton = vbRightButton) Then
        cUse = pBack.BackColor
    End If
    
    DrawArea.Line (xpos * m_cell, ypos * m_cell)-(m_cell - 1 + xpos * m_cell, m_cell - 1 + ypos * m_cell), cUse, BF
    ImageData(xpos, ypos) = cUse
    'Show the Image
    Call DrawPreviewImage
End Sub

Private Sub StoreImageData()
Dim X As Integer
Dim Y As Integer

    Erase ImageData
    
    For X = 0 To pSrc.ScaleWidth - 1
        For Y = 0 To pSrc.Height - 1
            ImageData(X, Y) = pSrc.Point(X, Y)
        Next
    Next
End Sub

Private Sub DrawImage()
Dim X As Integer
Dim Y As Integer
Dim cCol As Long

    For X = 0 To GridArea Step m_cell
        For Y = 0 To GridArea Step m_cell
            cCol = ImageData(X \ m_cell, Y \ m_cell)
            DrawArea.Line (X, Y)-(X + m_cell, Y + m_cell), cCol, BF
        Next
    Next
    
    'Draw the button preview Image
    Call DrawPreviewImage
    
End Sub

Private Sub DrawPreviewImage()
Dim X As Integer
Dim Y As Integer
    
    'This draws the main picture
    For X = 0 To ImgSize
        For Y = 0 To ImgSize
            pSrc.PSet (X, Y), ImageData(X, Y)
        Next
    Next
    
    'Next we Draw the button Preview
    pView.Cls
    pView.Line (0, 0)-(pView.ScaleWidth, 0), vbWhite
    pView.Line (0, 0)-(0, pView.ScaleHeight), vbWhite
    pView.Line (pView.ScaleWidth - 1, 0)-(pView.ScaleWidth - 1, pView.ScaleHeight), &H404040
    pView.Line (pView.ScaleWidth - 2, 1)-(pView.ScaleWidth - 2, pView.ScaleHeight - 1), &H808080
    pView.Line (0, pView.ScaleHeight - 1)-(pView.ScaleWidth, pView.ScaleHeight - 1), &H404040
    pView.Line (1, pView.ScaleHeight - 2)-(pView.ScaleWidth - 2, pView.ScaleHeight - 2), &H808080

    'Transfer the Image to the button image
    TransparentBlt pView.hdc, 4, 3, 16, 16, pSrc.hdc, 0, 0, 16, 16, pSrc.Point(0, 0)
    pView.Refresh
    DrawArea.Refresh
End Sub


Private Sub DrawArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button <> 0) Then Call PutPixel(Button)
End Sub

Private Sub DrawArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    xpos = (X \ m_cell)
    ypos = (Y \ m_cell)
    
    If (xpos < 0) Then xpos = 0
    If (xpos > ImgSize) Then xpos = ImgSize
    If (ypos < 0) Then ypos = 0
    If (ypos > ImgSize) Then ypos = ImgSize
    
    lblXY.Caption = "xPos=" & xpos & " ,yPos" & ypos
    
    
    If Not Shape1.Visible Then Shape1.Visible = True

    Shape1.Left = (xpos * m_cell)
    Shape1.Top = (ypos * m_cell)

    If (Button <> 0) Then Call PutPixel(Button)
    
End Sub

Private Sub Form_Load()
    DrawArea.Width = (m_cell * 16) + 4
    DrawArea.Height = (m_cell * 16) + 4
    GridArea = (pSrc.ScaleHeight * m_cell) - 1
    
    FraMove.Top = (DrawArea.Height + Screen.TwipsPerPixelX) + 10

    
    Call mnuNew_Click
    DrawArea_MouseMove 0, 0, 0, 0
    pColor_MouseDown 15, 1, 1, 1, 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.Visible = False
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = frmmain.ScaleWidth
    Line1(1).X2 = Line1(0).X2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub ImgMove_Click(Index As Integer)
    Call MoveImage(Index)
End Sub

Private Sub mnuabout_Click()
    MsgBox "My Toolbar Bitmap Editor V1.0" _
    & vbCrLf & vbTab & "By DreamVB", vbInformation, "About"
    
End Sub

Private Sub mnuBack_Click()
    Call NewImage(pBack.BackColor)
    Call DrawImage
End Sub

Private Sub mnuexit_Click()
    Call CleanUp
    Unload frmmain
End Sub

Private Sub MnuFlipH_Click()
    Call Flip(2)
End Sub

Private Sub MnuFlipV_Click()
    Call Flip(1)
End Sub

Private Sub mnuFore_Click()
    Call NewImage(pFore.BackColor)
    Call DrawImage
End Sub

Private Sub mnuInvert_Click()
Dim X As Integer, Y As Integer
Dim r As Integer, g As Integer, b As Integer

    For X = 0 To ImgSize
        For Y = 0 To ImgSize
            'Get and convert long color to rgb
            Call LongToRgb(ImageData(X, Y), r, g, b)
            'Invert the color values
            r = (255 - r)
            g = (255 - g)
            b = (255 - b)
            'Store back the inverted color
            ImageData(X, Y) = RGB(r, g, b)
        Next
    Next

    Call DrawImage

End Sub

Private Sub mnuNew_Click()
    Call NewImage
    Call DrawImage
End Sub

Private Sub mnuOpen_Click()
On Error GoTo OpenErr:
    
    With CD1
        .InitDir = App.Path
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "Bitmap Files(*.bmp)|*.bmp|"
        .ShowOpen
        'Check for 16x16 Bitmaps
        If Not CheckBitmapSize(.FileName) Then
            MsgBox "Only 16x16 bitmaps are supported.", vbInformation, "Size Not Supported"
        Else
            pSrc.Picture = LoadPicture(.FileName)
            Call StoreImageData
            Call DrawImage
        End If
    End With
    
    Exit Sub
OpenErr:
    If Err Then Err.Clear
    
End Sub

Private Sub mnuRotate_Click()
    Call Flip(3)
End Sub

Private Sub mnuSave_Click()
On Error GoTo SaveErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = "Save"
        .Filter = "Bitmap Files(*.bmp)|*.bmp|"
        .ShowSave
        SavePicture pSrc.Image, .FileName
        
    End With
    
    Exit Sub
SaveErr:
    If Err Then Err.Clear
End Sub

Private Sub pBack_Click()
Dim cCol As Long
    cCol = ShowColorDLG
    
    If (cCol <> -1) Then
        pBack.BackColor = cCol
    End If
End Sub

Private Sub pBar_Resize()
    pBar.Line (0, 0)-(pBar.ScaleWidth - 1, pBar.ScaleHeight - 1), &H808080, B
    pBar.Refresh
End Sub

Private Sub pColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pObj As Object
Dim tmp As String
Dim Cnt As Integer
Dim TmpCol As Long

    Set pObj = pColor(Index)
    'Get the backcolor of the picturebox
    TmpCol = Val(Mid(pObj.Tag, 3))
    
    If (Button = vbLeftButton) Then
        pFore.BackColor = TmpCol
    End If
    
    If (Button = vbRightButton) Then
        pBack.BackColor = TmpCol
    End If

    'Draw the selcted outline in the picturebox
    pObj.BackColor = vbWhite
    pObj.Line (1, 1)-(pObj.ScaleWidth - 2, pObj.ScaleHeight - 2), 0, B
    pObj.Line (2, 2)-(pObj.ScaleWidth - 3, pObj.ScaleHeight - 3), TmpCol, BF
    'Store the selected value
    pObj.Tag = "1," & TmpCol
    
    For Cnt = 0 To pColor.Count - 1
        If (Mid(pColor(Cnt).Tag, 1, 1) = "1") And (Cnt <> Index) Then
            tmp = pColor(Cnt).Tag
            Mid(tmp, 1, 1) = "0"
            pColor(Cnt).Tag = tmp
            pColor(Cnt).BackColor = Val(Mid(tmp, 3))
            Exit For
        End If
    Next
    
    Set pObj = Nothing
    tmp = vbNullString
    Set pObj = Nothing
    Cnt = 0
    
End Sub


Private Sub pFore_Click()
Dim cCol As Long
    cCol = ShowColorDLG
    
    If (cCol <> -1) Then
        pFore.BackColor = cCol
    End If
    
End Sub
