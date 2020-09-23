VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Bump mapping"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000002&
      Height          =   405
      Left            =   2295
      ScaleHeight     =   345
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   2925
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   600
      ScaleHeight     =   2265
      ScaleWidth      =   3525
      TabIndex        =   0
      Top             =   330
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   "Color :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1605
      TabIndex        =   1
      Top             =   3000
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Coded by : Safo
' Email : safo@zoznam.sk
' Program run faster when it is compile.

Const X = 240 ' picture width
Const Y = 150 ' picture height
Dim environmentmap(0 To 255, 0 To 255) As Single
Dim Pixel(-200 To 500, -100 To 500) As Integer ' pixels of picture
Dim Mouse_X As Integer
Dim Mouse_y As Integer
Dim Buf(-200 To 500, -200 To 500, 1 To 2) As Double

Private Function CreateEnvironmentMap()
    Dim i As Integer
    Dim j As Integer
    Dim nX As Double
    Dim nY As Double
    Dim nZ As Double
    
    For i = 0 To 255
        For j = 0 To 255
            nX = (i - 128) / 255
            nY = (j - 128) / 255
            nZ = 1 - Sqr(nX * nX + nY * nY)
            environmentmap(i, j) = Round(nZ * 256)
            If environmentmap(i, j) > 255 Then environmentmap(i, j) = 255
            If environmentmap(i, j) < 0 Then environmentmap(i, j) = 0
        Next j
    Next i
    
End Function

Private Function Init()

For i = 0 To X
        For j = 0 To Y
            nX = Pixel(i - 1, j) - Pixel(i + 1, j)
            nY = Pixel(i, j - 1) - Pixel(i, j + 1)
            Buf(i, j, 1) = nX - i + 128
            Buf(i, j, 2) = nY - j + 128
        Next j
    Next i

End Function

Public Function Draw_Bumpmap()
    Dim bm As Bitmap
    Dim r, g, b As Integer
    
    r = Color Mod 256
    g = (Color \ 256) Mod 256
    b = Color \ 256 \ 256
    GetObject Picture1.Image, Len(bm), bm
    Dim ImageData() As Byte
    ReDim ImageData(0 To (bm.bmBitsPixel \ 8) - 1, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)

    For i = 0 To X - 1
        For j = 0 To Y - 1
            konx = Buf(i, j, 1) + Mouse_X
            kony = Buf(i, j, 2) + Mouse_y
            If (konx < 0) Or (konx > 255) Then konx = 0
            If (kony < 0) Or (kony > 255) Then kony = 0
            ImageData(0, i, j) = environmentmap(konx, kony) * (b / 256)  ' blue
            ImageData(1, i, j) = environmentmap(konx, kony) * (g / 256) ' zelena
            ImageData(2, i, j) = environmentmap(konx, kony) * (r / 256) ' cervena
        Next j
    Next i
    
    SetBitmapBits Picture1.Image, bm.bmWidthBytes * bm.bmHeight, ImageData(0, 0, 0)
    Picture1.Refresh

End Function

Private Sub Form_Load()
    Picture1.Picture = LoadPicture("bump.bmp")
    CreateEnvironmentMap
    
    For i = 0 To X
        For j = 0 To Y
            Pixel(i, j) = GetPixel(Picture1.hdc, i, j) Mod 256
        Next j
    Next i
    
    Init
    Color = RGB(64, 128, 255)
    Picture2.BackColor = Color
    Mouse_X = 100
    Mouse_y = 90
    Draw_Bumpmap
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_X = X / 15
Mouse_y = Y / 15
Draw_Bumpmap
End Sub

Private Sub Picture2_Click()
    Colordialog.Show
End Sub
