VERSION 5.00
Begin VB.Form Colordialog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Dialog"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2940
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Cld1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H00000000&
      Height          =   2850
      Left            =   120
      ScaleHeight     =   2790
      ScaleWidth      =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   2700
   End
   Begin VB.PictureBox Cld3 
      BackColor       =   &H80000009&
      Height          =   825
      Left            =   120
      ScaleHeight     =   765
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Cld7 
      Height          =   240
      Left            =   2220
      MaxLength       =   3
      TabIndex        =   5
      Top             =   3120
      Width           =   570
   End
   Begin VB.TextBox Cld8 
      Height          =   240
      Left            =   2220
      MaxLength       =   3
      TabIndex        =   4
      Top             =   3420
      Width           =   570
   End
   Begin VB.TextBox Cld9 
      Height          =   240
      Left            =   2220
      MaxLength       =   3
      TabIndex        =   3
      Top             =   3720
      Width           =   570
   End
   Begin VB.CommandButton Cld10 
      Caption         =   "Ok"
      Height          =   360
      Left            =   300
      TabIndex        =   2
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Cld11 
      Caption         =   "Storno"
      Height          =   360
      Left            =   1725
      TabIndex        =   1
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Cld4 
      BackStyle       =   0  'Transparent
      Caption         =   "Red: "
      Height          =   240
      Left            =   1530
      TabIndex        =   0
      Top             =   3150
      Width           =   435
   End
   Begin VB.Label Cld5 
      BackStyle       =   0  'Transparent
      Caption         =   "Green: "
      Height          =   240
      Left            =   1515
      TabIndex        =   9
      Top             =   3435
      Width           =   480
   End
   Begin VB.Label Cld6 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue: "
      Height          =   240
      Left            =   1530
      TabIndex        =   8
      Top             =   3750
      Width           =   435
   End
End
Attribute VB_Name = "Colordialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cld1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Cld3.BackColor = GetPixel(Cld1.hdc, X / 15, Y / 15)
    Cld7.Text = GetPixel(Cld1.hdc, X / 15, Y / 15) Mod 256 ' red
    Cld8.Text = (GetPixel(Cld1.hdc, X / 15, Y / 15) \ 256) Mod 256 ' green
    Cld9.Text = GetPixel(Cld1.hdc, X / 15, Y / 15) \ 256 \ 256 ' blue
    
End Sub

Private Sub Cld1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And GetPixel(Cld1.hdc, X / 15, Y / 15) > 0 Then
    Cld3.BackColor = GetPixel(Cld1.hdc, X / 15, Y / 15)
    Cld7.Text = GetPixel(Cld1.hdc, X / 15, Y / 15) Mod 256 ' red
    Cld8.Text = (GetPixel(Cld1.hdc, X / 15, Y / 15) \ 256) Mod 256 ' green
    Cld9.Text = GetPixel(Cld1.hdc, X / 15, Y / 15) \ 256 \ 256 ' blue
    End If
    
End Sub

Private Sub Cld10_Click()
    Color = RGB(Cld7, Cld8, Cld9)
    Form1.Picture2.BackColor = Color
    Form1.Draw_Bumpmap
    Unload Me
End Sub

Private Sub Cld11_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Cld1.Picture = LoadPicture("Color.bmp")
    Cld3.BackColor = Color
    Cld7.Text = Color Mod 256 ' red
    Cld8.Text = (Color \ 256) Mod 256 ' green
    Cld9.Text = Color \ 256 \ 256 ' blue
    
End Sub

Private Sub Cld7_Change()
If Cld7 <> "" And Cld8 <> "" And Cld9 <> "" Then
Cld3.BackColor = RGB(Cld7, Cld8, Cld9)
End If
End Sub

Private Sub Cld8_Change()
If Cld7 <> "" And Cld8 <> "" And Cld9 <> "" Then
Cld3.BackColor = RGB(Cld7, Cld8, Cld9)
End If
End Sub

Private Sub Cld9_Change()
If Cld7 <> "" And Cld8 <> "" And Cld9 <> "" Then
Cld3.BackColor = RGB(Cld7, Cld8, Cld9)
End If
End Sub
