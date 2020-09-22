VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GfxFx"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Go - Option 5 - Under Water"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Go - Option 4 - Walk 2"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Go - Option 3 - Walk 1"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go - Option 2 - Frantic Blur 1"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5040
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go - Option 1 - Heat Blur 2"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go - Option 0 - Heat Blur 1"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   0
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5460
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   7260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020
Private Type Clip
x As Single
y As Single
offsetx As Single
offsety As Single
End Type
Dim Clips() As Clip, Sine1() As Single, Sine2() As Single, i As Integer, ii As Integer
Const TileSize = 16
Private Function RndRange(ByVal Min As Single, ByVal Max As Single) As Single
RndRange = (Rnd * (Max - Min + 1)) + Min
End Function

Private Sub Command1_Click()
Do: DoEvents
DoClip 0
Loop
End Sub

Sub DoClip(Op As Byte) 'Delete ones you do not use...
Select Case Op
Case 0
For i = 0 To Picture1.ScaleWidth / TileSize
For ii = 0 To Picture1.ScaleHeight / TileSize
Clips(i, ii).x = i * TileSize
Clips(i, ii).y = ii * TileSize
Clips(i, ii).offsetx = RndRange(-1, 1)
Clips(i, ii).offsety = RndRange(-1, 1)
BitBlt Picture1.hDC, Clips(i, ii).x + Clips(i, ii).offsetx, Clips(i, ii).y + Clips(i, ii).offsety, TileSize, TileSize, Picture2.hDC, i * TileSize, ii * TileSize, SRCCOPY
Next
Next
Case 1
For i = 0 To Picture1.ScaleWidth / TileSize
For ii = 0 To Picture1.ScaleHeight / TileSize
Clips(i, ii).x = i * TileSize
Clips(i, ii).y = ii * TileSize
Clips(i, ii).offsetx = RndRange(-2, 2)
Clips(i, ii).offsety = RndRange(-2, 2)
BitBlt Picture1.hDC, Clips(i, ii).x + Clips(i, ii).offsetx, Clips(i, ii).y + Clips(i, ii).offsety, TileSize, TileSize, Picture2.hDC, i * TileSize, ii * TileSize, SRCCOPY
Next
Next
Case 2
For i = 0 To Picture1.ScaleWidth / TileSize
For ii = 0 To Picture1.ScaleHeight / TileSize
Clips(i, ii).x = i * TileSize
Clips(i, ii).y = ii * TileSize
Clips(i, ii).offsetx = RndRange(-4, 4)
Clips(i, ii).offsety = RndRange(-4, 4)
BitBlt Picture1.hDC, Clips(i, ii).x + Clips(i, ii).offsetx, Clips(i, ii).y + Clips(i, ii).offsety, TileSize, TileSize, Picture2.hDC, i * TileSize, ii * TileSize, SRCCOPY
Next
Next
Case 3
For i = 0 To Picture1.ScaleWidth / TileSize
For ii = 0 To Picture1.ScaleHeight / TileSize
Clips(i, ii).x = i * TileSize
Clips(i, ii).y = ii * TileSize
Clips(i, ii).offsetx = Clips(i, ii).offsetx + RndRange(-2, 1)
Clips(i, ii).offsety = Clips(i, ii).offsety + RndRange(-2, 1)
BitBlt Picture1.hDC, Clips(i, ii).x + Clips(i, ii).offsetx, Clips(i, ii).y + Clips(i, ii).offsety, TileSize, TileSize, Picture2.hDC, i * TileSize, ii * TileSize, SRCCOPY
Next
Next
Case 4
For i = 0 To Picture1.ScaleWidth / TileSize
For ii = 0 To Picture1.ScaleHeight / TileSize
Clips(i, ii).x = i * TileSize
Clips(i, ii).y = ii * TileSize
Clips(i, ii).offsetx = Clips(i, ii).offsetx + RndRange(-4, 3)
Clips(i, ii).offsety = Clips(i, ii).offsety + RndRange(-4, 3)
BitBlt Picture1.hDC, Clips(i, ii).x + Clips(i, ii).offsetx, Clips(i, ii).y + Clips(i, ii).offsety, TileSize, TileSize, Picture2.hDC, i * TileSize, ii * TileSize, SRCCOPY
Next
Next
Case 5
For i = 0 To Picture1.ScaleWidth / TileSize
For ii = 0 To Picture1.ScaleHeight / TileSize
Clips(i, ii).x = i * TileSize
Clips(i, ii).y = ii * TileSize
Clips(i, ii).offsetx = Clips(i, ii).offsetx + 0.5 * Sin(Timer * i)
Clips(i, ii).offsety = Clips(i, ii).offsety + 0.5 * Sin(Timer * ii)
BitBlt Picture1.hDC, Clips(i, ii).x + Clips(i, ii).offsetx, Clips(i, ii).y + Clips(i, ii).offsety, TileSize, TileSize, Picture2.hDC, i * TileSize, ii * TileSize, SRCCOPY
Next
Next
End Select
End Sub

Private Sub Command2_Click()
Do: DoEvents
DoClip 1
Loop
End Sub

Private Sub Command3_Click()
Do: DoEvents
DoClip 2
Loop
End Sub

Private Sub Command4_Click()
Do: DoEvents
DoClip 3
Loop
End Sub

Private Sub Command5_Click()
Do: DoEvents
DoClip 4
Loop
End Sub

Private Sub Command6_Click()
Do: DoEvents
DoClip 5
Loop
End Sub

Private Sub Form_Load()
ReDim Clips(Picture1.ScaleWidth / TileSize, Picture1.ScaleHeight / TileSize)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
