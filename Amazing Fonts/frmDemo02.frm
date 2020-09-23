VERSION 5.00
Begin VB.Form frmDemo02 
   Caption         =   "Demo02"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCircle 
      Interval        =   100
      Left            =   4380
      Top             =   1800
   End
   Begin VB.PictureBox picCircle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2160
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   4
      Top             =   0
      Width           =   3615
      Begin prjMain.LabelBox lbxFont 
         Height          =   1335
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2355
         Alignment       =   2
         BackColor       =   12648447
         Caption         =   "ABC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SetWallpaper    =   2
         BottomFontTop   =   -35
         BottomFontDepth =   5
         TransparentArea =   3
         UpperFontTop    =   -25
         BottomFontColorStart=   12632319
         UpperFontColorStart=   0
         FontBold        =   -1  'True
         FontSize        =   72
         Enabled         =   0   'False
      End
   End
   Begin VB.Timer tmrTimer 
      Interval        =   10
      Left            =   2880
      Top             =   1800
   End
   Begin prjMain.LabelBox lbxMovingPicture 
      Height          =   2715
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4789
      Caption         =   "See the picture"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BottomFontColorEnd=   4210752
      BottomFontLeft  =   -5
      BottomFontDepth =   3
      UpperFontDepth  =   5
      WordWrap        =   -1  'True
      BottomFontColorStart=   12632256
      UpperFontColorStart=   12640511
      FontSize        =   48
   End
   Begin prjMain.LabelBox lbxSound 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1085
      BorderLine      =   5
      Caption         =   "Play sound"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BottomFontColorEnd=   8421504
      BottomFontLeft  =   1
      BottomFontTop   =   2
      BottomFontDepth =   2
      UpperFontDepth  =   2
      BottomFontColorStart=   8421504
      UpperFontColorStart=   12632064
      UpperFontColorEnd=   12632064
      FontBold        =   -1  'True
      FontSize        =   18
   End
   Begin prjMain.LabelBox lbxHyperlink 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1085
      Alignment       =   2
      BorderLine      =   5
      Caption         =   "Hyperlink"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BottomFontColorEnd=   8421504
      BottomFontLeft  =   1
      BottomFontTop   =   2
      BottomFontDepth =   2
      UpperFontDepth  =   2
      BottomFontColorStart=   8421504
      UpperFontColorStart=   12632064
      UpperFontColorEnd=   12632064
      FontBold        =   -1  'True
      FontSize        =   18
   End
   Begin prjMain.LabelBox lbxSizePicture 
      Height          =   2715
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4789
      Caption         =   "See the picture"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SetWallpaper    =   1
      BottomFontColorEnd=   4210752
      BottomFontLeft  =   -5
      BottomFontDepth =   3
      UpperFontDepth  =   5
      WordWrap        =   -1  'True
      BottomFontColorStart=   12632256
      UpperFontColorStart=   12640511
      FontSize        =   48
   End
End
Attribute VB_Name = "frmDemo02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===========================================================================================================================
Dim IsPlay As Boolean
Dim ImagePath As String
'===========================================================================================================================

'===========================================================================================================================
Private Sub Form_Load()
    IsPlay = True
            
    ImagePath = App.Path & "\Pictures"
    
    lbxSound.Sound.Path = App.Path & "\win.wav"
    Set lbxMovingPicture.Wallpaper = LoadPicture(ImagePath & "\sample.bmp")
    Set lbxSizePicture.Wallpaper = LoadPicture(ImagePath & "\sample.bmp")
    Set lbxFont.Wallpaper = LoadPicture(ImagePath & "\sample.bmp")
    lbxSizePicture.PictureWidth = 1
End Sub

'===========================================================================================================================
Private Sub lbxHyperlink_Click()
    lbxHyperlink.HyperLink.NavigateTo "http://www.planetsourcecode.com"
End Sub

'===========================================================================================================================
Private Sub lbxSound_Click()
    If IsPlay Then
        lbxSound.Caption = "Stop sound"
        lbxSound.Sound.BeginPlaySound
        IsPlay = False
    Else
        lbxSound.Caption = "Play sound"
        lbxSound.Sound.EndPlaySound
        IsPlay = True
    End If
End Sub

'===========================================================================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    IsPlay = False
    lbxSound_Click
End Sub

'===========================================================================================================================
Private Sub tmrCircle_Timer()
    Dim c As Long
    Dim i As Integer, r As Integer
    Dim x As Integer, y As Integer
    
    For i = 1 To 5
        x = picCircle.ScaleWidth * Rnd() + 1
        y = picCircle.ScaleHeight * Rnd() + 1
        c = RGB(255 * Rnd() + 1, 255 * Rnd + 1, 255 * Rnd() + 1)
        For r = 0 To 10
            picCircle.ForeColor = c
            picCircle.Circle (x, y), r
        Next r
    Next i
End Sub

'===========================================================================================================================
Private Sub tmrTimer_Timer()
    Static i As Integer
    Static mbPicture As Integer
    
    lbxMovingPicture.PictureLeft = mbPicture
    lbxSizePicture.PictureWidth = lbxSizePicture.PictureWidth + 5
    mbPicture = mbPicture + 5
    
    If mbPicture > lbxMovingPicture.Width / Screen.TwipsPerPixelX Then
        mbPicture = 0
        lbxSizePicture.PictureWidth = 1
    End If
    
    lbxMovingPicture.BackColor = RGB(i, i, i)
    
    i = i + 5
    If i > 255 Then i = 0
End Sub
'===========================================================================================================================
