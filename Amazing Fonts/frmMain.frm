VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10830
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4800
      Top             =   3720
   End
   Begin prjMain.LabelBox lbxTitle 
      Height          =   1155
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2037
      Caption         =   "Amazing Fonts"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TransparentArea =   4
      TransparentColor=   16777215
      UpperFontDepth  =   5
      UpperFontTop    =   -10
      PictureTop      =   15
      BottomFontColorStart=   0
      UpperFontColorStart=   16777088
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontSize        =   48
   End
   Begin VB.Frame fraFrame 
      Height          =   7215
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   10815
      Begin VB.PictureBox picBGround 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   6975
         Left            =   60
         ScaleHeight     =   463
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   711
         TabIndex        =   4
         Top             =   180
         Width           =   10695
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1155
            Index           =   17
            Left            =   3540
            TabIndex        =   22
            Top             =   5640
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2037
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   1
            BottomFontColorEnd=   16761087
            BottomFontLeft  =   -3
            BottomFontTop   =   -13
            BottomFontDepth =   9
            TransparentArea =   4
            TransparentColor=   16777215
            UpperFontDepth  =   2
            UpperFontLeft   =   5
            UpperFontTop    =   -15
            BottomFontColorStart=   8388736
            UpperFontColorStart=   12640511
            UpperFontColorEnd=   16761087
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1275
            Index           =   0
            Left            =   -60
            TabIndex        =   5
            Top             =   -120
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2249
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   128
            BottomFontTop   =   -15
            BottomFontDepth =   10
            TransparentArea =   1
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   12632319
            UpperFontColorStart=   33023
            UpperFontColorEnd=   192
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1155
            Index           =   1
            Left            =   -60
            TabIndex        =   6
            Top             =   840
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2037
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontLeft  =   -2
            BottomFontTop   =   -12
            BottomFontDepth =   6
            TransparentArea =   1
            UpperFontDepth  =   3
            UpperFontTop    =   -10
            BottomFontColorStart=   0
            UpperFontColorStart=   14737632
            UpperFontColorEnd=   128
            FontBold        =   -1  'True
            FontItalic      =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1155
            Index           =   2
            Left            =   -60
            TabIndex        =   7
            Top             =   1740
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2037
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   8421631
            BottomFontLeft  =   -3
            BottomFontTop   =   -13
            BottomFontDepth =   8
            TransparentArea =   1
            UpperFontTop    =   -10
            BottomFontColorStart=   12632319
            UpperFontColorStart=   255
            UpperFontColorEnd=   255
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1155
            Index           =   3
            Left            =   -60
            TabIndex        =   8
            Top             =   2580
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2037
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   4194304
            BottomFontTop   =   -10
            BottomFontDepth =   10
            TransparentArea =   1
            UpperFontTop    =   -10
            BottomFontColorStart=   16761024
            UpperFontColorStart=   12582912
            UpperFontColorEnd=   16711680
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1335
            Index           =   4
            Left            =   -60
            TabIndex        =   9
            Top             =   3600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   8388736
            BottomFontTop   =   -15
            BottomFontDepth =   15
            TransparentArea =   1
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   12648384
            UpperFontColorStart=   49344
            UpperFontColorEnd=   32896
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1335
            Index           =   5
            Left            =   0
            TabIndex        =   10
            Top             =   4560
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   192
            BottomFontLeft  =   -4
            BottomFontTop   =   -12
            BottomFontDepth =   12
            TransparentArea =   1
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   12632319
            UpperFontColorStart=   16576
            UpperFontColorEnd=   12640511
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1335
            Index           =   6
            Left            =   3420
            TabIndex        =   11
            Top             =   -120
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            BackColor       =   16777152
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   192
            BottomFontLeft  =   -4
            BottomFontTop   =   -12
            TransparentArea =   1
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   12632319
            UpperFontColorStart=   16777215
            UpperFontColorEnd=   8421504
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1275
            Index           =   7
            Left            =   3480
            TabIndex        =   12
            Top             =   780
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2249
            BackColor       =   16761087
            BorderLine      =   8
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   1
            BottomFontColorEnd=   49152
            BottomFontLeft  =   -4
            BottomFontTop   =   -12
            BottomFontDepth =   20
            TransparentArea =   1
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   12648384
            UpperFontColorStart=   16744576
            UpperFontColorEnd=   4194304
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1335
            Index           =   8
            Left            =   3540
            TabIndex        =   13
            Top             =   1860
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            BorderLine      =   8
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   1
            BottomFontColorEnd=   128
            BottomFontLeft  =   -10
            BottomFontTop   =   -15
            BottomFontDepth =   20
            TransparentArea =   1
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   49344
            UpperFontColorStart=   12640511
            UpperFontColorEnd=   16576
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1335
            Index           =   9
            Left            =   7200
            TabIndex        =   14
            Top             =   60
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            BackColor       =   16777152
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   192
            BottomFontLeft  =   -4
            BottomFontTop   =   -12
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   12632319
            UpperFontColorStart=   16777215
            UpperFontColorEnd=   8421504
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1395
            Index           =   10
            Left            =   7140
            TabIndex        =   15
            Top             =   2940
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2461
            BackColor       =   12632319
            BorderLine      =   5
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   1
            BottomFontColorEnd=   49152
            BottomFontLeft  =   -4
            BottomFontTop   =   -12
            BottomFontDepth =   20
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   12648384
            UpperFontColorStart=   16744576
            UpperFontColorEnd=   4194304
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1395
            Index           =   11
            Left            =   7200
            TabIndex        =   16
            Top             =   1440
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2461
            BorderLine      =   8
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   49152
            BottomFontTop   =   -20
            BottomFontDepth =   20
            TransparentColor=   16777215
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   16761087
            UpperFontColorStart=   8421631
            UpperFontColorEnd=   4194304
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1335
            Index           =   12
            Left            =   -60
            TabIndex        =   17
            Top             =   5520
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   12648384
            BottomFontTop   =   -12
            BottomFontDepth =   9
            TransparentArea =   1
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            BottomFontColorStart=   -2147483615
            UpperFontColorStart=   12648384
            UpperFontColorEnd=   16512
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1335
            Index           =   13
            Left            =   3480
            TabIndex        =   18
            Top             =   2820
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   8421376
            BottomFontTop   =   -1
            BottomFontDepth =   1
            TransparentArea =   1
            UpperFontDepth  =   3
            UpperFontTop    =   -10
            BottomFontColorStart=   8421376
            UpperFontColorStart=   16776960
            UpperFontColorEnd=   12632064
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1335
            Index           =   14
            Left            =   3540
            TabIndex        =   19
            Top             =   3900
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontColorEnd=   12640511
            BottomFontLeft  =   -10
            BottomFontTop   =   -19
            BottomFontDepth =   10
            TransparentArea =   1
            UpperFontDepth  =   3
            UpperFontTop    =   -10
            BottomFontColorStart=   16512
            UpperFontColorStart=   8454016
            UpperFontColorEnd=   32768
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   2535
            Index           =   15
            Left            =   7140
            TabIndex        =   20
            Top             =   4380
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   4471
            BackColor       =   12648447
            BorderLine      =   8
            Caption         =   "Fonts Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BottomFontColorEnd=   4210752
            BottomFontLeft  =   -5
            BottomFontTop   =   -15
            BottomFontDepth =   20
            TransparentArea =   4
            TransparentColor=   12648447
            UpperFontDepth  =   5
            UpperFontTop    =   -10
            PictureLeft     =   20
            PictureTop      =   45
            WordWrap        =   -1  'True
            PictureHeight   =   100
            BottomFontColorStart=   16576
            UpperFontColorStart=   8421631
            FontBold        =   -1  'True
            FontSize        =   48
         End
         Begin prjMain.LabelBox lbxLabelBox 
            Height          =   1155
            Index           =   16
            Left            =   3540
            TabIndex        =   21
            Top             =   4740
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2037
            Caption         =   "Fonts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SetWallpaper    =   2
            BottomFontLeft  =   -3
            BottomFontTop   =   -13
            BottomFontDepth =   9
            TransparentArea =   1
            UpperFontDepth  =   2
            UpperFontTop    =   -10
            BottomFontColorStart=   0
            UpperFontColorStart=   12648447
            UpperFontColorEnd=   32896
            FontBold        =   -1  'True
            FontSize        =   48
         End
      End
   End
   Begin prjMain.LabelBox LabelBox1 
      Height          =   735
      Index           =   0
      Left            =   8520
      TabIndex        =   0
      Top             =   60
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1296
      Caption         =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 2"
         Size            =   48
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentArea =   1
      UpperFontDepth  =   5
      UpperFontTop    =   -10
      BottomFontColorStart=   0
      UpperFontColorStart=   16777152
      FontBold        =   -1  'True
      FontSize        =   48
   End
   Begin prjMain.LabelBox LabelBox1 
      Height          =   735
      Index           =   1
      Left            =   1380
      TabIndex        =   2
      Top             =   60
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1296
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 2"
         Size            =   48
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentArea =   1
      UpperFontDepth  =   5
      UpperFontTop    =   -10
      BottomFontColorStart=   0
      UpperFontColorStart=   16777152
      FontBold        =   -1  'True
      FontSize        =   48
   End
   Begin VB.Menu mnuDemo 
      Caption         =   "&Demo"
      Begin VB.Menu mnuDemoDemo01 
         Caption         =   "Demo 0&1"
      End
      Begin VB.Menu mnuDemoDemo02 
         Caption         =   "Demo 0&2"
      End
      Begin VB.Menu mnuDemoExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===========================================================================================================================
Private Sub Form_Load()
    Dim Filepath As String
    
    Filepath = App.Path & "\Pictures"
                    
    Set lbxTitle.Wallpaper = LoadPicture(Filepath & "\Ball.bmp")
    
    lbxTitle.Sound.Path = App.Path & "\win.wav"
    
    Set lbxLabelBox(0).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(1).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(6).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(7).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(8).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(9).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(10).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(11).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(15).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
    Set lbxLabelBox(17).Wallpaper = LoadPicture(Filepath & "\Sample.bmp")
End Sub

'===========================================================================================================================
Private Sub mnuDemoDemo01_Click()
    Timer1.Enabled = False
    lbxTitle.Sound.EndPlaySound
    frmDemo01.Show vbModal
    Timer1.Enabled = True
End Sub

'===========================================================================================================================
Private Sub mnuDemoDemo02_Click()
    Timer1.Enabled = False
    lbxTitle.Sound.EndPlaySound
    frmDemo02.Show vbModal
    Timer1.Enabled = True
End Sub

'===========================================================================================================================
Private Sub mnuDemoExit_Click()
    Unload Me
End Sub

'===========================================================================================================================
Private Sub Timer1_Timer()
    Dim n As Integer
    Static MoveObject As Integer
    On Error Resume Next
    
    lbxTitle.PictureLeft = MoveObject
    n = lbxTitle.Width / Screen.TwipsPerPixelX - 36
    lbxTitle.Sound.BeginPlaySound
    
    If MoveObject > n Then
        MoveObject = 0
        
    Else
        MoveObject = MoveObject + 5
    End If
End Sub

'===========================================================================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    lbxTitle.Sound.EndPlaySound
End Sub
'===========================================================================================================================

