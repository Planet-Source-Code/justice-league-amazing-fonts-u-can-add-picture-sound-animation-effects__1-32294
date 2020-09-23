VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemo01 
   Caption         =   "Properties"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Transparent"
      Height          =   1335
      Left            =   5040
      TabIndex        =   53
      Top             =   3180
      Width           =   3075
      Begin VB.CommandButton cmdTransparaentColor 
         Height          =   255
         Left            =   2460
         TabIndex        =   59
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtTransparentColor 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   960
         Width           =   1515
      End
      Begin VB.ComboBox cmbTransparentArea 
         Height          =   315
         ItemData        =   "frmDemo01.frx":0000
         Left            =   840
         List            =   "frmDemo01.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "only if area is equal to Customize"
         Height          =   255
         Left            =   180
         TabIndex        =   57
         Top             =   720
         Width           =   2595
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Color : "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   56
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Area : "
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   54
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Upper Caption"
      Height          =   2415
      Left            =   0
      TabIndex        =   24
      Top             =   4560
      Width           =   4035
      Begin VB.CommandButton cmdUpperFontColor 
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   44
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtUpperFontColor 
         BackColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1020
         Width           =   795
      End
      Begin VB.CommandButton cmdUpperFontColor 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   42
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtUpperFontColor 
         BackColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   720
         Width           =   795
      End
      Begin VB.TextBox txtUpperFont 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   660
         TabIndex        =   32
         Text            =   "1"
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox txtUpperFont 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   660
         TabIndex        =   30
         Text            =   "0"
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txtUpperFont 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   660
         TabIndex        =   28
         Text            =   "0"
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Properties: UpperFontLeft, UpperFontTop, UpperFontDepth, UpperFonColorStart, UpperFontColorEnd"
         Height          =   615
         Index           =   1
         Left            =   60
         TabIndex        =   51
         Top             =   1740
         Width           =   3855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Color End  : "
         Height          =   195
         Index           =   5
         Left            =   1785
         TabIndex        =   40
         Top             =   1020
         Width           =   870
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Color Start  : "
         Height          =   195
         Index           =   4
         Left            =   1740
         TabIndex        =   39
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Depth : "
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   31
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Top : "
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   29
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Left : "
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   27
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label6 
         Caption         =   "to hide upper caption set  depth(UpperFontDept) to zero."
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bottom Caption"
      Height          =   2415
      Left            =   4080
      TabIndex        =   23
      Top             =   4560
      Width           =   4035
      Begin VB.TextBox txtBottomFontColor 
         BackColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   660
         Width           =   795
      End
      Begin VB.CommandButton cmdBottomrFontColor 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   47
         Top             =   660
         Width           =   255
      End
      Begin VB.TextBox txtBottomFontColor 
         BackColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton cmdBottomrFontColor 
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   45
         Top             =   1020
         Width           =   255
      End
      Begin VB.TextBox txtBottomFont 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   38
         Text            =   "0"
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox txtBottomFont 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   36
         Text            =   "0"
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txtBottomFont 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   34
         Text            =   "0"
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Properties: BottomFontLeft, BottomFontTop, BottomFontDepth, BottomFonColorStart, BottomFontColorEnd"
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   52
         Top             =   1740
         Width           =   3855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Color Start  : "
         Height          =   195
         Index           =   7
         Left            =   1740
         TabIndex        =   50
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Color End  : "
         Height          =   195
         Index           =   6
         Left            =   1785
         TabIndex        =   49
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Depth : "
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Top : "
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Left : "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "to hide bottom caption set  depth(BottomFontDepth) to zero."
         Height          =   435
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Picture"
      Height          =   3075
      Left            =   5040
      TabIndex        =   2
      Top             =   60
      Width           =   3015
      Begin VB.CommandButton cmdRemovePicture 
         Caption         =   "&Remove Picture"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbSetWallpaper 
         Height          =   315
         ItemData        =   "frmDemo01.frx":004B
         Left            =   1380
         List            =   "frmDemo01.frx":0058
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtPicturePos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1860
         TabIndex        =   12
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtPicturePos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1860
         TabIndex        =   10
         Text            =   "0"
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox txtPicturePos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   8
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtPicturePos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   4
         Text            =   "0"
         Top             =   1380
         Width           =   615
      End
      Begin VB.CommandButton cmdAddPicture 
         Caption         =   "&Add Picture"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Set Wallpaper : "
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   2100
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Properties: Wallpaper, PictureLeft, PictureTop, PictureWidth and PictureHeight, SetWallpaper"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   2475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPicturePos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Height : "
         Height          =   195
         Index           =   3
         Left            =   1260
         TabIndex        =   11
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label lblPicturePos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Width : "
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblPicturePos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Top : "
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1740
         Width           =   420
      End
      Begin VB.Label lblPictureDefault 
         Caption         =   "to use default width and height set it to zero."
         Height          =   555
         Left            =   240
         TabIndex        =   6
         Top             =   900
         Width           =   2115
      End
      Begin VB.Label lblPicturePos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Left : "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1380
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   4515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      Begin VB.ComboBox cmbBorderLine 
         Height          =   315
         ItemData        =   "frmDemo01.frx":0072
         Left            =   1020
         List            =   "frmDemo01.frx":0091
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   4140
         Width           =   1275
      End
      Begin VB.CommandButton cmdColor 
         Height          =   255
         Left            =   1980
         TabIndex        =   20
         Top             =   3840
         Width           =   315
      End
      Begin VB.TextBox txtBackColor 
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3840
         Width           =   915
      End
      Begin VB.CheckBox chkWordWrap 
         Caption         =   "Wordwrap"
         Height          =   195
         Left            =   3540
         TabIndex        =   14
         Top             =   3900
         Width           =   1215
      End
      Begin prjMain.LabelBox lbxLabelBox 
         Height          =   3555
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6271
         Caption         =   "Sample Sample"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentColor=   16777215
         BottomFontColorStart=   0
         UpperFontColorStart=   0
         FontSize        =   48
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BorderLine : "
         Height          =   195
         Left            =   75
         TabIndex        =   21
         Top             =   4200
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BackColor : "
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   3900
         Width           =   930
      End
   End
   Begin MSComDlg.CommonDialog dlgCommondialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuGetPicture 
      Caption         =   "GetPicture"
   End
   Begin VB.Menu mnuSavePicture 
      Caption         =   "SavePicture"
   End
End
Attribute VB_Name = "frmDemo01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===========================================================================================================================
Private Sub Form_Load()
    Set lbxLabelBox.Wallpaper = LoadPicture(App.Path & "\pictures\ball.bmp")
    cmbSetWallpaper.ListIndex = cmbSetWallpaper.TopIndex
    cmbBorderLine.ListIndex = cmbBorderLine.TopIndex
    cmbTransparentArea.ListIndex = cmbTransparentArea.TopIndex
End Sub

'===========================================================================================================================
Private Sub chkWordWrap_Click()
    lbxLabelBox.WordWrap = chkWordWrap.Value
End Sub

'===========================================================================================================================
Private Sub cmbSetWallpaper_Click()
    lbxLabelBox.SetWallpaper = cmbSetWallpaper.ListIndex
End Sub

'===========================================================================================================================
Private Sub cmdAddPicture_Click()
    On Error GoTo OpenErr
    
    With dlgCommondialog
        .Filter = "Bitmap (*.bmp) | *.bmp; | " & _
            "JPEG Filter (*.jpg) | *.jpg; | " & _
            "Graphics Interchange Format (*.gif) | *.gif; | " & _
            "All Files (*.*) | *.*"
        .FilterIndex = 1
        .Filename = vbNullString
        .Action = 1
        
        Set lbxLabelBox.Wallpaper = LoadPicture(.Filename)
    End With
    Exit Sub

OpenErr:
    If Err.Number <> 32755 Then MsgBox Err.Description
End Sub

'===========================================================================================================================
Private Sub cmdRemovePicture_Click()
    Set lbxLabelBox.Wallpaper = Nothing
End Sub

'===========================================================================================================================
Private Sub mnuGetPicture_Click()
    Set frmGetPicture.imgPicture.Picture = lbxLabelBox.GetPicture
    
    frmGetPicture.Show vbModal
End Sub

'===========================================================================================================================
Private Sub mnuSavePicture_Click()
    Dim Filename As String
    
    Filename = InputBox("Example: c:\Filename.bmp" & Chr(13) & "Filename : ", "Save")
    
    If Filename <> vbNullString Then lbxLabelBox.Save Filename
End Sub

'===========================================================================================================================
Private Sub txtPicturePos_Change(Index As Integer)
    On Error Resume Next
    
    lbxLabelBox.PictureLeft = CInt(txtPicturePos(0).Text)
    lbxLabelBox.PictureTop = CInt(txtPicturePos(1).Text)
    lbxLabelBox.PictureWidth = CInt(txtPicturePos(2).Text)
    lbxLabelBox.PictureHeight = CInt(txtPicturePos(3).Text)
End Sub

'===========================================================================================================================
Private Sub cmdColor_Click()
    On Error Resume Next
    
    dlgCommondialog.ShowColor
    
    txtBackColor.BackColor = dlgCommondialog.Color
    lbxLabelBox.BackColor = txtBackColor.BackColor
End Sub

'===========================================================================================================================
Private Sub cmbBorderLine_Click()
    Select Case cmbBorderLine.ListIndex
    Case 0 To 5
        lbxLabelBox.BorderLine = cmbBorderLine.ListIndex
    Case 6
        lbxLabelBox.BorderLine = 8
    Case 7
        lbxLabelBox.BorderLine = 10
    Case 8
        lbxLabelBox.BorderLine = 12
    End Select
End Sub

'===========================================================================================================================
Private Sub txtUpperFont_Change(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case Is = 0
        lbxLabelBox.UpperFontLeft = CInt(txtUpperFont(0).Text)
    Case Is = 1
        lbxLabelBox.UpperFontTop = CInt(txtUpperFont(1).Text)
    Case Is = 2
        lbxLabelBox.UpperFontDepth = CInt(txtUpperFont(2).Text)
    End Select
End Sub

'===========================================================================================================================
Private Sub cmdUpperFontColor_Click(Index As Integer)
    On Error Resume Next
    
    dlgCommondialog.ShowColor
    
    Select Case Index
    Case Is = 0
        txtUpperFontColor(0).BackColor = dlgCommondialog.Color
        lbxLabelBox.UpperFontColorStart = txtUpperFontColor(0).BackColor
    Case Is = 1
        txtUpperFontColor(1).BackColor = dlgCommondialog.Color
        lbxLabelBox.UpperFontColorEnd = txtUpperFontColor(1).BackColor
    End Select
End Sub

'===========================================================================================================================
Private Sub txtBottomFont_Change(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case Is = 0
        lbxLabelBox.BottomFontLeft = CInt(txtBottomFont(0).Text)
    Case Is = 1
        lbxLabelBox.BottomFontTop = CInt(txtBottomFont(1).Text)
    Case Is = 2
        lbxLabelBox.BottomFontDepth = CInt(txtBottomFont(2).Text)
    End Select
End Sub

'===========================================================================================================================
Private Sub cmdBottomrFontColor_Click(Index As Integer)
    On Error Resume Next
    
    dlgCommondialog.ShowColor
    
    Select Case Index
    Case Is = 0
        txtBottomFontColor(0).BackColor = dlgCommondialog.Color
        lbxLabelBox.BottomFontColorStart = txtBottomFontColor(0).BackColor
    Case Is = 1
        txtBottomFontColor(1).BackColor = dlgCommondialog.Color
        lbxLabelBox.BottomFontColorEnd = txtBottomFontColor(1).BackColor
    End Select
End Sub

'===========================================================================================================================
Private Sub cmbTransparentArea_Click()
    lbxLabelBox.TransparentArea = cmbTransparentArea.ListIndex
    
    ' refresh the picture
    Set lbxLabelBox.Wallpaper = LoadPicture(App.Path & "\pictures\ball.bmp")
End Sub

'===========================================================================================================================
Private Sub cmdTransparaentColor_Click()
    On Error Resume Next
    
    dlgCommondialog.ShowColor
    
    txtTransparentColor.BackColor = dlgCommondialog.Color
    lbxLabelBox.TransparentColor = txtTransparentColor.BackColor
    
    ' refresh the picture
    Set lbxLabelBox.Wallpaper = LoadPicture(App.Path & "\pictures\ball.bmp")
End Sub
'===========================================================================================================================

