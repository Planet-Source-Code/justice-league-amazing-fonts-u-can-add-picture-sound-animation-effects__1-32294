VERSION 5.00
Begin VB.UserControl LabelBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   Enabled         =   0   'False
   PropertyPages   =   "LabelBox.ctx":0000
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   96
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2340
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   0
      Top             =   1980
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "LabelBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'*******************************************************************************'
'*                                                                             *##'
'*                  Programmed by: Aris J. Buenaventura                        *##'
'*                      email: AJB2001LG@YAHOO.COM                             *##'
'*                                                                             *##'
'*******************************************************************************##'
'  ###############################################################################'

' My next version will be much more better and faster than this,
' because I know the technique now, but I don't have enough time
' to do this all over again! If you like it please vote.
'
' To increase the speed use ACTIVEX CONTROL (Required).
'
'      If you can't understand how to use this program, you can mail me.
' SORRY for my properties (lack of time to analyze)
'
' TIP OF THE DAY!
'   Don't you know that Microsft Office contains VB Help.
'
'     1. Open Microsoft Word
'     2. Goto Menu
'     3. Click View
'     4. Select Toolbar
'     5. Check Visual Basic
'     6. Visual Basic Toolbar appears
'     7. Click Visual Basic Editor
'     8. Press F1 (Help)
'
' To search for more hidden VB Help, just search for
'   VB*.hlp and VB*.chm
 
 '===========================================================================================================================
Public Enum BorderLineConstants
    [Flat] = &H0
    [Raised Outer] = &H1
    [Sunken Outer] = &H2
    [Outer] = &H3
    [Raised Inner] = &H4
    [Raised] = &H5
    [Sunken Inner] = &H8
    [Sunken] = &HA
    [Inner] = &HC
End Enum
'===========================================================================================================================

'===========================================================================================================================
Public Enum SetWallpaperConstant
    None = 0
    Stretch = 1
    Tiled = 2
End Enum
'===========================================================================================================================

'===========================================================================================================================
Public Enum TransparentAreaConstant
    None = 0
    Background = 1
    BottomFont = 2
    UpperFont = 3
    Customize = 4
End Enum
'===========================================================================================================================

'===========================================================================================================================
Dim m_Alignment As AlignmentConstants
Dim m_BorderLine As BorderLineConstants
Dim m_BottomFontColorStart As OLE_COLOR
Dim m_UpperFontColorStart As OLE_COLOR
Dim m_UpperFontColorEnd As OLE_COLOR
Dim m_UpperFontDepth As Integer
Dim m_UpperFontLeft As Integer
Dim m_UpperFontTop As Integer
Dim m_Caption As String
Dim m_SetWallpaper As SetWallpaperConstant
Dim m_BottomFontColorEnd As OLE_COLOR
Dim m_BottomFontLeft As Integer
Dim m_BottomFontTop As Integer
Dim m_BottomFontDepth As Integer
Dim m_TransparentArea As TransparentAreaConstant
Dim m_TransparentColor As Variant
Dim m_PictureLeft As Integer
Dim m_PictureTop As Integer
Dim m_PictureWidth As Integer
Dim m_PictureHeight As Integer
Dim m_Wallpaper As Picture
Dim m_WordWrap As Boolean
'===========================================================================================================================

'===========================================================================================================================
Dim NewWallpaper As New StdPicture
'===========================================================================================================================

'===========================================================================================================================
' LabelBox.Sound.Path
' LabelBox.Sound.BeginPlaySound
' LabelBox.Sound.EndPlaySound
' When creating ACTIVEX CONTROL the Instancing (Property clsSound) must
' be set to PublicNotCreatable
Public Sound As New clsSound
'===========================================================================================================================

'===========================================================================================================================
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'===========================================================================================================================

'===========================================================================================================================
Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    Dim OldAlignment As AlignmentConstants
    
    OldAlignment = Alignment
    
    If OldAlignment <> New_Alignment Then
        m_Alignment = New_Alignment
        PropertyChanged "Alignment"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Dim OldBackColor As Long
    
    OldBackColor = UserControl.BackColor()
    
    If OldBackColor <> New_BackColor Then
        UserControl.BackColor = New_BackColor
        PropertyChanged "BackColor"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get BorderLine() As BorderLineConstants
    BorderLine = m_BorderLine
End Property

Public Property Let BorderLine(ByVal New_BorderLine As BorderLineConstants)
    Dim OldBorderLine As BorderLineConstants
    
    OldBorderLine = BorderLine
    
    If OldBorderLine <> New_BorderLine Then
        m_BorderLine = New_BorderLine
        PropertyChanged "BorderLine"
        If BorderLine <> 0 Then
            UserControl_Paint
        Else
            UserControl.Cls
        End If
    End If
End Property

'===========================================================================================================================
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Dim OldCaption As String
    
    OldCaption = Caption
    
    If OldCaption <> New_Caption Then
        m_Caption = New_Caption
        PropertyChanged "Caption"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawLabelBox
End Property

'===========================================================================================================================
Public Property Get PictureLeft() As Integer
    PictureLeft = m_PictureLeft
End Property

Public Property Let PictureLeft(ByVal New_PictureLeft As Integer)
    Dim OldPictureLeft As Integer
    
    OldPictureLeft = PictureLeft
    
    If OldPictureLeft <> New_PictureLeft Then
        m_PictureLeft = New_PictureLeft
        PropertyChanged "PictureLeft"
        If NewWallpaper Then DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get PictureTop() As Variant
    PictureTop = m_PictureTop
End Property

Public Property Let PictureTop(ByVal New_PictureTop As Variant)
    Dim OldPictureTop As Integer
    
    OldPictureTop = PictureTop
    
    If OldPictureTop <> New_PictureTop Then
        m_PictureTop = New_PictureTop
        PropertyChanged "PictureTop"
        If NewWallpaper Then DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get PictureWidth() As Integer
    PictureWidth = m_PictureWidth
End Property

Public Property Let PictureWidth(ByVal New_PictureWidth As Integer)
    Dim OldPictureWidth As Integer
    
    OldPictureWidth = PictureWidth
    
    If OldPictureWidth <> New_PictureWidth Then
        m_PictureWidth = New_PictureWidth
        PropertyChanged "PictureWidth"
        If NewWallpaper Then DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get PictureHeight() As Integer
    PictureHeight = m_PictureHeight
End Property

Public Property Let PictureHeight(ByVal New_PictureHeight As Integer)
    Dim OldPictureHeight As Integer
    
    OldPictureHeight = PictureHeight
    
    If OldPictureHeight <> New_PictureHeight Then
        m_PictureHeight = New_PictureHeight
        PropertyChanged "PictureHeight"
        If NewWallpaper Then DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get SetWallpaper() As SetWallpaperConstant
    SetWallpaper = m_SetWallpaper
End Property

Public Property Let SetWallpaper(ByVal New_SetWallpaper As SetWallpaperConstant)
    Dim OldSetWallpaper As SetWallpaperConstant
    
    OldSetWallpaper = SetWallpaper
    
    If (OldSetWallpaper <> New_SetWallpaper) Then
        m_SetWallpaper = New_SetWallpaper
        PropertyChanged "SetWallpaper"
        
        If NewWallpaper Then Set Wallpaper = m_Wallpaper
    End If
End Property

'===========================================================================================================================
Public Property Get BottomFontColorEnd() As OLE_COLOR
    BottomFontColorEnd = m_BottomFontColorEnd
End Property

Public Property Let BottomFontColorEnd(ByVal New_BottomFontColorEnd As OLE_COLOR)
    Dim OldBottomFontColorEnd As Long
    
    OldBottomFontColorEnd = BottomFontColorEnd
    
    If OldBottomFontColorEnd <> New_BottomFontColorEnd Then
        m_BottomFontColorEnd = New_BottomFontColorEnd
        PropertyChanged "BottomFontColorEnd"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get BottomFontLeft() As Integer
    BottomFontLeft = m_BottomFontLeft
End Property

Public Property Let BottomFontLeft(ByVal New_BottomFontLeft As Integer)
    Dim OldBottomFontLeft As Integer
    
    OldBottomFontLeft = BottomFontLeft
    
    If OldBottomFontLeft <> New_BottomFontLeft Then
        m_BottomFontLeft = New_BottomFontLeft
        PropertyChanged "BottomFontLeft"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get BottomFontTop() As Integer
    BottomFontTop = m_BottomFontTop
End Property

Public Property Let BottomFontTop(ByVal New_BottomFontTop As Integer)
    Dim OldBottomFontTop As Integer
    
    OldBottomFontTop = BottomFontTop
    
    If OldBottomFontTop <> New_BottomFontTop Then
        m_BottomFontTop = New_BottomFontTop
        PropertyChanged "BottomFontTop"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get BottomFontDepth() As Integer
    BottomFontDepth = m_BottomFontDepth
End Property

Public Property Let BottomFontDepth(ByVal New_BottomFontDepth As Integer)
    Dim OldBottomFontDepth As Integer
    
    OldBottomFontDepth = BottomFontDepth
    
    If OldBottomFontDepth <> New_BottomFontDepth Then
        m_BottomFontDepth = New_BottomFontDepth
        PropertyChanged "BottomFontDepth"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get UpperFontDepth() As Integer
    UpperFontDepth = m_UpperFontDepth
End Property

Public Property Let UpperFontDepth(ByVal New_UpperFontDepth As Integer)
    Dim OldUpperFontDepth As Integer
    
    OldUpperFontDepth = UpperFontDepth
    
    If OldUpperFontDepth <> New_UpperFontDepth Then
        m_UpperFontDepth = New_UpperFontDepth
        PropertyChanged "UpperFontDepth"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get UpperFontLeft() As Integer
    UpperFontLeft = m_UpperFontLeft
End Property

Public Property Let UpperFontLeft(ByVal New_UpperFontLeft As Integer)
    Dim OldUpperFontLeft As Integer
    
    OldUpperFontLeft = UpperFontLeft
    
    If OldUpperFontLeft <> New_UpperFontLeft Then
        m_UpperFontLeft = New_UpperFontLeft
        PropertyChanged "UpperFontLeft"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get UpperFontTop() As Integer
    UpperFontTop = m_UpperFontTop
End Property

Public Property Let UpperFontTop(ByVal New_UpperFontTop As Integer)
    Dim OldUpperFontTop As Integer
    
    OldUpperFontTop = UpperFontTop
    
    If OldUpperFontTop <> New_UpperFontTop Then
        m_UpperFontTop = New_UpperFontTop
        PropertyChanged "UpperFontTop"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get TransparentArea() As TransparentAreaConstant
    TransparentArea = m_TransparentArea
End Property

Public Property Let TransparentArea(ByVal New_TransparentArea As TransparentAreaConstant)
    Dim OldTransparentArea As TransparentAreaConstant
    
    OldTransparentArea = TransparentArea
    
    If OldTransparentArea <> New_TransparentArea Then
        m_TransparentArea = New_TransparentArea
        PropertyChanged "TransparentArea"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = m_TransparentColor
End Property

Public Property Let TransparentColor(ByVal New_TransparentColor As OLE_COLOR)
    Dim OldTransparentColor As Long
    
    OldTransparentColor = TransparentColor
     
    If OldTransparentColor <> New_TransparentColor Then
        m_TransparentColor = New_TransparentColor
        PropertyChanged "TransparentColor"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get BottomFontColorStart() As OLE_COLOR
    BottomFontColorStart = m_BottomFontColorStart
End Property

Public Property Let BottomFontColorStart(ByVal New_BottomFontColorStart As OLE_COLOR)
    Dim OldBottomFontColorStart As Long
    
    OldBottomFontColorStart = BottomFontColorStart
    
    If OldBottomFontColorStart <> New_BottomFontColorStart Then
        m_BottomFontColorStart = New_BottomFontColorStart
        PropertyChanged "BottomFontColorStart"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get UpperFontColorStart() As OLE_COLOR
    UpperFontColorStart = m_UpperFontColorStart
End Property

Public Property Let UpperFontColorStart(ByVal New_UpperFontColorStart As OLE_COLOR)
    Dim OldUpperFontColorStart As Long
    
    OldUpperFontColorStart = UpperFontColorStart
    
    If OldUpperFontColorStart <> New_UpperFontColorStart Then
        m_UpperFontColorStart = New_UpperFontColorStart
        PropertyChanged "UpperFontColorStart"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get UpperFontColorEnd() As OLE_COLOR
    UpperFontColorEnd = m_UpperFontColorEnd
End Property

Public Property Let UpperFontColorEnd(ByVal New_UpperFontColorEnd As OLE_COLOR)
    Dim OldUpperFontColorEnd As Long
    
    OldUpperFontColorEnd = UpperFontColorEnd
    
    If OldUpperFontColorEnd <> New_UpperFontColorEnd Then
        m_UpperFontColorEnd = New_UpperFontColorEnd
        PropertyChanged "UpperFontColorEnd"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get Wallpaper() As Picture
    Set Wallpaper = m_Wallpaper
End Property

Public Property Set Wallpaper(ByVal New_Wallpaper As Picture)
    Dim pict As New StdPicture
    
    Set m_Wallpaper = New_Wallpaper
    PropertyChanged "Wallpaper"
    
    Set pict = Wallpaper
    
    If pict Then
        Set NewWallpaper = GetWallpaper(Wallpaper, SetWallpaper)
    Else
        Set NewWallpaper = Nothing
        Set UserControl.Picture = Nothing
    End If
    
    DrawLabelBox
End Property

'===========================================================================================================================
Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    Dim OldWordWrap As Boolean
    
    OldWordWrap = WordWrap
    
    If OldWordWrap <> New_WordWrap Then
        m_WordWrap = New_WordWrap
        PropertyChanged "WordWrap"
    
        picText.Width = UserControl.ScaleWidth
        picText.Height = UserControl.ScaleHeight
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = UserControl.FontBold()
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Dim OldFontBold As Boolean
    
    OldFontBold = FontBold
    
    If OldFontBold <> New_FontBold Then
        UserControl.FontBold() = New_FontBold
        PropertyChanged "FontBold"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = UserControl.FontItalic()
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Dim OldFontItalic As Boolean
    
    OldFontItalic = UserControl.FontItalic()
    
    If OldFontItalic <> New_FontItalic Then
        UserControl.FontItalic() = New_FontItalic
        PropertyChanged "FontItalic"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = UserControl.FontName()
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Dim OldFontName As String
    
    OldFontName = UserControl.FontName()
    
    If OldFontName <> New_FontName Then
        UserControl.FontName() = New_FontName
        PropertyChanged "FontName"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get FontSize() As Single
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = UserControl.FontSize()
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Dim OldFontSize As Single
    
    OldFontSize = UserControl.FontSize()
    
    If OldFontSize <> New_FontSize Then
        UserControl.FontSize() = New_FontSize
        PropertyChanged "FontSize"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = UserControl.FontStrikethru()
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Dim OldFontStrikethru As Boolean
    
    OldFontStrikethru = UserControl.FontStrikethru()
    
    If OldFontStrikethru <> New_FontStrikethru Then
        UserControl.FontStrikethru() = New_FontStrikethru
        PropertyChanged "FontStrikethru"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = UserControl.FontUnderline()
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Dim OldFontUnderline As Boolean
    
    OldFontUnderline = UserControl.FontUnderline()
    
    If OldFontUnderline <> New_FontUnderline Then
        UserControl.FontUnderline() = New_FontUnderline
        PropertyChanged "FontUnderline"
        DrawLabelBox
    End If
End Property

'===========================================================================================================================
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled()
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'===========================================================================================================================
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon()
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon() = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'===========================================================================================================================
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer()
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'===========================================================================================================================
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'===========================================================================================================================
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'===========================================================================================================================
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'===========================================================================================================================
Private Sub UserControl_Initialize()
    UpperFontDepth = 1
End Sub

'===========================================================================================================================

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'===========================================================================================================================
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'===========================================================================================================================
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'===========================================================================================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'===========================================================================================================================
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'===========================================================================================================================
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'===========================================================================================================================
Public Property Get HyperLink() As HyperLink
    Set HyperLink = UserControl.HyperLink
End Property

Private Sub UserControl_Paint()
    If BorderLine <> Flat Then
        DrawBorderLine UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BorderLine
    End If
End Sub

'===========================================================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Alignment = PropBag.ReadProperty("Alignment", DT_LEFT)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_BorderLine = PropBag.ReadProperty("BorderLine", 0)
    m_Caption = PropBag.ReadProperty("Caption", vbNullString)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_SetWallpaper = PropBag.ReadProperty("SetWallpaper", 0)
    m_BottomFontColorEnd = PropBag.ReadProperty("BottomFontColorEnd", &H0)
    m_BottomFontLeft = PropBag.ReadProperty("BottomFontLeft", 0)
    m_BottomFontTop = PropBag.ReadProperty("BottomFontTop", 0)
    m_BottomFontDepth = PropBag.ReadProperty("BottomFontDepth", 0)
    m_TransparentArea = PropBag.ReadProperty("TransparentArea", 0)
    m_TransparentColor = PropBag.ReadProperty("TransparentColor", &H80000012)
    m_UpperFontDepth = PropBag.ReadProperty("UpperFontDepth", 1)
    m_UpperFontLeft = PropBag.ReadProperty("UpperFontLeft", 0)
    m_UpperFontTop = PropBag.ReadProperty("UpperFontTop", 0)
    Set m_Wallpaper = PropBag.ReadProperty("Wallpaper", Nothing)
    m_PictureLeft = PropBag.ReadProperty("PictureLeft", 0)
    m_PictureTop = PropBag.ReadProperty("PictureTop", 0)
    m_PictureWidth = PropBag.ReadProperty("PictureWidth", 0)
    m_PictureHeight = PropBag.ReadProperty("PictureHeight", 0)
    m_WordWrap = PropBag.ReadProperty("WordWrap", False)
    m_BottomFontColorStart = PropBag.ReadProperty("BottomFontColorStart", &H80000012)
    m_UpperFontColorStart = PropBag.ReadProperty("UpperFontColorStart", &H80000012)
    m_UpperFontColorEnd = PropBag.ReadProperty("UpperFontColorEnd", 0)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", False)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", False)
    UserControl.FontName = PropBag.ReadProperty("FontName", UserControl.FontName)
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 8)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", False)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

'===========================================================================================================================
Private Sub UserControl_Resize()
    On Error Resume Next
     
    Set Wallpaper = m_Wallpaper
    picText.Width = UserControl.ScaleWidth
    picText.Height = UserControl.ScaleHeight
    DrawLabelBox
    UserControl_Paint
End Sub

'===========================================================================================================================
Private Sub UserControl_Show()
    Set Wallpaper = UserControl.Extender.Wallpaper
    BackColor = UserControl.BackColor
    BottomFontDepth = UserControl.Extender.BottomFontDepth
    UpperFontDepth = UserControl.Extender.UpperFontDepth
End Sub

'===========================================================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", m_Alignment, DT_LEFT)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BorderLine", m_BorderLine, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, vbNullString)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("SetWallpaper", m_SetWallpaper, 0)
    Call PropBag.WriteProperty("BottomFontColorEnd", m_BottomFontColorEnd, &H0)
    Call PropBag.WriteProperty("BottomFontLeft", m_BottomFontLeft, 0)
    Call PropBag.WriteProperty("BottomFontTop", m_BottomFontTop, 0)
    Call PropBag.WriteProperty("BottomFontDepth", m_BottomFontDepth, 0)
    Call PropBag.WriteProperty("TransparentArea", m_TransparentArea, 0)
    Call PropBag.WriteProperty("TransparentColor", m_TransparentColor, &H80000012)
    Call PropBag.WriteProperty("UpperFontDepth", m_UpperFontDepth, 1)
    Call PropBag.WriteProperty("UpperFontLeft", m_UpperFontLeft, 0)
    Call PropBag.WriteProperty("UpperFontTop", m_UpperFontTop, 0)
    Call PropBag.WriteProperty("Wallpaper", m_Wallpaper, Nothing)
    Call PropBag.WriteProperty("PictureLeft", m_PictureLeft, 0)
    Call PropBag.WriteProperty("PictureTop", m_PictureTop, 0)
    Call PropBag.WriteProperty("WordWrap", m_WordWrap, False)
    Call PropBag.WriteProperty("PictureWidth", m_PictureWidth, 0)
    Call PropBag.WriteProperty("PictureHeight", m_PictureHeight, 0)
    Call PropBag.WriteProperty("BottomFontColorStart", m_BottomFontColorStart, &H80000012)
    Call PropBag.WriteProperty("UpperFontColorStart", m_UpperFontColorStart, &H80000012)
    Call PropBag.WriteProperty("UpperFontColorEnd", m_UpperFontColorEnd, &H0)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, False)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, False)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, UserControl.FontName)
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 8)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, False)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, False)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

'===========================================================================================================================
Private Function GetWallpaper(ByVal wp As StdPicture, ByVal opt As SetWallpaperConstant) As StdPicture
    Dim BM As BITMAP
    Dim cx As Integer, cy As Integer
    Dim SaveWallpaper As New StdPicture
    On Error Resume Next
    
    Set SaveWallpaper = wp
    
    If GetObject(SaveWallpaper, Len(BM), BM) Then ' Get the width and height of the picture
        With UserControl
            .AutoRedraw = True
            .Picture = Nothing
            
            Select Case opt
            Case Is = 0 ' None
                .PaintPicture SaveWallpaper, 0, 0, , , , , , , SRCCOPY  ' copy the picture
            Case Is = 1 ' Stretch
                .PaintPicture SaveWallpaper, 0, 0, .ScaleWidth, .ScaleHeight, , , , , SRCCOPY  ' copy the picture
            Case Is = 2 ' Tiled
                For cy = 0 To .ScaleHeight / BM.bmHeight
                    For cx = 0 To .ScaleWidth / BM.bmWidth
                        .PaintPicture SaveWallpaper, 0, 0, , , _
                            -(cx * BM.bmWidth), -(cy * BM.bmHeight), , , SRCCOPY ' copy the picture
                    Next cx
                Next cy
            End Select
            
            Set GetWallpaper = .Image
            .AutoRedraw = False
        End With
    Else
        Set GetWallpaper = Nothing
    End If
End Function

'===========================================================================================================================
Public Function GetPicture() As StdPicture
    Set GetPicture = UserControl.Picture
End Function

'===========================================================================================================================
Public Sub Save(ByVal Filename As String)
    SavePicture UserControl.Picture, Filename
End Sub
'===========================================================================================================================
Private Sub DrawLabelBox()
    Dim TColor As Long
    Dim PosX As Integer, PosY As Integer
    Dim PicWidth As Integer, PicHeight As Integer
    On Error Resume Next
    
    With UserControl
        .AutoRedraw = True
        Set .Picture = Nothing
        Set picText = Nothing
        Set picText.Font = .Font
        
        picText.BackColor = .BackColor
        
        ' Bottom Caption
        DrawText picText, Caption, BottomFontLeft, BottomFontTop, .ScaleWidth, _
            .ScaleHeight, BottomFontDepth, BottomFontColorStart, BottomFontColorEnd
        
        ' Upper Caption
        DrawText picText, Caption, UpperFontLeft, UpperFontTop, .ScaleWidth, _
            .ScaleHeight, UpperFontDepth, UpperFontColorStart, UpperFontColorEnd
                    
        If NewWallpaper Then
            Set .Picture = picText.Image
            
            PosX = IIf(PictureLeft < 0, Abs(PictureLeft), -PictureLeft)
            PosY = IIf(PictureTop < 0, Abs(PictureTop), -PictureTop)
            
            PicWidth = IIf(PictureWidth > .ScaleWidth, .ScaleWidth, PictureWidth)
            PicHeight = IIf(PictureHeight > .ScaleHeight, .ScaleHeight, PictureHeight)
            
            PicWidth = PicWidth + Abs(PictureLeft)
            PicHeight = PicHeight + Abs(PictureTop)
            
            If (PictureWidth = 0) And (PictureHeight = 0) Then
                ' use default width and height
                .PaintPicture NewWallpaper, 0, 0, , , PosX, PosY, , , SRCAND ' combines foreground and background colors
            ElseIf (PictureWidth = 0) And (PictureHeight <> 0) Then
                ' set user width: use default height
                .PaintPicture NewWallpaper, 0, 0, , , PosX, PosY, , PicHeight, SRCAND ' combines foreground and background colors
            ElseIf (PictureWidth <> 0) And (PictureHeight = 0) Then
                ' use default width: set user height
                .PaintPicture NewWallpaper, 0, 0, , , PosX, PosY, PicWidth, , SRCAND ' combines foreground and background colors
            Else
                ' set user width and height
                .PaintPicture NewWallpaper, 0, 0, , , PosX, PosY, PicWidth, PicHeight, SRCAND ' combines foreground and background colors
            End If
                
           .Picture = .Image ' convert image to picture
        Else
            Set .Picture = .picText.Image
        End If

        Select Case TransparentArea
        Case Is = 0 ' None
        Case Is = 1 ' Background
            TColor = .BackColor
        Case Is = 2, 3
            TColor = _
                IIf(TransparentArea = UpperFont, UpperFontColorStart, BottomFontColorStart)
        Case Is = 4
            TColor = TransparentColor
        End Select
        
        If TransparentArea Then
            .BackStyle = 0 ' Transparent
            .MaskColor = TColor ' select the color to be mask(Transparent)
            .MaskPicture = _
                IIf(TransparentArea <> Customize, picText.Image, .Image)
        Else
            .BackStyle = 1 ' Opaque
        End If
        
        .AutoRedraw = False
    End With
End Sub

'===========================================================================================================================
Private Sub DrawText(ByVal pBox As PictureBox, ByVal Text As String, Optional ByVal sx As Integer = 0, _
    Optional ByVal sy As Integer = 0, Optional ByVal dx As Integer = 0, Optional ByVal dy As Integer = 0, _
    Optional ByVal Depth As Integer = 0, Optional ByVal TColor As Long = &HFFFFF, Optional ByVal BColor As Long = &H0)
    
    Dim sd As Integer
    Dim X As Integer, Y As Integer
    Dim R1 As Integer, G1 As Integer, B1 As Integer
    Dim R2 As Integer, G2 As Integer, B2 As Integer
    Dim Sr As Integer, Sg As Integer, Sb As Integer
            
    ' This will draw 3D Text
    ' Split the color to be able to draw 3D Text
    RGBSplit BColor, R1, G1, B1
    RGBSplit TColor, R2, G2, B2
        
    Sr = (R2 - R1) / Depth
    Sg = (G2 - G1) / Depth
    Sb = (B2 - B1) / Depth
        
    For sd = 0 To Depth - 1
        X = (sx + sd + 4)
        Y = (sy + sd + 4)
        
        ' The values must >= 0
        R1 = IIf(R1 < 0, 0, R1 + Sr)
        G1 = IIf(G1 < 0, 0, G1 + Sg)
        B1 = IIf(B1 < 0, 0, B1 + Sb)
        
        If sd <> Depth - 1 Then
            pBox.ForeColor = RGB(R1, G1, B1)
        Else
            pBox.ForeColor = TColor
        End If
        
        PutText pBox.hDC, Text, X, Y, dx - 4, dy - 4, Alignment, WordWrap
    Next sd
End Sub

'===========================================================================================================================
Public Sub Refresh()
    UserControl.Refresh
End Sub
'===========================================================================================================================

'=============================================================================================
Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Display the copyright dialog."
Attribute ShowAbout.VB_UserMemId = -552
    MsgBox "LabelBox ver 1.0" & Chr(13) & "Programmed by: Aris Buenaventura" _
        & Chr(13) & "Email : AJB2001LG@YAHOO.COM" & Chr(13) _
        & "February 18 2002 - March 2 2002", , "LabelBox"
End Sub
'=============================================================================================
