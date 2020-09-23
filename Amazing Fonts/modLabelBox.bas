Attribute VB_Name = "modLabelBox"
Option Explicit

'===========================================================================================================================
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
'===========================================================================================================================

'===========================================================================================================================
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'===========================================================================================================================

'===========================================================================================================================
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_WORDBREAK = &H10
'===========================================================================================================================

'===========================================================================================================================
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
'===========================================================================================================================

'===========================================================================================================================
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'===========================================================================================================================

'===========================================================================================================================
Public Sub PutText(ByVal hDC As Long, ByVal Caption As String, ByVal sx As Long, ByVal sy As Long, ByVal dx As Long, ByVal dy As Long, _
    Optional ByVal Alignment As Long = DT_LEFT, Optional ByVal WordWrap As Boolean = False)
    
    Dim wFormat As Long
    Dim WinRect As RECT
    
    
    If Alignment = 2 Then
        Alignment = DT_CENTER
    ElseIf Alignment = 1 Then
        Alignment = DT_RIGHT
    End If
    
    If Not WordWrap Then
        wFormat = Alignment
    Else
        wFormat = Alignment Or DT_WORDBREAK
    End If
    
    WinRect.Left = sx
    WinRect.Top = sy
    WinRect.Right = dx
    WinRect.Bottom = dy
    
    Call DrawText(hDC, Caption, Len(Caption), WinRect, wFormat)
End Sub

'===========================================================================================================================
Public Function Hex2Int(ByVal HexNum As String) As Integer
    Dim ch As String
    Dim d As Integer, dd As Integer
    
    HexNum = UCase$(HexNum)
    
    ch = Val(Right$(HexNum, 1))
    d = IIf(IsNumeric(ch), Val(ch), 10 + Abs(65 - Asc(UCase$(ch))))
    
    ch = Left$(HexNum, 1)
    dd = IIf(IsNumeric(ch), Val(ch), 10 + Abs(65 - Asc(UCase$(ch))))
    
    Hex2Int = d + 16 * dd
End Function

'===========================================================================================================================
Public Sub RGBSplit(ByVal RGBColor As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
    Dim HexR As String
    Dim HexG As String
    Dim HexB As String
    Dim HexRGB As String
    
    HexRGB = Hex$(RGBColor)
    If Len(HexRGB) < 6 Then HexRGB = String(6 - Len(HexRGB), "0") + HexRGB

    HexR = Right(HexRGB, 2)
    HexG = Mid(HexRGB, 3, 2)
    HexB = Left(HexRGB, 2)
    
    R = Hex2Int(HexR)
    G = Hex2Int(HexG)
    B = Hex2Int(HexB)
End Sub

'===========================================================================================================================
Public Sub DrawBorderLine(ByVal hDC As Long, ByVal sx As Long, ByVal sy As Long, ByVal dx As Long, ByVal dy As Long, _
    Optional ByVal BS = 0)
    
    Dim WinRect As RECT
    
    WinRect.Left = sx
    WinRect.Top = sy
    WinRect.Right = dx
    WinRect.Bottom = dy
    
    Call DrawEdge(hDC, WinRect, BS, &H100F)
End Sub
'===========================================================================================================================
