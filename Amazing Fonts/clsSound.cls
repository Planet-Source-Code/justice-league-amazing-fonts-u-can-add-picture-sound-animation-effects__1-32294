VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===========================================================================================================================
Public Enum ConstantsPlayOption
    Asynchronously = &H0        '  play asynchronously
    Synchrounously = &H1        '  play synchronously (default)
End Enum
'===========================================================================================================================

'===========================================================================================================================
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_FILENAME = &H20000      '  name is a file name
'===========================================================================================================================

'===========================================================================================================================
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'===========================================================================================================================

'===========================================================================================================================
Public Path As String
'===========================================================================================================================

'===========================================================================================================================
Public Sub BeginPlaySound(Optional ByVal PlayOption As ConstantsPlayOption = Synchrounously)
    sndPlaySound Path, PlayOption Or SND_FILENAME
End Sub

'===========================================================================================================================
Public Sub EndPlaySound()
    sndPlaySound ByVal vbNullString, &H0
End Sub
'===========================================================================================================================


