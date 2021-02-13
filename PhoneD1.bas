Attribute VB_Name = "Module1"
Option Explicit

Public lSound As Long

Declare Function sndSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const BM_GETCHECK = &HF0
Public Const BM_SETSTATE = &HF3
Public Const BM_SETCHECK = &HF1
Public Const BM_GETSTATE = &HF2


