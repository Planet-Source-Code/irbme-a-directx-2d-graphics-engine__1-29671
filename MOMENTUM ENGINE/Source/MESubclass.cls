VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MESubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4

Private NeedUnhook As Boolean
Private PrevWndProc As Long
'-----------------------------------------------------------------------------


Public Sub Hook(ByVal WndProcAddr As Long)
   AddLog "Hooking window"
    PrevWndProc = SetWindowLong(mvarWindowHandle, GWL_WNDPROC, WndProcAddr)
    NeedUnhook = True
End Sub
'-----------------------------------------------------------------------------


Public Sub UnHook()
   AddLog "Unhooking iwndow"
   SetWindowLong mvarWindowHandle, GWL_WNDPROC, PrevWndProc
   NeedUnhook = False
End Sub
'-----------------------------------------------------------------------------


Private Sub Class_Terminate()
    If NeedUnhook Then UnHook
End Sub
'-----------------------------------------------------------------------------


Public Sub HandleMessages(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
   Call CallWindowProc(PrevWndProc, hw, uMsg, wParam, lParam)
End Sub
'-----------------------------------------------------------------------------

