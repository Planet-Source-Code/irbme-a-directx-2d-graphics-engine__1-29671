VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "METimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
'-----------------------------------------------------------------------------


Private NeedDestroy As Boolean

Public Sub Delay(ByVal Milliseconds As Long)
   Sleep Milliseconds
   AddLog "Delaying by " & Milliseconds & "ms"
End Sub
'-----------------------------------------------------------------------------


Public Function IsCounterSupported() As Boolean

Dim liFrequency As LARGE_INTEGER

AddLog "Querying counter"

If QueryPerformanceFrequency(liFrequency) = 0 Then
   IsCounterSupported = False
   AddLog "Counter not supported"
Else
   IsCounterSupported = True
   AddLog "Counter supported"
End If

End Function
'-----------------------------------------------------------------------------


Public Function GetPerformanceFrequency() As Currency

Dim liFrequency As LARGE_INTEGER

AddLog "Querying frequency of counter"

QueryPerformanceFrequency liFrequency
GetPerformanceFrequency = LargeIntToCurrency(liFrequency)

AddLog "Frequency is " & GetPerformanceFrequency

End Function
'-----------------------------------------------------------------------------


Public Function GetPerformanceCounter() As Currency

Dim liFrequency As LARGE_INTEGER

AddLog "Querying performance of counter"
QueryPerformanceCounter liFrequency
GetPerformanceCounter = LargeIntToCurrency(liFrequency)

AddLog "Performance is " & GetPerformanceCounter

End Function
'-----------------------------------------------------------------------------


Public Sub CreateTimer(ByVal Frequency As Long, ByVal TmrProcAddr As Long)
   AddLog "Creating timer"
   SetTimer mvarWindowHandle, 0, Frequency, TmrProcAddr
   NeedDestroy = True
End Sub
'-----------------------------------------------------------------------------


Public Sub DestroyTimer()
   AddLog "Destroying timer"
   KillTimer mvarWindowHandle, 0
   NeedDestroy = False
End Sub
'-----------------------------------------------------------------------------


Public Function QueryTickCount() As Long
   AddLog "Querying tickcount"
   QueryTickCount = GetTickCount
   AddLog "Tickcount is " & QueryTickCount
End Function
'-----------------------------------------------------------------------------


Private Sub Class_Terminate()
   If NeedDestroy Then DestroyTimer
End Sub
'-----------------------------------------------------------------------------

