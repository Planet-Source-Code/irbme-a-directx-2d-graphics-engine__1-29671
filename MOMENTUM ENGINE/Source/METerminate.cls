VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "METerminate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub TerminateME()

On Local Error Resume Next

Dim Counter As Long

AddLog "Terminating engine"

AddLog "Destroying objects"
Call DDraw.SetCooperativeLevel(mvarWindowHandle, DDSCL_NORMAL)
DDraw.RestoreDisplayMode

Set BackBufferSurf = Nothing
Set ScreenSurf = Nothing
Set FontX = Nothing

For Counter = 1 To UBound(surface)
   Set surface(Counter) = Nothing
Next Counter
Erase surface

Erase SurfaceDesc

Set D3DDevice = Nothing
Set DDraw = Nothing
Set DX = Nothing

End Sub
'-----------------------------------------------------------------------------

