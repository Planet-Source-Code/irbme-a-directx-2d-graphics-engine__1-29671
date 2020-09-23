VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MEBackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Event ErrorOccur(ByVal Procedure As String, ByVal ErrorCode As Long)
'-----------------------------------------------------------------------------


Public Property Let BackBufferCount(ByVal lData As Long)
   mvarBackBufferCount = lData
   AddLog "BackbufferCount = " & lData
End Property

Public Property Get BackBufferCount() As Long
   BackBufferCount = mvarBackBufferCount
End Property
'-----------------------------------------------------------------------------


Public Property Let ForeColor(ByVal lData As Long)
    mvarForeColor = lData
    AddLog "Forecolor of backbuffer  = " & lData
End Property

Public Property Get ForeColor() As Long
    ForeColor = mvarForeColor
End Property
'-----------------------------------------------------------------------------


Public Property Let BackColor(ByVal lData As Long)
   mvarBackColor = lData
   AddLog "Backcolor of backbuffer = " & lData
End Property

Public Property Get BackColor() As Long
    BackColor = mvarBackColor
End Property
'-----------------------------------------------------------------------------


Friend Sub CreateBackBuffer()

On Local Error GoTo Error

Dim BackBufferCaps As DDSCAPS2

AddLog "Createbackbuffer sub called:"
AddLog "Setting screen attributes"

ScreenDesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
ScreenDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_3DDEVICE
ScreenDesc.lBackBufferCount = mvarBackBufferCount

AddLog "Setting screen surface"
Set ScreenSurf = DDraw.CreateSurface(ScreenDesc)

AddLog "Creating backbuffer"

BackBufferCaps.lCaps = DDSCAPS_BACKBUFFER
Set BackBufferSurf = ScreenSurf.GetAttachedSurface(BackBufferCaps)
BackBufferSurf.GetSurfaceDesc BackbufferDesc
BackBufferSurf.SetForeColor mvarForeColor

AddLog "Clearing backbuffer"
ClearBackBuffer

AddLog "Create backbuffer sub was successful"

Exit Sub
Error:
   AddLog "Create backbuffe sub failed"
   RaiseEvent ErrorOccur("MEBackBuffer: CreateBackBuffer", Err.Number)

End Sub
'-----------------------------------------------------------------------------


Public Sub ClearBackBuffer()

On Local Error GoTo Error

Dim Dest As RECT

AddLog "Clear backbuffer sub was called"

With Dest
   .Top = 0
   .Left = 0
   .Bottom = mvarResolutionY
   .Right = mvarResolutionX
End With

AddLog "Clearing backbuffer"
BackBufferSurf.BltColorFill Dest, mvarBackColor

AddLog "Clearbackbuffer sub was successful"

Exit Sub
Error:
   AddLog "Clearbackbuffer sub failed"
   RaiseEvent ErrorOccur("MEBackBuffer: ClearBackBuffer", Err.Number)

End Sub
'-----------------------------------------------------------------------------


Public Sub Flip()

On Local Error GoTo Error

ScreenSurf.Flip Nothing, DDFLIP_WAIT

Exit Sub
Error:
   AddLog "flip sub failed"
   RaiseEvent ErrorOccur("MEbackBuffer: Flip", Err.Number)

End Sub
'-----------------------------------------------------------------------------

