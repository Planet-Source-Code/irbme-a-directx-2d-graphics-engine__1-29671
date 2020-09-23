VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MEInitialise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ErrorOccur(ByVal Procedure As String, ByVal ErrorCode As Long)
'-----------------------------------------------------------------------------


Public Function InitialiseME() As Boolean

On Local Error GoTo Error

Dim RenderState As String
Dim DevEnum As Direct3DEnumDevices
Dim Scrn As New MEScreen
Dim Backbuffer As New MEBackBuffer
Dim Counter As Long

AddLog "Initialisation sub called"

Set DX = New DirectX7

AddLog "Testing screen mode"

If Not Scrn.TestDisplayMode(mvarResolutionX, mvarResolutionY, mvarColorDepth) Then
   InitialiseME = False
   AddLog "Screen mode unsupported"
   Err.Number = DDERR_UNSUPPORTEDMODE
   GoTo Error
End If

AddLog "Creating Device"
Set DDraw = DX.DirectDrawCreate("")

AddLog "Setting cooperate levels and display modes"
Call DDraw.SetCooperativeLevel(mvarWindowHandle, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
Call DDraw.SetDisplayMode(mvarResolutionX, mvarResolutionY, mvarColorDepth, 0, DDSDM_DEFAULT)

AddLog "Creating backbuffer"
Backbuffer.CreateBackBuffer

AddLog "Initialising 3D device"
Set D3D = DDraw.GetDirect3D
AddLog "Creating hardware enumeration"
Set DevEnum = D3D.GetDevicesEnum
     
AddLog "enumerating hardware     "
For Counter = 1 To DevEnum.GetCount
   If DevEnum.GetGuid(Counter) = "IID_IDirect3DRGBDevice" Then
      RenderState = "IID_IDirect3DRGBDevice"
   ElseIf DevEnum.GetGuid(Counter) = "IID_IDirect3DHALDevice" Then
      RenderState = "IID_IDirect3DHALDevice"
      Exit For
   End If
Next Counter

AddLog "Device used = " & RenderState

AddLog "Creating 3D device"
Set D3DDevice = D3D.CreateDevice(RenderState, BackBufferSurf)

InitialiseME = True

Set Scrn = Nothing
Set Backbuffer = Nothing

AddLog "Initialsiation was successful"

Exit Function
Error:
   AddLog "Initialisation failed"
   On Local Error Resume Next
   Set Backbuffer = Nothing
   Set Scrn = Nothing
   Set ScreenSurf = Nothing
   Set DDraw = Nothing
   Set DX = Nothing
   RaiseEvent ErrorOccur("MEInitialise: InitialiseME", Err.Number)

End Function
'-----------------------------------------------------------------------------

