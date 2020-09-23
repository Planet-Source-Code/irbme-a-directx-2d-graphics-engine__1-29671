VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MEScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Event ErrorOccur(ByVal Procedure As String, ByVal ErrorCode As Long)
'-----------------------------------------------------------------------------


Public Property Let ColorDepth(ByVal lData As Long)
    mvarColorDepth = lData
    AddLog "Screen depth = " & lData
End Property

Public Property Get ColorDepth() As Long
    ColorDepth = mvarColorDepth
End Property
'-----------------------------------------------------------------------------


Public Property Let WindowHandle(ByVal lData As Long)
    mvarWindowHandle = lData
   AddLog "Window handle = " & lData
End Property

Public Property Get WindowHandle() As Long
    WindowHandle = mvarWindowHandle
End Property
'-----------------------------------------------------------------------------


Public Property Let ResolutionY(ByVal lData As Long)
    mvarResolutionY = lData
   AddLog "ResolutionY = " & lData
End Property

Public Property Get ResolutionY() As Long
    ResolutionY = mvarResolutionY
End Property
'-----------------------------------------------------------------------------


Public Property Let ResolutionX(ByVal lData As Long)
    mvarResolutionX = lData
    AddLog "ResolutionX = " & lData
End Property

Public Property Get ResolutionX() As Long
    ResolutionX = mvarResolutionX
End Property
'-----------------------------------------------------------------------------


Public Sub EnumDispModes(ByRef ResolutionX() As Long, ByRef ResolutionY() As Long, ByRef ColorDepth() As Long, ByRef lCount As Long)

On Local Error GoTo Error

Dim TempDDraw As DirectDraw7
Dim DisplayModesEnum As DirectDrawEnumModes
Dim DispMode As DDSURFACEDESC2
Dim Counter As Integer

AddLog "Enumerating display modes"

AddLog "Creating temporary object"
Set TempDDraw = DX.DirectDrawCreate("")
Call TempDDraw.SetCooperativeLevel(mvarWindowHandle, DDSCL_NORMAL)

AddLog "Rediminsioning dynamic arrays"
ReDim DispModes(1)
Set DisplayModesEnum = TempDDraw.GetDisplayModesEnum(0, DispMode)
ReDim DispModes(1 To DisplayModesEnum.GetCount)

AddLog "Looping through display modes"

lCount = DisplayModesEnum.GetCount

For Counter = 1 To DisplayModesEnum.GetCount
   DisplayModesEnum.GetItem Counter, DispMode
   ReDim Preserve ResolutionX(1 To Counter)
   ReDim Preserve ResolutionY(1 To Counter)
   ReDim Preserve ColorDepth(1 To Counter)

   ResolutionX(Counter) = DispMode.lWidth
   ResolutionY(Counter) = DispMode.lHeight
   ColorDepth(Counter) = DispMode.ddpfPixelFormat.lRGBBitCount
Next Counter

AddLog "Deleting temporary object"
Set TempDDraw = Nothing

AddLog "Enumeration of Display modes was successful"

Exit Sub
Error:
   AddLog "Enumeration of display modes failed"
   On Local Error Resume Next
   Set TempDDraw = Nothing
   RaiseEvent ErrorOccur("MEScreen: EnumDispModes", Err.Number)

End Sub
'-----------------------------------------------------------------------------


Public Function TestDisplayMode(ByVal ResolutionX As Long, ByVal ResolutionY As Long, ByVal ColorDepth As Long) As Boolean

On Local Error GoTo Error

Dim TempDDraw As DirectDraw7
Dim DisplayModesEnum As DirectDrawEnumModes
Dim DispMode As DDSURFACEDESC2
Dim Counter As Integer

AddLog "Testing display mode " & ResolutionX & "X" & ResolutionY & "X" & ColorDepth

Set TempDDraw = DX.DirectDrawCreate("")
Call TempDDraw.SetCooperativeLevel(mvarWindowHandle, DDSCL_NORMAL)
Set DisplayModesEnum = TempDDraw.GetDisplayModesEnum(0, DispMode)

TestDisplayMode = False

For Counter = 1 To DisplayModesEnum.GetCount
   DisplayModesEnum.GetItem Counter, DispMode
   If ResolutionY = DispMode.lHeight And ResolutionX = DispMode.lWidth And ColorDepth = DispMode.ddpfPixelFormat.lRGBBitCount Then
      TestDisplayMode = True
      Exit For
   End If
Next Counter

AddLog "Display mode " & IIf(TestDisplayMode, "Supported", "Unsupported")

Set TempDDraw = Nothing

Exit Function
Error:
   AddLog "Error in testing. Failed"
   On Local Error Resume Next
   Set TempDDraw = Nothing
   RaiseEvent ErrorOccur("MEScreen: TestDisplayMode", Err.Number)

End Function
'-----------------------------------------------------------------------------


Public Function GetAdapterInfo() As DirectDrawIdentifier

On Local Error GoTo Error

Dim TempDDraw As DirectDraw7

Set TempDDraw = DX.DirectDrawCreate("")
Call TempDDraw.SetCooperativeLevel(mvarWindowHandle, DDSCL_NORMAL)

AddLog "Getting adapter info"

Set GetAdapterInfo = TempDDraw.GetDeviceIdentifier(DDGDI_DEFAULT)

Set TempDDraw = Nothing

Exit Function
Error:
   AddLog "Error getting adapters"
   If Not TempDDraw Is Nothing Then Set TempDDraw = Nothing
   RaiseEvent ErrorOccur("MEScreen: GetAdapterInfo", Err.Number)

End Function
'-----------------------------------------------------------------------------
