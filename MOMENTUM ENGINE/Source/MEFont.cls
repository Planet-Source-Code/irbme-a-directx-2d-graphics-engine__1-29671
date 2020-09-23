VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MEFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Event ErrorOccur(ByVal Procedure As String, ByVal ErrorCode As Long)
'-----------------------------------------------------------------------------


Public Property Let FontName(ByVal lData As String)
    mvarFontName = lData
    AddLog "Font name = " & lData
End Property

Public Property Get FontName() As String
    FontName = mvarFontName
End Property
'-----------------------------------------------------------------------------


Public Property Let FontSize(ByVal lData As Long)
   mvarFontSize = lData
   AddLog "Font size = " & lData
End Property

Public Property Get FontSize() As Long
    FontSize = mvarFontSize
End Property
'-----------------------------------------------------------------------------


Public Property Let FontTransparent(ByVal bData As Boolean)
   mvarFontTransparent = bData
   AddLog "Font transparent = " & bData
End Property

Public Property Get FontTransparent() As Boolean
    FontTransparent = mvarFontTransparent
End Property
'-----------------------------------------------------------------------------


Public Property Let FontBackColor(ByVal lData As Long)
    mvarFontBackColor = lData
    AddLog "Font backcolor = " & lData
End Property

Public Property Get FontBackColor() As Long
    FontTransparent = mvarFontBackColor
End Property
'-----------------------------------------------------------------------------


Public Sub CreateFont()

On Local Error GoTo Error

AddLog "Create font sub called"

FontX.Size = mvarFontSize
FontX.Name = mvarFontName
BackBufferSurf.SetFont FontX
BackBufferSurf.SetFontTransparency mvarFontTransparent
BackBufferSurf.SetFontBackColor mvarFontBackColor

AddLog "Create font sub was sucessful"

Exit Sub
Error:
   AddLog "Create font sub failed"
   RaiseEvent ErrorOccur("MEFont: CreateFont", Err.Number)

End Sub
'-----------------------------------------------------------------------------



Public Sub DrawFont(ByVal x As Long, ByVal Y As Long, ByVal Text As String)

On Local Error GoTo Error

Call BackBufferSurf.DrawText(x, Y, Text, False)

Exit Sub
Error:
   AddLog "Draw font sub failed"
   RaiseEvent ErrorOccur("MEFont: DrawFont", Err.Number)

End Sub
'-----------------------------------------------------------------------------

