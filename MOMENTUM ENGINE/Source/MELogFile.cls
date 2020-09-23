VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MELogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Event ErrorOccur(ByVal Procedure As String, ByVal ErrorCode As Long)
'-----------------------------------------------------------------------------


Friend Property Let FileNumber(ByVal iData As Integer)
    mvarFileNumber = iData
End Property

Friend Property Get FileNumber() As Integer
    FileNumber = mvarFileNumber
End Property
'-----------------------------------------------------------------------------


Public Property Let Filename(ByVal sData As String)
    mvarFileName = sData
End Property

Public Property Get Filename() As String
    Filename = mvarFileName
End Property
'-----------------------------------------------------------------------------


Public Property Let AutoLog(ByVal bData As Boolean)
   mvarAutoLog = bData
End Property

Public Property Get AutoLog() As Boolean
   AutoLog = mvarAutoLog
End Property
'-----------------------------------------------------------------------------


Public Sub AddLogLine(ByVal LogLine As String)
   Print #mvarFileNumber, LogLine
End Sub
'-----------------------------------------------------------------------------


Public Sub StartLog()

AutoLog = True
mvarFileNumber = FreeFile

If Not Filename = "" Then
   Open Filename For Binary As #mvarFileNumber: Close #mvarFileNumber
   Open Filename For Output As #mvarFileNumber
      Print #mvarFileNumber, "Log file opened at " & Time & " on " & Date & " as #" & FileNumber
Else
   RaiseEvent ErrorOccur("MELogFile: StartLogging", 52)
End If
   
End Sub
'-----------------------------------------------------------------------------


Private Sub StopLog()
   On Error Resume Next
   Print #mvarFileNumber, "Log file closed at " & Time & " on " & Date
   Close #mvarFileNumber
End Sub
'-----------------------------------------------------------------------------


Public Sub WriteINI(ByVal Filename As String, ByVal SectionName As String, ByVal DataName As String, ByVal Value As String)
   WritePrivateProfileString SectionName, DataName, Value, Filename
   DoEvents
End Sub
'-----------------------------------------------------------------------------


Public Function ReadINI(Filename As String, ByVal SectionName As String, ByVal DataName As String) As String

Static GetStr As String * 256, LenStr As Long

LenStr = GetPrivateProfileString(SectionName, DataName, "Error reading: " & Filename, GetStr, 256, Filename)
ReadINI = Left(GetStr, LenStr)

End Function
'-----------------------------------------------------------------------------


Private Sub Class_Terminate()
   StopLog
End Sub
'-----------------------------------------------------------------------------

