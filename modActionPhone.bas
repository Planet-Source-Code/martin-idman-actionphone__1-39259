Attribute VB_Name = "modActionPhone"
' Force variable declaration
Option Explicit

' Public API functions
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

' Public subs and functions

' Sub to sleep x seconds
Public Function Sleep(ByVal lngSleepSeconds As Long)
   ' Dim temp variable
    Dim lngSleepEnd As Long
   ' Set end time
   lngSleepEnd = GetTickCount + lngSleepSeconds * 1000
   ' Loop until end-time
   While GetTickCount <= lngSleepEnd
      ' Leave resources to your computer to do other things
      DoEvents
   Wend
End Function
