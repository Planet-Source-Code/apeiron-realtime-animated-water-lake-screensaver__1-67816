Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean

'Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Const AC_SRC_OVER = &H0

Public Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetTickCount& Lib "kernel32" ()

Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

'Public Type POINTAPI
'    X As Long
'    Y As Long
'End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SW_SHOWNORMAL = 1
Public Const ULW_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_OPAQUE = &H4
Public Const LWA_ALPHA = &H2

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HWNDPARENT = (-8)

Public Const WS_CHILD = &H40000000
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_WINDOWEDGE = &H100&

Public Const AC_SRC_ALPHA As Long = &H1

Public Const HWND_TOPMOST As Long = -1
Public Const HWND_TOP            As Long = 0

'Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long)
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SETSCREENSAVEACTIVE = 17

Public bCancel As Boolean
Public DisplayMode As String
Public PreviewHWND As Long

' Configuration variables
Public birdAnim As Boolean
Public fishAnim As Boolean
Public shipAnim As Boolean
Public waterTransparency As Byte
Public showBack As Boolean
Public showBeach As Boolean

Public Function GetHwndFromCmd(Cmd As String) As Long

  Dim i As Long

  For i = 1 To Len(Cmd)
    If Not IsNumeric(Right$(Cmd, i + 1)) Then
      GetHwndFromCmd = CLng(Right$(Cmd, i))
      Exit For
    End If
  Next i

End Function

Public Sub UnloadAll()
    On Error Resume Next
    Dim f As Form
    For Each f In Forms
        Unload f
    Next
    End
End Sub

Public Sub ReadIni()
    On Error Resume Next
    Dim iniPath As String
    Dim fFile As Integer
    Dim iniString As String
    Dim iniSplit() As String
    Dim tempString As String
    
    fFile = FreeFile
      
    If Right(App.Path, 1) <> "\" Then
        iniPath = App.Path & "\LakeScreenSaver.ini"
    Else
        iniPath = App.Path & "LakeScreenSaver.ini"
    End If
    
    
    If Dir$(iniPath) <> "" Then
        Open iniPath For Input As #fFile
            Do While Not EOF(fFile)
                Input #fFile, tempString
                iniString = iniString & tempString
            Loop
        Close #fFile
    
        iniSplit = Split(iniString, vbCrLf)
        
        birdAnim = CBool(Mid$(iniSplit(0), InStr(1, iniSplit(0), "=") + 1))
        fishAnim = CBool(Mid$(iniSplit(1), InStr(1, iniSplit(1), "=") + 1))
        shipAnim = CBool(Mid$(iniSplit(2), InStr(1, iniSplit(2), "=") + 1))
        waterTransparency = CByte(Mid$(iniSplit(3), InStr(1, iniSplit(3), "=") + 1))
        showBack = CBool(Mid$(iniSplit(4), InStr(1, iniSplit(4), "=") + 1))
        showBeach = CBool(Mid$(iniSplit(5), InStr(1, iniSplit(5), "=") + 1))
    Else ' defaults
        birdAnim = True
        fishAnim = True
        shipAnim = True
        waterTransparency = 200
        showBack = True
        showBeach = True
    End If
    
End Sub
