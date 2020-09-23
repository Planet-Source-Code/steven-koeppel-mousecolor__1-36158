Attribute VB_Name = "modMain"
Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Type RGBCOLOR
    R As Byte
    G As Byte
    B As Byte
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Function LongToRGB(Color As Long) As RGBCOLOR
    LongToRGB.R = Abs(Color) Mod 256
    LongToRGB.G = ((Abs(Color) And &HFF00) / 256&) Mod 256&
    LongToRGB.B = (Abs(Color) And &HFF0000) / 65536
End Function

