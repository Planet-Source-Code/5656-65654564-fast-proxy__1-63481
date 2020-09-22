Attribute VB_Name = "ModuleTrans"

Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204


' allow you to pause a timer in real time (very accurate)
Public Sub Pause(Secs As Long)
     Dim secStart As Variant
    Dim secNow As Variant
    Dim secDiff As Variant
    Dim Temp%
    
    exitPause = False 'this is our early way out out of the pause
    
    secStart = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'get the starting seconds
    


    Do While secDiff < Secs
        If exitPause = True Then Exit Do
        secNow = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'this is the current time and Date at any itteration of the Loop
        secDiff = DateDiff("s", secStart, secNow) 'this compares the start time With the current time
        Temp% = DoEvents
    Loop
End Sub

Public Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)

Dim hRgn As Long, tRgn As Long
Dim X As Integer, Y As Integer, X0 As Integer
Dim hDC As Long, BM As BITMAP

hDC = CreateCompatibleDC(0)
If hDC Then
    SelectObject hDC, cPicture

    GetObject cPicture, Len(BM), BM
    hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
    For Y = 0 To BM.bmHeight
        For X = 0 To BM.bmWidth
            While X <= BM.bmWidth And GetPixel(hDC, X, Y) <> cTransparent
                X = X + 1
            Wend
            X0 = X
            While X <= BM.bmWidth And GetPixel(hDC, X, Y) = cTransparent
                X = X + 1
            Wend
            If X0 < X Then
                tRgn = CreateRectRgn(X0, Y, X, Y + 1)
                CombineRgn hRgn, hRgn, tRgn, 4
                DeleteObject tRgn
            End If
        Next X
    Next Y
    GetBitmapRegion = hRgn
    DeleteObject SelectObject(hDC, cPicture)
End If

DeleteDC hDC

End Function

