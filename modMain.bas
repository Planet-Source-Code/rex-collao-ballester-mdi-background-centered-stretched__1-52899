Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, _
                                             ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
                                             
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
                                                ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent&, ByVal hWndChildAfter&, ByVal lpClassName$, ByVal lpWindowName$) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd&, lpRect As RECT) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd&, lpRect As RECT, ByVal bErase&) As Long

Public Const API_FALSE As Long = 0&
Public Const API_TRUE As Long = 1&

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public fMain As frmMain

Public currentImageFile As String
Public mdibg_mode As Boolean    'True is stretched mode, false is center mode

Public Sub Main()
    
    Set fMain = New frmMain
    Load fMain
    fMain.Show
    
End Sub

Public Sub CreateFormPic(hWnd As Long, pic1 As PictureBox, pic2 As PictureBox, mode As Boolean)
    
    Dim hCleintArea&, rc As RECT
  
    hCleintArea = FindWindowEx(hWnd, 0&, "MDIClient", vbNullChar)
    Call GetClientRect(hCleintArea, rc)
  
    pic2.Width = rc.right * 15 + 75
    pic2.Height = rc.bottom * 15 + 75

    Call InvalidateRect(hCleintArea, rc, API_TRUE)
    Call CenterPic(pic2, pic1, mode)
    Call InvalidateRect(hCleintArea, rc, API_TRUE)
    
End Sub

Public Sub CenterPic(picDest As PictureBox, picSource As PictureBox, mode As Boolean)

    On Error GoTo err1
    Dim left As Long
    Dim top As Long
    
    fMain.Picture = Nothing
    picDest.Picture = Nothing
    
    If picDest.Width > picSource.Width Then
        left = picDest.ScaleWidth \ 2 - picSource.ScaleWidth \ 2
    End If
    
    If picDest.Height > picSource.Height Then
        top = picDest.ScaleHeight \ 2 - picSource.ScaleHeight \ 2
    End If
    
    If mode Then
        Dim a, b, c, d As Long
        a = fMain.ScaleWidth \ 15
        b = fMain.ScaleHeight \ 15
        c = picSource.ScaleWidth
        d = picSource.ScaleHeight
        
        StretchBlt picDest.hdc, 0, 0, a, b, picSource.hdc, 0, 0, c, d, vbSrcCopy
    Else
        BitBlt picDest.hdc, left, top, 1024, 768, picSource.hdc, 0, 0, vbSrcCopy
    End If
    
    fMain.Picture = picDest.Image
    
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description
End Sub
