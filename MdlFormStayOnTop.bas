Attribute VB_Name = "MdlFormStayOnTop"
'***********************************************
' Name: StayOnTop
' Description: Keep a form always on top
'***********************************************

'***********************************************
' Windows API/Global Declarations for :StayOnTop
'***********************************************

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
    
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer

    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .width / Screen.TwipsPerPixelX
        iHeight = .height / Screen.TwipsPerPixelY
    End With

    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub

