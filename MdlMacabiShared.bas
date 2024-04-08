Attribute VB_Name = "MdlMacabiShared"
Private ftextcontrol As New Dictionary
 Private lpPrevWndProc As New Dictionary
   Declare Function CallWindowProc Lib "user32" Alias _
      "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
      ByVal hwnd As Long, ByVal Msg As Long, _
      ByVal wParam As Long, ByVal lParam As Long) As Long

      Declare Function SetWindowLong Lib "user32" Alias _
      "SetWindowLongA" (ByVal hwnd As Long, _
      ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

      Private Const GWL_WNDPROC = -4
'      Public IsHooked As Boolean
     ' Global lpPrevWndProc As Long
'      Global gHW As Long
      Const WM_CHAR = &H102
      Const WM_KEYUP = &H101
      Const WM_KEYDOWN = &H100
      Const WM_PARENTNOTIFY = &H210
      Const WM_UPDATUISTATE = &H128
     
Public Function nte(e As Variant) As Variant
2500      nte = IIf(IsNull(e), "", e)
End Function
      Public Sub Hook(hw As Long, ftxt As FreeTextTemplateCtrl)
      '          If IsHooked Then
      '          MsgBox "Don't hook it twice without " & _
      '            "unhooking, or you will be unable to unhook it."
      '          Else
                  Dim temp As Long
2510              If Not ftextcontrol.Exists(hw) Then
2520                  ftextcontrol.Add hw, ftxt
                      'Set ftextcontrol(hw) = ftxt
2530                  temp = SetWindowLong(hw, GWL_WNDPROC, _
                      AddressOf WindowProc)
2540                  lpPrevWndProc.Add hw, temp
2550              End If
                
                
                
      '          IsHooked = True
      '          End If
      End Sub

      Public Sub Unhook(hw As Long)
                Dim temp As Long
                
2560            DoEvents
2570            If lpPrevWndProc.Exists(hw) Then
2580              temp = SetWindowLong(hw, GWL_WNDPROC, lpPrevWndProc(hw))
2590              lpPrevWndProc.Remove (hw)
2600              If ftextcontrol.Exists(hw) Then
2610                  ftextcontrol.Remove (hw)
2620              End If
                  
2630            End If
                
      '          IsHooked = False
      End Sub

      Function WindowProc(ByVal hw As Long, ByVal uMsg As _
      Long, ByVal wParam As Long, ByVal lParam As Long) As Long
                'Debug.Print "Message: "; hw, uMsg, wParam, lParam
      '          If (uMsg <> WM_CHAR And uMsg <> WM_KEYUP And _
      '          uMsg <> WM_KEYDOWN And uMsg <> WM_PARENTNOTIFY And uMsg <> WM_UPDATUISTATE) Then
2640              If (uMsg = WM_CHAR And wParam >= 224 And wParam <= 250) Then
      '            MsgBox uMsg & " " & wParam & " " & lParam
2650                  If ftextcontrol.Exists(hw) Then
2660                      ftextcontrol(hw).HandleHebrewChar (wParam)
2670                  End If
      '            MsgBox hw & " " & wParam
      '                WindowProc = CallWindowProc(lpPrevWndProc, hw, _
      '                uMsg, wParam - 96, lParam)
                      
2680               ElseIf (uMsg <> WM_KEYDOWN) Then
2690                  If ftextcontrol.Exists(hw) Then
2700                      WindowProc = CallWindowProc(lpPrevWndProc(hw), hw, _
                          uMsg, wParam, lParam)
2710                  End If
2720               Else
2730                 DoEvents
2740                 If ftextcontrol.Exists(hw) Then
2750                      WindowProc = CallWindowProc(lpPrevWndProc(hw), hw, _
                          uMsg, wParam, lParam)
2760                  End If
2770              End If
      End Function



