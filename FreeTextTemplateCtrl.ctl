VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl FreeTextTemplateCtrl 
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   ScaleHeight     =   4185
   ScaleWidth      =   4905
   Begin VB.Timer tmrBackup 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1800
      Top             =   1800
   End
   Begin VB.Frame fraFreeTextTemplate 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin RichTextLib.RichTextBox TxtFreeTextTemplate 
         Height          =   3855
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   6800
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"FreeTextTemplateCtrl.ctx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu MnLists 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnListName 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "FreeTextTemplateCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CTRL_G = 7

Public Connection As ADODB.Connection
Private rst As ADODB.Recordset
Public InitContent As String
Public Lists As String

'Maps a list name (header) to a collection (the list items);
'Each such list item maps the text to be selected to a Snomed-M value;
Private ListsNames As New Dictionary
'Private ListsNames As New Collection

Public Event DblClick()
Public Event OnChange()
Public Event ShowList(ListIndex As Integer)
Public Event BackupRecordExists()
Private Const StrReplace = "[]"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'catch keyup and transmit another key
' This is used with the SetWindowLong API function.
'Private Const GWL_WNDPROC = -4
'Private Const WM_KEYUP = &H101
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private OldWindowProc As Long


Private Const WM_CHAR = &H102
Private InFindMode As Boolean

Private hadFocus As Boolean
Private strResultId As String

Private Const csHeBrEw As String = "iso-8859-8" ' Hebrew character set

  

Public Sub AssignList2RTF(Docklist As DockListCtrl, ListIndex As Integer)
2890  On Error GoTo ERR_AssignList2RTF

          Dim ListName As String
          Dim arr As Variant
          Dim i As Integer
          Dim d As Dictionary
          
2900      ListName = MnListName(ListIndex).Caption
2910      If ListName = "" Then Exit Sub
2920      With Docklist
2930          .setFreeText TxtFreeTextTemplate
2940          .ListFontName = TxtFreeTextTemplate.Font.Name
2950          .ListFontSize = TxtFreeTextTemplate.Font.Size
2960          .RemoveAllItems

2970          Set d = ListsNames(ListName)
2980          For i = 0 To d.Count - 1
2990               Call .ListAddItemAndSnomed(CStr(d.Keys(i)), CStr(d.Items(i)))
                  '.ListAddItem CStr(d.Keys(i))
3000              .ToolTipText i + 1, CStr(d.Keys(i))
3010          Next i
      '        For i = 1 To ListsNames(ListName).Count
      '            .ListAddItem ListsNames(ListName).Item(i)
      '            .ToolTipText i, ListsNames(ListName).Item(i)
      '        Next i
3020      End With

3030      Exit Sub
ERR_AssignList2RTF:
3040  MsgBox "Error on line:" & Erl & "  in AssignList2RTF" & vbCrLf & Err.DESCRIPTION
End Sub

Private Sub MnListName_Click(Index As Integer)
3050      RaiseEvent ShowList(Index)
End Sub

Public Sub TxtFreeTextTemplate_RightMouseUp()
3060  On Error GoTo ERR_TxtFreeTextTemplate_RightMouseUp

3070     If MnListName.Count = 1 And MnListName(0).Caption = "" Then Exit Sub
3080     PopupMenu MnLists

3090      Exit Sub
ERR_TxtFreeTextTemplate_RightMouseUp:
3100  MsgBox "Error on line:" & Erl & "  in TxtFreeTextTemplate_RightMouseUp" & vbCrLf & Err.DESCRIPTION
End Sub



Private Sub tmrBackup_Timer()
3110      Call BackupResult
End Sub

Private Sub TxtFreeTextTemplate_Change()
3120  On Error GoTo ERR_TxtFreeTextTemplate_Change
          
3130      RaiseEvent OnChange
          
3140      Exit Sub
ERR_TxtFreeTextTemplate_Change:
3150  MsgBox "Error on line:" & Erl & "  in TxtFreeTextTemplate_Change" & vbCrLf & Err.DESCRIPTION
End Sub

Public Property Get Lines() As Integer
3160      Lines = IIf(Len(TxtFreeTextTemplate.Text) = 0, 0, TxtFreeTextTemplate.GetLineFromChar(Len(TxtFreeTextTemplate.Text)) + 1)
End Property

Private Sub TxtFreeTextTemplate_DblClick()
3170      RaiseEvent DblClick
End Sub

Private Sub TxtFreeTextTemplate_GotFocus()
3180  On Error GoTo ERR_TxtFreeTextTemplate_GotFocus

          Dim i As Integer

3190      FindAndReplaceText
          
          'position the cursor in the 2nd line, but not
          'if focus was on this control before:
3200      If hadFocus = False Then
          
3210          i = InStr(1, TxtFreeTextTemplate.Text, vbCrLf)
3220          TxtFreeTextTemplate.SelStart = i + 1
              
3230      End If
          
3240      tmrBackup.Enabled = True
          
3250      Exit Sub
ERR_TxtFreeTextTemplate_GotFocus:
3260  MsgBox "Error on line:" & Erl & "  in TxtFreeTextTemplate_GotFocus" & vbCrLf & Err.DESCRIPTION
End Sub

Private Sub TxtFreeTextTemplate_KeyPress(KeyAscii As Integer)
3270  On Error GoTo ERR_TxtFreeTextTemplate_KeyPress

3280      If KeyAscii = CTRL_G Then
3290          Call TxtFreeTextTemplate_RightMouseUp
3300      ElseIf KeyAscii = vbKeyReturn And InFindMode Then
3310          Call FindAndReplaceText
3320          KeyAscii = 0
3330      End If


3340      Exit Sub
ERR_TxtFreeTextTemplate_KeyPress:
3350  MsgBox "Error on line:" & Erl & "  in TxtFreeTextTemplate_KeyPress" & vbCrLf & Err.DESCRIPTION
End Sub

Private Sub TxtFreeTextTemplate_KeyUp(KeyCode As Integer, Shift As Integer)
3360  On Error GoTo ERR_TxtFreeTextTemplate_KeyUp
          
3370      If KeyCode = vbKeyF2 And Shift = 0 Then
3380          Call FindAndReplaceText
3390      End If
          
3400      Exit Sub
ERR_TxtFreeTextTemplate_KeyUp:
3410  MsgBox "Error on line:" & Erl & "  in TxtFreeTextTemplate_KeyUp" & vbCrLf & Err.DESCRIPTION
End Sub

Private Sub TxtFreeTextTemplate_LostFocus()
3420      hadFocus = True
End Sub

Private Sub TxtFreeTextTemplate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
3430  On Error GoTo ERR_TxtFreeTextTemplate_MouseUp

3440     If Button <> vbRightButton Then Exit Sub
3450     If MnListName.Count = 1 And MnListName(0).Caption = "" Then Exit Sub
3460     If MnListName.Count = 1 And MnListName(0).Caption <> "" Then
             
3470        RaiseEvent ShowList(0)
3480     Else
3490        PopupMenu MnLists
3500     End If
         
3510      Exit Sub
ERR_TxtFreeTextTemplate_MouseUp:
3520  MsgBox "Error on line:" & Erl & "  in TxtFreeTextTemplate_MouseUp" & vbCrLf & Err.DESCRIPTION
End Sub

Private Sub UserControl_GotFocus()
3530      TxtFreeTextTemplate.SetFocus
End Sub

Private Sub UserControl_Initialize()
3540      InFindMode = False
3550      TxtFreeTextTemplate.Left = 0
3560      TxtFreeTextTemplate.Top = 0
3570      TxtFreeTextTemplate.width = UserControl.width
3580      TxtFreeTextTemplate.height = UserControl.height
3590      fraFreeTextTemplate.width = UserControl.width
3600      fraFreeTextTemplate.height = UserControl.height
3610      If fraFreeTextTemplate.Caption <> "" Then
3620          fraFreeTextTemplate.width = UserControl.width - 80
3630          TxtFreeTextTemplate.width = UserControl.width - 80
3640          TxtFreeTextTemplate.height = UserControl.height - 250
3650          TxtFreeTextTemplate.Top = 250
3660      End If
      '    OldWindowProc = SetWindowLong( _
      '    fraFreeTextTemplate.hwnd, GWL_WNDPROC, _
      '    AddressOf NewWindowProc)

             ' Hook (TxtFreeTextTemplate.hwnd)


End Sub


Public Sub Initialize()

3670      TxtFreeTextTemplate.TextRTF = FormateInitContent(InitContent)
      '    If Mid(TxtFreeTextTemplate.Text, 1, 5) = "אבחנה" Then

3680          Hook TxtFreeTextTemplate.hwnd, Me

      '    End If
          
3690      BuildListsMenu (Lists)
3700      hadFocus = False
End Sub

Public Sub HandleHebrewChar(c As Long)
          'MsgBox c
3710      TxtFreeTextTemplate.SelText = Chr(c)
End Sub

Public Sub UpdateRTF()
3720      TxtFreeTextTemplate.TextRTF = FormateInitContent(InitContent)
End Sub

Public Function GetContent() As String
3730      GetContent = Replace(TxtFreeTextTemplate.Text, vbCrLf, "<BR>")
End Function

Public Function GetRTFContent() As String
3740      GetRTFContent = TxtFreeTextTemplate.TextRTF
End Function


Public Function FormateInitContent(ParamStr As String) As String
          Dim SQLStr As String
          Dim temp As String
          Dim RstResult As String
          Dim i As Integer

3750      SQLStr = ParseSQL(ParamStr)
3760      If SQLStr = "" Then
3770          FormateInitContent = ParamStr
3780          Exit Function
3790      End If
3800      On Error GoTo BadSQL
3810      Set rst = Connection.Execute(SQLStr)
3820      On Error GoTo 0

3830      RstResult = ""
3840      While Not rst.EOF
3850          For i = 0 To rst.Fields.Count - 1
3860              RstResult = RstResult & rst(i) & " "
3870          Next i
3880          RstResult = RstResult & vbCrLf
3890          rst.MoveNext
3900      Wend
          
3910      FormateInitContent = FormateInitContent(Replace(ParamStr, "~^" & SQLStr & "~^", RstResult, 1, 1))
3920      Exit Function
BadSQL:
3930      Call SaveRtfTemplateError(Err.DESCRIPTION, ParamStr)
3940      FormateInitContent = ParamStr
End Function


Private Sub SaveRtfTemplateError(strError As String, strReceivedText As String)
3950  On Error GoTo ERR_SaveRtfTemplateError

          Dim rs As Recordset
          Dim sql As String
          
3960      Call Connection.BeginTrans
          
3970      Set rs = Connection.Execute("select lims.sq_u_rtf_template_error.nextval from dual")

3980      sql = " insert into lims_sys.u_rtf_template_error"
3990      sql = sql & " ("
4000      sql = sql & " u_rtf_template_error_id,"
4010      sql = sql & " name,"
4020      sql = sql & " description,"
4030      sql = sql & " version,"
4040      sql = sql & " version_status"
4050      sql = sql & " )"
4060      sql = sql & " values"
4070      sql = sql & "("
4080      sql = sql & " '" & rs(0) & "',"
4090      sql = sql & " '" & rs(0) & "',"
4100      sql = sql & " '" & Left(strError, 2000) & "',"
4110      sql = sql & " '1',"
4120      sql = sql & " 'A'"
4130      sql = sql & ")"
4140      Call Connection.Execute(sql)

4150      sql = " insert into lims_sys.u_rtf_template_error_user"
4160      sql = sql & " ("
4170      sql = sql & " u_rtf_template_error_id,"
4180      sql = sql & " u_created_on,"
4190      sql = sql & " u_received_text"
4200      sql = sql & " )"
4210      sql = sql & " values"
4220      sql = sql & "("
4230      sql = sql & " '" & rs(0) & "',"
4240      sql = sql & " sysdate,"
4250      sql = sql & " '" & Left(Replace(strReceivedText, "'", "''"), 2000) & "'"
4260      sql = sql & ")"
4270      Call Connection.Execute(sql)

4280      Call Connection.CommitTrans


4290      Exit Sub
ERR_SaveRtfTemplateError:
      'MsgBox "Error on line:" & Erl & "  in SaveRtfTemplateError" & vbCrLf & Err.DESCRIPTION
End Sub


Private Function ParseSQL(ParamStr As String, Optional Separator As String = "~^") As String
          Dim p1 As Integer
          Dim p2 As Integer
4300      p1 = InStr(1, ParamStr, Separator)
4310      p2 = InStr(p1 + 2, ParamStr, Separator)
4320      If p2 = 0 Then
4330          ParseSQL = ""
4340      Else
4350          ParseSQL = Mid(ParamStr, p1 + 2, p2 - (p1 + 2))
4360      End If
End Function

Private Sub UserControl_Resize()
4370      TxtFreeTextTemplate.Left = 0
4380      TxtFreeTextTemplate.Top = 0
4390      TxtFreeTextTemplate.width = UserControl.width
4400      TxtFreeTextTemplate.height = UserControl.height
4410      fraFreeTextTemplate.Left = 0
4420      fraFreeTextTemplate.Top = 0
4430      fraFreeTextTemplate.width = UserControl.width
4440      fraFreeTextTemplate.height = UserControl.height
4450      If fraFreeTextTemplate.Caption <> "" Then
4460          fraFreeTextTemplate.width = UserControl.width - 80
4470          TxtFreeTextTemplate.width = UserControl.width - 80
4480          TxtFreeTextTemplate.height = UserControl.height - 250
4490          TxtFreeTextTemplate.Top = 250
4500      End If
End Sub

Private Sub UserControl_Show()
4510      TxtFreeTextTemplate.Left = 0
4520      TxtFreeTextTemplate.Top = 0
4530      TxtFreeTextTemplate.width = UserControl.width
4540      TxtFreeTextTemplate.height = UserControl.height
4550      fraFreeTextTemplate.width = UserControl.width
4560      fraFreeTextTemplate.height = UserControl.height
4570      If fraFreeTextTemplate.Caption <> "" Then
4580          fraFreeTextTemplate.width = UserControl.width - 80
4590          TxtFreeTextTemplate.width = UserControl.width - 80
4600          TxtFreeTextTemplate.height = UserControl.height - 250
4610          TxtFreeTextTemplate.Top = 250
4620      End If
End Sub

Private Sub BuildListsMenu(ParamStr As String)
4630  On Error GoTo ERR_BuildListsMenu
          
          Dim SQLStr As String
          Dim temp As String
          Dim newListName As String
          Dim oldListName As String
          Dim ListItem As String
          Dim SnomedM As String
          Dim ListsItems As Dictionary
      '    Dim ListsItems As Collection
          Dim i As Integer
              
4640      ListsNames.RemoveAll
4650      i = 0
4660      Do While ParamStr <> ""
4670          SQLStr = ParseSQL(ParamStr, "$$")
4680          ParamStr = Mid(ParamStr, Len(SQLStr) + 2)
4690          If SQLStr = "" Then
4700              Exit Sub
4710          End If
4720          Set rst = Connection.Execute(SQLStr)
4730          If Not rst.EOF Then
4740              newListName = oldListName = ""
4750          End If
4760          While Not rst.EOF
4770              newListName = rst(0)
4780              If newListName <> oldListName Then
4790                  oldListName = newListName
4800                  Set ListsItems = New Dictionary
                      'Set ListsItems = New Collection
4810                  ListsNames.Add newListName, ListsItems
                      'ListsNames.Add ListsItems, newListName
4820                  If i > 0 Then
4830                      Load MnListName(i)
4840                  End If
4850                  MnListName(i).Caption = newListName
4860                  i = i + 1
4870              End If
4880              ListItem = rst(1)
4890              If rst.Fields.Count < 3 Then
4900                  SnomedM = ""
4910              Else
4920                  SnomedM = nte(rst(2))
4930              End If
                  
4940              If ListsItems.Exists(ListItem) = False Then
4950                  ListsItems.Add ListItem, SnomedM
4960              End If
                  
4970              rst.MoveNext
4980          Wend
4990      Loop
          
5000      Exit Sub
ERR_BuildListsMenu:
5010  MsgBox "Error on line:" & Erl & "  in BuildListsMenu" & vbCrLf & Err.DESCRIPTION & vbCrLf & "ParamStr:" & ParamStr
End Sub

'Private Function ntes(e As Variant) As String
'    ntes = IIf(IsNull(e), "", e)
'End Function


Public Sub Terminate()
          Dim i As Integer

5020      Call ListsNames.RemoveAll
      '    While ListsNames.Count > 0
      '        While ListsNames(ListsNames.Count - 1).Count > 0
      '            ListsNames(ListsNames.Count - 1).Remove (ListsNames(ListsNames.Count - 1).Count - 1)
      '        Wend
      '        ListsNames.Remove (ListsNames.Count - 1)
      '    Wend

      '    While ListsNames.Count > 0
      '        While ListsNames(ListsNames.Count).Count > 0
      '            ListsNames(ListsNames.Count).Remove (ListsNames(ListsNames.Count).Count)
      '        Wend
      '        ListsNames.Remove (ListsNames.Count)
      '    Wend
5030      Set ListsNames = Nothing
5040      While MnListName.Count > 1
5050          Unload MnListName(MnListName.Count - 1)
5060      Wend
5070      MnListName(MnListName.Count - 1).Caption = ""
5080      Unhook TxtFreeTextTemplate.hwnd
5090      Set rst = Nothing
          
          
          'Call RemoveBackupResult
End Sub

Public Property Let Locked(ByVal vNewValue As Boolean)
5100      TxtFreeTextTemplate.Locked = vNewValue
End Property

Public Property Let Rtl(ByVal vNewValue As Boolean)
5110      If vNewValue = True Then
      '        TxtFreeTextTemplate.Alignment = vbRightJustify
      '        TxtFreeTextTemplate.RightToLeft = True
5120      Else
      '        TxtFreeTextTemplate.Alignment = vbLeftJustify
      '        TxtFreeTextTemplate.RightToLeft = False
5130      End If
5140      fraFreeTextTemplate.RightToLeft = vNewValue
End Property

Public Property Get Rtl() As Boolean
5150      Rtl = fraFreeTextTemplate.RightToLeft
End Property

Public Property Let FontName(fname As String)
5160      TxtFreeTextTemplate.Font.Name = fname
End Property

Public Property Let RightMargin(RMargin As Long)
5170      TxtFreeTextTemplate.RightMargin = RMargin
End Property

Public Property Get RTBHandle() As Long
5180      RTBHandle = TxtFreeTextTemplate.hwnd
End Property

Public Sub FindAndReplaceText()
          Dim StrStart As String
          Dim StrEnd As String
          Dim idxStart As Integer
          Dim idxEnd As Integer
          Dim ReplaceText As String
          Dim FindText As String
          Dim Text As String
5190      StrStart = Left(StrReplace, 1)
5200      StrEnd = Right(StrReplace, 1)
5210      Text = TxtFreeTextTemplate.Text
5220      idxStart = TxtFreeTextTemplate.Find(StrStart, TxtFreeTextTemplate.SelStart)
5230      If idxStart > 0 Then
5240          idxEnd = TxtFreeTextTemplate.Find(StrEnd, idxStart)
5250          If idxEnd < 0 Then
5260              InFindMode = False
5270              Exit Sub
5280          End If
5290          TxtFreeTextTemplate.SelStart = idxStart
5300          TxtFreeTextTemplate.SelLength = idxEnd - idxStart + 1
5310          InFindMode = True
5320      Else
5330          InFindMode = False
5340      End If
End Sub

Private Sub TypeText(str As String)
5350      If Trim(str) = "" Then SendKeys (vbKeySpace), True
5360      While Len(str) > 0
5370          Call SendMessage(TxtFreeTextTemplate.hwnd, WM_CHAR, Asc(Mid(str, 1, 1)), Null)
5380          str = Mid(str, 2, Len(str) - 1)
5390      Wend
End Sub

Public Property Let Caption(ByVal vNewValue As String)
5400      fraFreeTextTemplate.Caption = vNewValue
End Property

Public Property Get Caption() As String
5410      Caption = fraFreeTextTemplate.Caption
End Property

Public Property Let ResultId(strResultIdInit)

      'msgbox "resultid 1" & strResultIdInit
5420      strResultId = strResultIdInit
      'msgbox "resultid 1"
5430      If BackupExists = True Then
          
              'notify the user program that a backup record exists
              'which means that the last work on this result
              'ended in a crash:
              'msgbox "resultid 1eeeee22"
5440          RaiseEvent BackupRecordExists
          'msgbox "resultid 1eeee"
5450      End If
          
End Property


Private Function BackupExists() As Boolean
5460  On Error GoTo ERR_BackupExists


      'msgbox "BackupExists 1"
          Dim rs As Recordset
          Dim sql As String


5470      If strResultId = "" Then Exit Function
      'msgbox "BackupExists 2"
5480      sql = "select rtf_result_id from lims_sys.rtf_result_backup " & _
                   "where rtf_result_id = '" & strResultId & "'"
          'msgbox "BackupExists 3"
5490      Set rs = Connection.Execute(sql)
          'msgbox "BackupExists 4"
5500      If rs.EOF Then
          
5510          BackupExists = False
              
5520      Else
          
5530          BackupExists = True
              
5540      End If
          
      'msgbox "BackupExists 5"
5550      Exit Function
ERR_BackupExists:
5560  MsgBox "Error on line:" & Erl & "  in BackupExists" & vbCrLf & Err.DESCRIPTION
End Function

'delete the RTF backup record:
Public Sub RemoveBackupResult()
5570  On Error GoTo ERR_RemoveBackupResult
      'MsgBox "1"
          Dim sql As String
          
5580      If strResultId = "" Then Exit Sub

5590      sql = " delete lims_sys.rtf_result_backup rrb"
5600      sql = sql & " where rrb.rtf_result_id = '" & strResultId & "'"
      '    MsgBox "2"
5610      Call Connection.Execute(sql)
      'MsgBox "23"
5620      Exit Sub
ERR_RemoveBackupResult:
5630  MsgBox "Error on line:" & Erl & "  in RemoveBackupResult" & vbCrLf & Err.DESCRIPTION
End Sub


'update the RTF contents
'from the backup table:
Public Sub ReadFromBackup()
5640  On Error GoTo ERR_ReadFromBackup

          Dim rs As New Recordset


5650      If strResultId = "" Then Exit Sub

5660      Call rs.Open(" select rtf_text from lims_sys.rtf_result_backup " & _
                       " where rtf_result_id = '" & strResultId & "'", _
                       Connection, adOpenStatic, adLockOptimistic)
          
          
5670      If rs.EOF = False Then
          
5680          InitContent = ReadClob(rs("rtf_text"))
5690          Call UpdateRTF
          
5700      End If

5710      Exit Sub
ERR_ReadFromBackup:
5720  MsgBox "Error on line:" & Erl & "  in ReadFromBackup" & vbCrLf & Err.DESCRIPTION
End Sub


Private Function ReadClob(pFld As ADODB.Field) As String

          ' Function read a the clob data from the field
          ' using the stream object of the ADODB library

          Dim lStream As ADODB.Stream
          Dim lstData As String

5730      Set lStream = New ADODB.Stream
5740      lStream.Charset = csHeBrEw
5750      lStream.Type = adTypeText
5760      lStream.Open

5770      lStream.WriteText nte(pFld.Value)
5780      lStream.Position = 0
5790      lstData = lStream.ReadText

5800      lStream.Close
5810      Set lStream = Nothing
          
5820      ReadClob = lstData
          
End Function


'save the RTF contents to the
'backup RTF table:
Private Sub BackupResult() '(RtfResultId As String, FreeTextCtrl As FreeTextTemplateCtrl)
5830      On Error GoTo ErrEnd
          Dim RtfResult As ADODB.Recordset
      '    Dim RtfResultId As String
          Dim ResSTR As String
          Dim lStream As ADODB.Stream

5840      If strResultId = "" Then
5850          Exit Sub
5860      End If

              
5870      If BackupExists = False Then
                  
5880          ResSTR = "insert into lims_sys.rtf_result_backup (rtf_result_id) values ('" & _
                        strResultId & "')"
5890          Call Connection.Execute(ResSTR)
          
5900      End If
          
5910      Set RtfResult = New ADODB.Recordset

5920      Call RtfResult.Open("select rtf_text from lims_sys.rtf_result_backup where rtf_result_id = " & strResultId, Connection, adOpenStatic, adLockOptimistic)

5930      Set lStream = New ADODB.Stream
5940      lStream.Charset = csHeBrEw
5950      lStream.Type = adTypeText
5960      lStream.Open

5970      lStream.WriteText TxtFreeTextTemplate.TextRTF 'TxtFreeText.GetRTFContent
5980      lStream.Position = 0
5990      RtfResult("RTF_TEXT").Value = lStream.ReadText
          
6000      lStream.Close
6010      Set lStream = Nothing
6020      RtfResult.Update

6030      RtfResult.Close
6040      Set RtfResult = Nothing
          
          

6050      Exit Sub
ErrEnd:
6060      MsgBox "BackupResult... " & vbCrLf & _
                  Err.DESCRIPTION
End Sub

Private Sub UserControl_Terminate()
    'Call RemoveBackupResult
    
End Sub
