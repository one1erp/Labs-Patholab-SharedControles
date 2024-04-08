VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl DockListCtrl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   Picture         =   "DockListCtrl.ctx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   6840
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   6855
   End
   Begin VB.TextBox txtLine 
      Height          =   615
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   6735
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3495
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      WordWrap        =   -1  'True
      RightToLeft     =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.ListBox List 
      Height          =   1500
      ItemData        =   "DockListCtrl.ctx":0342
      Left            =   0
      List            =   "DockListCtrl.ctx":0344
      TabIndex        =   4
      Top             =   2040
      Width           =   6855
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   325
      TabIndex        =   1
      ToolTipText     =   "לחץ לעזרה"
      Top             =   60
      Width           =   255
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   50
      TabIndex        =   2
      ToolTipText     =   "לחץ לסגירה"
      Top             =   60
      Width           =   255
   End
   Begin MSComctlLib.ListView lstListItems 
      Height          =   1680
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   2963
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "List Item"
         Object.Width           =   17639
      EndProperty
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "רשימת הערות לטקסט חופשי "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "DockListCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
    
Private Const EM_GETLINECOUNT = &HBA

Private mX As Single
Private mY As Single
Private FreeTxt As Object
Private dicTextToSnomedM As New Dictionary
Private dicTextToSnomedMSelected As New Dictionary
Private Const LIGHT_YELLOW = &HC0FFFF

Public Event CloseList()
Public Event KeyPress(KeyAscii As Integer)
Public Event DblClick()


Public Function GetSelectedList() As Dictionary
7420      Set GetSelectedList = dicTextToSnomedMSelected
End Function


'return the Snomed-M associated with the
'last item selected from the list;
Public Function GetSnomedM() As String
7430  On Error GoTo ERR_GetSnomedM

          'GetSnomedM = dicTextToSnomedM.Items(List.ListIndex)
7440      GetSnomedM = dicTextToSnomedM.Items(grid.Row)

7450      Exit Function
ERR_GetSnomedM:
7460  MsgBox "Error on line:" & Erl & "  in GetSnomedM" & vbCrLf & Err.DESCRIPTION
End Function


Private Sub CmdClose_Click()
7470      dicTextToSnomedMSelected.RemoveAll
7480      RaiseEvent CloseList
End Sub

Private Sub CmdHelp_Click()
          Dim HelpStr As String

7490      HelpStr = "בחר מהרשימה את הטקסט להוספה" & vbCrLf & _
                    "להזזת החלון לחץ על הכותרת וגרור למקום הרצוי"
7500      MsgBox HelpStr, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, "רשימון"
End Sub

Private Sub cmdOK_Click()
7510  On Error GoTo ERR_cmdOK_Click

          Dim i As Integer
          Dim strTotal As String
          
7520      For i = 0 To dicTextToSnomedMSelected.Count - 1
7530          FreeTxt.SelIndent = 0
7540          FreeTxt.SelHangingIndent = 130
7550          If (InStr(1, dicTextToSnomedMSelected.Keys(i), "+")) = 1 Then
7560              FreeTxt.SelText = Replace(dicTextToSnomedMSelected.Keys(i), "+", "") & vbCrLf
7570          Else
7580              FreeTxt.SelText = "- " & dicTextToSnomedMSelected.Keys(i) & vbCrLf
7590          End If
              'FreeTxt.SelBullet = 1
              'FreeTxt.SelHangingIndent = 130
7600      Next i
          
          
      '    For i = 0 To dicTextToSnomedMSelected.Count - 1
      '        strTotal = strTotal & "- " & dicTextToSnomedMSelected.Keys(i) & vbCrLf
      '    Next i
      '
      '    If strTotal <> "" Then
      '        'FreeTxt.SelBullet = 1
      '        FreeTxt.SelText = strTotal
      '    End If
          
      '    For i = 0 To dicTextToSnomedMSelected.Count - 1
      '        If Trim(FreeTxt.Text) = "" Then
      '            FreeTxt.SelText = "- " & dicTextToSnomedMSelected.Keys(i)
      '        Else
      '            FreeTxt.SelText = vbCrLf & "- " & dicTextToSnomedMSelected.Keys(i)
      '        End If
      '    Next i
          
7610      RaiseEvent DblClick
7620      RaiseEvent CloseList

7630      Exit Sub
ERR_cmdOK_Click:
7640  MsgBox "Error on line:" & Erl & "  in cmdOK_Click" & vbCrLf & Err.DESCRIPTION
End Sub

Private Sub grid_Click()
7650  On Error GoTo ERR_grid_Click

7660      If Trim(grid.Text) = "" Then Exit Sub
          
          
7670      grid.Col = 1
7680      If grid.CellBackColor <> LIGHT_YELLOW Then
7690          grid.CellBackColor = LIGHT_YELLOW
7700          Call dicTextToSnomedMSelected.Add(dicTextToSnomedM.Keys(grid.Row), _
                                                dicTextToSnomedM.Items(grid.Row))
7710      Else
7720          grid.CellBackColor = vbWhite
7730          Call dicTextToSnomedMSelected.Remove(dicTextToSnomedM.Keys(grid.Row))
7740      End If
          
          
      '    If Trim(FreeTxt.Text) = "" Then
      '        FreeTxt.SelText = "- " & grid.TextMatrix(grid.Row, 1)
      '    Else
      '        FreeTxt.SelText = vbCrLf & "- " & grid.TextMatrix(grid.Row, 1)
      '    End If
      '    RaiseEvent DblClick

7750      Exit Sub
ERR_grid_Click:
7760  MsgBox "Error on line:" & Erl & "  in grid_Click" & vbCrLf & Err.DESCRIPTION
End Sub



Private Sub LblHeader_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
7770      If Button = vbLeftButton Then
7780          mX = x
7790          mY = y
7800      End If
End Sub

Private Sub LblHeader_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
7810      If Button = vbLeftButton Then
7820          UserControl.Extender.Left = UserControl.Extender.Left - (mX - x)
7830          UserControl.Extender.Top = UserControl.Extender.Top - (mY - y)
7840          UserControl.Extender.Top = IIf(UserControl.Extender.Top > 0, UserControl.Extender.Top, 0)
7850      End If
End Sub

Private Sub List_DblClick()
7860  On Error GoTo ERR_List_DblClick

      '    MsgBox dicTextToSnomed.Items(List.ListIndex)

7870      If Trim(List.Text) = "" Then Exit Sub
7880      If Trim(FreeTxt.Text) = "" Then
7890          FreeTxt.SelText = List.Text
7900      Else
7910          FreeTxt.SelText = vbCrLf & List.Text
7920      End If
7930      RaiseEvent DblClick
          
7940      Exit Sub
ERR_List_DblClick:
7950  MsgBox "Error on line:" & Erl & "  in List_DblClick" & vbCrLf & Err.DESCRIPTION
End Sub

Private Sub List_KeyPress(KeyAscii As Integer)
7960      If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
7970          FreeTxt.SelText = List.Text
7980      End If
7990      RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lstListItems_DblClick()
8000      If Trim(lstListItems.SelectedItem.Text) = "" Then Exit Sub
8010      If Trim(FreeTxt.GetContent) = "" Then
8020          FreeTxt.SelText = lstListItems.SelectedItem.Text
8030      Else
8040          FreeTxt.SelText = vbCrLf & lstListItems.SelectedItem.Text
8050      End If
8060      RaiseEvent DblClick
End Sub


Public Sub setFreeText(FreeText As Object)
8070      Set FreeTxt = FreeText
End Sub

Private Sub lstListItems_GotFocus()
   ' lstListItems.ListItems(1).Selected = True
End Sub

Private Sub lstListItems_KeyPress(KeyAscii As Integer)
8080      If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
8090          FreeTxt.SelText = lstListItems.SelectedItem.Text
8100      End If
8110      RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lstListItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
8120      lstListItems.HitTest(x, y).Selected = True
End Sub

Private Sub UserControl_GotFocus()
'    Call lstListItems.SetFocus
End Sub

Private Sub UserControl_Initialize()
'    InitWidth = UserControl.width
'    InitHeight = UserControl.height
'    LblHeader.Top = 0
'    LblHeader.Left = 0
'    LblHeader.width = UserControl.width - 10
'    lstListItems.Top = LblHeader.height
'    lstListItems.Left = 0
'    lstListItems.width = UserControl.width - 10
'    lstListItems.height = UserControl.height - lstListItems.Top
End Sub

Private Sub UserControl_Resize()
8130      LblHeader.Top = 0
8140      LblHeader.Left = 0
8150      LblHeader.width = UserControl.width - 50
      '    lstListItems.Top = LblHeader.height
      '    lstListItems.Left = 0
      '    lstListItems.width = UserControl.width - 50
      '    lstListItems.height = (UserControl.height - lstListItems.Top - 50) / 2
8160      List.Top = LblHeader.height
8170      List.Left = 0
8180      List.width = UserControl.width - 50
8190      List.height = UserControl.height + 25 ' - 50
          
8200      grid.Top = LblHeader.height
8210      grid.Left = 0
8220      grid.width = UserControl.width - 50
8230      grid.height = UserControl.height - LblHeader.height - cmdOK.height ' - 50
8240      grid.RightToLeft = False

8250      txtLine.Left = 0
8260      txtLine.width = grid.width - 1100
          
8270      cmdOK.Top = grid.Top + grid.height
8280      cmdOK.width = grid.width

          
End Sub

Private Sub UserControl_Terminate()
8290      Set FreeTxt = Nothing
End Sub

Public Property Let ListFontName(FontName As String)
8300      lstListItems.Font.Name = Font
8310      List.Font.Name = Font
8320      grid.Font.Name = Font
8330      txtLine.Font.Name = FontName
End Property

Public Property Let ListFontSize(FontSize As Long)
8340      lstListItems.Font.Size = FontSize
8350      List.Font.Size = FontSize
8360      grid.Font.Size = FontSize
8370      txtLine.Font.Size = FontSize
End Property


Public Sub ListAddItemAndSnomed(ListItem As String, SnomedM As String)
8380  On Error GoTo ERR_ListAddItemAndSnomed

8390      lstListItems.ListItems.Add , , ListItem
8400      List.AddItem ListItem
8410      Call AddGridRow(GetLinesOfTextBox(txtLine, ListItem), ListItem)
          
8420      Call dicTextToSnomedM.Add(ListItem, SnomedM)

8430      Exit Sub
ERR_ListAddItemAndSnomed:
8440  MsgBox "Error on line:" & Erl & "  in ListAddItemAndSnomed" & vbCrLf & Err.DESCRIPTION
End Sub

Public Sub ListAddItem(ListItem As String)
8450      lstListItems.ListItems.Add , , ListItem
8460      List.AddItem ListItem
End Sub

Public Sub RemoveAllItems()
8470      lstListItems.ListItems.Clear
8480      List.Clear
8490      Call dicTextToSnomedM.RemoveAll
8500      Call dicTextToSnomedMSelected.RemoveAll
8510      grid.Rows = 0
End Sub

Public Sub ToolTipText(i As Integer, ToolTip As String)
8520      lstListItems.ListItems(i).ToolTipText = ToolTip
End Sub

Public Property Let Rtl(ByVal vNewValue As Boolean)
8530      UserControl.RightToLeft = vNewValue
8540      List.RightToLeft = True
8550      grid.RightToLeft = vNewValue
End Property

Private Sub AddGridRow(iHeightRows As Integer, strText As String)
8560  On Error GoTo ERR_AddGridRow
          
8570      grid.Rows = grid.Rows + 1
8580      grid.Row = grid.Rows - 1
          
          '1st column: the text number
8590      grid.Col = 0
8600      grid.ColWidth(0) = 500
8610      If grid.RightToLeft = True Then
8620          grid.CellAlignment = vbRightJustify
8630      Else
8640          grid.CellAlignment = vbLeftJustify
8650      End If
8660      grid.Text = grid.Rows
          
          '2nd column: the selection text;
          'the width is set by the txtLine Text Box,
          'since it is used to indicate the number of rows needed
          'to contain the specific text:
8670      grid.Col = 1
8680      grid.ColWidth(1) = txtLine.width + 200
8690      grid.CellAlignment = vbLeftJustify
8700      grid.Text = strText
          
8710      grid.RowHeight(grid.Row) = 39 * grid.Font.Size * ((15 / 24) * iHeightRows + (1 / 3))

8720      Exit Sub
ERR_AddGridRow:
8730  MsgBox "Error on line:" & Erl & "  in AddGridRow" & vbCrLf & Err.DESCRIPTION
End Sub

Private Function GetLinesOfTextBox(tb As TextBox, str As String) As Integer
8740  On Error GoTo ERR_GetLinesOfTextBox

8750      tb.Text = str
8760      GetLinesOfTextBox = SendMessage(tb.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)

8770      Exit Function
ERR_GetLinesOfTextBox:
8780  MsgBox "Error on line:" & Erl & "  in GetLinesOfTextBox" & vbCrLf & Err.DESCRIPTION
End Function




