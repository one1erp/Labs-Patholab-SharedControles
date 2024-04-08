VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl FindPatientCtrl 
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9885
   ScaleHeight     =   6795
   ScaleMode       =   0  'User
   ScaleWidth      =   10146.75
   Begin MSComctlLib.ListView LsPatient 
      Height          =   3975
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imglstIcons"
      ColHdrIcons     =   "imglstHeaderIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Last Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "First Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame FramePatient 
      BackColor       =   &H80000016&
      Caption         =   "Patient Details"
      Height          =   2295
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   7560
         TabIndex        =   11
         Top             =   1260
         Width           =   1695
      End
      Begin VB.TextBox TxtFirstName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox TxtLastName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   980
         Width           =   2535
      End
      Begin VB.TextBox TxtIDNbr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton CmdFind 
         Caption         =   "Find Now"
         Height          =   375
         Left            =   7560
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   7560
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label LblFirstName 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
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
         Left            =   360
         TabIndex        =   10
         Top             =   400
         Width           =   1215
      End
      Begin VB.Label LblLastName 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label LblIDNbr 
         AutoSize        =   -1  'True
         Caption         =   "ID No.:"
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
         Left            =   360
         TabIndex        =   8
         Top             =   1740
         Width           =   735
      End
      Begin VB.Image ImageFind 
         Height          =   480
         Left            =   5880
         Picture         =   "FindPatientCtrl.ctx":0000
         Top             =   960
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList imglstIcons 
      Left            =   8880
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FindPatientCtrl.ctx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNumOfRecords 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   2655
   End
End
Attribute VB_Name = "FindPatientCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Connection As ADODB.Connection
Private PatientID As String
Private ID As String
Private DESCRIPTION As String
Public Event CloseClick()
Public Event CancelClick()

Public Sub Initialize(Con As ADODB.Connection)
1120      Set Connection = Con
1130      Call imglstIcons.ListImages.Add(, "L1", LoadPicture("Resource\Client.ico"))

1140      Call zLang.Hebrew
1150      TxtFirstName.Alignment = vbRightJustify
1160      TxtFirstName.RightToLeft = True
1170      TxtLastName.Alignment = vbRightJustify
1180      TxtLastName.RightToLeft = True
End Sub

Private Sub CmdClear_Click()
1190      LsPatient.ListItems.Clear
1200      lblNumOfRecords.Caption = ""
1210      TxtFirstName.Text = ""
1220      TxtLastName.Text = ""
1230      TxtIDNbr.Text = ""
End Sub

Private Sub CmdClose_Click()
1240      RaiseEvent CancelClick
End Sub

Private Sub FillList()
          Dim RstPatient As ADODB.Recordset
          Dim WhereStr As String
          Dim InitWhereStr As String
          Dim Flag As Boolean
          Dim li As ListItem
          Dim i As Integer
          
          'show the hourglass mous pointer
1250      MousePointer = 11
          
1260      Flag = False
1270      InitWhereStr = "where client_user.client_id = client.client_id "
1280      WhereStr = ""

1290      If Trim(TxtFirstName.Text) <> "" Then
1300          WhereStr = WhereStr & "U_FIRST_NAME like '" & Replace(TxtFirstName.Text, "'", "''") & "%' "
1310          Flag = True
1320      End If
1330      If Trim(TxtLastName.Text) <> "" Then
1340          If Flag Then WhereStr = WhereStr & "and "
1350          WhereStr = WhereStr & "U_LAST_NAME like '" & Replace(TxtLastName.Text, "'", "''") & "%' "
1360          Flag = True
1370      End If
1380      If Trim(TxtIDNbr.Text) <> "" Then
1390          If Flag Then WhereStr = WhereStr & "and "
1400          WhereStr = WhereStr & "NAME like '%" & TxtIDNbr.Text & "%' "
1410          Flag = True
1420      End If

1430      If Trim(WhereStr) <> "" Then
1440          WhereStr = InitWhereStr & "and " & WhereStr
1450      Else
1460          WhereStr = InitWhereStr
1470      End If

1480      Set RstPatient = Connection.Execute("select * from lims_sys.client_user, lims_sys.client " & WhereStr)

1490      LsPatient.ListItems.Clear
1500      If Not RstPatient.EOF Then
1510          RstPatient.MoveFirst
1520          i = 0
          
1530          While Not RstPatient.EOF
1540              Set li = LsPatient.ListItems.Add(, , nte(RstPatient("NAME")), , 1)
1550              li.Tag = nte(RstPatient("CLIENT_ID"))
1560              li.SubItems(1) = nte(RstPatient("U_LAST_NAME"))
1570              li.SubItems(2) = nte(RstPatient("U_FIRST_NAME"))
1580              RstPatient.MoveNext
1590              i = i + 1
1600          Wend
                                      
1610          lblNumOfRecords.ForeColor = vbBlack
1620          lblNumOfRecords.Caption = " נמצאו " & i & " רשומות "
          
1630      Else
1640          lblNumOfRecords.RightToLeft = True
1650          lblNumOfRecords.ForeColor = vbRed
1660          lblNumOfRecords.Caption = " לא נמצאו רשומות "
1670      End If
1680      RstPatient.Close

          'show the regular mouse pointer
1690      MousePointer = 0

End Sub

Private Sub CmdFind_Click()
1700      If TxtFirstName.Text = "" And _
             TxtLastName.Text = "" And _
             TxtIDNbr = "" Then
1710          MsgBox " לא הוכנסו קריטריונים לחיפוש "
1720          Exit Sub
1730      End If
          
          
1740      FillList
End Sub

Private Sub CmdUpdate_Click()
          Dim IDtemp As String
1750      If LsPatient.ListItems.Count > 0 Then
1760          If LsPatient.SelectedItem.Tag <> "" Then
1770              Set FrmCheckConfigPatientGeneral.aConnection = Connection
1780              IDtemp = LsPatient.SelectedItem.Tag
1790              FrmCheckConfigPatientGeneral.Client_ID = IDtemp
1800              FrmCheckConfigPatientGeneral.Show vbModal
1810          End If
1820      End If
End Sub

Private Sub LsPatient_DblClick()
1830      CloseForm
End Sub

Private Sub CloseForm()
1840      If LsPatient.ListItems.Count > 0 Then
1850          PatientID = LsPatient.SelectedItem.Tag
1860          DESCRIPTION = LsPatient.SelectedItem.SubItems(1) & " " & _
                            LsPatient.SelectedItem.SubItems(2)
1870          ID = PatientID
1880      End If
1890      Call zLang.SetOrigLang
1900      RaiseEvent CloseClick
End Sub

Private Function nte(e As Variant) As Variant
1910      nte = IIf(IsNull(e), "", e)
End Function

Public Function GetPatientID() As String
1920      If LsPatient.ListItems.Count > 0 Then
1930          GetPatientID = PatientID
1940      Else
1950          GetPatientID = ""
1960      End If
End Function

Public Function GetID() As String
1970      If LsPatient.ListItems.Count > 0 Then
1980          GetID = ID
1990      Else
2000          GetID = ""
2010      End If
End Function

Public Function GetDescription() As String
2020      If LsPatient.ListItems.Count > 0 Then
2030          GetDescription = DESCRIPTION
2040      Else
2050          GetDescription = ""
2060      End If
End Function

Private Sub TxtFirstName_KeyDown(KeyCode As Integer, Shift As Integer)
2070      If KeyCode = 13 Then
2080          FillList
2090      End If
End Sub

Private Sub TxtLastName_KeyDown(KeyCode As Integer, Shift As Integer)
2100      If KeyCode = 13 Then
2110          FillList
2120      End If
End Sub

Private Sub TxtIDNbr_KeyDown(KeyCode As Integer, Shift As Integer)
2130      If KeyCode = 13 Then
2140          FillList
2150      End If
End Sub


