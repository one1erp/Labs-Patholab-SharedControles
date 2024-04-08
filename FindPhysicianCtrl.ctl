VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl FindPhysicianCtrl 
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9885
   ScaleHeight     =   6795
   ScaleWidth      =   9885
   Begin MSComctlLib.ImageList imglstIcons 
      Left            =   8880
      Top             =   5760
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
            Picture         =   "FindPhysicianCtrl.ctx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FramePhysician 
      BackColor       =   &H80000016&
      Caption         =   "Physician Details"
      Height          =   2295
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   7560
         TabIndex        =   13
         Top             =   1250
         Width           =   1695
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7560
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton CmdFind 
         Caption         =   "Find Now"
         Height          =   375
         Left            =   7560
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtLicenseNbr 
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
         TabIndex        =   3
         Top             =   1800
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
         Top             =   1320
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
         Height          =   375
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   2535
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
      Begin VB.Image ImageFind 
         Height          =   480
         Left            =   5880
         Picture         =   "FindPhysicianCtrl.ctx":031A
         Top             =   960
         Width           =   480
      End
      Begin VB.Label LblLicenseNbr 
         AutoSize        =   -1  'True
         Caption         =   "License No.:"
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
         TabIndex        =   12
         Top             =   1850
         Width           =   1335
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
         TabIndex        =   11
         Top             =   1375
         Width           =   735
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
         TabIndex        =   10
         Top             =   900
         Width           =   1215
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
         TabIndex        =   9
         Top             =   400
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView LstPhysician 
      Height          =   3975
      Left            =   240
      TabIndex        =   7
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "H1"
         Text            =   "License"
         Object.Width           =   6844
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "First Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblNumOfRecords 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   2895
   End
End
Attribute VB_Name = "FindPhysicianCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Connection As ADODB.Connection
Private SupplierID As String
Private ID As String
Private DESCRIPTION As String
Private MDoc As String
Public Event CloseClick()
Public Event CancelClick()

Public Sub Initialize(Con As ADODB.Connection)
10        Set Connection = Con
20        Call imglstIcons.ListImages.Add(, "L1", LoadPicture("Resource\Supplier.ico"))

30        Call zLang.Hebrew
40        TxtFirstName.Alignment = vbRightJustify
50        TxtFirstName.RightToLeft = True
60        TxtLastName.Alignment = vbRightJustify
70        TxtLastName.RightToLeft = True
End Sub

Private Sub CmdClear_Click()
80        LstPhysician.ListItems.Clear
90        TxtFirstName.Text = ""
100       TxtLastName.Text = ""
110       TxtLicenseNbr.Text = ""
120       TxtIDNbr.Text = ""
130       lblNumOfRecords.Caption = ""
End Sub

Private Sub CmdClose_Click()
140       RaiseEvent CancelClick
End Sub

Private Sub FillList()
          Dim RstPhysician As ADODB.Recordset
          Dim WhereStr As String
          Dim Flag As Boolean
          Dim li As ListItem
          Dim i As Integer
150       Flag = False
160       WhereStr = ""
          
          
          
          'show the hourglass mous pointer
170       MousePointer = 11

180       If Trim(TxtFirstName.Text) <> "" Then
190           WhereStr = "U_FIRST_NAME like '" & Replace(TxtFirstName.Text, "'", "''") & "%' "
200           Flag = True
210       End If

220       If Trim(TxtLastName.Text) <> "" Then
230           If Flag Then WhereStr = WhereStr & "and "
240           WhereStr = WhereStr & "U_LAST_NAME like '" & Replace(TxtLastName.Text, "'", "''") & "%' "
250           Flag = True
260       End If

270       If Trim(TxtIDNbr.Text) <> "" Then
280           If Flag Then WhereStr = WhereStr & "and "
290           WhereStr = WhereStr & "U_ID_NBR like '%" & TxtIDNbr.Text & "%' "
300           Flag = True
310       End If

320       If Trim(TxtLicenseNbr.Text) <> "" Then
330           If Flag Then WhereStr = WhereStr & "and "
340           WhereStr = WhereStr & "U_LICENSE_NBR like '%" & TxtLicenseNbr.Text & "%'"
350           Flag = True
360       End If

370       If Trim(MDoc) <> "" Then
380           If Flag Then WhereStr = WhereStr & "and "
390           WhereStr = WhereStr & "U_M_DOC = '" & MDoc & "'"
400           Flag = True
410       End If

420       If Trim(WhereStr) <> "" Then WhereStr = "where " & WhereStr

430       Set RstPhysician = Connection.Execute("select * from lims_sys.supplier_user " & WhereStr)

440       LstPhysician.ListItems.Clear
450       If Not RstPhysician.EOF Then
460           RstPhysician.MoveFirst
470           i = 0
          
480           While Not RstPhysician.EOF
490               Set li = LstPhysician.ListItems.Add(, , nte(RstPhysician("U_LICENSE_NBR")), , 1)
500               li.Tag = nte(RstPhysician("SUPPLIER_ID"))
510               li.SubItems(1) = nte(RstPhysician("U_ID_NBR"))
520               li.SubItems(2) = nte(RstPhysician("U_LAST_NAME"))
530               li.SubItems(3) = nte(RstPhysician("U_FIRST_NAME"))
540               RstPhysician.MoveNext
550               i = i + 1
560           Wend
              
570           lblNumOfRecords.ForeColor = vbBlack
580           lblNumOfRecords.Caption = " נמצאו " & i & " רשומות "
          
590       Else
600           lblNumOfRecords.RightToLeft = True
610           lblNumOfRecords.ForeColor = vbRed
620           lblNumOfRecords.Caption = " לא נמצאו רשומות "
630       End If
          
640       RstPhysician.Close
          
          'show the regular mouse pointer
650       MousePointer = 0
End Sub

Private Sub CmdFind_Click()
660       FillList
End Sub

Private Sub CmdUpdate_Click()
          Dim IDtemp As String
670       If LstPhysician.ListItems.Count > 0 Then
680           If LstPhysician.SelectedItem.Tag <> "" Then
690               Set FrmConfigDoctor.aConnection = Connection
700               IDtemp = LstPhysician.SelectedItem.Tag
710               FrmConfigDoctor.SupplierID = IDtemp
720               FrmConfigDoctor.Show vbModal
730           End If
740       End If
End Sub

Private Sub LstPhysician_DblClick()
750       CloseForm
End Sub

Private Sub CloseForm()
760       If LstPhysician.ListItems.Count > 0 Then
770           SupplierID = LstPhysician.SelectedItem.Tag
780           ID = SupplierID
790           DESCRIPTION = LstPhysician.SelectedItem.Text & " - " & _
                            LstPhysician.SelectedItem.SubItems(2) & " " & _
                            LstPhysician.SelectedItem.SubItems(3)
800       End If
810       Call zLang.SetOrigLang
820       RaiseEvent CloseClick
End Sub

Private Function nte(e As Variant) As Variant
830       nte = IIf(IsNull(e), "", e)
End Function

Public Function GetSupplierID() As String
840       If LstPhysician.ListItems.Count > 0 Then
850           GetSupplierID = SupplierID
860       Else
870           GetSupplierID = ""
880       End If
End Function

Public Function GetID() As String
890       If LstPhysician.ListItems.Count > 0 Then
900           GetID = ID
910       Else
920           GetID = ""
930       End If
End Function

Public Function GetDescription() As String
940       If LstPhysician.ListItems.Count > 0 Then
950           GetDescription = DESCRIPTION
960       Else
970           GetDescription = ""
980       End If
End Function

Private Sub TxtFirstName_KeyDown(KeyCode As Integer, Shift As Integer)
990       If KeyCode = 13 Then
1000          FillList
1010      End If
End Sub

Private Sub TxtLastName_KeyDown(KeyCode As Integer, Shift As Integer)
1020      If KeyCode = 13 Then
1030          FillList
1040      End If
End Sub

Private Sub TxtLicenseNbr_KeyDown(KeyCode As Integer, Shift As Integer)
1050      If KeyCode = 13 Then
1060          FillList
1070      End If
End Sub

Private Sub TxtIDNbr_KeyDown(KeyCode As Integer, Shift As Integer)
1080      If KeyCode = 13 Then
1090          FillList
1100      End If
End Sub


Public Property Let M_Doc(ByVal vNewValue As String)
1110      MDoc = vNewValue
End Property
