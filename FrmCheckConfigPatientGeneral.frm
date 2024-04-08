VERSION 5.00
Object = "{A6CE5300-67D0-401F-9352-1E8DEEA88C0F}#40.0#0"; "ConfigPatientGeneral.ocx"
Begin VB.Form FrmCheckConfigPatientGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "עידכון פצייאנט"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14385
   Icon            =   "FrmCheckConfigPatientGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtIDNbr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11760
      TabIndex        =   1
      ToolTipText     =   "הקש מס. תעודת זהות"
      Top             =   120
      Width           =   1650
   End
   Begin ConfigPatientGeneral.ConfigPatientGCtrl ConfigPatientGCtrl 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12726
   End
   Begin VB.Label LblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "הנבדק:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "FrmCheckConfigPatientGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public aConnection As New ADODB.Connection
Public Client_ID As String
Public Patient As ADODB.Recordset

Private Sub ConfigPatientGCtrl_CloseClick()
          Dim MBRes As VbMsgBoxResult

6790      MBRes = MsgBox("? האם את/ה בטוח שברצונך לצאת המסך", vbYesNo, "Nautilus - עידכון פצייאנט")
6800      If MBRes = vbNo Then Exit Sub

6810      Unload Me
End Sub

Private Sub ConfigPatientGCtrl_SaveClick()
6820      Unload Me
End Sub

Private Sub Form_Load()
6830      Set ConfigPatientGCtrl.Connection = aConnection
6840      Call ConfigPatientGCtrl.GetClient(Client_ID)
End Sub

Private Sub TxtIDNbr_KeyDown(KeyCode As Integer, Shift As Integer)
6850      If Not KeyCode = vbKeyReturn Then Exit Sub

          Dim StrSaveID As String
          Dim IDFlag As Boolean

6860      If Len(TxtIDNbr.Text) > 10 Then Exit Sub
6870      StrSaveID = lpad(TxtIDNbr.Text, "0", 10)

6880      Set Patient = aConnection.Execute("select * from lims_sys.client, lims_sys.client_user where client.client_id = client_user.client_id and name = '" & StrSaveID & "'")

6890      If Patient.EOF Then
6900          IDFlag = CheckIDNo(StrSaveID)
6910          If IDFlag = False Then
6920              MsgBox " ! המספר שהוקש שגוי, סיפרת ביקורת לא תקינה ", , "Nautilus - קלט פצייאנט"
6930              Call TxtIDNbr.SetFocus
6940              Exit Sub
6950          Else
6960              MsgBox " ! הפצייאנט אינו קיים במערכת, ירשם פצייאנט חדש ", , "Nautilus - קלט פצייאנט"
6970              CreatePatient (StrSaveID)
6980              Call ConfigPatientGCtrl.GetClient(nte(Trim(Patient("CLIENT_ID"))))
6990          End If
7000      Else
7010          Call ConfigPatientGCtrl.GetClient(nte(Trim(Patient("CLIENT_ID"))))
7020      End If
7030      TxtIDNbr.Text = ""
End Sub

Private Function lpad(s As String, c As String, leng As Integer) As String
7040      lpad = String(leng - Len(s), c) & s
End Function

Private Function CheckIDNo(s As String) As Boolean
          Dim errOn As Boolean
          Dim i As Integer
          Dim j As Integer
          Dim sum As Integer
          Dim currSum As Integer
          Dim res As Boolean

7050      j = 2
7060      sum = 0
7070      errOn = False
7080      For i = Len(s) - 1 To 1 Step -1
          'On Error GoTo err1
7090          currSum = Mid(s, i, 1) * j
7100          If currSum >= 10 Then
7110              currSum = currSum + 1
7120          End If
7130          sum = sum + currSum
7140          If j = 2 Then
7150              j = 1
7160          Else
7170              j = 2
7180          End If
              'GoTo ok
              'err1:
              'errOn = True
              'ok:
7190      Next i

7200      sum = 10 - sum Mod 10
7210      If sum = 10 Then
7220          sum = 0
7230      End If

7240      If sum = Mid(s, Len(s), 1) Then
7250          res = True
7260      Else
7270          res = False
7280      End If

7290      CheckIDNo = res

End Function

Private Function nte(e As Variant) As Variant
7300      nte = IIf(IsNull(e), "", e)
End Function

Private Sub CreatePatient(IDnumber As String)
          Dim SQClient As Recordset
          Dim SQAddress As Recordset
          Dim StrSql As String
          Dim ClientID
          Dim AddressID

7310      Set SQClient = aConnection.Execute("select lims.sq_client.nextval from dual")

7320      Set SQAddress = aConnection.Execute("select lims.sq_address.nextval from dual")

7330      ClientID = SQClient(0)
7340      AddressID = SQAddress(0)

          ' Insert new record into client table
7350      StrSql = "insert into lims_sys.client " & _
                  "(CLIENT_ID, NAME, VERSION, VERSION_STATUS) " & _
                  "values (" & ClientID & ", " & _
                  "'" & IDnumber & "', " & _
                  "'1', " & _
                  "'A')"

7360      Call aConnection.Execute(StrSql)

          ' Insert new record into client_user table
7370      StrSql = "insert into lims_sys.client_user " & _
                  "(CLIENT_ID) " & _
                  "values (" & ClientID & ")"

7380      Call aConnection.Execute(StrSql)

          ' Insert new record into address table
7390      StrSql = "insert into lims_sys.address " & _
                  "(ADDRESS_ID, ADDRESS_TABLE_NAME, ADDRESS_ITEM_ID, ADDRESS_LINE_1, ADDRESS_TYPE) " & _
                  "values (" & AddressID & ", " & _
                  "'CLIENT', " & _
                  ClientID & ", " & _
                  "'0', " & _
                  "'C')"

7400      Call aConnection.Execute(StrSql)

          ' For using in the next routine (Update Routine)
7410      Set Patient = aConnection.Execute("select * from lims_sys.client, lims_sys.client_user where client.client_id = client_user.client_id and name = '" & IDnumber & "'")
End Sub



