VERSION 5.00
Object = "{5CD7A57E-9286-4870-B439-7253EA161B09}#419.1#0"; "ConfigDoctors.ocx"
Begin VB.Form FrmConfigDoctor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "עידכון רופא"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "FrmConfigDoctor.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtLicenseNbr 
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
      Left            =   6000
      TabIndex        =   0
      ToolTipText     =   "הקש מס. רישיון"
      Top             =   90
      Visible         =   0   'False
      Width           =   1650
   End
   Begin ConfigDoctors.ConfigDoctorsCtrl ConfigDoctorsCtrl 
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13785
   End
   Begin VB.Label LblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "מס. רישיון רופא חדש "
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
      Left            =   8055
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2445
   End
End
Attribute VB_Name = "FrmConfigDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public aConnection As ADODB.Connection
Public SupplierID As String
Public Doctor As ADODB.Recordset

Private Sub ConfigDoctorsCtrl_CloseClick()
          Dim MBRes As VbMsgBoxResult

6380      MBRes = MsgBox("? האם את/ה בטוח שברצונך לצאת המסך", vbYesNo, "Nautilus - עידכון רופא")
6390      If MBRes = vbNo Then Exit Sub

6400      Unload Me
End Sub

Private Sub ConfigDoctorsCtrl_SaveClick()
6410      Unload Me
End Sub

Private Sub Form_Load()
6420      Set ConfigDoctorsCtrl.Connection = aConnection
          
          ' ----------------------
          ' 933
          ' In order to activate the New Dr feature in this control without breaking compatabilty,
          ' The SupplierId field is used, and when passed a "#" - it enables the TxtLicenseNbr field.
6430      If InStr(1, SupplierID, "#") <> 0 Then
              
6440          Me.TxtLicenseNbr.Visible = True
6450          Me.LblName.Visible = True
6460          Me.Caption = "הוספת רופא"
              
6470      Else
          
6480          Call ConfigDoctorsCtrl.GetSupplier(SupplierID)
              
6490      End If
End Sub

Private Sub TxtLicenseNbr_KeyDown(KeyCode As Integer, Shift As Integer)
6500      If Not KeyCode = vbKeyReturn Then Exit Sub
          
6510      If Trim(TxtLicenseNbr.Text) <> "" Then
6520          Foo
6530      End If
          
End Sub

Private Sub TxtLicenseNbr_LostFocus()
6540      If Trim(TxtLicenseNbr.Text) <> "" Then
6550          Foo
6560      End If
End Sub

Private Sub Foo()
          Dim StrSaveID As String
          Dim IDFlag As Boolean

6570      If Len(Trim(TxtLicenseNbr.Text)) > 6 Then Exit Sub
          
6580      StrSaveID = lpad(Trim(TxtLicenseNbr.Text), "0", 6)
6590      Set Doctor = aConnection.Execute("select * from lims_sys.supplier, lims_sys.supplier_user where supplier.supplier_id = supplier_user.supplier_id and supplier_user.U_LICENSE_NBR = '" & StrSaveID & "'")
6600      If Not Doctor.EOF Then
6610          MsgBox " ! קיים כבר רופא לרישיון שהוקש ", , "Nautilus - קלט רופא"
6620      Else
6630          CreateDoctor (StrSaveID)
6640          TxtLicenseNbr.Enabled = False
              ' 933
              ' The substring #ENABLE_ID_EDIT" is passed to the ConfigDoctorsControl
              ' in order to signal it to enable the TxtIdNbr field without adding or changing functions and
              ' breaking the compatabilty.
6650          Call ConfigDoctorsCtrl.GetSupplier(Doctor("SUPPLIER_ID") & "#ENABLE_ID_EDIT")
              
6660      End If

6670      TxtLicenseNbr.Text = ""

End Sub


Private Function lpad(s As String, c As String, leng As Integer) As String
6680      lpad = String(leng - Len(s), c) & s
End Function

Private Function nte(e As Variant) As Variant
6690      nte = IIf(IsNull(e), "", e)
End Function

Private Sub CreateDoctor(IDnumber As String)
          Dim SQSupplier As Recordset
          Dim StrSql As String
          Dim Supplier_ID

6700      Set SQSupplier = aConnection.Execute("select lims.sq_supplier.nextval from dual")

6710      Supplier_ID = SQSupplier(0)
          
6720      Call aConnection.BeginTrans
          
          ' Insert new record into client table
6730      StrSql = "insert into lims_sys.supplier " & _
                  "(SUPPLIER_ID, NAME, VERSION, VERSION_STATUS) " & _
                  "values (" & Supplier_ID & ", " & _
                  "'" & Supplier_ID & "', " & _
                  "'1', " & _
                  "'A')"

6740      Call aConnection.Execute(StrSql)

          ' Insert new record into client_user table
6750      StrSql = "insert into lims_sys.supplier_user " & _
                  "(SUPPLIER_ID, U_LICENSE_NBR) " & _
                  "values (" & Supplier_ID & ", " & _
                  "'" & IDnumber & "')"

6760      Call aConnection.Execute(StrSql)
          
6770      Call aConnection.CommitTrans
          
          ' For using in the next routine (Update Routine)
6780      Set Doctor = aConnection.Execute("select * from lims_sys.supplier, lims_sys.supplier_user where supplier.supplier_id = supplier_user.supplier_id and supplier.supplier_id = '" & Supplier_ID & "'")
End Sub

