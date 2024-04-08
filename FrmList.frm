VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "FrmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstListItems 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7435
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
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
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FreeTxt As Object
Private InitWidth As Long
Private InitHeight As Long

Private Sub Form_Initialize()
2410      InitWidth = Me.width
2420      InitHeight = Me.height
End Sub

Private Sub Form_Unload(Cancel As Integer)
2430      Set FreeTxt = Nothing
End Sub

Private Sub lstListItems_DblClick()
2440      FreeTxt.SelText = lstListItems.SelectedItem.Text
      '    Set FreeTxt = Nothing
      '    Unload Me
End Sub

Public Sub setFreeText(FreeText As Object)
2450      Set FreeTxt = FreeText
End Sub

Private Sub lstListItems_KeyPress(KeyAscii As Integer)
2460      If KeyAscii = vbKeyReturn Then
2470          FreeTxt.SelText = lstListItems.SelectedItem.Text
2480          Unload Me
2490      End If
End Sub
