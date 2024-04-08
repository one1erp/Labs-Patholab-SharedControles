VERSION 5.00
Begin VB.Form FrmFindPhysician 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Physician"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   Icon            =   "FrmFindPhysician.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340.5
   ScaleMode       =   0  'User
   ScaleWidth      =   10417.84
   StartUpPosition =   3  'Windows Default
   Begin MacabiShared.FindPhysicianCtrl FindPhysicianCtrl 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11880
   End
End
Attribute VB_Name = "FrmFindPhysician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As ADODB.Connection
Public SupplierID As String
Public ID As String
Public DESCRIPTION As String
Public MDoc As String

Private Sub FindPhysicianCtrl_CloseClick()
2280      SupplierID = FindPhysicianCtrl.GetSupplierID
2290      ID = FindPhysicianCtrl.GetID
2300      DESCRIPTION = FindPhysicianCtrl.GetDescription
2310      Unload Me
End Sub

Private Sub FindPhysicianCtrl_CancelClick()
2320      SupplierID = ""
2330      ID = ""
2340      DESCRIPTION = ""
2350      Unload Me
End Sub

Private Sub Form_Load()
          'clean former data:
2360      SupplierID = ""
2370      ID = ""
2380      DESCRIPTION = ""
          
          ''''''''''''''''''''''''''''''''''''''
          
2390      FindPhysicianCtrl.M_Doc = MDoc
2400      Call FindPhysicianCtrl.Initialize(Con)
End Sub

